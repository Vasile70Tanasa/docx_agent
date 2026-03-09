"""mapper.py - Page-Based LLM Field-to-JSON-Key Mapper

Groups fields into page-like sections based on paragraph ranges, builds
full page text with [[PH:id]] markers, and uses key_selector to provide
per-placeholder candidate keys+values to the LLM.
"""
from __future__ import annotations

import io
import json
import os
import re
import sys
import hashlib

if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from typing import Any, Dict, List, Optional, Tuple

from dotenv import load_dotenv
load_dotenv()

try:
    import anthropic
except ImportError:
    anthropic = None

from key_selector import top_keys


def _sha(s: str) -> str:
    return hashlib.sha256(s.encode('utf-8')).hexdigest()[:16]


def _para_index_from_location(location: str) -> Optional[int]:
    m = re.findall(r'p:(\d+)', location or '')
    return int(m[-1]) if m else None


def _effective_para_index(ent: Dict) -> int:
    """Return body-level para index for any entity (body or table cell)."""
    if '_body_para_index' in ent:
        return ent['_body_para_index']
    return _para_index_from_location(ent.get('location', '')) or 999999


class _JsonCache:
    def __init__(self, path: str):
        self._path = path
        self._data: Dict[str, Any] = {}
        self._load()

    def _load(self):
        if self._path and os.path.exists(self._path):
            try:
                with open(self._path, 'r', encoding='utf-8') as f:
                    self._data = json.load(f)
            except Exception:
                self._data = {}

    def _save(self):
        if not self._path:
            return
        path = os.path.abspath(self._path)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(self._data, f, ensure_ascii=False, indent=2)

    def get(self, key: str) -> Optional[Dict]:
        return self._data.get(key)

    def set(self, key: str, value: Dict):
        self._data[key] = value
        self._save()


# ── Page text building ────────────────────────────────────────────────────────

def _build_page_text(paras_elements, p_start: int, p_end: int,
                     fields_by_para: Dict[int, List[Dict]],
                     data: Dict, n_cand: int = 10,
                     extra_fields: Optional[List[Dict]] = None):
    """Build page text with [[PH:id]] markers + per-placeholder candidate lists.

    paras_elements: list of lxml w:p elements (direct body children).
    extra_fields: table-cell fields not in body paragraphs but included for candidates.
    Returns (page_text, page_fields, cand_section).
    """
    from docx.text.paragraph import Paragraph as P

    page_fields = []
    lines = []

    for i in range(p_start, min(p_end + 1, len(paras_elements))):
        para = P(paras_elements[i], None)
        txt = para.text.strip()
        if not txt:
            continue
        if i in fields_by_para:
            flds = sorted(fields_by_para[i], key=lambda f: f['start'])
            marked = txt
            offset = 0
            for fld in flds:
                s = fld['start'] + offset
                e = fld['end'] + offset
                fid = fld['field_id'][:8]
                marker = f'[[PH:{fid}]]'
                marked = marked[:s] + marker + marked[e:]
                offset += len(marker) - (fld['end'] - fld['start'])
                page_fields.append(fld)
            lines.append(marked)
        else:
            lines.append(txt)

    # Add table-cell fields to page text as separate context lines
    if extra_fields:
        for fld in extra_fields:
            fid = fld['field_id'][:8]
            ctx = fld.get('ctx_before', '') + f' [[PH:{fid}]] ' + fld.get('ctx_after', '')
            prev = fld.get('ctx_prev_para', '')
            if prev:
                ctx = prev + '\n' + ctx
            lines.append(ctx.strip())
            page_fields.append(fld)

    page_text = '\n'.join(lines)

    # Build candidate lists per field
    cand_section = "\nCANDIDATE KEYS PER PLACEHOLDER (pick from these):\n"
    for fld in page_fields:
        fid = fld['field_id'][:8]
        candidates = top_keys(fld, data, n=n_cand)
        cand_lines = ', '.join(f'"{k}": "{v}"' for k, _, v in candidates)
        cand_section += f"  {fid}: {{{cand_lines}}}\n"

    return page_text, page_fields, cand_section


def _find_section_title(paras_elements, para_index: int, max_look_back: int = 20) -> str:
    """Scan paragraphs above para_index for a section title."""
    from docx.text.paragraph import Paragraph as P

    for i in range(para_index - 1, max(0, para_index - max_look_back) - 1, -1):
        para = P(paras_elements[i], None)
        txt = para.text.strip()
        if not txt or len(txt) < 3:
            continue
        if 'formular' in txt.lower():
            title = txt
            for j in range(i + 1, min(i + 4, len(paras_elements))):
                next_txt = P(paras_elements[j], None).text.strip()
                if next_txt and len(next_txt) > 3:
                    title += ' - ' + next_txt
                    break
            return title
        upper_ratio = sum(1 for c in txt if c.isupper()) / max(len(txt), 1)
        if upper_ratio > 0.6 and len(txt) > 5 and len(txt) < 100:
            return txt
    return ''


# ── Page chunking ─────────────────────────────────────────────────────────────

def _page_chunks(entities: List[Dict], gap: int = 15, max_fields: int = 40) -> List[Tuple[int, int, List[Dict]]]:
    """Split entities into page-like chunks.

    A new chunk starts when:
    - gap between consecutive para indices > `gap`, OR
    - chunk already has `max_fields` entities.

    Returns list of (p_start, p_end, [entities]).
    """
    if not entities:
        return []

    chunks = []
    current = [entities[0]]
    cur_start = _effective_para_index(entities[0])

    for ent in entities[1:]:
        pid = _effective_para_index(ent)
        prev_pid = _effective_para_index(current[-1])

        if (pid - prev_pid > gap) or len(current) >= max_fields:
            cur_end = prev_pid
            chunks.append((cur_start, cur_end, current))
            current = [ent]
            cur_start = pid
        else:
            current.append(ent)

    if current:
        cur_end = _effective_para_index(current[-1])
        chunks.append((cur_start, cur_end, current))

    return chunks


# ── LLM call ──────────────────────────────────────────────────────────────────

def _llm_map_page(client: Any, page_text: str, cand_section: str,
                  section_title: str, model: str) -> Dict:
    """Send page text + candidate keys to LLM, return parsed results."""

    prompt = f"""You are a document data-entry system for Romanian procurement forms.

SECTION TITLE: {section_title}

Below is a page from the form with placeholders marked as [[PH:xxxxxxxx]].
For each placeholder, select the BEST matching key from its candidate list below.

RULES:
1. Select an exact key from the placeholder's candidate list, or null if none fits.
2. When both a composite key (e.g. "Asociat 1 - denumire, sediu, telefon") and its derived sub-key (e.g. "Asociat 1 - denumire") appear in the candidate list, prefer the sub-key if the placeholder asks for only that piece of information. Do NOT treat shorter keys as automatically more specific — "Suma (in litere si cifre)" and "Suma de ... lei (in litere si cifre)" are BOTH atomic keys, not parent-child.
3. Use the FULL page context to understand which key belongs where.
4. When the same structure repeats (e.g. first = "Asociat 1", second = "Asociat 2"), use ordering.
5. When multiple placeholders appear in the same sentence, consider them TOGETHER — they often represent parts of the same concept (e.g., name + role, amount + currency).
6. Return ONLY a JSON array.

DOCUMENT PAGE:
{page_text}
{cand_section}
Return ONLY a JSON array:
[{{"id": "xxxxxxxx", "selected_key": "key"|null, "extracted_value": "subset of value for this specific placeholder, or null if the full value should be used", "confidence": 0.0-1.0, "reasoning": "brief"}}]
"""

    try:
        resp = client.messages.create(
            model=model, max_tokens=8192, temperature=0,
            messages=[{"role": "user", "content": prompt}],
        )
        content = (resp.content[0].text or "").strip()
        content = re.sub(r"^```(?:json)?\s*", "", content)
        content = re.sub(r"\s*```$", "", content).strip()
        m = re.search(r"(\[[\s\S]*\])", content)
        items = json.loads(m.group(1)) if m else json.loads(content)

        result = {}
        for item in items:
            eid = item.get("id")
            if eid:
                result[eid] = item
        return result
    except Exception as e:
        print(f"LLM Mapping error: {e}")
        return {}


# ── Main entry point ──────────────────────────────────────────────────────────

def build_mapping(
    template_fingerprint: str,
    fields: List[Dict],
    tables: List[Dict],
    data: Dict[str, Any],
    cache_path: str = 'cache/mapping_cache.json',
    model: str = 'claude-sonnet-4-20250514',
    docx_path: str = 'input/sample_forms.docx',
) -> Dict[str, Any]:

    json_keys = list(data.keys())
    cache = _JsonCache(cache_path)

    if anthropic is None:
        raise RuntimeError("anthropic package not installed: pip install anthropic")
    api_key = os.environ.get('ANTHROPIC_API_KEY', '').strip()
    if not api_key:
        raise RuntimeError("ANTHROPIC_API_KEY environment variable is not set")

    mapping: Dict[str, Any] = {}
    entities: List[Dict] = []

    # 1. Group checkboxes
    group_options: Dict[str, List[str]] = {}
    group_ctx: Dict[str, Dict] = {}

    for f in fields:
        gid = f.get('group_id')
        ftype = f['field_type']
        if ftype == 'CHECKBOX' and gid:
            if gid not in group_options:
                group_options[gid] = []
                group_ctx[gid] = {
                    'ctx_before':    f.get('ctx_before', ''),
                    'ctx_after':     f.get('ctx_after',  ''),
                    'ctx_prev_para': f.get('ctx_prev_para', ''),
                    'ctx_next_para': f.get('ctx_next_para', ''),
                    'location':      f.get('location', ''),
                    'start':         f.get('start', 0),
                    'end':           f.get('end', 0),
                }
            opt = f.get('option_label') or f.get('label', '')
            if opt:
                group_options[gid].append(opt)
        elif ftype != 'CHECKBOX':
            entities.append(f)

    for gid, options in group_options.items():
        ctx = group_ctx[gid]
        entities.append({
            'field_id':      gid,
            'ctx_before':    ctx.get('ctx_before', ''),
            'ctx_after':     ctx.get('ctx_after',  ''),
            'ctx_prev_para': ctx.get('ctx_prev_para', ''),
            'ctx_next_para': ctx.get('ctx_next_para', ''),
            'field_type':    'CHECKBOX_GROUP',
            'location':      ctx.get('location', ''),
            'start':         ctx.get('start', 0),
            'end':           ctx.get('end', 0),
        })

    # 1.5. Process TABLE entities with mathematical similarity
    from difflib import SequenceMatcher

    def _norm(s: str) -> str:
        s = s.lower().replace('_', ' ')
        for a, b in [('ă','a'),('â','a'),('î','i'),('ș','s'),('ş','s'),('ț','t'),('ţ','t')]:
            s = s.replace(a, b)
        return re.sub(r'\s+', ' ', re.sub(r'[^\w\s]', ' ', s)).strip()

    def _sim(a: str, b: str) -> float:
        return SequenceMatcher(None, _norm(a), _norm(b)).ratio()

    array_keys = []
    for k, v in data.items():
        if isinstance(v, list):
            array_keys.append(k)
        elif isinstance(v, str) and v.strip().startswith('['):
            try:
                if isinstance(json.loads(v), list):
                    array_keys.append(k)
            except Exception:
                pass

    for t in tables:
        col_info = ' '.join(h for h in t.get('col_headers', []) if h.strip())
        best_k, best_s = None, 0.0
        for ak in array_keys:
            sc = _sim(col_info, ak)
            if sc > best_s:
                best_s, best_k = sc, ak
        if best_k and best_s > 0.15:
            mapping[t['field_id']] = {
                'json_key': best_k,
                'confidence': best_s,
                'source': 'direct_table',
                'para_index': _para_index_from_location(t.get('location', '')),
                'start': 0, 'end': 0
            }
        else:
            mapping[t['field_id']] = {'json_key': None, 'confidence': 0.0, 'source': 'unmatched_table'}

    # 2. Load DOCX paragraphs (direct body children only)
    from docx import Document
    from docx.oxml.ns import qn
    doc = Document(docx_path)
    body = doc.element.body
    paras_elements = [el for el in body if el.tag == qn('w:p')]

    # Assign body-level para index to table-cell fields based on table position
    tbl_para_index = {}
    p_count = 0
    t_count = 0
    for el in body:
        tag = el.tag.split('}')[-1]
        if tag == 'p':
            p_count += 1
        elif tag == 'tbl':
            tbl_para_index[t_count] = p_count  # para index just before this table
            t_count += 1

    for ent in entities:
        loc = ent.get('location', '')
        if '/t:' in loc:
            m = re.search(r't:(\d+)', loc)
            if m:
                tidx = int(m.group(1))
                ent['_body_para_index'] = tbl_para_index.get(tidx, 999999)

    # Sort entities by para_index, then start
    def _sort_key(e):
        if '_body_para_index' in e:
            return (e['_body_para_index'], e.get('start', 0))
        pid = _para_index_from_location(e.get('location', '')) or 999999
        return (pid, e.get('start', 0))
    entities.sort(key=_sort_key)

    # 3. Split into page chunks and process each
    client = anthropic.Anthropic(api_key=api_key)

    for p_start, p_end, chunk_ents in _page_chunks(entities):
        # Check cache first
        to_process = []
        for ent in chunk_ents:
            eid = ent['field_id']
            cache_key = _sha(f"{template_fingerprint}|{eid}")
            hit = cache.get(cache_key)
            if hit and (hit.get('json_key') is None or hit.get('json_key') in json_keys):
                mapping[eid] = hit
            else:
                to_process.append(ent)

        if not to_process:
            continue

        section_title = _find_section_title(paras_elements, p_start)

        # If few fields to process, skip full page text — use compact per-field context
        if len(to_process) <= 3:
            print(f"  Chunk p:{p_start}-{p_end}: {len(to_process)} fields (compact), "
                  f"title='{section_title[:40]}'")
            compact_lines = []
            compact_cand = "\nCANDIDATE KEYS PER PLACEHOLDER (pick from these):\n"
            for fld in to_process:
                fid8 = fld['field_id'][:8]
                ctx = (fld.get('ctx_before', '') + f' [[PH:{fid8}]] ' +
                       fld.get('ctx_after', '')).strip()
                label = fld.get('label', '')
                if label:
                    ctx = f"[label: {label}] {ctx}"
                prev = fld.get('ctx_prev_para', '')
                nxt = fld.get('ctx_next_para', '')
                if prev:
                    ctx = prev + '\n' + ctx
                if nxt:
                    ctx = ctx + '\n' + nxt
                compact_lines.append(ctx)
                candidates = top_keys(fld, data, n=10)
                cand_str = ', '.join(f'"{k}": "{v}"' for k, _, v in candidates)
                compact_cand += f"  {fid8}: {{{cand_str}}}\n"
            page_text = '\n---\n'.join(compact_lines)
            cand_section = compact_cand
            results = _llm_map_page(client, page_text, cand_section, section_title, model)
        else:
            # Build fields_by_para for this chunk; table-cell fields go to extra_fields
            fields_by_para: Dict[int, List[Dict]] = {}
            extra_fields: List[Dict] = []
            for ent in chunk_ents:
                if '/t:' in ent.get('location', ''):
                    extra_fields.append(ent)
                else:
                    pid = _para_index_from_location(ent.get('location', ''))
                    if pid is not None:
                        fields_by_para.setdefault(pid, []).append(ent)

            # Expand range by a few paragraphs for context
            expanded_start = max(0, p_start - 5)
            expanded_end = min(len(paras_elements) - 1, p_end + 3)

            page_text, page_fields, cand_section = _build_page_text(
                paras_elements, expanded_start, expanded_end,
                fields_by_para, data, n_cand=10,
                extra_fields=extra_fields
            )

            print(f"  Chunk p:{p_start}-{p_end}: {len(to_process)} fields to map, "
                  f"title='{section_title[:40]}', {len(page_text)} chars")

            results = _llm_map_page(client, page_text, cand_section, section_title, model)

        # Collect null results for fallback retry
        null_fields = []

        for ent in to_process:
            eid = ent['field_id']
            fid8 = eid[:8]
            r = results.get(fid8, {})
            sel = r.get('selected_key')
            ext = r.get('extracted_value')
            reasoning = r.get('reasoning')
            conf = float(r.get('confidence', 0.0))

            if sel and sel not in data:
                sel = None

            if sel is None:
                null_fields.append(ent)
                continue

            res_obj = {
                'json_key': sel,
                'extracted_value': ext,
                'reasoning': reasoning,
                'confidence': conf,
                'source': 'llm_page_ks',
                'para_index': _para_index_from_location(ent.get('location', '')),
                'start': ent.get('start', 0),
                'end': ent.get('end', 0),
            }
            mapping[eid] = res_obj

            cache_key = _sha(f"{template_fingerprint}|{eid}")
            cache.set(cache_key, res_obj)

        # Fallback: retry null fields with full key list
        if null_fields:
            all_keys_str = ', '.join(f'"{k}"' for k in json_keys)
            fallback_section = f"\nALL AVAILABLE KEYS:\n{all_keys_str}\n"
            # Build minimal page text with only null fields' context
            fallback_lines = []
            for fld in null_fields:
                fid8 = fld['field_id'][:8]
                ctx = (fld.get('ctx_before', '') + f' [[PH:{fid8}]] ' +
                       fld.get('ctx_after', '')).strip()
                label = fld.get('label', '')
                if label:
                    ctx = f"[label: {label}] {ctx}"
                prev = fld.get('ctx_prev_para', '')
                nxt = fld.get('ctx_next_para', '')
                if prev:
                    ctx = prev + '\n' + ctx
                if nxt:
                    ctx = ctx + '\n' + nxt
                fallback_lines.append(ctx)

            fallback_text = '\n---\n'.join(fallback_lines)
            fallback_cand = fallback_section

            print(f"    Fallback retry: {len(null_fields)} fields with all keys")
            fallback_results = _llm_map_page(
                client, fallback_text, fallback_cand, section_title, model)

            for ent in null_fields:
                eid = ent['field_id']
                fid8 = eid[:8]
                r = fallback_results.get(fid8, {})
                sel = r.get('selected_key')
                ext = r.get('extracted_value')
                reasoning = r.get('reasoning', '')
                conf = float(r.get('confidence', 0.0))

                if sel and sel not in data:
                    sel = None

                res_obj = {
                    'json_key': sel,
                    'extracted_value': ext,
                    'reasoning': f'fallback: {reasoning}' if reasoning else 'fallback',
                    'confidence': conf,
                    'source': 'llm_fallback',
                    'para_index': _para_index_from_location(ent.get('location', '')),
                    'start': ent.get('start', 0),
                    'end': ent.get('end', 0),
                }
                mapping[eid] = res_obj

                cache_key = _sha(f"{template_fingerprint}|{eid}")
                cache.set(cache_key, res_obj)

    # Fill unmatched
    for ent in entities:
        if ent['field_id'] not in mapping:
            mapping[ent['field_id']] = {'json_key': None, 'confidence': 0.0, 'source': 'unmatched'}

    return mapping
