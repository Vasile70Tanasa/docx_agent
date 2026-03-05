"""mapper.py - Single-Pass LLM Field-to-JSON-Key Mapper

This mapper discards complex Python heuristics (like Jaccard) and label enrichment.
Instead, it groups fields into sequential text chunks and asks the LLM to fill the 
placeholders by choosing directly from the available JSON keys.

No keys are removed from the pool after being used (to allow multiple same-key mappings).
Only the JSON keys (not the values) are sent, to keep the context clean.
"""
from __future__ import annotations

import json
import os
import re
import hashlib
from typing import Any, Dict, List, Optional, Tuple

from dotenv import load_dotenv
load_dotenv()

try:
    import anthropic
except ImportError:
    anthropic = None


def _sha(s: str) -> str:
    return hashlib.sha256(s.encode('utf-8')).hexdigest()[:16]


def _para_index_from_location(location: str) -> Optional[int]:
    m = re.findall(r'p:(\d+)', location or '')
    return int(m[-1]) if m else None


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


def _llm_map_chunk(client: Any, chunk: List[Dict], data: Dict[str, Any], model: str) -> Dict[str, Tuple[Optional[str], float, Optional[str], Optional[str]]]:
    """Sends a single chunk of fields to the LLM (Anthropic) to be mapped against json_keys."""
    
    # Build text representation of the chunk
    chunk_text = ""
    for ent in chunk:
        eid = ent['id']
        t_type = ent.get('field_type', 'TEXT')
        
        ctx_before = ent.get('ctx_before', '') or ''
        ctx_after  = ent.get('ctx_after',  '') or ''
        prev_p     = ent.get('ctx_prev_para', '') or ''
        next_p     = ent.get('ctx_next_para', '') or ''
        
        # DYNAMIC CONTEXT PRESERVATION:
        # Even if the line context is long, if adjacent paragraphs contain a hint '(',
        # we MUST include them to avoid "context blindness".
        ctx = ""
        if '(' in prev_p or len(ctx_before) < 50:
            ctx += prev_p + "\n"
            
        ctx += ctx_before + f" [[PLACEHOLDER_ID: {eid}]] " + ctx_after
        
        if '(' in next_p or len(ctx_after) < 50:
            ctx += "\n" + next_p
            
        chunk_text += f"---\nField Type: {t_type}\nContext:\n{ctx.strip()}\n\n"

    payload = {
        "task": "You are a strict data entry mapping system. Read the provided text chunks, each containing a [[PLACEHOLDER_ID: ...]]. For each placeholder, select the EXACT matching JSON key from the 'available_data'.",
        "rules": [
            "1. You MUST select an exact string key from 'available_data', or return null if none is a good fit.",
            "2. Pay close attention to the surrounding text (before and after the placeholder).",
            "3. CRITICAL: If multiple keys are valid (e.g. variants of the same concept), STRICTLY count the exact overlapping words between the JSON key and the immediate context ('ctx_before' and 'ctx_after') and ALWAYS choose the key with the highest literal word overlap.",
            "4. If the context surrounding the placeholder specifically asks for a part of the value (e.g., just the 'ziua/day' or 'luna/month' from a full date, or just one item from a list of partners), you MUST provide that exact partial substring (from the data value) in the 'extracted_value' field.",
            "5. If no partial extraction is needed, leave 'extracted_value' as null.",
            "6. Return a JSON array: [{\"id\": \"<PLACEHOLDER_ID>\", \"selected_key\": \"<JSON_KEY>\" | null, \"reasoning\": \"<MATCHING_LOGIC>\", \"extracted_value\": \"<SUBSTRING>\" | null, \"confidence\": 0.0-1.0}]"
        ],
        "available_data": data,
        "document_chunk": chunk_text
    }

    try:
        # Move system instructions into user prompt to avoid "multiple user roles" error
        # which can happen if some proxies/versions mangle the system/user boundary.
        system_instr = "You output ONLY valid JSON arrays. No markdown, no conversational text."
        user_content = f"{system_instr}\n\nTask:\n{json.dumps(payload, ensure_ascii=False)}"

        msg_list = [
            {"role": "user", "content": user_content}
        ]
        
        # Anthropic Message API call
        resp = client.messages.create(
            model=model,
            max_tokens=4096,
            temperature=0,
            messages=msg_list,
        )
        content = (resp.content[0].text or "").strip()
        
        # Robust JSON Array Extraction
        try:
            # First, clean markdown blocks if they exist
            content = re.sub(r"^```(?:json)?\s*", "", content)
            content = re.sub(r"\s*```$", "", content).strip()
            
            # Use regex to find the outermost [...] block
            m = re.search(r"(\[[\s\S]*\])", content)
            if m:
                items = json.loads(m.group(1))
            else:
                items = json.loads(content)
        except Exception as json_err:
            print(f"LLM Mapping error (JSON parsing failed): {json_err}\nRaw Content: {content[:100]}...")
            return {}
        
        result = {}
        for item in items:
            eid = item.get("id")
            sel = item.get("selected_key")
            ext = item.get("extracted_value")
            reasoning = item.get("reasoning")
            conf = float(item.get("confidence", 0.0))
            if sel not in data:
                sel = None
            if eid:
                result[eid] = (sel, conf, ext, reasoning)
        return result
    except Exception as e:
        print(f"LLM Mapping error (Anthropic): {e}")
        return {}


def build_mapping(
    template_fingerprint: str,
    fields: List[Dict],
    tables: List[Dict],
    data: Dict[str, Any],
    cache_path: str = 'cache/mapping_cache.json',
    model: str = 'claude-3-7-sonnet-latest',
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
            entities.append({
                'id':            f['field_id'],
                'ctx_before':    f.get('ctx_before', ''),
                'ctx_after':     f.get('ctx_after',  ''),
                'ctx_prev_para': f.get('ctx_prev_para', ''),
                'ctx_next_para': f.get('ctx_next_para', ''),
                'field_type':    ftype,
                'location':      f.get('location', ''),
                'start':         f.get('start', 0),
                'end':           f.get('end', 0),
            })

    for gid, options in group_options.items():
        ctx = group_ctx[gid]
        entities.append({
            'id':            gid,
            'ctx_before':    ctx.get('ctx_before', ''),
            'ctx_after':     ctx.get('ctx_after',  ''),
            'ctx_prev_para': ctx.get('ctx_prev_para', ''),
            'ctx_next_para': ctx.get('ctx_next_para', ''),
            'field_type':    'CHECKBOX_GROUP',
            'location':      ctx.get('location', ''),
            'start':         ctx.get('start', 0),
            'end':           ctx.get('end', 0),
        })

    for t in tables:
        col_info = ' '.join(h for h in t.get('col_headers', []) if h.strip())
        entities.append({
            'id':         t['field_id'],
            'ctx_before': '',
            'ctx_after':  col_info[:300],
            'ctx_prev_para': '',
            'ctx_next_para': '',
            'field_type': 'TABLE',
            'location':   t.get('location', ''),
            'start':      0,
            'end':        0,
        })

    # Sort entities logically (by para_index, then start)
    def _sort_key(e):
        pid = _para_index_from_location(e.get('location', '')) or 999999
        return (pid, e.get('start', 0))
    entities.sort(key=_sort_key)

    # --- 1.5. Process TABLE entities with mathematical similarity ---
    from difflib import SequenceMatcher
    
    def _norm(s: str) -> str:
        s = s.lower().replace('_', ' ')
        for a, b in [('ă','a'),('â','a'),('î','i'),('ș','s'),('ş','s'),('ț','t'),('ţ','t')]:
            s = s.replace(a, b)
        return re.sub(r'\s+', ' ', re.sub(r'[^\w\s]', ' ', s)).strip()

    def _sim(a: str, b: str) -> float:
        return SequenceMatcher(None, _norm(a), _norm(b)).ratio()

    # Identify JSON keys that lead to Lists/Arrays
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

    non_table_entities = []
    for ent in entities:
        if ent['field_type'] == 'TABLE':
            col_info = ent.get('ctx_after', '')
            best_k, best_s = None, 0.0
            
            for ak in array_keys:
                sc = _sim(col_info, ak)
                if sc > best_s:
                    best_s, best_k = sc, ak
                    
            if best_k and best_s > 0.15: # Lower threshold since column names concatenated vs single key
                mapping[ent['id']] = {
                    'json_key': best_k,
                    'confidence': best_s,
                    'source': 'direct_table',
                    'para_index': _para_index_from_location(ent.get('location', '')),
                    'start': 0, 'end': 0
                }
            else:
                mapping[ent['id']] = {'json_key': None, 'confidence': 0.0, 'source': 'unmatched_table'}
        else:
            non_table_entities.append(ent)
            
    entities = non_table_entities

    # 2. Process textual chunks using LLM
    CHUNK_SIZE = 10
    for i in range(0, len(entities), CHUNK_SIZE):
        chunk = entities[i : i + CHUNK_SIZE]
        
        # Filter what needs LLM vs what's in cache
        to_process = []
        for ent in chunk:
            eid = ent['id']
            cache_key = _sha(f"{template_fingerprint}|{eid}")
            hit = cache.get(cache_key)
            if hit and hit.get('json_key') in json_keys:
                mapping[eid] = hit
            else:
                to_process.append(ent)
                
        if to_process:
            client = anthropic.Anthropic(api_key=api_key)
            results = _llm_map_chunk(client, to_process, data, model)
            for ent in to_process:
                eid = ent['id']
                sel, conf, ext, reasoning = results.get(eid, (None, 0.0, None, None))
                
                res_obj = {
                    'json_key': sel, 
                    'extracted_value': ext,
                    'reasoning': reasoning,
                    'confidence': conf, 
                    'source': 'llm_v4',
                    'para_index': _para_index_from_location(ent.get('location', '')),
                    'start': ent.get('start', 0),
                    'end': ent.get('end', 0),
                    'full_text': ent.get('full_text', ''),
                    'label': ent.get('label', ''),
                    'label_source': ent.get('label_source', ''),
                    'ctx_before': ent.get('ctx_before', ''),
                    'ctx_after': ent.get('ctx_after', ''),
                    'ctx_prev_para': ent.get('ctx_prev_para', ''),
                    'ctx_next_para': ent.get('ctx_next_para', '')
                }
                mapping[eid] = res_obj
                
                cache_key = _sha(f"{template_fingerprint}|{eid}")
                cache.set(cache_key, res_obj)

    # Fill unmatched
    for ent in entities:
        if ent['id'] not in mapping:
            mapping[ent['id']] = {'json_key': None, 'confidence': 0.0, 'source': 'unmatched'}

    return mapping
