"""key_selector.py - Pre-select top N JSON keys for a placeholder using combined similarity.

Combines:
  - SequenceMatcher ratio on normalized text
  - Token overlap (Jaccard) between context words and key words

Usage:
    python key_selector.py
"""
from __future__ import annotations

import json
import re
from difflib import SequenceMatcher
from typing import Dict, List, Tuple


def _norm(s: str) -> str:
    s = s.lower()
    for a, b in [('ă','a'),('â','a'),('î','i'),('ș','s'),('ş','s'),('ț','t'),('ţ','t')]:
        s = s.replace(a, b)
    return re.sub(r'\s+', ' ', re.sub(r'[^\w\s]', ' ', s)).strip()


def _seq_sim(a: str, b: str) -> float:
    return SequenceMatcher(None, _norm(a), _norm(b)).ratio()


def _token_overlap(a: str, b: str) -> float:
    ta = set(_norm(a).split())
    tb = set(_norm(b).split())
    if not ta or not tb:
        return 0.0
    # Fuzzy Jaccard: count tokens as matching if SequenceMatcher ratio > 0.8
    matched = set()
    for t1 in ta:
        for t2 in tb:
            if t1 == t2 or SequenceMatcher(None, t1, t2).ratio() > 0.8:
                matched.add(t1)
                matched.add(t2)
                break
    # Proportion of matched tokens relative to all unique tokens
    all_tokens = ta | tb
    return len(matched) / len(all_tokens) if all_tokens else 0.0


def score_key(field: Dict, key: str) -> float:
    """Combined similarity score between a field's context and a JSON key.

    Weights parts differently:
    - label (highest): direct hint
    - ctx_next_para (high): hint in parentheses
    - ctx_before/after (medium): immediate context
    - ctx_prev_para (low): previous paragraph
    """
    label = field.get('label', '')
    ctx_before = field.get('ctx_before', '')
    ctx_after = field.get('ctx_after', '')
    ctx_prev = field.get('ctx_prev_para', '')
    ctx_next = field.get('ctx_next_para', '')

    # Score each context separately
    score_label = _token_overlap(label, key) if label else 0
    score_next = _token_overlap(ctx_next, key) if ctx_next else 0
    score_before = _token_overlap(ctx_before, key) if ctx_before else 0
    score_after = _token_overlap(ctx_after, key) if ctx_after else 0
    score_immed = (score_before + score_after) / 2 if (ctx_before or ctx_after) else 0
    score_prev = _token_overlap(ctx_prev, key) if ctx_prev else 0

    # Combined context: catch keys that match across multiple context parts
    all_ctx = ' '.join(filter(None, [label, ctx_before, ctx_after, ctx_prev, ctx_next]))
    score_combined = _token_overlap(all_ctx, key) if all_ctx else 0

    # If label came from next_para (not parentheses), reduce its weight
    # and boost ctx_before/after which are closer to the placeholder
    label_src = field.get('label_source', '')
    if label_src == 'next_para':
        w_label, w_next, w_immed = 0.20, 0.20, 0.30
    else:
        w_label, w_next, w_immed = 0.45, 0.25, 0.10

    weighted = w_label * score_label + w_next * score_next + w_immed * score_immed + 0.05 * score_prev
    return weighted + 0.15 * score_combined


def top_keys(field: Dict, data: Dict, n: int = 10) -> List[Tuple[str, float, str]]:
    """Return top N (key, score, value) tuples for a field."""
    scored = []
    for key, val in data.items():
        s = score_key(field, key)
        scored.append((key, s, str(val)[:80]))
    scored.sort(key=lambda x: -x[1])
    return scored[:n]


def select_all(fields: list, data: Dict, n: int = 10) -> Dict:
    """Run top_keys for every field. Returns {field_id: {location, label, candidates: [{key, value}, ...]}}."""
    result = {}
    for fld in fields:
        if fld.get('field_type') == 'CHECKBOX':
            continue
        fid = fld['field_id']
        candidates = []
        for key, score, val in top_keys(fld, data, n):
            candidates.append({'key': key, 'value': val})
        result[fid] = {
            'location': fld.get('location', ''),
            'label': fld.get('label', ''),
            'candidates': candidates
        }
    return result


if __name__ == '__main__':
    import os, sys, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    project_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    with open(os.path.join(project_dir, 'debug/parser_result.json'), encoding='utf-8') as f:
        pr = json.load(f)
    with open(os.path.join(project_dir, 'cache/input_date_expanded.json'), encoding='utf-8') as f:
        data = json.load(f)

    result = select_all(pr['fields'], data)

    # Save to JSON (without scores — LLM should not see them)
    out_path = os.path.join(project_dir, 'debug/key_selector_result.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"Saved {len(result)} fields to {out_path}")

    # Also print summary
    for fid, entry in result.items():
        loc = entry.get('location', '')
        label = entry.get('label', '')[:40]
        candidates = entry.get('candidates', [])
        top_key = candidates[0]['key'] if candidates else '???'
        print(f"  {loc:15s} [{label:40s}] top: {top_key}")
