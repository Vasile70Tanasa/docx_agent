"""run_pipeline.py - Single-Pass pipeline: parse → map → fill

Usage:
    python run_pipeline.py --docx input/sample_forms.docx --data input/input_date.json
"""
from __future__ import annotations

import argparse
import json
import os
import sys

from dotenv import load_dotenv
load_dotenv()

# Add src/ to path for local imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from parser import parse_document
from mapper import build_mapping
from filler import fill_document


def main() -> None:
    ap = argparse.ArgumentParser(description='Fill DOCX form from JSON data (Single-Pass pipeline)')
    ap.add_argument('--docx', default='input/sample_forms.docx', help='Input DOCX template')
    ap.add_argument('--data', default='input/input_date.json', help='JSON data file (will auto-expand if needed)')
    ap.add_argument('--out', default=None, help='Output DOCX path (default: output/<name>.filled.docx)')
    ap.add_argument('--cache', default='cache/mapping_cache.json')
    ap.add_argument('--model', default='claude-sonnet-4-20250514', help='Model for mapping (default: Sonnet 4)')
    ap.add_argument('--vtop', action='store_true', help='Set vertical alignment=Top to prevent Formularul shift')
    args = ap.parse_args()

    project_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    data_path = args.data
    if not os.path.isabs(data_path):
        data_path = os.path.join(project_dir, data_path)
    if not os.path.exists(data_path):
        raise FileNotFoundError(f"Data file not found: {args.data}")

    docx_path = args.docx
    if not os.path.isabs(docx_path):
        docx_path = os.path.join(project_dir, docx_path)
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"DOCX not found: {args.docx}")

    if args.vtop:
        from pathlib import Path
        from set_vertical_alignment import set_valign_top
        vtop_path = Path(docx_path).with_stem(Path(docx_path).stem + '_vtop')
        print(f'Applying vertical alignment=Top -> {vtop_path.name}')
        set_valign_top(Path(docx_path), vtop_path)
        docx_path = str(vtop_path)

    with open(data_path, encoding='utf-8') as f:
        data = json.load(f)

    # Auto-expand composite keys if expanded file doesn't exist or is older than source
    cache_dir = os.path.join(project_dir, 'cache')
    os.makedirs(cache_dir, exist_ok=True)
    expanded_name = os.path.splitext(os.path.basename(data_path))[0] + '_expanded.json'
    expanded_path = os.path.join(cache_dir, expanded_name)
    need_expand = not os.path.exists(expanded_path)
    if not need_expand:
        need_expand = os.path.getmtime(data_path) > os.path.getmtime(expanded_path)
    if need_expand:
        print('Expanding composite keys ...')
        from expand_keys import expand_keys
        data = expand_keys(data)
        with open(expanded_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f'  Saved {len(data)} keys to {os.path.basename(expanded_path)}')
    else:
        print(f'Using cached expanded keys: {os.path.basename(expanded_path)}')
        with open(expanded_path, encoding='utf-8') as f:
            data = json.load(f)

    # Step 1: Parse
    print('Parsing document ...')
    parsed = parse_document(docx_path)
    fields = parsed['fields']
    tables = parsed['tables']
    template_fp = parsed['template_fingerprint']

    debug_dir = os.path.join(project_dir, 'debug')
    os.makedirs(debug_dir, exist_ok=True)
    output_dir = os.path.join(project_dir, 'output')
    os.makedirs(output_dir, exist_ok=True)
    parser_out = os.path.join(debug_dir, 'parser_result.json')

    with open(parser_out, 'w', encoding='utf-8') as f:
        json.dump(parsed, f, ensure_ascii=False, indent=2)

    cb_groups = len({f['group_id'] for f in fields if f['field_type'] == 'CHECKBOX' and f.get('group_id')})
    print(f'  Fields: {len(fields)}  Checkbox groups: {cb_groups}  Tables: {len(tables)}')

    # Step 2: Map
    print('Mapping fields (LLM Single-Pass) ...')
    mapping = build_mapping(
        template_fingerprint=template_fp,
        fields=fields,
        tables=tables,
        data=data,
        cache_path=os.path.join(project_dir, args.cache),
        model=args.model,
        docx_path=docx_path,
    )

    mapped = sum(1 for v in mapping.values() if v.get('json_key'))
    by_src = {}
    for v in mapping.values():
        s = v.get('source', '?')
        by_src[s] = by_src.get(s, 0) + 1
    print(f'  Mapped: {mapped}/{len(mapping)}  by source: {by_src}')

    # Output mapping result
    mapping_out = os.path.join(debug_dir, 'mapping.json')
    with open(mapping_out, 'w', encoding='utf-8') as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)

    # Step 3: Fill
    print('Filling document ...')
    if args.out:
        out_path = args.out if os.path.isabs(args.out) else os.path.join(project_dir, args.out)
    else:
        stem = os.path.splitext(os.path.basename(docx_path))[0]
        out_path = os.path.join(output_dir, f'{stem}.filled.docx')
    result = fill_document(
        docx_in=docx_path,
        docx_out=out_path,
        fields=fields,
        tables=tables,
        mapping=mapping,
        data=data,
    )
    print(f'  Filled: {result["filled"]}  Skipped: {result["skipped"]}  Failed: {result["failed"]}')
    print(f'Output: {out_path}')


if __name__ == '__main__':
    main()
