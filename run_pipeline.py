"""run_pipeline.py - Single-Pass pipeline: parse → map → fill

Usage:
    python run_pipeline.py --docx sample_forms.docx --data input_date.json
"""
from __future__ import annotations

import argparse
import json
import os

from dotenv import load_dotenv
load_dotenv()

from parser import parse_document
from mapper import build_mapping
from filler import fill_document


def main() -> None:
    ap = argparse.ArgumentParser(description='Fill DOCX form from JSON data (Single-Pass pipeline)')
    ap.add_argument('--docx', default='sample_forms.docx', help='Input DOCX template')
    ap.add_argument('--data', default='input_date.json', help='JSON data file')
    ap.add_argument('--out', default='filled.docx', help='Output DOCX path')
    ap.add_argument('--cache', default='cache/mapping_cache.json')
    ap.add_argument('--model', default='claude-haiku-4-5-20251001', help='Model for mapping')
    ap.add_argument('--vtop', action='store_true', help='Set vertical alignment=Top to prevent Formularul shift')
    args = ap.parse_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))

    data_path = args.data
    if not os.path.isabs(data_path):
        data_path = os.path.join(script_dir, data_path)
    if not os.path.exists(data_path):
        raise FileNotFoundError(f"Data file not found: {args.data}")

    docx_path = args.docx
    if not os.path.isabs(docx_path):
        docx_path = os.path.join(script_dir, docx_path)
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

    # Step 1: Parse
    print('Parsing document ...')
    parsed = parse_document(docx_path)
    fields = parsed['fields']
    tables = parsed['tables']
    template_fp = parsed['template_fingerprint']

    parser_out = os.path.join(script_dir, 'parser_result.json')
    os.makedirs(os.path.dirname(parser_out) or '.', exist_ok=True)

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
        cache_path=os.path.join(script_dir, args.cache),
        model=args.model,
    )

    mapped = sum(1 for v in mapping.values() if v.get('json_key'))
    by_src = {}
    for v in mapping.values():
        s = v.get('source', '?')
        by_src[s] = by_src.get(s, 0) + 1
    print(f'  Mapped: {mapped}/{len(mapping)}  by source: {by_src}')

    # Output mapping result
    mapping_out = os.path.join(script_dir, 'mapping.json')
    with open(mapping_out, 'w', encoding='utf-8') as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)

    # Step 3: Fill
    print('Filling document ...')
    out_path = args.out if os.path.isabs(args.out) else os.path.join(script_dir, args.out)
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
