"""expand_keys.py - Expand composite JSON keys into atomic sub-keys.

Uses a powerful LLM to decompose composite keys (those containing multiple
values like "name, address, phone") into individual atomic keys.

Usage:
    python expand_keys.py [--input input_date.json] [--output input_date_expanded.json] [--model claude-sonnet-4-20250514]
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys

from dotenv import load_dotenv
load_dotenv()

import anthropic


PROMPT = """You are a data pre-processor for Romanian procurement forms.

I have a JSON dictionary where some keys are COMPOSITE — they describe multiple pieces of information bundled into one value.

Your task: identify composite keys and split them into atomic sub-keys.

RULES:
1. Keep ALL original keys unchanged with their original values.
2. For each composite key, add new atomic keys AFTER it.
3. New key names must start with the relevant prefix from the original key.
4. Each new key must have exactly ONE atomic value (not comma-separated, not combined).
5. For array values (JSON strings starting with "["), expand each element with an index.
6. Do NOT split keys that are already atomic (single value).
7. Do NOT invent data — only decompose what exists in the value.
8. When a key name mentions sub-components (e.g. "zi / luna / an", "serie, numar, CNP"), split the value into those exact sub-components.

EXAMPLES of splits:

Original: "Asociat 1 - denumire, sediu, telefon": "SC Partener 1 SRL, Str. Parteneriat 1, Bucuresti, 0210000000"
Add:
  "Asociat 1 - denumire": "SC Partener 1 SRL"
  "Asociat 1 - sediu": "Str. Parteneriat 1, Bucuresti"
  "Asociat 1 - telefon": "0210000000"

Original: "Reprezentat prin - nume si calitate": "Ionescu Mihai, Administrator"
Add:
  "Reprezentat prin - nume": "Ionescu Mihai"
  "Reprezentat prin - calitate": "Administrator"

Original: "Contributie fiecare parte (%)": "[\"60% SC Partener 1 SRL\", \"40% SC Partener 2 SRL\"]"
Add:
  "Contributie parte 1 - procent": "60%"
  "Contributie parte 1 - denumire": "SC Partener 1 SRL"
  "Contributie parte 2 - procent": "40%"
  "Contributie parte 2 - denumire": "SC Partener 2 SRL"

Original: "Data parafare (zi / luna / an)": "2025-10-01"
Add:
  "Data parafare - zi": "01"
  "Data parafare - luna": "10"
  "Data parafare - an": "2025"

Original: "Data": "2025-10-08"
DO NOT split (already atomic — the key name has no sub-components like zi/luna/an).

INPUT JSON:
{input_json}

Return ONLY the expanded JSON dictionary (all original keys + new atomic keys). No markdown, no explanation."""


def expand_keys(input_data: dict, model: str = 'claude-sonnet-4-20250514') -> dict:
    """Send data to LLM for key expansion."""
    client = anthropic.Anthropic(api_key=os.environ['ANTHROPIC_API_KEY'])

    input_json = json.dumps(input_data, ensure_ascii=False, indent=2)

    resp = client.messages.create(
        model=model,
        max_tokens=16384,
        temperature=0,
        messages=[{"role": "user", "content": PROMPT.format(input_json=input_json)}],
    )

    content = (resp.content[0].text or "").strip()
    # Strip markdown fences if present
    content = re.sub(r"^```(?:json)?\s*", "", content)
    content = re.sub(r"\s*```$", "", content).strip()

    result = json.loads(content)
    return result


def main() -> None:
    if not isinstance(sys.stdout, type(None)):
        import io
        if hasattr(sys.stdout, 'buffer'):
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    ap = argparse.ArgumentParser(description='Expand composite JSON keys into atomic sub-keys')
    ap.add_argument('--input', default='input_date.json', help='Input JSON file')
    ap.add_argument('--output', default='input_date_expanded.json', help='Output JSON file')
    ap.add_argument('--model', default='claude-sonnet-4-20250514', help='Model to use')
    args = ap.parse_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))

    input_path = args.input
    if not os.path.isabs(input_path):
        input_path = os.path.join(script_dir, input_path)

    with open(input_path, encoding='utf-8') as f:
        data = json.load(f)

    print(f'Input: {len(data)} keys')
    print(f'Calling LLM ({args.model}) to expand composite keys...')

    expanded = expand_keys(data, model=args.model)

    new_keys = set(expanded.keys()) - set(data.keys())
    print(f'Output: {len(expanded)} keys (+{len(new_keys)} new)')

    if new_keys:
        print(f'\nNew atomic keys:')
        for k in sorted(new_keys):
            print(f'  "{k}": "{expanded[k]}"')

    output_path = args.output
    if not os.path.isabs(output_path):
        output_path = os.path.join(script_dir, output_path)

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(expanded, f, ensure_ascii=False, indent=2)

    print(f'\nSaved to {output_path}')


if __name__ == '__main__':
    main()
