# Demo Script — DOCX Form Filler (3 minutes)

## Overview
This script guides you through a 3-minute Loom demo of the DOCX Form Filler system. The demo flows through the pipeline stages with concrete examples from the actual codebase artifacts.

---

## [0:00–0:25] Intro + Problem Statement (25 sec)

### What You Say:
> "Salut! Asta e sistemul meu pentru completarea automată a formularelor DOCX. Problema e simplă: ai un template cu ~200 de câmpuri marcate cu puncte și underscore-uri, și ai date în JSON. Cum mapezi datele la câmpuri corect, fără a schimba layout-ul? Răspunsul: o pipeline LLM în 5 etape."

### What's on Screen:
- **Show the 5-stage pipeline diagram** from README.md:
  ```
  JSON data (composite keys)
    ↓
  expand_keys (LLM) → JSON (atomic keys)
    ↓
  parser → Fields + context (~195 fields)
    ↓
  key_selector → Top-10 candidates per field
    ↓
  mapper (LLM) → Best key per field (cached)
    ↓
  filler → DOCX values in-place
    ↓
  filled.docx
  ```
- Briefly highlight each arrow and mention that two stages use the LLM (expand_keys, mapper)

### Timing Notes:
- **Text appears slowly** — don't rush; let viewer absorb the flow
- **Mouse pointer** guides eyes from top to bottom

---

## [0:25–1:05] Key Expansion Example (40 sec)

### What You Say:
> "Primul pas: expand_keys. JSON-ul de intrare are chei complexe, cum ar fi 'Asociat 1 - denumire, sediu, telefon'. Sistemul folosește LLM-ul pentru a descompune asta în chei atomice separate: 'Asociat 1 - denumire', 'Asociat 1 - sediu', 'Asociat 1 - telefon'. De ce? Pentru că un câmp nu cere de obicei trei lucruri dintr-o dată. Asta face mapping-ul mai precis."

### What's on Screen:

**Show the pipeline diagram first** (from README):
```
input_date.json (composite keys) + sample_forms.docx
  ↓
expand_keys (LLM) → JSON (atomic keys)
  ↓
parser → Fields + context (~195 fields)
  ↓
key_selector → Top-10 candidates per field
  ↓
mapper (LLM) → Best key per field (cached)
  ↓
filler → DOCX values in-place
  ↓
sample_forms.filled.docx
```

**Panel 1: Input JSON**
```json
{
  "Asociat 1 - denumire, sediu, telefon": "SC ACME SRL, Str. X nr. 5, 0123456789",
  "Valabilitate oferta - ... zile (in litere si cifre)": "90 (nouazeci)",
  "Lista persoane cu functii de decizie in autoritatea contractanta": "dl. Popa Mihai, dna. Ionescu Ana"
}
```

**Panel 2: Expanded JSON (after expand_keys)**
```json
{
  "Asociat 1 - denumire": "SC ACME SRL",
  "Asociat 1 - sediu": "Str. X nr. 5",
  "Asociat 1 - telefon": "0123456789",
  "Valabilitate oferta - ... zile (in litere si cifre)": "90 (nouazeci)",
  "Lista persoane - 1": "dl. Popa Mihai",
  "Lista persoane - 2": "dna. Ionescu Ana"
}
```

**Highlight & Explain:**
- The first key is split into 3 atomic keys
- The "Valabilitate..." key is NOT split (because it asks for BOTH litere AND cifre together)
- The "Lista persoane..." is split because the value contains comma-separated entities and the key implies multiple items
- Total: from ~100 keys to ~140+ atomic keys

### Timing Notes:
- **Side-by-side panels** make the transformation clear
- **Use arrow (→)** or animation between panels to show transformation
- **Speak slowly** through the three examples

---

## [1:05–1:50] Parser + Key Selector Example (45 sec)

### What You Say:
> "Pasul 2: parser citește DOCX-ul și detectează ~195 de câmpuri. Pentru fiecare câmp, extrage contextul din jur — labelul, textul înainte și după. Apoi, pasul 3, key_selector compară contextul cu toate cheile din JSON și clasează cele mai relevante ~10 candidați pentru fiecare câmp. Asta e important: dacă trimiteam toate 140+ cheile la LLM pentru fiecare câmp, ar fi fost prea lent și prea zgomotos. key_selector ține prompt-urile mari și concentrate."

### What's on Screen:

**Panel 1: Parser Output** (from `debug/parser_result.json`)
```json
{
  "field_id": "c1773d69-...",
  "field_type": "TEXT",
  "location": "body/p:55",
  "label": "Denumirea / numele:",
  "ctx_before": "Denumirea / numele ofertantului:",
  "ctx_after": "....................................................................",
  "start": 35,
  "end": 45
}
```

**Explanation:**
- Point out the **label**, **ctx_before**, **ctx_after** fields
- Mention that these form the "context" used for matching

**Panel 2: Key Selector Output** (show a ranked list)
```
Top-10 candidates for field c1773d69:
1. "OPERATOR ECONOMIC - denumirea / numele"          (score: 0.92)
2. "Denumirea / numele ofertantului"                 (score: 0.88)
3. "Operator Economic - denumire"                    (score: 0.81)
4. "Denumirea ofertantului"                          (score: 0.79)
5. "Denumire / Nume"                                 (score: 0.75)
... (5 more candidates)
```

**Highlight:**
- The **score** shows confidence; top-1 is clearly the best match
- This list is small (10 items) vs. all 140+ keys
- The mapper will send this compact list (not all 140+ keys) to the LLM

### Timing Notes:
- **Scroll through** the candidates slowly
- **Point to the score column** — visual confirmation that similarity ranking works
- **Emphasize the contrast**: "10 candidates here vs. 140+ keys total"

---

## [1:50–2:30] Mapper LLM Example (40 sec)

### What You Say:
> "Pasul 4 e mapperul — LLM-ul adevărat. Pentru fiecare pagină, sistemul construiește un prompt cu: textul paginii, câmpurile marcate cu `[[PH:id]]`, și pentru fiecare câmp, top-10 candidații. LLM-ul vede contextul complet al paginii și alege cel mai potrivit. Rezultatele se cache-ază cu o amprentă a template-ului — dacă rulezi din nou, nu se mai apelează LLM-ul."

### What's on Screen:

**Panel 1: Mapper Output** (from `debug/mapping.json` — show 2-3 entries)
```json
{
  "c1773d69...": {
    "json_key": "OPERATOR ECONOMIC - denumirea / numele",
    "extracted_value": null,
    "confidence": 0.95,
    "source": "llm_page_ks",
    "reasoning": "Matches exact label context '(denumirea/numele ofertantului)'"
  },
  "e884e1ec...": {
    "json_key": "Suma de ... lei (in litere si cifre)",
    "extracted_value": null,
    "confidence": 0.92,
    "source": "llm_page_ks",
    "reasoning": "Full context mentions 'suma de ... lei' with both literal and numeric forms"
  },
  "a3b9c2d1...": {
    "json_key": "Lista persoane cu functii de decizie in autoritatea contractanta",
    "extracted_value": "dl. Popa Mihai",
    "confidence": 0.88,
    "source": "llm_page_ks",
    "reasoning": "Multiple people listed; extracted 'dl. Popa Mihai' for this specific placeholder"
  }
}
```

**Highlight:**
- **json_key**: The chosen key from the candidates
- **extracted_value**: If non-null, only this subset of the value is used (e.g., one person from a list)
- **confidence**: 0.88–0.95 is excellent
- **reasoning**: How the LLM justified its choice

**Panel 2: Cache visualization**
```
Cache structure (mapping_cache.json):
{
  "b97d52efc40d5e76": { "json_key": "...", ... },
  "83e34f258f6ddfe8": { "json_key": "...", ... },
  ...
  (193 more entries)
}

When you run again:
  ✓ All 195 results are cached
  → 0 LLM calls needed
  → Instant (< 1 second)
```

### Timing Notes:
- **Zoom in on the third entry** to show `extracted_value` in action — this handles multi-person lists
- **Point to reasoning field** — shows LLM's decision process was sound
- **Show cache stats** — "First run: LLM calls. Next runs: instant."

---

## [2:30–3:00] Filler + Output Example (30 sec)

### What You Say:
> "Pasul 5 final: filler. Ia rezultatele din mapping și scrie valorile în DOCX-ul original. Nu creează paragrafe noi — modifică XML-ul existent, în-place. Asta e cheia: layout-ul rămâne intact. Apoi deschid DOCX-ul completat și arăt că placeholder-urile sunt înlocuite cu date, diacriticele românești sunt păstrate, și formatul e același ca în original."

### What's on Screen:

**Panel 1: Filler Code Snippet** (from `src/filler.py`)
```python
def _write_cell(cell, value, font_name, size_pt):
    """Write value in-place without creating new paragraphs."""
    cell.text = value  # Modifies existing run, preserves formatting
```

**Explanation:**
- In-place modification of XML runs
- Font, size, bold, italic are preserved
- No new paragraphs created → layout stays intact

**Panel 2: Before → After comparison**

**Before (Template):**
```
Denumirea ofertantului: ...................................................................
Suma de ... lei: ___________________________________________________________
```

**After (Filled):**
```
Denumirea ofertantului: SC ACME SRL
Suma de ... lei: 100,000 (o sută de mii)
```

**Highlight:**
- Placeholder dots/underscores → replaced with actual text
- Font size and style unchanged
- Diacritics preserved (ă, î, ș, ț visible if any in the data)
- Page breaks, tables, headers/footers untouched

**Optional: Side-by-side with Word open**
- Open `output/sample_forms.filled.docx` in Word
- Scroll to show 2-3 filled fields
- Point to a diacritic (ă, î) to confirm preservation
- Show a table row with multiple columns filled

### Timing Notes:
- **Keep code snippet brief** — just show it's simple (in-place modification)
- **Before/After side-by-side is most impactful** — visual proof
- **If using Word**: close it gently at the end (don't linger)

---

## Wrap-up (implicit, no extra time)

### Implied Message:
The system is deterministic, caches results, and respects the original DOCX structure. First run populates the cache; all subsequent runs are instant and produce identical output.

---

## 📋 Pre-Recording Checklist

- [ ] **Terminal**: Maximized, font size ≥ 14pt
- [ ] **Text editors/viewers** (VS Code, notepad++, etc.): Open and docked
  - [ ] README.md (with diagram visible)
  - [ ] input/input_date.json ready to show
  - [ ] debug/parser_result.json open
  - [ ] debug/mapping.json open
- [ ] **File explorer**: Ready to navigate folders
- [ ] **Word/LibreOffice**: Closed (or minimized); will open sample_forms.filled.docx near the end
- [ ] **Microphone**: Tested and levels checked
- [ ] **Camera/Screen**: Recording at 1080p or higher
- [ ] **Cache status**: Pipeline has been run once; cache is populated (so no long wait)
- [ ] **Mental rehearsal**: Read through this script 2-3 times before recording

---

## 💡 Delivery Tips

1. **Speak clearly and slowly** — 3 minutes is short; every word matters
2. **Use mouse pointer** — Guide the viewer's eye
3. **Pause briefly** at each diagram or code snippet — let viewers read
4. **Avoid technical jargon** unless you explain it immediately
5. **Show confidence** — you've built something solid; let that shine through
6. **No ad-libs** — stick to the script; practice so it flows naturally

---

## 📹 Loom Recording Steps

1. **Start recording**
2. **Show diagram** (0:00–0:25)
3. **Show JSON before/after expand** (0:25–1:05)
4. **Show parser and key_selector examples** (1:05–1:50)
5. **Show mapping JSON and cache concept** (1:50–2:30)
6. **Show filler code, before/after, and filled.docx** (2:30–3:00)
7. **Stop recording**

---

## Final Notes

- **Total time**: Exactly 3 minutes if you follow the timings
- **Flexibility**: If you finish a section early, pause briefly before moving to the next
- **No edits needed after recording** — everything is linear and sequential
- **The demo tells a complete story**: Problem → Solution → Pipeline → Proof

Good luck! 🎬
