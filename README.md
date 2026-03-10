# DOCX Form Filler — Agentic AI Pipeline

Automatically fills multi-page DOCX forms from JSON data using an LLM-powered pipeline.
Given a template DOCX with visual blanks (dots, underscores, empty spaces) and a JSON file with field values, the system identifies each placeholder, maps it to the correct JSON key, and produces a filled copy — preserving the original layout, fonts, and styles.

📹 **[Watch the demo (~5 min)](https://www.loom.com/share/c5be301b0ff844e08cab330936f67815)**

## Architecture

The pipeline runs in five stages:

```
input_date.json (composite keys) + sample_forms.docx
  ↓
expand_keys (LLM) → JSON (atomic keys)
  ↓
parser → Fields + context (~200 fields)
  ↓
key_selector → Top-10 candidates per field
  ↓
mapper (LLM) → Best key per field (cached)
  ↓
filler → DOCX values in-place
  ↓
sample_forms.filled.docx
```

1. **expand_keys** — Pre-processes the input JSON using an LLM call to split composite keys (e.g. `"Asociat 1 - denumire, sediu, telefon": "..."`) into atomic sub-keys (`"Asociat 1 - denumire"`, `"Asociat 1 - sediu"`, etc.). This makes downstream mapping more precise.

2. **parser** (`parser.py`) — Reads the DOCX template and detects all fillable fields: text placeholders (dot/underscore sequences), checkboxes, and table cells. Extracts surrounding context (labels, adjacent text) for each field. Produces a structured representation with ~200 fields.

3. **key_selector** (`key_selector.py`) — Bridges parser output and mapper input. For each field, calculates a similarity score against all JSON keys using token overlap (Jaccard distance with fuzzy matching) and sequence similarity (SequenceMatcher ratio). Ranks keys by score and returns the top ~10 candidates per field. This pre-filtering is crucial: sending all 140+ keys to the LLM for every field would be expensive and noisy; the key_selector keeps the mapper's prompts focused on realistic candidates, speeding up both processing and inference.

4. **mapper** (`mapper.py`) — Groups fields into page-like chunks and sends each chunk to Claude (Sonnet 4) along with the pre-selected candidate keys. The LLM sees the full page context with `[[PH:id]]` markers and selects the best JSON key for each placeholder. A two-pass strategy handles edge cases: fields that receive no match in the first pass are retried with the complete key list. Results are cached by `(template_fingerprint, field_id)` to ensure subsequent runs are deterministic and instant.

5. **filler** (`filler.py`) — Writes values into the DOCX template in-place, replacing placeholder characters with actual data while preserving the original run formatting (font, size, bold, italic). Handles text fields, checkboxes (Unicode ☒/☐), and table rows.

### Key Design Decisions

- **LLM for mapping, not rules** — Romanian procurement forms have diverse labeling patterns. Rule-based matching would require endless special cases; the LLM generalizes naturally.
- **Key pre-selection** — Sending all 140 JSON keys to the LLM for every field would be expensive and noisy. The `key_selector` narrows it to ~10 candidates using fast string similarity, keeping LLM prompts focused.
- **Deterministic caching** — Mapping results are cached by `(template_fingerprint, field_id)`. Subsequent runs on the same inputs skip all LLM calls, producing identical output instantly.
- **In-place filling** — The filler modifies existing XML runs rather than creating new paragraphs, so Word's layout (margins, page breaks, headers/footers, tables) remains unchanged.

### Exercise Compliance

- **Output format** — Produces `<original_name>.filled.docx` (e.g. `sample_forms.filled.docx`).
- **Paragraph wrapping** — Values that exceed the placeholder width use Word's native text wrapping. The filler does not truncate; it replaces placeholder characters in-place, so Word applies its standard line-break rules. Multi-line JSON values (with `\n`) are normalized to spaces; for true paragraph breaks, the source would need separate fields.
- **Multiple occurrences** — When a JSON key maps to several placeholders (e.g. header + body), the mapper assigns the same key to each. The filler then writes the same value to all mapped locations, so all instances are filled consistently.
- **Unicode & special characters** — No stripping of characters. Romanian diacritics (ă, â, î, ș, ț) and other Unicode text are preserved in both mapping and filling.
- **Determinism** — Running the program multiple times on the same inputs produces identical outputs. The first run populates the cache; all subsequent runs use cached mappings and skip LLM calls entirely.

## Setup

```bash
# 1. Clone and enter the project
git clone <repo-url>
cd docx_agent

# 2. Create a virtual environment
python -m venv .venv
.venv\Scripts\activate   # Windows
# source .venv/bin/activate  # Linux/Mac

# 3. Install dependencies
pip install -r requirements.txt

# 4. Set your Anthropic API key
echo ANTHROPIC_API_KEY=sk-ant-... > .env
```

### Dependencies

| Package | Purpose |
|---------|---------|
| `python-docx` | Read/write DOCX files |
| `anthropic` | Claude API for key expansion and field mapping |
| `python-dotenv` | Load API key from `.env` file |

## Usage

```bash
# Run with defaults (uses input/sample_forms.docx + input/input_date.json)
python src/run_pipeline.py

# Custom files
python src/run_pipeline.py --docx input/my_form.docx --data input/my_data.json

# Specify output path
python src/run_pipeline.py --out output/custom_name.filled.docx

# Use a different model
python src/run_pipeline.py --model claude-haiku-4-5-20241001
```

Output is written to `output/<template_name>.filled.docx` by default.

## Project Structure

```
docx_agent/
├── src/
│   ├── run_pipeline.py    # Main entry point — orchestrates all steps
│   ├── expand_keys.py     # LLM-based JSON key expansion
│   ├── parser.py          # DOCX field detection and context extraction
│   ├── key_selector.py    # Fuzzy similarity ranking for candidate keys
│   ├── mapper.py          # LLM-based field-to-key mapping with caching
│   └── filler.py          # In-place DOCX value writing
├── input/                 # Provided by exercise
│   ├── sample_forms.docx  # Template DOCX form (~200 fields)
│   └── input_date.json    # Source data (JSON, ~100 composite keys)
├── output/                # Generated after running pipeline
│   └── sample_forms.filled.docx  # Filled DOCX (final output)
├── cache/                 # Generated after first run
│   ├── mapping_cache.json        # Cached LLM mapping results (193 entries)
│   └── input_date_expanded.json  # Expanded JSON keys (~140 atomic keys)
├── debug/                 # Intermediate analysis files (generated)
│   ├── parser_result.json         # Field detection results
│   ├── mapping.json               # Field-to-key mapping details
│   └── key_selector_result.json   # Candidate ranking per field
├── requirements.txt
├── README.md
├── DEMO_SCRIPT.md         # Script for 3-minute video demo
└── .env                   # API key (not committed)
```

**Legend:**
- **input/** — Provided by the exercise (template + data)
- **output/** — Final deliverable: `sample_forms.filled.docx`
- **cache/** — Persistent cache for deterministic reruns
- **debug/** — Intermediate artifacts for analysis/verification

## Known Limitations

1. **LLM non-determinism on first run** — The first execution calls Claude to map fields. While `temperature=0` is used, LLM outputs are not guaranteed to be bit-identical across runs. However, once cached, all subsequent runs are fully deterministic.

2. **Checkbox groups** — Checkboxes are detected and grouped, but matching them to JSON values depends on exact label correspondence. Ambiguous checkbox labels may not map correctly.

3. **Complex table structures** — The system handles simple data tables (header row + data rows) well. Nested tables or merged cells may not fill correctly.

4. **Layout edge cases** — Very long values that exceed the placeholder width rely on Word's native text wrapping. The filler does not attempt to resize cells or adjust spacing.

5. **Language assumption** — The LLM prompt and key_selector heuristics are tuned for Romanian procurement forms. Adapting to other languages would require prompt adjustments.

---

# DOCX Form Filler — Pipeline AI Agentic (RO)

Completează automat formulare DOCX multi-pagină din date JSON folosind un pipeline bazat pe LLM.
Dat fiind un template DOCX cu spații vizuale de completat (puncte, underscore-uri, spații goale) și un fișier JSON cu valorile câmpurilor, sistemul identifică fiecare placeholder, îl asociază cu cheia JSON corectă și produce o copie completată — păstrând layout-ul original, fonturile și stilurile.

## Arhitectură

Pipeline-ul rulează în cinci etape:

```
input_date.json (chei compozite) + sample_forms.docx
  ↓
expand_keys (LLM) → Date JSON (chei atomice)
  ↓
parser → Câmpuri + context (~200 câmpuri)
  ↓
key_selector → Top-10 candidați per câmp
  ↓
mapper (LLM) → Cea mai bună cheie per câmp (cached)
  ↓
filler → Valori DOCX in-place
  ↓
sample_forms.filled.docx
```

1. **expand_keys** — Pre-procesează JSON-ul de intrare folosind un apel LLM pentru a descompune cheile compozite (ex. `"Asociat 1 - denumire, sediu, telefon": "..."`) în sub-chei atomice (`"Asociat 1 - denumire"`, `"Asociat 1 - sediu"`, etc.).

2. **parser** (`parser.py`) — Citește template-ul DOCX și detectează toate câmpurile completabile: placeholder-e text (secvențe de puncte/underscore), checkbox-uri și celule de tabel. Extrage contextul din jur (etichete, text adiacent) pentru fiecare câmp. Produce o reprezentare structurată cu ~200 de câmpuri.

3. **key_selector** (`key_selector.py`) — Leagă output-ul parser-ului și input-ul mapper-ului. Pentru fiecare câmp, calculează un scor de similaritate față de toate cheile JSON folosind suprapunere de tokeni (Jaccard fuzzy) și similaritate de secvență (SequenceMatcher). Clasează cheile după scor și returnează top 10 candidate per câmp. Această pre-filtrare e crucială: trimiterea tuturor 140+ chei la LLM pentru fiecare câmp ar fi costisitoare și zgomotoasă; key_selector ține prompt-urile mapper-ului concentrate pe candidații realiști, accelerând atât procesarea cât și inferența.

4. **mapper** (`mapper.py`) — Grupează câmpurile pe secțiuni de pagină și trimite fiecare secțiune la Claude (Sonnet 4) împreună cu cheile pre-selectate. LLM-ul vede contextul complet al paginii cu markeri `[[PH:id]]` și selectează cea mai potrivită cheie JSON pentru fiecare placeholder. O strategie cu două treceri tratează cazurile limită: câmpurile care nu primesc niciun match la prima trecere se re-procesează cu lista completă de chei. Rezultatele se cache-ază după `(template_fingerprint, field_id)` pentru a asigura că rulările ulterioare sunt deterministe și instantanee.

5. **filler** (`filler.py`) — Scrie valorile în template-ul DOCX in-place, înlocuind caracterele placeholder cu datele reale, păstrând formatarea originală (font, mărime, bold, italic). Gestionează câmpuri text, checkbox-uri (Unicode ☒/☐) și rânduri de tabel.

### Decizii de Design

- **LLM pentru mapping, nu reguli** — Formularele românești de achiziții publice au tipare diverse de etichetare. Un matching bazat pe reguli ar necesita cazuri speciale interminabile; LLM-ul generalizează natural.
- **Pre-selecție de chei** — Trimiterea tuturor celor 140 de chei JSON la LLM pentru fiecare câmp ar fi costisitoare și zgomotoasă. `key_selector` le restrânge la ~10 candidate folosind similaritate rapidă pe string-uri.
- **Cache determinist** — Rezultatele mapping-ului sunt cache-uite după `(template_fingerprint, field_id)`. Rulările ulterioare pe aceleași intrări sar peste toate apelurile LLM, producând output identic instantaneu.
- **Completare in-place** — Filler-ul modifică run-urile XML existente în loc să creeze paragrafe noi, astfel încât layout-ul Word (margini, page breaks, headers/footers, tabele) rămâne neschimbat.

### Conformitate cu Cerințele

- **Format output** — Produce `<nume_original>.filled.docx` (ex. `sample_forms.filled.docx`).
- **Paragraph wrapping** — Valorile care depășesc lățimea placeholder-ului folosesc word wrap-ul nativ al Word. Filler-ul nu trunchiază; înlocuiește caracterele placeholder in-place, astfel Word aplică regulile standard de line-break. Valorile JSON multi-linie (cu `\n`) sunt normalizate la spații; pentru paragrafe reale, sursa ar trebui să aibă câmpuri separate.
- **Apariții multiple** — Când o cheie JSON se mapează la mai multe placeholder-uri (ex. header + body), mapper-ul atribuie aceeași cheie fiecăruia. Filler-ul scrie apoi aceeași valoare în toate locațiile mapate, astfel toate instanțele sunt completate consistent.
- **Unicode și caractere speciale** — Nu se elimină caractere. Diacriticele românești (ă, â, î, ș, ț) și alte texte Unicode sunt păstrate atât la mapping cât și la completare.
- **Determinism** — Rularea programului de mai multe ori pe aceleași intrări produce output identic. Prima rulare populează cache-ul; toate rulările ulterioare folosesc mapping-urile din cache și sar peste apelurile LLM.

## Instalare

```bash
# 1. Clonare
git clone <repo-url>
cd docx_agent

# 2. Creare mediu virtual
python -m venv .venv
.venv\Scripts\activate   # Windows

# 3. Instalare dependențe
pip install -r requirements.txt

# 4. Configurare cheie API
echo ANTHROPIC_API_KEY=sk-ant-... > .env
```

## Utilizare

```bash
# Rulare cu valorile implicite
python src/run_pipeline.py

# Fișiere personalizate
python src/run_pipeline.py --docx input/formular.docx --data input/date.json
```

Rezultatul se salvează în `output/<nume_template>.filled.docx`.

## Structura Proiectului

```
docx_agent/
├── src/
│   ├── run_pipeline.py    # Punct de intrare — orchestrează toți pașii
│   ├── expand_keys.py     # Expandare compozite → chei atomice (LLM)
│   ├── parser.py          # Detectare câmpuri + extragere context
│   ├── key_selector.py    # Clasificare similaritate pentru candidații de chei
│   ├── mapper.py          # Mapare câmp → cheie (LLM + cache)
│   └── filler.py          # Scriere valori in-place în DOCX
├── input/                 # Furnizat de exercițiu
│   ├── sample_forms.docx  # Template DOCX (~200 câmpuri)
│   └── input_date.json    # Date JSON (~100 chei compozite)
├── output/                # Generat după rularea pipeline-ului
│   └── sample_forms.filled.docx  # DOCX completat (deliverable final)
├── cache/                 # Generat după prima rulare
│   ├── mapping_cache.json        # Rezultate cache-uite (193 intrări)
│   └── input_date_expanded.json  # Chei expandate (~140 chei atomice)
├── debug/                 # Fișiere intermediare de analiză (generate)
│   ├── parser_result.json         # Rezultate detectare câmpuri
│   ├── mapping.json               # Detalii mapare câmp→cheie
│   └── key_selector_result.json   # Clasificare candidați per câmp
├── requirements.txt
├── README.md
├── DEMO_SCRIPT.md         # Script pentru demo video (3 minute)
└── .env                   # Cheie API (nu se commitează)
```

**Semnificație:**
- **input/** — Furnizat de exercițiu (template + date)
- **output/** — Deliverable final: `sample_forms.filled.docx`
- **cache/** — Cache persistent pentru rulări deterministe
- **debug/** — Artefacte intermediare pentru analiză/verificare

## Limitări Cunoscute

1. **Non-determinism LLM la prima rulare** — Prima execuție apelează Claude pentru mapping. Deși se folosește `temperature=0`, output-urile LLM nu sunt garantat identice bit-cu-bit între rulări. Odată cache-uite, toate rulările ulterioare sunt complet deterministe.

2. **Grupuri de checkbox-uri** — Checkbox-urile sunt detectate și grupate, dar asocierea lor cu valori JSON depinde de corespondența exactă a etichetelor.

3. **Structuri complexe de tabele** — Sistemul gestionează bine tabelele simple (rând header + rânduri date). Tabelele imbricate sau celulele unite pot să nu se completeze corect.

4. **Cazuri limită de layout** — Valorile foarte lungi care depășesc lățimea placeholder-ului se bazează pe word wrap-ul nativ al Word. Filler-ul nu încearcă să redimensioneze celulele.

5. **Presupunere de limbă** — Prompt-ul LLM și euristicile key_selector sunt calibrate pentru formulare românești de achiziții publice.
