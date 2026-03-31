# Survey Pipeline Pilot — Roadmap
**Auberle 2026 Participant Survey · 10-survey pilot**
_Last updated: March 2026_

---

## Goal
Validate AI-assisted extraction, QA, and report generation against manual workflow. Target: ≥50% time reduction, ≥90% field accuracy on clean scans.

---

## Directory structure

Two survey types are supported: IL (Auberle, s001–s011) and 412YZ (412 Youth Zone, s012+). Survey type is auto-detected from the filename. Each type has its own extraction prompt, output directory, and compiled CSV. All raw PDFs and extracted JSONs share flat directories.

```
survey-pipeline/
├── data/
│   ├── raw/          # Scanned PDFs from phone (s001.pdf … s010.pdf)
│   └── extracted/    # Per-survey JSON outputs from Claude
├── output/
│   ├── review_{id}.xlsx      # Per-survey human review workbook
│   ├── survey_data.csv       # Compiled extraction results
│   ├── flagged_review.csv    # Low-confidence / exception rows
│   └── report.docx           # Final auto-generated report
├── scripts/
│   ├── 01_extract.py         # Claude Vision API → JSON per survey
│   ├── review.py             # Single-survey review workbook generator
│   ├── 02_compile.py         # Merge JSONs → survey_data.csv (applies reviewer corrections)
│   ├── 03_qa.py              # Validation rules → flagged_review.csv
│   ├── 04_analyze.py         # Descriptive stats from clean data
│   └── 05_report.py          # Claude API → report.docx
├── prompts/
│   ├── extraction_prompt.txt # Field schema + extraction instructions
│   └── report_prompt.txt     # Report narrative instructions
├── config.py                 # API key, model, paths, thresholds
├── requirements.txt
└── README.md
```

---

## Phases

### Phase 1 — Extraction (01_extract.py)
- Pass each PDF page as base64 JPEG (150 DPI) to Claude Sonnet 4.6
- Returns structured JSON matching survey field schema
- Includes per-field confidence score (0.0–1.0)
- Accepts optional filename argument: `python scripts/01_extract.py s001`
- Processes all PDFs in `data/raw/` when run without argument

### Phase 1b — Human Review (review.py)
- Run after extraction to spot-check individual surveys before compiling
- Produces `output/review_{id}.xlsx` with one row per field, sorted in question order
- Two-tier confidence flagging based purely on score:
  - Red `#FFCCCC` — `⚠ Must review`: confidence < 75%
  - Yellow `#FFF2CC` — `⚠ Check`: confidence 75–89%
  - Blank: confidence ≥ 90% (clean)
- Rows sorted by canonical question order (cover page → Q1–Q18 → demographics)
- Summary row: `N fields | Must review (<75%): X | Check (75-89%): X | Clean (≥90%): X`
- Blank columns E–F for reviewer corrections and notes
- Usage: `python scripts/review.py s001`

### Phase 2 — Compile (02_compile.py)
- Merges all JSONs from `data/extracted/` into `output/survey_data.csv`
- Applies reviewer corrections from matching `review_{id}.xlsx` workbooks if present
- Accepts corrections as JSON array or comma-separated plain text
- Array fields serialized as pipe-separated values in CSV
- Appends `_conf` columns for every field

### Phase 3 — QA (03_qa.py)
- Validate field types and allowed values
- Flag missing required fields
- Flag low-confidence extractions (score < 0.9)
- Output exceptions to `flagged_review.csv` for human review

### Phase 4 — Analysis (04_analyze.py)
- Frequency distributions for all categorical fields
- Likert scale summaries (Q1, Q15)
- NPS score calculation (Q17)
- Cross-tabs: coach name × satisfaction items

### Phase 5 — Report generation (05_report.py)
- Feed summary stats to Claude Sonnet 4.6
- Auto-populate `report.docx` with narrative + tables

---

## Configuration (config.py)

| Setting | Value |
|---|---|
| `EXTRACTION_MODEL` | `claude-sonnet-4-6` |
| `REPORT_MODEL` | `claude-sonnet-4-6` |
| `CONFIDENCE_THRESHOLD` | `0.9` |
| `MUST_REVIEW_THRESHOLD` | `0.75` (used by review.py) |

---

## Pilot success criteria

| Metric | Target |
|---|---|
| Field extraction accuracy | ≥ 90% on clear checkboxes |
| Exception rate (fields needing human review) | ≤ 15% of total fields |
| End-to-end time for 10 surveys | ≤ 45 min |
| API cost for 10 surveys | ≤ $0.25 |

---

## Progress check (after pilot)
Review `flagged_review.csv` to assess:
- Which question types have highest error rates
- Whether scan quality or handwriting is the primary failure mode
- Prompt adjustments needed before scaling to 100 surveys

---

## Stack
- Python 3.12 (miniconda3) — use `python -m pip install` for packages
- `anthropic` SDK — Sonnet 4.6 for both extraction and report
- `pandas` for data compilation and QA
- `python-docx` for report generation
- `pdf2image` + `Pillow` for PDF → image conversion (requires Poppler on PATH)
- `openpyxl` for review workbook generation
