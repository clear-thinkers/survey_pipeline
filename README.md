# Survey OCR

Pipeline for extracting, compiling, QA-ing, analyzing, and reporting on survey data using Claude.

## Structure

```
data/raw/          # Input survey files (PDFs, images, docs)
data/extracted/    # Extracted JSON/CSV per survey
output/            # Final reports and analysis
scripts/           # Processing pipeline scripts
prompts/           # Claude prompt templates
```

## Pipeline

| Script | Purpose |
|--------|---------|
| `01_extract.py` | OCR + field extraction via Claude |
| `01b_review.py` | Generate per-survey human review workbooks |
| `02_compile.py` | Merge extracted records into survey CSVs |
| `02b_ingest_online_412YZ.py` / `02b_ingest_online_IL.py` | Append SurveyMonkey exports |
| `03_qa_412YZ.py` / `03_qa_IL.py` | Generate QA logs and reviewer workbooks |
| `03b_apply_corrections_412YZ.py` / `03b_apply_corrections_IL.py` | Apply QA workbook corrections |
| `03c_standardize_fields_412YZ.py` | Standardize 412YZ DOB and coach names |
| `04_analyze_412YZ.py` / `04_analyze_IL.py` | Produce the analysis workbook and charts |
| `05_report_412YZ.js` / `05_report_IL.js` | Generate the active report docx |

## Setup

1. Copy `.env` and add your Anthropic API key:
   ```
   ANTHROPIC_API_KEY=your_key_here
   ```
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
