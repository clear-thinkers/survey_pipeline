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
| `02_compile.py` | Merge extracted records into a dataset |
| `03_qa.py` | Flag low-confidence or missing fields |
| `04_analyze.py` | Summarize and aggregate responses |
| `05_report.py` | Generate final report via Claude |

## Setup

1. Copy `.env` and add your Anthropic API key:
   ```
   ANTHROPIC_API_KEY=your_key_here
   ```
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
