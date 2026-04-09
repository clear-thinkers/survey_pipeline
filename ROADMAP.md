# Survey OCR Roadmap

This document is the current operating view of the repository: what exists, what is active, and what still needs work. It is not a historical log.

## Current Pipeline

### Shared intake and review

| Step | Script | Status | Purpose |
| --- | --- | --- | --- |
| 01 | `scripts/01_extract.py` | Implemented | Extract survey responses from scanned PDFs into per-survey JSON. |
| 01b | `scripts/01b_review.py` | Implemented | Generate a human review workbook for one extracted survey before compilation. |
| 02 | `scripts/02_compile.py` | Implemented | Compile reviewed survey JSON into a wide CSV. |

### 412YZ pipeline

| Step | Script | Status | Purpose |
| --- | --- | --- | --- |
| 02b | `scripts/02b_ingest_online_412YZ.py` | Implemented | Append online SurveyMonkey responses into the compiled 412YZ dataset. |
| 03 | `scripts/03_qa_412YZ.py` | Implemented | Generate QA flags and reviewer workbook for 412YZ. |
| 03b | `scripts/03b_apply_corrections_412YZ.py` | Implemented | Apply reviewed QA corrections back to the 412YZ CSV. |
| 03c | `scripts/03c_standardize_fields_412YZ.py` | Implemented | Standardize selected 412YZ fields after QA corrections. |
| 04 | `scripts/04_analyze_412YZ.py` | Implemented | Build descriptive summaries and chart-ready outputs for 412YZ. |
| 05 | `scripts/05_report_412YZ.js` | Active | Generate the 412YZ narrative report. |

### IL pipeline

| Step | Script | Status | Purpose |
| --- | --- | --- | --- |
| 02b | `scripts/02b_ingest_online_IL.py` | Implemented | Append online SurveyMonkey responses into the compiled IL dataset. |
| 03 | `scripts/03_qa_IL.py` | Implemented | Generate QA flags and reviewer workbook for IL. |
| 03b | `scripts/03b_apply_corrections_IL.py` | Implemented | Apply reviewed QA corrections back to the IL CSV. |
| 04 | `scripts/04_analyze_IL.py` | Implemented | Build descriptive summaries and chart-ready outputs for IL. |
| 05 | `scripts/05_report_IL.js` | Active | Generate the IL narrative report. |

## Current QA Workflow

### Reviewer actions

Reviewer-facing QA actions are now:

- `clear`: blank out noise, placeholders, or invalid extracted values
- `recode`: replace a value with the normalized value in `corrected_value`
- `accept`: keep the extracted value as-is

`exclude` has been removed from reviewer-facing workflows because it duplicated `clear` without changing downstream behavior. The 412YZ correction-apply script still maps legacy `exclude` entries to `clear` for backward compatibility.

### Demographic handling

The QA scripts now handle demographic normalization directly instead of relying on reviewers to type corrections manually.

Implemented behavior:

- Known demographic variants for gender, race, and sexual orientation are prefilled as `recode`
- Plausible self-describe responses can be prefilled as `accept`
- Placeholder text and obvious OCR noise can be prefilled as `clear`
- Prefilled informational rows that do not need reviewer input are moved off the active QA sheet

This is implemented in both:

- `scripts/03_qa_412YZ.py`
- `scripts/03_qa_IL.py`

## Data and Outputs

### Source data

- `data/raw/`: scanned input files
- `data/extracted/`: extracted per-survey JSON
- `data/online/`: online-response source files when present

### Generated outputs

- `output/412YZ/survey_data_412YZ.csv`: compiled and corrected 412YZ dataset
- `output/412YZ/flagged_412YZ.csv`: machine-generated QA flags for 412YZ
- `output/IL/survey_data_IL.csv`: compiled IL dataset
- `output/IL/flagged_IL.csv`: machine-generated QA flags for IL
- `output/412YZ/charts/`: 412YZ analysis charts
- `report/412YZ/`: generated 412YZ report artifacts

## Current Priorities

### 1. Maintain the IL downstream pipeline

Implemented IL downstream steps now include:

- correction apply after `scripts/03_qa_IL.py`
- analysis workbook and chart generation
- report generation

Remaining work, if needed, is limited to any future IL-specific standardization refinements.

### 2. Keep prompts aligned with QA normalization

The extraction prompts should still aim for clean demographic values, but QA remains the enforcement layer for normalization. If prompt changes are made, they should reduce noise without assuming perfect OCR or perfect model adherence.

### 3. Reduce legacy ambiguity

The 412YZ report path is now `scripts/05_report_412YZ.js`. Documentation and operator flow should treat that as the only maintained 412YZ report generator.

## Operating Notes

- Static type warnings around workbook sheet access in the editor have been false positives; syntax checks and script runs have been the real validation path.
- One-off workbook inspection should be done with an ad hoc snippet when needed rather than a maintained numbered script.
- When rerunning QA, reviewer workbooks are regenerated from the current CSV state, so any manual review process should treat the generated workbook as disposable output rather than a source of truth.
- The authoritative operational docs should stay consistent across this file and `README.md`.
