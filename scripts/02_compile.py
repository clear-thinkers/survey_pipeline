"""
02_compile.py
Merge all per-survey JSONs from data/extracted/ into type-specific CSVs:
  output/IL/survey_data_IL.csv
  output/412YZ/survey_data_412YZ.csv

If a matching review Excel file exists in the type-specific output folder
(output/<type>/review_{survey_id}.xlsx), any non-blank "Reviewer Correction"
values override the extracted field value before the row is written to the CSV.

Array field corrections are accepted in two formats:
  - Valid JSON array:      ["education", "drivers_license"]
  - Comma-separated text:  education, drivers_license

Usage:
    python scripts/02_compile.py
"""

import json
import sys
from pathlib import Path

import pandas as pd
import openpyxl

sys.path.insert(0, str(Path(__file__).parent.parent))
import config

BASE_DIR = Path(__file__).parent.parent


def get_survey_type(json_path: Path) -> str:
    """Infer survey type from filename when the JSON lacks a survey_type field."""
    import re
    match = re.search(r's(\d+)', json_path.stem, re.IGNORECASE)
    if match:
        return "IL" if int(match.group(1)) <= 11 else "412YZ"
    raise ValueError(f"Cannot determine survey type from filename: {json_path.stem}")


# Fields whose values are lists in the JSON schema (covers both IL and 412YZ)
ARRAY_FIELDS = {
    # IL multi-select fields
    "q6b_job_types",
    "q7_barriers",
    "q8_left_job_reasons",
    "q8a_quit_reasons",
    "q9_bank_account",
    "q9a_no_account_reasons",
    "q11_program_helped",
    # 412YZ multi-select fields
    "q7a_not_registered_reasons",
    "q10_job_barriers",
    "q11_left_job_reasons",
    "q11a_quit_reasons",
    "q14_housing_instability_reasons",
    "q15a_visit_reasons",
    "q15b_visit_barriers",
    "q17_program_helped",
    "q13_sleeping_location",
    "q24_money_methods",
    "q25_bank_account",
    "q26a_account_setup",
    "q26b_account_usage",
    # shared
    "race_ethnicity",
}


# ---------------------------------------------------------------------------
# Reviewer correction parsing
# ---------------------------------------------------------------------------

def parse_correction(raw: str, field: str):
    """
    Parse a reviewer correction cell value.
    - For array fields: accept JSON array or comma-separated plain text.
    - For all other fields: return the stripped string as-is,
      or None if the cell is blank.
    """
    value = str(raw).strip()
    if not value:
        return None

    if field in ARRAY_FIELDS:
        if value.startswith("["):
            try:
                parsed = json.loads(value)
                if isinstance(parsed, list):
                    return parsed
            except json.JSONDecodeError:
                pass
        # Comma-separated plain text fallback
        return [item.strip() for item in value.split(",") if item.strip()]

    # For non-array fields: treat "[]" as a blank/null correction
    if value == "[]":
        return ""

    # If reviewer accidentally entered multiple comma-separated values for a
    # single-select field, take only the first one and warn.
    if "," in value:
        first = value.split(",")[0].strip()
        print(f"  [WARN] {field}: multi-value correction '{value}' for single-select "
              f"field -- using first value '{first}'")
        return first

    return value


def load_review_corrections(survey_id: str, survey_type: str) -> dict:
    """
    Read output/<type>/review_{survey_id}.xlsx and return a dict of
    {field_name: corrected_value} for any row where column E is non-blank.
    Columns: A=Field, B=Extracted Value, C=Confidence, D=Flagged,
             E=Reviewer Correction, F=Notes
    Data starts at row 4 (rows 1–3 are header, summary, blank).
    """
    output_dir = BASE_DIR / config.SURVEY_TYPES[survey_type]["output_dir"]
    xlsx_path = output_dir / f"review_{survey_id}.xlsx"
    if not xlsx_path.exists():
        return {}

    wb = openpyxl.load_workbook(str(xlsx_path), data_only=True)
    ws = wb.active

    corrections = {}
    for row in ws.iter_rows(min_row=4, values_only=True):
        field = row[0]   # col A
        raw   = row[4]   # col E — Reviewer Correction
        if not field or raw is None or str(raw).strip() == "":
            continue
        parsed = parse_correction(str(raw), str(field))
        if parsed is not None:
            corrections[str(field)] = parsed

    return corrections


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    json_files = sorted(config.EXTRACTED_DIR.glob("*.json"))
    if not json_files:
        print(f"No extracted JSON files found in {config.EXTRACTED_DIR}")
        sys.exit(0)

    # Bucket JSONs by survey_type; fall back to filename inference if field is absent
    buckets: dict[str, list] = {t: [] for t in config.SURVEY_TYPES}
    for json_path in json_files:
        data = json.loads(json_path.read_text(encoding="utf-8"))
        survey_type = data.get("survey_type")
        if survey_type not in buckets:
            try:
                survey_type = get_survey_type(json_path)
                print(f"  [INFO] {json_path.name}: no survey_type field, inferred '{survey_type}' from filename.")
            except ValueError:
                print(f"  [WARN] {json_path.name}: cannot determine survey type, skipping.")
                continue
        buckets[survey_type].append((json_path, data))

    for survey_type, entries in buckets.items():
        if not entries:
            print(f"No JSONs found for type '{survey_type}' — skipping.")
            continue

        output_dir = BASE_DIR / config.SURVEY_TYPES[survey_type]["output_dir"]
        output_dir.mkdir(parents=True, exist_ok=True)

        records = []
        for json_path, data in entries:
            survey_id = json_path.stem
            fields = data.get("fields", {})
            confidence = data.get("confidence", {})

            # Apply reviewer corrections if a review workbook exists
            corrections = load_review_corrections(survey_id, survey_type)
            if corrections:
                fields.update(corrections)
                print(f"  [{survey_id}] Applied {len(corrections)} reviewer correction(s): "
                      f"{list(corrections.keys())}")

            # Flatten: array fields → pipe-separated string; None → empty string
            row = {"survey_id": survey_id}
            for field, value in fields.items():
                if isinstance(value, list):
                    row[field] = " | ".join(str(v) for v in value)
                elif value is None:
                    row[field] = ""
                else:
                    row[field] = value

            # Append confidence columns with suffix _conf
            for field, conf in confidence.items():
                row[f"{field}_conf"] = conf

            records.append(row)

        df = pd.DataFrame(records)

        # Put survey_id first, then fields in canonical order, then _conf columns
        conf_cols = sorted(c for c in df.columns if c.endswith("_conf"))
        field_cols = [c for c in df.columns if c != "survey_id" and not c.endswith("_conf")]
        df = df[["survey_id"] + field_cols + conf_cols]

        out_path = output_dir / f"survey_data_{survey_type}.csv"
        df.to_csv(str(out_path), index=False, encoding="utf-8-sig")

        print(f"\nCompiled {len(records)} {survey_type} survey(s) -> {out_path}")
        print(f"Shape: {df.shape[0]} rows x {df.shape[1]} columns")


if __name__ == "__main__":
    main()
