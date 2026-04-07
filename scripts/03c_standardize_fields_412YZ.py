"""
standardize_fields.py
Pre-analysis standardization for output/412YZ/survey_data_412YZ.csv.

1. dob — normalize all values to MM/DD/YYYY
   - Converts YYYY-MM-DD HH:MM:SS (pandas/Excel datetime strings)
   - Expands two-digit years MM/DD/YY → MM/DD/20YY (all participants are youth)
   - Auto-corrects two known data-entry errors:
       s012: 09/20/2026 → 09/20/2006  (future year, obvious typo)
       s106: 07/12/26   → 07/12/2006  (two-digit year, confirmed by age_range 21-23)

2. coach_name_corrected — new column inserted immediately after coach_name
   - Pre-populated with best-guess canonical staff names
   - Blank where the raw value is too ambiguous to guess reliably
   - Intended for manual reviewer confirmation / correction before analysis

Usage:
    python scripts/03c_standardize_fields_412YZ.py
"""

import re
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).parent.parent
CSV_PATH = BASE_DIR / "output" / "412YZ" / "survey_data_412YZ.csv"


# ---------------------------------------------------------------------------
# DOB standardization
# ---------------------------------------------------------------------------

# Explicit per-survey overrides for values that cannot be resolved by rules
DOB_OVERRIDES = {
    "s012": "09/20/2006",   # raw '09/20/2026' — future year, almost certainly 2006
    "s106": "07/12/2006",   # raw '07/12/26' — '26' cannot mean 2026 for a youth; corrected to 2006
}


def standardize_dob(survey_id: str, raw: str) -> str:
    """Return DOB in MM/DD/YYYY, applying rules and explicit overrides."""
    raw = raw.strip() if isinstance(raw, str) else ""
    if not raw:
        return raw

    sid = survey_id.lower()

    # 1. Explicit override table
    if sid in DOB_OVERRIDES:
        return DOB_OVERRIDES[sid]

    # 2. YYYY-MM-DD ... (pandas/Excel datetime string)
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", raw)
    if m:
        return f"{m.group(2)}/{m.group(3)}/{m.group(1)}"

    # 3. MM/DD/YYYY — already correct
    if re.match(r"^\d{2}/\d{2}/\d{4}$", raw):
        return raw

    # 4. MM/DD/YY — expand to MM/DD/20YY (all participants are youth)
    m = re.match(r"^(\d{2})/(\d{2})/(\d{2})$", raw)
    if m:
        return f"{m.group(1)}/{m.group(2)}/20{m.group(3)}"

    # Fallback — return as-is so nothing is silently lost
    return raw


# ---------------------------------------------------------------------------
# Coach name canonical mapping
# ---------------------------------------------------------------------------

# Maps raw OCR / handwritten variants to a best-guess canonical staff name.
# Names absent from this dict are passed through unchanged (already clean or
# ambiguous — reviewer should verify).
COACH_NAME_MAP = {
    # --- Ariella ---
    "Arillea":              "Ariella",
    "ARiela":               "Ariella",
    "Ari":                  "Ariella",
    "Mrs. Ariella":         "Ariella",

    # --- Tamika ---
    "Tameeka":              "Tamika",
    "Ms. Tamikia":          "Tamika",

    # --- Morgan Stewart ---
    "Morgan Stiest":        "Morgan Stewart",   # OCR noise on last name
    "Morgan":               "Morgan Stewart",   # first-name-only entries

    # --- Megan Monroe Ambrose ---
    "Megan":                "Megan Monroe Ambrose",
    "Megan Mo":             "Megan Monroe Ambrose",
    "Meg M":                "Megan Monroe Ambrose",
    "Meg Monroe Ambrose":   "Megan Monroe Ambrose",

    # --- Will Witt ---
    "Will W":               "Will Witt",

    # --- Will C (different coach from Will Witt) ---
    "Will C.":              "Will C",
    "Mr. Will C":           "Will C",

    # --- Noelle ---
    "Nolle":                "Noelle",

    # --- Kalei ---
    "Kalel":                "Kalei",
    "Mis Lacalei":          "Kalei",

    # --- Cedric ---
    "Cedric rudph":         "Cedric",
    "Mr Cedric":            "Cedric",

    # --- Allicia Brayard ---
    "Allicia":              "Allicia Brayard",
    "Alicia B":             "Allicia Brayard",

    # --- Shawna ---
    "Shawnn":               "Shawna",
    "Ms. Shawna":           "Shawna",

    # --- Breanne ---
    "Bri":                  "Breanne",
    "Brie":                 "Breanne",
    "Brieanna":             "Breanne",

    # --- Nalon (OCR uncertainty on this name) ---
    "Nahon":                "Nalon",

    # --- Ambiguous / garbled — leave blank for reviewer ---
    "D":                    "",   # only an initial; unclear
    "Hcg":                  "",   # garbled OCR
    "Aci":                  "",   # garbled OCR
    "Mcy":                  "",   # likely 'Meg' but too uncertain
    "Will":                 "",   # could be Will Witt or Will C
    "Sharon Shannon/Shaung": "",  # garbled; may be 'Shawna'
}


def suggest_coach(name: str) -> str:
    """Return canonical name if in map, else return the original (already clean)."""
    name = name.strip() if isinstance(name, str) else ""
    if not name:
        return ""
    if name in COACH_NAME_MAP:
        return COACH_NAME_MAP[name]
    return name  # pass through — already looks canonical


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    df = pd.read_csv(CSV_PATH, encoding="utf-8-sig", dtype=str)

    # ---- DOB ----------------------------------------------------------------
    dob_changes: list[str] = []
    new_dobs: list[str] = []

    for _, row in df.iterrows():
        sid = str(row["survey_id"]).strip()
        old = str(row["dob"]).strip() if pd.notna(row["dob"]) else ""
        new = standardize_dob(sid, old)
        new_dobs.append(new)
        if old != new:
            dob_changes.append(f"  [DOB   ] {sid}: '{old}' -> '{new}'")

    df["dob"] = new_dobs

    # ---- coach_name_corrected -----------------------------------------------
    # Always recompute from coach_name so the script is idempotent.
    # Once a reviewer starts editing this column, don't re-run this script.
    coach_col_added = "coach_name_corrected" not in df.columns
    df["coach_name_corrected"] = df["coach_name"].apply(suggest_coach)

    if coach_col_added:
        # Move to immediately after coach_name
        cols = list(df.columns)
        cols.remove("coach_name_corrected")
        insert_pos = cols.index("coach_name") + 1
        cols.insert(insert_pos, "coach_name_corrected")
        df = df[cols]

    # ---- Save ---------------------------------------------------------------
    df.to_csv(CSV_PATH, index=False, encoding="utf-8-sig")

    # ---- Report -------------------------------------------------------------
    print("DOB changes:")
    if dob_changes:
        for msg in dob_changes:
            print(msg)
    else:
        print("  (none)")

    if coach_col_added:
        print(f"\nAdded 'coach_name_corrected' column after 'coach_name'.")
    else:
        print(f"\nRecomputed 'coach_name_corrected' column from current mapping.")

    print(f"\nSaved: {CSV_PATH}")


if __name__ == "__main__":
    main()
