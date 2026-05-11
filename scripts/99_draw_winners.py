"""
99_draw_winners.py
Randomly draw annual winners from paper surveys and named online responses.

Usage:
    python scripts/99_draw_winners.py
"""

from __future__ import annotations

import random
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).parent.parent
YZ_CSV_PATH = BASE_DIR / "output" / "412YZ" / "survey_data_412YZ.csv"
IL_CSV_PATH = BASE_DIR / "output" / "IL" / "survey_data_IL.csv"
YZ_ONLINE_PATH = BASE_DIR / "data" / "online" / "Youth Zone Survey - Feb 2026.xlsx"
IL_ONLINE_PATH = BASE_DIR / "data" / "online" / "Crawford IL Participant Survey 2026.xlsx"
YZ_HAS_NAME_PATH = BASE_DIR / "data" / "online" / "has_name" / "Youth Zone Survey - Feb 2026.xlsx"
IL_HAS_NAME_PATH = BASE_DIR / "data" / "online" / "has_name" / "Crawford IL Participant Survey 2026.xlsx"
WINNERS_DIR = BASE_DIR / "output" / "winners"


def make_pool_summary(pool_label: str, total_surveys: int, eligible_surveys: int) -> dict[str, object]:
    return {
        "survey_type": pool_label,
        "total_surveys": total_surveys,
        "total_surveys_eligible_for_drawing": eligible_surveys,
    }


def normalize_cell(value: object) -> str:
    if value is None:
        return ""
    return " ".join(str(value).split())


def load_workbook_rows(path: Path, data_start_row: int) -> list[tuple[object, ...]]:
    if not path.exists():
        print(f"Missing required workbook: {path}")
        sys.exit(1)

    workbook = pd.ExcelFile(path, engine="openpyxl")
    sheet_name = workbook.sheet_names[0]
    frame = pd.read_excel(
        workbook,
        sheet_name=sheet_name,
        header=None,
        skiprows=data_start_row - 1,
        dtype=object,
    )
    rows: list[tuple[object, ...]] = []
    for row in frame.itertuples(index=False, name=None):
        if any(normalize_cell(value) for value in row):
            rows.append(row)
    return rows


def load_named_online_lookup(label: str, online_path: Path, has_name_path: Path) -> dict[str, dict[str, str]]:
    online_rows = load_workbook_rows(online_path, data_start_row=3)
    named_rows = load_workbook_rows(has_name_path, data_start_row=8)

    named_by_respondent_id: dict[str, str] = {}
    for row in named_rows:
        respondent_id = normalize_cell(row[0] if len(row) > 0 else None)
        raffle_name = normalize_cell(row[2] if len(row) > 2 else None)
        if respondent_id and raffle_name:
            named_by_respondent_id[respondent_id] = raffle_name

    lookup: dict[str, dict[str, str]] = {}
    for idx, row in enumerate(online_rows, start=1):
        respondent_id = normalize_cell(row[0] if len(row) > 0 else None)
        raffle_name = named_by_respondent_id.get(respondent_id)
        if raffle_name:
            lookup[f"o{idx:03d}"] = {
                "respondent_id": respondent_id,
                "raffle_name": raffle_name,
            }

    if not lookup and named_by_respondent_id:
        print(f"Could not match any named online {label} responses back to compiled survey IDs.")
        sys.exit(1)

    return lookup


def prompt_nonnegative_int(label: str) -> int:
    while True:
        raw = input(label).strip()
        try:
            value = int(raw)
        except ValueError:
            print("Please enter a whole number.")
            continue

        if value < 0:
            print("Please enter 0 or a positive whole number.")
            continue

        return value


def load_pool(label: str, csv_path: Path, online_path: Path, has_name_path: Path) -> tuple[pd.DataFrame, dict[str, object]]:
    if not csv_path.exists():
        print(f"Missing {label} survey CSV: {csv_path}")
        sys.exit(1)

    df = pd.read_csv(csv_path, encoding="utf-8-sig", dtype=str).fillna("")
    if "survey_id" not in df.columns:
        print(f"{label} survey CSV is missing required column: survey_id")
        sys.exit(1)
    if "source" not in df.columns:
        print(f"{label} survey CSV is missing required column: source")
        sys.exit(1)

    paper_df = df[df["source"].str.strip().str.lower() == "paper"].copy().reset_index(drop=True)
    online_df = df[df["source"].str.strip().str.lower() == "online"].copy().reset_index(drop=True)

    paper_df["respondent_id"] = ""
    paper_df["raffle_name"] = ""

    eligible_online_df = online_df.head(0).copy()
    if not online_df.empty:
        named_online_lookup = load_named_online_lookup(label, online_path, has_name_path)
        eligible_online_df = online_df[online_df["survey_id"].isin(named_online_lookup)].copy().reset_index(drop=True)
        eligible_online_df["respondent_id"] = eligible_online_df["survey_id"].map(
            lambda survey_id: named_online_lookup[survey_id]["respondent_id"]
        )
        eligible_online_df["raffle_name"] = eligible_online_df["survey_id"].map(
            lambda survey_id: named_online_lookup[survey_id]["raffle_name"]
        )

    excluded_online_count = len(online_df) - len(eligible_online_df)
    eligible_df = pd.concat([paper_df, eligible_online_df], ignore_index=True, sort=False).fillna("")
    summary = make_pool_summary(label, total_surveys=len(df), eligible_surveys=len(eligible_df))
    print(
        f"Loaded {len(paper_df)} paper and {len(eligible_online_df)} named online {label} survey(s); "
        f"excluded {excluded_online_count} anonymous online row(s)."
    )
    return eligible_df, summary


def draw_winners(pool_label: str, df: pd.DataFrame, count: int) -> pd.DataFrame:
    if count > len(df):
        print(f"Cannot draw {count} {pool_label} winner(s); only {len(df)} survey(s) are available.")
        sys.exit(1)

    if count == 0:
        winners = df.head(0).copy()
    else:
        selected_indexes = random.sample(list(df.index), count)
        winners = df.loc[selected_indexes].copy().reset_index(drop=True)

    winners.insert(0, "draw_rank", range(1, len(winners) + 1))
    winners.insert(0, "survey_pool", pool_label)
    return winners


def annual_output_path(draw_year: int, drawn_at: datetime) -> Path:
    out_path = WINNERS_DIR / f"winners_{draw_year}.csv"
    if not out_path.exists():
        return out_path

    stamp = drawn_at.strftime("%Y%m%d_%H%M%S")
    return WINNERS_DIR / f"winners_{draw_year}_{stamp}.csv"


def write_output_csv(out_path: Path, winners: pd.DataFrame, summary_rows: list[dict[str, object]]) -> None:
    summary_df = pd.DataFrame(summary_rows)
    with out_path.open("w", encoding="utf-8-sig", newline="") as handle:
        winners.to_csv(handle, index=False)
        handle.write("\n")
        summary_df.to_csv(handle, index=False)


def main() -> None:
    yz_df, yz_summary = load_pool("412YZ", YZ_CSV_PATH, YZ_ONLINE_PATH, YZ_HAS_NAME_PATH)
    il_df, il_summary = load_pool("IL", IL_CSV_PATH, IL_ONLINE_PATH, IL_HAS_NAME_PATH)

    print(f"Available eligible 412YZ surveys: {len(yz_df)}")
    print(f"Available eligible IL surveys: {len(il_df)}")

    yz_count = prompt_nonnegative_int("How many 412YZ/YZ winners should be drawn? ")
    il_count = prompt_nonnegative_int("How many IL winners should be drawn? ")

    drawn_at = datetime.now().astimezone()
    draw_year = drawn_at.year

    yz_winners = draw_winners("412YZ", yz_df, yz_count)
    il_winners = draw_winners("IL", il_df, il_count)

    winners = pd.concat([yz_winners, il_winners], ignore_index=True, sort=False).fillna("")
    winners.insert(0, "drawn_at", drawn_at.isoformat(timespec="seconds"))
    winners.insert(0, "draw_year", draw_year)
    summary_rows = [yz_summary, il_summary]

    WINNERS_DIR.mkdir(parents=True, exist_ok=True)
    out_path = annual_output_path(draw_year, drawn_at)
    write_output_csv(out_path, winners, summary_rows)

    print(f"\nSaved {len(winners)} winner(s) -> {out_path}")
    if not winners.empty:
        print("\nDrawn winners:")
        for _, row in winners.iterrows():
            display_name = row.get("raffle_name", "")
            name_suffix = f" ({display_name})" if display_name else ""
            print(f"  {row['survey_pool']} #{row['draw_rank']}: {row['survey_id']}{name_suffix}")


if __name__ == "__main__":
    main()
