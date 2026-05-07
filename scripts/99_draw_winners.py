"""
99_draw_winners.py
Randomly draw annual winners from paper surveys in the compiled 412YZ and IL CSVs.

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
WINNERS_DIR = BASE_DIR / "output" / "winners"


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


def load_pool(label: str, csv_path: Path) -> pd.DataFrame:
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
    online_count = len(df) - len(paper_df)
    print(f"Loaded {len(paper_df)} paper {label} survey(s); excluded {online_count} online/anonymous row(s).")
    return paper_df


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


def main() -> None:
    yz_df = load_pool("412YZ", YZ_CSV_PATH)
    il_df = load_pool("IL", IL_CSV_PATH)

    print(f"Available paper 412YZ surveys: {len(yz_df)}")
    print(f"Available paper IL surveys: {len(il_df)}")

    yz_count = prompt_nonnegative_int("How many 412YZ/YZ winners should be drawn? ")
    il_count = prompt_nonnegative_int("How many IL winners should be drawn? ")

    drawn_at = datetime.now().astimezone()
    draw_year = drawn_at.year

    yz_winners = draw_winners("412YZ", yz_df, yz_count)
    il_winners = draw_winners("IL", il_df, il_count)

    winners = pd.concat([yz_winners, il_winners], ignore_index=True, sort=False).fillna("")
    winners.insert(0, "drawn_at", drawn_at.isoformat(timespec="seconds"))
    winners.insert(0, "draw_year", draw_year)

    WINNERS_DIR.mkdir(parents=True, exist_ok=True)
    out_path = annual_output_path(draw_year, drawn_at)
    winners.to_csv(out_path, index=False, encoding="utf-8-sig")

    print(f"\nSaved {len(winners)} winner(s) -> {out_path}")
    if not winners.empty:
        print("\nDrawn winners:")
        for _, row in winners.iterrows():
            print(f"  {row['survey_pool']} #{row['draw_rank']}: {row['survey_id']}")


if __name__ == "__main__":
    main()
