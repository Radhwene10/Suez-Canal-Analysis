"""Suez Canal Shipments ETL (Excel cleaning)

Usage:
  python main.py \
    --input suez_canal_shipments_RAW_before_ETL.xlsx \
    --output suez_canal_shipments_cleaned.xlsx

If you omit args, defaults match the original notebook.
"""

from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd


def clean_dataframe(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()

    # Normalize column names
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(" ", "_", regex=False)
    )

    # Clean text columns
    text_cols = df.select_dtypes(include="object").columns
    for col in text_cols:
        df[col] = (
            df[col]
            .astype(str)
            .str.strip()
            .str.replace(r"[\n\t]", "", regex=True)
            .str.replace(r"\*+|\?+", "", regex=True)
            .replace("nan", np.nan)
        )

    # Parse date-like columns (any column containing 'date')
    date_cols = [c for c in df.columns if "date" in c]
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], errors="coerce")
        # Fill missing dates with the mode if available
        if df[col].notna().any():
            df[col] = df[col].fillna(df[col].mode()[0])

    # Try converting object columns that are numeric strings
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = pd.to_numeric(df[col], errors="ignore")

    # Fill missing values
    num_cols = df.select_dtypes(include=["int64", "float64"]).columns
    cat_cols = df.select_dtypes(include="object").columns

    if len(num_cols) > 0:
        df[num_cols] = df[num_cols].fillna(df[num_cols].median(numeric_only=True))
    if len(cat_cols) > 0:
        df[cat_cols] = df[cat_cols].fillna("Unknown")

    # Clip outliers using IQR
    for col in num_cols:
        Q1 = df[col].quantile(0.25)
        Q3 = df[col].quantile(0.75)
        IQR = Q3 - Q1

        lower = Q1 - 1.5 * IQR
        upper = Q3 + 1.5 * IQR

        df[col] = df[col].clip(lower, upper)

    # Remove rows containing 'Unknown' anywhere
    if len(df) > 0:
        df = df[~df.apply(lambda row: row.astype(str).str.lower().eq("unknown").any(), axis=1)]

    # Drop duplicates
    df = df.drop_duplicates()

    return df


def main() -> None:
    parser = argparse.ArgumentParser(description="Clean Suez Canal shipments Excel dataset")
    parser.add_argument(
        "--input",
        default="suez_canal_shipments_RAW_before_ETL.xlsx",
        help="Path to the raw input Excel file",
    )
    parser.add_argument(
        "--output",
        default="suez_canal_shipments_cleaned.xlsx",
        help="Path to write the cleaned Excel file",
    )
    args = parser.parse_args()

    in_path = Path(args.input)
    out_path = Path(args.output)

    if not in_path.exists():
        raise FileNotFoundError(
            f"Input file not found: {in_path}. Place it next to main.py or pass --input <path>."
        )

    df_raw = pd.read_excel(in_path)
    df_clean = clean_dataframe(df_raw)

    # Use openpyxl engine for .xlsx writing
    df_clean.to_excel(out_path, index=False, engine="openpyxl")

    print(f"Dataset successfully cleaned and saved to {out_path}")


if __name__ == "__main__":
    main()
