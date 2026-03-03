#!/usr/bin/env python3
"""Read an Excel file, trim spaces with pandas, and extract SQL table names."""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd
from sql_metadata import Parser


def trim_spaces(df: pd.DataFrame) -> pd.DataFrame:
    """Trim leading/trailing spaces from column names and string cells."""
    cleaned = df.copy()
    cleaned.columns = [str(col).strip() for col in cleaned.columns]

    object_columns = cleaned.select_dtypes(include=["object", "string"]).columns
    for col in object_columns:
        cleaned[col] = cleaned[col].apply(lambda value: value.strip() if isinstance(value, str) else value)

    return cleaned


def extract_table_names(df: pd.DataFrame, sql_column: str) -> pd.DataFrame:
    """Extract unique SQL table names from each SQL statement in ``sql_column``."""
    if sql_column not in df.columns:
        available_columns = ", ".join(df.columns)
        raise ValueError(
            f"SQL column '{sql_column}' not found. Available columns: {available_columns}"
        )

    table_lists = []
    unique_tables: set[str] = set()

    for value in df[sql_column]:
        if isinstance(value, str) and value.strip():
            tables = Parser(value).tables
        else:
            tables = []
        table_lists.append(tables)
        unique_tables.update(tables)

    result = df.copy()
    result["table_names"] = [", ".join(tables) for tables in table_lists]

    print("Unique tables found:")
    for table in sorted(unique_tables):
        print(f"- {table}")

    return result


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Read an Excel sheet, trim spaces, and extract SQL table names."
    )
    parser.add_argument("excel_file", type=Path, help="Path to input Excel file")
    parser.add_argument(
        "--sheet",
        default=0,
        help="Sheet name or index to read (default: first sheet)",
    )
    parser.add_argument(
        "--sql-column",
        default="sql",
        help="Column containing SQL statements (default: sql)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("processed_output.xlsx"),
        help="Output Excel file path",
    )
    args = parser.parse_args()

    df = pd.read_excel(args.excel_file, sheet_name=args.sheet)
    cleaned = trim_spaces(df)
    result = extract_table_names(cleaned, sql_column=args.sql_column.strip())
    result.to_excel(args.output, index=False)

    print(f"\nProcessed file saved to: {args.output}")


if __name__ == "__main__":
    main()
