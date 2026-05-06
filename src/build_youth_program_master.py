#!/usr/bin/env python3
"""
Identification
--------------
File name: build_youth_program_master.py
Description: Cleans, standardizes, validates, combines, and exports any number of
    Youth Program source datasets from a raw input folder into one master dataframe
    that follows a shared 27-column schema.
Author name(s): Aria Shahpari
Creation date: 2026-05-04
Last modified date: 2026-05-04
Version number: 2.0.0

Usage
-----
How to run the script:
    python build_youth_program_master.py --source-dir PATH_TO_RAW_FOLDER --output-dir PATH_TO_OUTPUT_FOLDER

Required arguments/parameters:
    --source-dir
        Path to the folder containing the raw Youth Program source files.
        The script will process all supported files in this folder.

Optional flags and defaults:
    --output-dir
        Folder where the cleaned master files and quality report will be saved.
        Default: ../data/processed

    --output-prefix
        File name prefix for output files, without extension.
        Default: youth_program_master_q1_q2_2026

    --file-pattern
        Glob pattern used to select files from --source-dir.
        Default: *
        Examples: *.csv, *.xlsx, youth_*.csv

Example invocations:
    python build_youth_program_master.py \
        --source-dir ../data/raw

    python build_youth_program_master.py \
        --source-dir ../data/raw \
        --output-dir ../data/processed \
        --output-prefix youth_program_master_q1_q2_2026

    python build_youth_program_master.py \
        --source-dir ../data/raw \
        --file-pattern "*.csv"

Dependencies
------------
Required language version:
    Python 3.10+

External libraries or tools needed:
    pandas
    numpy
    openpyxl      # required for reading/writing .xlsx files
    chardet       # required for detecting CSV file encodings
    pandera       # required for schema validation

Installation example:
    pip install pandas numpy openpyxl chardet pandera

Environment variables that must be set:
    None.

Context
-------
Why this script exists / what problem it solves:
    The Youth Program source datasets come from separate spreadsheets with overlapping
    but inconsistent structures. This script lets the user place all raw source files
    into one folder, then automatically loads, cleans, standardizes, validates, combines,
    and exports them as one schema-aligned master dataset for reporting, analysis,
    dashboarding, or later automation work.

Assumptions:
    - Source files are stored in one raw folder provided through --source-dir.
    - Supported input formats are .csv, .xlsx, and .xls.
    - CSV file encodings may vary; the script uses chardet to detect encoding before
      loading each CSV file.
    - Excel files are read with pandas.read_excel().
    - Each source should contain First Name and/or Last Name after universal column
      mapping is applied.
    - Source files may use different names for the same master fields. For example,
      Birth Country and Country of Birth can both map to Nationality.
    - Program participation columns use 1 for participation and blank/missing for
      non-participation or unknown.
    - Blank or whitespace-only cells should be treated as missing values.
    - Missing master columns should be added as empty columns during alignment.

Known limitations / edge cases:
    - This script does not deduplicate records; it reports duplicate counts only.
    - This script does not correct inconsistent category spellings such as gender,
      nationality, language, or education values.
    - Date parsing uses pandas format="mixed" with errors="coerce", so invalid dates
      become missing values.
    - Program columns are converted to nullable Int8. Values other than valid integers
      may become missing during numeric coercion.
    - If a future source uses different column names, update UNIVERSAL_COLUMN_MAPPING.
    - If multiple source columns map to the same master column in a single file, the
      script keeps the first non-missing value across those duplicate mapped columns.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable

import chardet
import pandas as pd
import pandera.pandas as pa
from pandera.errors import SchemaErrors
from pandera.pandas import Column, DataFrameSchema


MASTER_COLUMNS = [
    "Site",
    "Participant Site Identifier",
    "Subject Identifier",
    "Last Name",
    "First Name",
    "Full Name",
    "Primary Language",
    "Nationality",
    "Gender",
    "Alien Number",
    "DOB",
    "Arrival Date",
    "Household Income",
    "Household Size",
    "Highest Education Completed",
    "Immigration Status",
    "Res",
    "WD",
    "S&W",
    "FinCap",
    "MicroE",
    "Immigration",
    "Youth",
    "New Roots",
    "Other",
    "Most Recent TouchPoint",
    "Date Taken",
]

STRING_COLUMNS = [
    "Site",
    "Participant Site Identifier",
    "Subject Identifier",
    "Last Name",
    "First Name",
    "Full Name",
    "Primary Language",
    "Nationality",
    "Gender",
    "Alien Number",
    "Highest Education Completed",
    "Immigration Status",
]

PROGRAM_COLUMNS = [
    "Res",
    "WD",
    "S&W",
    "FinCap",
    "MicroE",
    "Immigration",
    "Youth",
    "New Roots",
    "Other",
]

DATE_COLUMNS = [
    "DOB",
    "Arrival Date",
    "Most Recent TouchPoint",
    "Date Taken",
]

NUMERIC_COLUMNS = [
    "Household Income",
    "Household Size",
]

SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xls"}

UNIVERSAL_COLUMN_MAPPING = {
    # Site and identifiers
    "Site": "Site",
    "Participant Site Identifier": "Participant Site Identifier",
    "Subject Identifier": "Subject Identifier",
    "Subject ID": "Subject Identifier",
    "Alien Number": "Alien Number",
    # Names
    "Last Name": "Last Name",
    "Last": "Last Name",
    "Surname": "Last Name",
    "First Name": "First Name",
    "First": "First Name",
    "Full Name": "Full Name",
    "Name": "Full Name",
    # Demographics
    "Primary Language": "Primary Language",
    "Language": "Primary Language",
    "Preferred Language": "Primary Language",
    "Nationality": "Nationality",
    "Birth Country": "Nationality",
    "Country of Birth": "Nationality",
    "Gender": "Gender",
    "Sex": "Gender",
    "Highest Education Completed": "Highest Education Completed",
    "Highest Education": "Highest Education Completed",
    "Immigration Status": "Immigration Status",
    "Immigration Category": "Immigration Status",
    # Dates
    "DOB": "DOB",
    "Date of Birth": "DOB",
    "Birth Date": "DOB",
    "Arrival Date": "Arrival Date",
    "Date of Arrival": "Arrival Date",
    "DOEntree": "Arrival Date",
    "Date of Entry": "Arrival Date",
    "Most Recent TouchPoint": "Most Recent TouchPoint",
    "Most Recent Touchpoint": "Most Recent TouchPoint",
    "Most Recent Touch Point": "Most Recent TouchPoint",
    "Date Taken": "Date Taken",
    # Household fields
    "Household Income": "Household Income",
    "HH Income": "Household Income",
    "Income": "Household Income",
    "Household Size": "Household Size",
    "HH Size": "Household Size",
    # Program fields
    "Res": "Res",
    "Resettlement": "Res",
    "WD": "WD",
    "Workforce Development": "WD",
    "S&W": "S&W",
    "Safety and Wellness": "S&W",
    "FinCap": "FinCap",
    "Financial Capability": "FinCap",
    "MicroE": "MicroE",
    "Microenterprise": "MicroE",
    "Immigration": "Immigration",
    "Youth": "Youth",
    "New Roots": "New Roots",
    "Other": "Other",
}


# Case-insensitive lookup for universal column mapping.
NORMALIZED_COLUMN_MAPPING = {
    source_column.strip().casefold(): target_column
    for source_column, target_column in UNIVERSAL_COLUMN_MAPPING.items()
}


def build_master_schema() -> DataFrameSchema | None:
    """Create the Pandera schema used to validate the standardized master dataframe."""
    schema_columns = {}

    for column in STRING_COLUMNS:
        schema_columns[column] = Column(pa.String, nullable=True)

    for column in DATE_COLUMNS:
        schema_columns[column] = Column(pa.DateTime, nullable=True)

    for column in NUMERIC_COLUMNS:
        schema_columns[column] = Column(pa.Float, nullable=True)

    for column in PROGRAM_COLUMNS:
        schema_columns[column] = Column(pd.Int8Dtype(), nullable=True)

    return DataFrameSchema(schema_columns, strict=True, coerce=False)


def detect_csv_encoding(path: Path, sample_size: int = 100_000) -> str:
    """Detect a CSV file's encoding using chardet."""
    with path.open("rb") as file:
        raw_sample = file.read(sample_size)

    result = chardet.detect(raw_sample)
    encoding = result.get("encoding")

    if encoding is None:
        return "utf-8"

    return encoding


def read_source_file(path: Path) -> pd.DataFrame:
    """Read a supported source file into a dataframe."""
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    extension = path.suffix.lower()

    if extension == ".csv":
        encoding = detect_csv_encoding(path)
        try:
            return pd.read_csv(path, encoding=encoding)
        except UnicodeDecodeError:
            return pd.read_csv(path, encoding="latin1")

    if extension in {".xlsx", ".xls"}:
        return pd.read_excel(path)

    raise ValueError(f"Unsupported file type: {path.name}")


def get_source_files(source_dir: Path, file_pattern: str) -> list[Path]:
    """Return supported files from the source directory using the supplied glob pattern."""
    if not source_dir.exists():
        raise FileNotFoundError(f"Source directory not found: {source_dir}")

    if not source_dir.is_dir():
        raise NotADirectoryError(f"Source path is not a directory: {source_dir}")

    source_files = [
        path
        for path in sorted(source_dir.glob(file_pattern))
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS
    ]

    if not source_files:
        raise FileNotFoundError(
            f"No supported source files found in {source_dir} using pattern {file_pattern!r}. "
            f"Supported extensions: {sorted(SUPPORTED_EXTENSIONS)}"
        )

    return source_files


def standardize_missing_values(df: pd.DataFrame) -> pd.DataFrame:
    """Replace blank or whitespace-only cells with pandas NA."""
    return df.replace(r"^\s*$", pd.NA, regex=True)


def strip_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """Remove leading/trailing whitespace from column names."""
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    return df


def apply_universal_column_mapping(df: pd.DataFrame) -> pd.DataFrame:
    """Rename known source columns to master column names using a universal mapping."""
    rename_mapping = {}

    for column in df.columns:
        normalized_column = column.strip().casefold()
        if normalized_column in NORMALIZED_COLUMN_MAPPING:
            rename_mapping[column] = NORMALIZED_COLUMN_MAPPING[normalized_column]

    return df.rename(columns=rename_mapping)


def combine_duplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Combine duplicate columns created by column mapping.

    If multiple source columns map to the same master column, this keeps the first
    non-missing value from left to right.
    """
    if not df.columns.duplicated().any():
        return df

    combined = pd.DataFrame(index=df.index)

    for column in pd.Index(df.columns).unique():
        matching_columns = df.loc[:, df.columns == column]

        if matching_columns.shape[1] == 1:
            combined[column] = matching_columns.iloc[:, 0]
        else:
            combined[column] = matching_columns.bfill(axis=1).iloc[:, 0]

    return combined


def build_full_name(df: pd.DataFrame) -> pd.DataFrame:
    """Create Full Name from First Name and Last Name when Full Name is missing."""
    df = df.copy()

    full_name_from_parts = (
        df["First Name"].fillna("").astype(str).str.strip()
        + " "
        + df["Last Name"].fillna("").astype(str).str.strip()
    ).str.strip()

    df["Full Name"] = df["Full Name"].fillna(full_name_from_parts)
    df["Full Name"] = df["Full Name"].replace("", pd.NA)

    return df


def apply_constant_values(df: pd.DataFrame) -> pd.DataFrame:
    """Set constant values for San Diego Youth pipeline."""
    df = df.copy()
    df["Youth"] = 1
    df["Site"] = "San Diego"
    return df


def drop_empty_rows_and_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Drop fully empty rows and fully empty columns."""
    df = df.dropna(axis=0, how="all")
    df = df.dropna(axis=1, how="all")
    return df


def drop_rows_missing_full_name(df: pd.DataFrame, source_name: str) -> pd.DataFrame:
    """Drop rows where Full Name is missing after name fields have been standardized."""
    if "Full Name" not in df.columns:
        raise KeyError(f"{source_name} is missing Full Name after schema alignment.")

    return df.dropna(subset=["Full Name"])


def align_to_master_schema(df: pd.DataFrame) -> pd.DataFrame:
    """Force a dataframe to use the master column order, adding missing columns as NA."""
    return df.reindex(columns=MASTER_COLUMNS)


def convert_column_types(df: pd.DataFrame) -> pd.DataFrame:
    """Convert columns to the standard data types expected by the master schema."""
    df = df.copy()

    for column in STRING_COLUMNS:
        df[column] = df[column].astype("string")

    for column in PROGRAM_COLUMNS:
        df[column] = pd.to_numeric(df[column], errors="coerce").astype(pd.Int8Dtype())

    for column in DATE_COLUMNS:
        df[column] = pd.to_datetime(df[column], errors="coerce", format="mixed")

    for column in NUMERIC_COLUMNS:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    return df


def clean_source(source_df: pd.DataFrame, source_name: str) -> pd.DataFrame:
    """Clean and standardize one Youth Program source dataframe."""
    df = source_df.copy()
    df = strip_column_names(df)
    df = standardize_missing_values(df)
    df = drop_empty_rows_and_columns(df)
    df = apply_universal_column_mapping(df)
    df = combine_duplicate_columns(df)
    df = align_to_master_schema(df)
    df = build_full_name(df)
    df = drop_rows_missing_full_name(df, source_name=source_name)
    df = apply_constant_values(df)
    df = convert_column_types(df)
    return df


def validate_dataframe(
    df: pd.DataFrame, schema: DataFrameSchema | None, label: str
) -> pd.DataFrame:
    """Validate a dataframe with Pandera and raise a readable error if validation fails."""
    if schema is None:
        return df

    try:
        return schema.validate(df, lazy=True)
    except SchemaErrors as error:
        failure_cases = error.failure_cases
        print(f"\nValidation failed for {label}.")
        print(failure_cases.head(50).to_string(index=False))
        raise


def combine_sources(source_dfs: Iterable[pd.DataFrame]) -> pd.DataFrame:
    """Stack cleaned source dataframes into one master dataframe."""
    return pd.concat(list(source_dfs), ignore_index=True)


def build_missing_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Return missing counts and percentages by column."""
    return (
        pd.DataFrame(
            {
                "missing_count": df.isna().sum(),
                "missing_percent": (df.isna().mean() * 100).round(2),
            }
        )
        .sort_values("missing_percent", ascending=True)
        .reset_index(names="column")
    )


def build_future_date_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Return counts of future dates in each date column."""
    today = pd.Timestamp.today()
    rows = []

    for column in DATE_COLUMNS:
        future_count = int((df[column] > today).sum())
        rows.append({"column": column, "future_date_count": future_count})

    return pd.DataFrame(rows)


def build_duplicate_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Return duplicate counts for full rows and Full Name + DOB combinations."""
    return pd.DataFrame(
        [
            {
                "check": "Full row duplicates",
                "count": int(df.duplicated().sum()),
            },
            {
                "check": "Duplicate Full Name + DOB",
                "count": int(df.duplicated(subset=["Full Name", "DOB"]).sum()),
            },
        ]
    )


def build_unique_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Return unique value counts by column."""
    return (
        df.nunique(dropna=True)
        .sort_values(ascending=False)
        .reset_index(name="unique_count")
        .rename(columns={"index": "column"})
    )


def build_source_summary(source_rows: list[dict[str, object]]) -> pd.DataFrame:
    """Return a summary of source files processed by the script."""
    return pd.DataFrame(source_rows)


def write_quality_report(
    df: pd.DataFrame,
    output_path: Path,
    source_rows: list[dict[str, object]],
) -> None:
    """Write a lightweight Excel quality report for the final master dataframe."""
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        build_source_summary(source_rows).to_excel(
            writer, sheet_name="source_summary", index=False
        )
        build_missing_summary(df).to_excel(
            writer, sheet_name="missing_summary", index=False
        )
        build_duplicate_summary(df).to_excel(
            writer, sheet_name="duplicate_summary", index=False
        )
        build_future_date_summary(df).to_excel(
            writer, sheet_name="future_dates", index=False
        )
        build_unique_summary(df).to_excel(
            writer, sheet_name="unique_counts", index=False
        )

        for column in [
            "Gender",
            "Primary Language",
            "Nationality",
            "Highest Education Completed",
        ]:
            value_counts = df[column].value_counts(dropna=False).reset_index()
            value_counts.columns = [column, "count"]
            sheet_name = column[:31]
            value_counts.to_excel(writer, sheet_name=sheet_name, index=False)


def save_outputs(
    df: pd.DataFrame,
    output_dir: Path,
    output_prefix: str,
    source_rows: list[dict[str, object]],
) -> dict[str, Path]:
    """Save the final master dataframe and quality report."""
    output_dir.mkdir(parents=True, exist_ok=True)

    csv_path = output_dir / f"{output_prefix}.csv"
    report_path = output_dir / f"{output_prefix}_quality_report.xlsx"

    df.to_csv(csv_path, index=False)

    write_quality_report(df, report_path, source_rows)

    return {
        "csv": csv_path,
        "quality_report": report_path,
    }


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description="Clean, validate, combine, and export Youth Program source datasets."
    )
    parser.add_argument(
        "--source-dir",
        required=True,
        type=Path,
        help="Folder containing raw Youth Program source files.",
    )
    parser.add_argument(
        "--output-dir",
        default=Path("../data/processed"),
        type=Path,
        help="Folder where output files will be saved. Default: ../data/processed",
    )
    parser.add_argument(
        "--output-prefix",
        default="youth_program_master_q1_q2_2026",
        help="Output file prefix without extension. Default: youth_program_master_q1_q2_2026",
    )
    parser.add_argument(
        "--file-pattern",
        default="*",
        help="Glob pattern used to select files from --source-dir. Default: *",
    )
    return parser.parse_args()


def main() -> None:
    """Run the Youth Program master dataset build process."""
    args = parse_args()

    schema = build_master_schema()

    source_paths = get_source_files(args.source_dir, args.file_pattern)
    if not source_paths:
        raise FileNotFoundError(
            f"No supported source files found in: {args.source_dir}"
        )

    cleaned_sources = []
    source_rows = []

    for source_path in source_paths:
        print(f"Processing source file: {source_path.name}")

        source_df = read_source_file(source_path)
        cleaned_df = clean_source(source_df, source_name=source_path.name)
        cleaned_df = validate_dataframe(cleaned_df, schema, source_path.name)

        source_rows.append(
            {
                "source_file": source_path.name,
                "raw_rows": len(source_df),
                "cleaned_rows": len(cleaned_df),
                "rows_dropped": len(source_df) - len(cleaned_df),
            }
        )

        cleaned_sources.append(cleaned_df)

    master_df = combine_sources(cleaned_sources)
    master_df = validate_dataframe(master_df, schema, "final master dataframe")

    outputs = save_outputs(
        df=master_df,
        output_dir=args.output_dir,
        output_prefix=args.output_prefix,
        source_rows=source_rows,
    )

    print("\nYouth Program master build complete.")
    print(f"Processed source files: {len(source_paths)}")
    print(f"Final shape: {master_df.shape[0]} rows x {master_df.shape[1]} columns")
    print(f"Full row duplicates: {int(master_df.duplicated().sum())}")
    print(
        "Duplicate Full Name + DOB records: "
        f"{int(master_df.duplicated(subset=['Full Name', 'DOB']).sum())}"
    )

    print("\nSaved files:")
    for label, path in outputs.items():
        print(f"- {label}: {path}")


if __name__ == "__main__":
    main()
