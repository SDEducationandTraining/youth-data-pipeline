# Youth Data Cleaning and Merging Pipeline

Cleans, standardizes, validates, combines, and deduplicates Youth Program source datasets into a single master CSV file.

## User Guide

For step-by-step instructions on how to run this pipeline, see the user guide:

[Youth Data Pipeline User Guide](https://docs.google.com/document/d/1RNqfM8Lbh2fI0FLMrmlbzGVyQvp87M9TpoLkBGTPo-s/edit?usp=sharing)

## Overview

This project converts multiple Youth Program source files into one standardized master dataset for analysis and reporting.

The script reads all supported files from a raw input folder, applies a 27-column schema, validates the cleaned data with Pandera, combines all sources, merges duplicates, and exports both a final CSV and an Excel quality report.


## Project Structure

```text
Youth_Data_Cleaning_Merging_Pipeline/
├── data/
│   ├── raw/
│   │   └── .gitkeep
│   └── processed/
│       └── .gitkeep
├── src/
│   └── build_youth_program_master.py
├── requirements.txt
├── .gitignore
└── README.md
```

## Requirements

Python 3.10+

Install dependencies from `requirements.txt`:

```bash
pip install -r requirements.txt
```

## Setup

From the project root:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Input Files

Place raw Youth Program source files in:

```bash
data/raw
```

Supported file types:

```text
.csv
.xlsx
.xls
```

Raw data files are intentionally ignored by Git and should not be committed.

## Usage

Run the script from the project root:

```bash
python src/build_youth_program_master.py --source-dir data/raw --output-dir data/processed --output-prefix youth_program_master_q1_q2_2026
```

To view available arguments:

```bash
python src/build_youth_program_master.py --help
```

## Outputs

The script writes two files to the output processed folder:

```text
data/processed/<output_prefix>.csv
data/processed/<output_prefix>_quality_report.xlsx
```

The final CSV contains the combined master dataset.

The quality report includes:

- source file row counts  
- missing value summary  
- duplicate summary (before and after merging)  
- duplicate records (pre-merge)  
- deduplication summary (rows before/after merge)  
- future date checks  
- unique value counts  
- selected categorical value counts  

Processed outputs are intentionally ignored by Git and should not be committed.

## Key Data Rules

The script:

- uses a shared 27-column master schema  
- standardizes known source column names using a universal column mapping  
- builds `Full Name` from `First Name` and `Last Name` when missing  
- keeps partial names  
- drops rows only when `Full Name` is still missing after cleaning
- identifies and merges duplicate records using Full Name + DOB
- forces `Site = "San Diego"`  
- forces `Youth = 1`  
- converts date, numeric, string, and program indicator columns to standard types  
- validates each cleaned source and the final combined master dataframe with Pandera  

## Deduplication Logic

After combining all source datasets, the script identifies duplicate records using:

```text
Full Name + DOB
```

The script performs two steps:

1. **Duplicate Reporting (Pre-Merge)**
   - Counts duplicate records
   - Outputs all duplicate rows in the quality report

2. **Duplicate Merging**
   - Only records with both `Full Name` and `DOB` are eligible for merging
   - Within each duplicate group:
     - Rows are sorted by completeness, with the most complete rows first
     - Each column keeps the first non-missing value across rows

## Known Limitations

This script merges duplicate records based on Full Name and DOB. Records missing either field are not merged to avoid incorrectly combining different individuals.

Duplicate detection assumes that Full Name + DOB uniquely identifies an individual. If this assumption does not hold, incorrect merges may occur.

This script does not standardize inconsistent category spellings such as gender, nationality, language, or education values.

Invalid dates are converted to missing values during date parsing.

If a future source file uses new column names, update `UNIVERSAL_COLUMN_MAPPING` in:

```text
src/build_youth_program_master.py
```

## Git Notes

The repository should track code and documentation only.

Do commit:

```text
README.md
requirements.txt
.gitignore
src/build_youth_program_master.py
data/raw/.gitkeep
data/processed/.gitkeep
```

Do not commit:

```text
.venv/
data/raw/*
data/processed/*
.csv
.xlsx
.xls
```

## Owner

Author: Aria Shahpari  
Maintained by: IRC San Diego Data Team