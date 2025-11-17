# Data Comparison Tool

This repository contains a Python script for comparing two CSV files and generating an Excel report summarizing the differences.
I am also adding a bit configured R script, just in case. 

## Features

- Reads two CSV files and normalizes column names.
- Removes specified technical columns.
- Identifies rows unique to each file.
- Compares common columns for value differences.
- Generates an Excel report with:
    - Summary of differences
    - Rows only in version 1
    - Rows only in version 2
    - Rows with column-level differences

## Requirements

- Python 3.x
- pandas
- openpyxl

Install dependencies:
```bash
pip install pandas openpyxl
```

## Usage

1. Update the file paths for `v1_path` and `v2_path` in the script.
2. Run the script:
     ```bash
     python compare.py
     ```
3. The Excel report will be saved as `Differences_Report_<filename>.xlsx` in your local environment. Press Shift + Alt + R or right click "Reveal in finder".
4. If you get an error stating you dont have permision to write. Make sure you don't have an excel file with the same name opened. 

## Output

- **Summary**: Overview of differences.
- **Only_in_v1**: Rows present only in the first CSV.
- **Only_in_v2**: Rows present only in the second CSV.
- **Column_Differences**: Rows with differing values in common columns.

## Customization

- Adjust `cols_to_remove` to exclude additional columns.
- Change row limits for performance as needed.
