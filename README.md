# Excel Report Generator

## Overview
The `ExcelReportGenerator` class is designed to create standardized Excel reports from Python. It utilizes the `openpyxl` library to generate Excel workbooks with multiple sheets, each containing formatted table versions of Pandas DataFrames.

- This package was originally designed to be tightly integrated with SQL queries. If you want, you can use the `sql-report-template` repo to simply write SQL queries; `ExcelReportGenerator` will handle the report generation for you.
- `ExcelReportGenerator` does not rely on SQL in any way. You may use it as a standalone tool for turning Pandas DataFrames into Excel reports.

## Installation
```bash
# OPTIONAL, but recommended: Make a new virtual environment, if you don't already have one
conda create -n my_new_environment
# Install the package from GitHub
# Change the version number after the "@" symbol to get different versions
pip install git+https://github.com/GSU-Analytics/excel-report-maker@v0.1.1
```

## Usage
### Initialization
Create an instance of the `ExcelReportGenerator` class by providing the results dictionary and introduction text.

```python
from excel_report_generator import ExcelReportGenerator

results = {
    "Sheet1": {
        "First Table Name": df1,
        "Second, Different Table Name": df2
    },
    "Sheet2": {
        "Another table name!": df3
    }
}

intro_text = [
    "This report contains the results of an analysis I conducted.",
    "Each sheet represents a different collection of results."
]

report_generator = ExcelReportGenerator(results, intro_text)
```

### Generating the Workbook
Call the `generate_workbook` method to create the Excel workbook and save it to the specified path.

```python
output_path = "path/to/save/report.xlsx"
report_generator.generate_workbook(output_path)
```

## Class Methods
### `__init__(self, results, intro_text)`
Initializes the Excel report generator.

- `results` (dict): A dictionary where each key is a sheet name and each value is a dictionary mapping table names to DataFrames.
- `intro_text` (list of str): A list of text lines to be added to the introduction sheet.

### `create_introduction_sheet(self)`
Creates the Introduction worksheet and populates it with the intro text.

### `append_df_as_table(self, ws, df, table_title, start_row)`
Appends a Pandas DataFrame to a worksheet as a formatted table.

- `ws` (Worksheet): The worksheet object.
- `df` (DataFrame): The Pandas DataFrame to add.
- `table_title` (str): Title to display above the table.
- `start_row` (int): The starting row number in the worksheet.

Returns:
- `int`: The row number after the inserted table.

### `create_results_sheets(self)`
Creates a separate worksheet for each SQL file (based on the file name) and populates it with all tables derived from that file’s queries.

### `generate_workbook(self, output_path)`
Creates the workbook with the Introduction and all results sheets and saves it.

- `output_path` (str): The file path where the workbook will be saved.

## Example
```python
from excel_report_maker import ExcelReportGenerator
import pandas as pd

# Sample data
df1 = pd.DataFrame({
    "Column1": [1, 2, 3],
    "Column2": [4, 5, 6]
})

df2 = pd.DataFrame({
    "Column1": [7, 8, 9],
    "Column2": [10, 11, 12]
})

results = {
    "Sheet1": {
        "Query1": df1,
        "Query2": df2
    }
}

intro_text = [
    "This report contains the results of various tables.",
    "Each sheet represents a different grouping of tables."
]

# Create the report generator
report_generator = ExcelReportGenerator(results, intro_text)

# Generate the workbook
output_path = "report.xlsx"
report_generator.generate_workbook(output_path)
```
