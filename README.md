# Excel Report Generator

<!-- markdown-toc start - Don't edit this section. Run M-x markdown-toc-refresh-toc -->
**Table of Contents**

- [Excel Report Generator](#excel-report-generator)
  - [Quickstart](#quickstart)
  - [Installation](#installation)
  - [Overview](#overview)
    - [SQL Integration](#sql-integration)
    - [Historical Notes for Backwards Compatibility](#historical-notes-for-backwards-compatibility)
  - [User Guide](#user-guide)
    - [Initialization](#initialization)
    - [Making Changes](#making-changes)
    - [Generating the Workbook](#generating-the-workbook)
  - [Manual Workflow](#manual-workflow)
  - [Image Embedding](#image-embedding)
    - [Installation for Image Support](#installation-for-image-support)
    - [Adding Matplotlib Figures](#adding-matplotlib-figures)
    - [Adding PNG Files](#adding-png-files)
    - [Combining Tables and Images](#combining-tables-and-images)
  - [DEPRECATED - ExcelReportGenerator](#deprecated---excelreportgenerator)
    - [Class Methods](#class-methods)
      - [`__init__(self, results, intro_text)`](#__init__self-results-intro_text)
      - [`create_introduction_sheet(self)`](#create_introduction_sheetself)
      - [`append_df_as_table(self, ws, df, table_title, start_row)`](#append_df_as_tableself-ws-df-table_title-start_row)
      - [`create_results_sheets(self)`](#create_results_sheetsself)
      - [`generate_workbook(self, output_path)`](#generate_workbookself-output_path)
    - [Example](#example)

<!-- markdown-toc end -->

## Quickstart
The `ExcelReport` class is designed to create standardized Excel reports from Python. It utilizes the `openpyxl` library to generate Excel workbooks with multiple sheets, each containing formatted table versions of Pandas DataFrames.

```python
import pandas as pd
from excel_report_maker import ExcelReport

dummy_data: pd.DataFrame = pd.read_csv('tests/dummy.csv')

dummy_dict: dict[str, dict[str, pd.DataFrame]] = {
    'My first topic': {
        'Informative Table': dummy_data,
        'Subset Table': dummy_data.iloc[:, 1:3]
    },
    'Deep Dive Topic' :{
        'Focused Table': dummy_data.iloc[1:4, :]
    }
}

report = ExcelReport.from_dict(dummy_dict)

output_path = "path/to/save/report.xlsx"
report.generate_workbook(output_path)
```

## Installation
```bash
# OPTIONAL, but recommended: Make a new virtual environment, if you don't already have one
conda create -n my_new_environment
# Install the package from GitHub
# Change the version number after the "@" symbol to get different versions
pip install git+https://github.com/GSU-Analytics/excel-report-maker@v0.2.0
```

## Overview

### SQL Integration
- This package was originally designed to facilitate writing SQL queries and exporting the results to Excel.
    - This functionality still exists through SQL-Runner, present in [sql-report-template](https://github.com/GSU-Analytics/sql-report-template). `ExcelReport` will handle the report generation for you.
- `ExcelReport` does not rely on SQL in any way. You may use it as a standalone tool for turning Pandas DataFrames into Excel reports.

### Historical Notes for Backwards Compatibility

- There are two classes exposed by this package: `ExcelReport` and `ExcelReportGenerator`.
    - `ExcelReport` is the current working class. It should be used for future projects.
    - `ExcelReportGenerator` is preserved for historical analyses which rely on it. It will not receive further development.

## User Guide

### Initialization
Create an instance of the `ExcelReport` class by providing the results dictionary and introduction text.

```python
import pandas as pd
from excel_report_maker import ExcelReport

dummy_data: pd.DataFrame = pd.read_csv('tests/dummy.csv')

dummy_dict: dict[str, dict[str, pd.DataFrame]] = {
    'My first topic': {
        'Informative Table': dummy_data,
        'Subset Table': dummy_data.iloc[:, 1:3]
    },
    'Deep Dive Topic' :{
        'Focused Table': dummy_data.iloc[1:4, :]
    }
}

report = ExcelReport.from_dict(dummy_dict)
```

### Making Changes

You can access the sheets and individual tables in the report to change their appearance or properties.

```python
from excel_report_maker import ReportSheet, ReportTable

# Sheets are stored in the report as ReportSheets
all_sheets = report.sheets
first_sheet = report.sheets[0]
assert isinstance(first_sheet, ReportSheet)
assert first_sheet.sheet_name == 'My first topic'

# Tables are stored in Sheets as ReportTables
first_table = first_sheet.tables[0]
assert isinstance(first_table, ReportTable)
assert first_table.title.text == 'Informative Table'
```

You can use the `ReportSheet` and `ReportTable` properties and methods to change how the tables and sheets are displayed.

```python
# Setting options for sheets are in the settings property of a ReportSheet
first_sheet.settings
# Setting options for tables are in the settings property of a ReportTable
first_table.settings

# Diable gridlines
first_sheet.settings.gridlines = False
# Change the table theme
first_table.settings.set_table_style(name='TableStyleLight3', showFirstColumn=True)
# Change the title
first_table.settings.title.text = 'Some brand new title'
```

### Generating the Workbook
Call the `.generate_workbook()` method to create the Excel workbook and save it to the specified path.

```python
output_path = "path/to/save/report.xlsx"
report.generate_workbook(output_path)
```

## Manual Workflow

If you want to manually customize each table and sheet as you work, you can use a more imperative design flow:

```python
# Make report ---------------------------------------------------------------------------------
report = ExcelReport()

# Make and register a sheet -------------------------------------------------------------------
first_topic_sheet = ReportSheet("First Topic")
report.register_sheet(first_topic_sheet)

# Do some things to process data, then register the data with the sheet
data1 = dummy_data
rtable1 = ReportTable.from_df('Full Data', data1)
first_topic_sheet.register_table(rtable1)
# Do some more things to process data, then register the data with the sheet
data2 = dummy_data.iloc[1:4, 1:4]
rtable2 = ReportTable.from_df('Subset', data2)
first_topic_sheet.register_table(rtable2)

# Build a second sheet ------------------------------------------------------------------------
second_topic_sheet = ReportSheet('Second Topic')
report.register_sheet(second_topic_sheet)

# Change some settings
second_topic_sheet.settings.gridlines = False
# Register a table
rtable3 = ReportTable.from_df('Custom Table', data2)
rtable3.settings.set_table_style(name='TableStyleLight3', showFirstColumn=True)
# Changing the text styling
# See the openpyxl documentation for more details:
# https://openpyxl.readthedocs.io/en/stable/styles.html
from openpyxl.styles import Font
rtable3.settings.set_text_style('title', 'font', Font(size=36))
second_topic_sheet.register_table(rtable3)

# Build the report ----------------------------------------------------------------------------
report.generate_workbook('tests/demo_workflow.xlsx')
```

## Image Embedding

Add matplotlib figures and PNG images to your Excel reports alongside tables.

### Installation for Image Support

To use image embedding features, install the optional matplotlib dependency:

```bash
pip install git+https://github.com/GSU-Analytics/excel-report-maker@v0.2.0[images]
```

### Adding Matplotlib Figures

Embed matplotlib figures directly into your reports:

```python
import matplotlib.pyplot as plt
from excel_report_maker import ExcelReport, ReportSheet, ReportTable, ReportImage

fig, ax = plt.subplots(figsize=(8, 5))
ax.plot([1, 2, 3, 4], [10, 20, 15, 25], marker='o')
ax.set_title('Sales Trend')
ax.set_xlabel('Quarter')
ax.set_ylabel('Revenue')

sheet = ReportSheet("Sales Analysis")
sheet.register_image(ReportImage.from_figure(fig, title='Quarterly Revenue', width=600, height=400))

report = ExcelReport([sheet])
report.generate_workbook('report_with_chart.xlsx')
```

### Adding PNG Files

Embed PNG images from files:

```python
from excel_report_maker import ReportImage

sheet = ReportSheet("Company Overview")
sheet.register_image(ReportImage.from_file('logo.png', title='Company Logo', width=300, height=200))

report = ExcelReport([sheet])
report.generate_workbook('report_with_logo.xlsx')
```

### Combining Tables and Images

Combine tables and images in a single sheet:

```python
import pandas as pd
import matplotlib.pyplot as plt
from excel_report_maker import ExcelReport, ReportSheet, ReportTable, ReportImage

data = pd.DataFrame({
    'Quarter': ['Q1', 'Q2', 'Q3', 'Q4'],
    'Revenue': [100, 150, 120, 180],
    'Growth_Rate': [0.10, 0.15, 0.12, 0.18]
})

fig, ax = plt.subplots(figsize=(8, 5))
ax.bar(data['Quarter'], data['Revenue'])
ax.set_title('Quarterly Revenue')

sheet = ReportSheet("Financial Report")
sheet.register_table(ReportTable.from_df('Sales Data', data))
sheet.register_image(ReportImage.from_figure(fig, title='Revenue Chart', width=600, height=400))
sheet.register_image(ReportImage.from_file('company_logo.png', width=250, height=150))

report = ExcelReport([sheet])
report.generate_workbook('financial_report.xlsx')
```

Images are rendered after tables in the order they are registered. You can customize image dimensions using the `width` and `height` parameters. The `title` parameter is optional.

## DEPRECATED - ExcelReportGenerator

The notes below are for those who want to use `ExcelReportGenerator`. It is not advised to use it for new projects.

### Class Methods
#### `__init__(self, results, intro_text)`
Initializes the Excel report generator.

- `results` (dict): A dictionary where each key is a sheet name and each value is a dictionary mapping table names to DataFrames.
- `intro_text` (list of str): A list of text lines to be added to the introduction sheet.

#### `create_introduction_sheet(self)`
Creates the Introduction worksheet and populates it with the intro text.

#### `append_df_as_table(self, ws, df, table_title, start_row)`
Appends a Pandas DataFrame to a worksheet as a formatted table.

- `ws` (Worksheet): The worksheet object.
- `df` (DataFrame): The Pandas DataFrame to add.
- `table_title` (str): Title to display above the table.
- `start_row` (int): The starting row number in the worksheet.

Returns:
- `int`: The row number after the inserted table.

#### `create_results_sheets(self)`
Creates a separate worksheet for each SQL file (based on the file name) and populates it with all tables derived from that file’s queries.

#### `generate_workbook(self, output_path)`
Creates the workbook with the Introduction and all results sheets and saves it.

- `output_path` (str): The file path where the workbook will be saved.

### Example
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
