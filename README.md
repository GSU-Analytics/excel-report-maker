# Excel Report Generator

## Overview
The `ExcelReportGenerator` class is designed to create standardized Excel reports from Python. It utilizes the `openpyxl` library to generate Excel workbooks with multiple sheets, each containing formatted table versions of Pandas DataFrames.

**New in v0.2.0**: Image embedding support for matplotlib figures and PNG files.

- This package was originally designed to be tightly integrated with SQL queries. If you want, you can use the `sql-report-template` repo to simply write SQL queries; `ExcelReportGenerator` will handle the report generation for you.
- `ExcelReportGenerator` does not rely on SQL in any way. You may use it as a standalone tool for turning Pandas DataFrames into Excel reports.

## Installation
```bash
# OPTIONAL, but recommended: Make a new virtual environment, if you don't already have one
conda create -n my_new_environment
# Install the package from GitHub
# Change the version number after the "@" symbol to get different versions
pip install git+https://github.com/GSU-Analytics/excel-report-maker@v0.2.0

# For image embedding support, also install matplotlib
pip install git+https://github.com/GSU-Analytics/excel-report-maker@v0.2.0[charts]
```

## Usage
### Basic Usage (Backward Compatible)
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
output_path = "path/to/save/report.xlsx"
report_generator.generate_workbook(output_path)
```

### New: Image Embedding Support
Embed matplotlib charts and PNG images directly into your Excel reports:

```python
import matplotlib.pyplot as plt
import seaborn as sns

def create_charts(df, table_title):
    """Example function to create charts for a given DataFrame."""
    figures = []
    
    # Create a simple bar chart
    fig, ax = plt.subplots(figsize=(10, 6))
    df.plot(kind='bar', ax=ax)
    ax.set_title(f"{table_title} - Bar Chart")
    figures.append(fig)
    
    # Create a histogram if numeric data exists
    numeric_cols = df.select_dtypes(include=['number']).columns
    if len(numeric_cols) > 0:
        fig, ax = plt.subplots(figsize=(10, 6))
        df[numeric_cols[0]].hist(ax=ax, bins=20)
        ax.set_title(f"{table_title} - Distribution")
        figures.append(fig)
    
    return figures

# Define image functions for specific sheets
image_functions = {
    "Sheet1": create_charts,
    "Sheet2": create_charts
}

# Generate workbook with embedded images
report_generator = ExcelReportGenerator(results, intro_text)
report_generator.generate_workbook(output_path, image_functions=image_functions)
```

### Manual Image Addition
You can also manually add images to worksheets:

```python
import matplotlib.pyplot as plt

# Create your workbook as usual
report_generator = ExcelReportGenerator(results, intro_text)
report_generator.create_introduction_sheet()
report_generator.create_results_sheets()

# Get a worksheet and add an image
ws = report_generator.wb["Sheet1"]

# Add a matplotlib figure
fig, ax = plt.subplots()
ax.plot([1, 2, 3, 4], [1, 4, 2, 3])
ax.set_title("Sample Chart")

current_row = report_generator.add_image_from_figure(ws, fig, start_row=10)

# Add a PNG file
current_row = report_generator.add_image_from_file(
    ws, "path/to/chart.png", start_row=current_row, width=500, height=300
)

# Save the workbook
report_generator.wb.save(output_path)
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

### `add_image_from_figure(self, ws, fig, start_row, start_col='A', width=600, height=400)` **NEW**
Embeds a matplotlib figure as an image in the worksheet.

- `ws` (Worksheet): The worksheet object.
- `fig`: Matplotlib figure object.
- `start_row` (int): Row to place the image.
- `start_col` (str or int): Column to place the image (default 'A').
- `width` (int): Image width in pixels (default 600).
- `height` (int): Image height in pixels (default 400).

Returns:
- `int`: Row number after the inserted image.

### `add_image_from_file(self, ws, image_path, start_row, start_col='A', width=600, height=400)` **NEW**
Embeds a PNG file as an image in the worksheet.

- `ws` (Worksheet): The worksheet object.
- `image_path` (str): Path to the PNG file.
- `start_row` (int): Row to place the image.
- `start_col` (str or int): Column to place the image.
- `width` (int): Image width in pixels.
- `height` (int): Image height in pixels.

Returns:
- `int`: Row number after the inserted image.

### `create_results_sheets(self)`
Creates a separate worksheet for each SQL file (based on the file name) and populates it with all tables derived from that file's queries.

### `create_results_sheets_with_images(self, image_functions=None)` **NEW**
Creates results sheets with optional image embedding after each table.

- `image_functions` (dict): Optional dictionary mapping sheet names to image generation functions.

### `generate_workbook(self, output_path, image_functions=None)`
Creates the workbook with the Introduction and all results sheets and saves it.

- `output_path` (str): The file path where the workbook will be saved.
- `image_functions` (dict): Optional dictionary of image generation functions for embedding charts. **NEW**

## Example with Image Embedding
```python
from excel_report_maker import ExcelReportGenerator
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Sample data
df1 = pd.DataFrame({
    "Category": ["A", "B", "C", "D"],
    "Values": [25, 35, 30, 20],
    "Percentage": [0.25, 0.35, 0.30, 0.20]
})

df2 = pd.DataFrame({
    "Month": ["Jan", "Feb", "Mar", "Apr"],
    "Sales": [100, 120, 140, 110]
})

results = {
    "Analysis": {
        "Category Breakdown": df1,
        "Monthly Sales": df2
    }
}

def create_analysis_charts(df, table_title):
    """Create charts for analysis data."""
    figures = []
    
    if "Category" in df.columns and "Values" in df.columns:
        # Bar chart
        fig, ax = plt.subplots(figsize=(8, 5))
        sns.barplot(data=df, x="Category", y="Values", ax=ax)
        ax.set_title(f"{table_title} - Bar Chart")
        figures.append(fig)
        
        # Pie chart
        fig, ax = plt.subplots(figsize=(6, 6))
        ax.pie(df["Values"], labels=df["Category"], autopct='%1.1f%%')
        ax.set_title(f"{table_title} - Distribution")
        figures.append(fig)
    
    elif "Month" in df.columns and "Sales" in df.columns:
        # Line chart
        fig, ax = plt.subplots(figsize=(8, 5))
        sns.lineplot(data=df, x="Month", y="Sales", marker='o', ax=ax)
        ax.set_title(f"{table_title} - Trend")
        figures.append(fig)
    
    return figures

intro_text = [
    "Sales Analysis Report with Embedded Visualizations",
    "This report contains both data tables and corresponding charts.",
    "Charts are embedded directly in the Excel file for easy sharing."
]

# Define which sheets should have charts
image_functions = {
    "Analysis": create_analysis_charts
}

# Create the report generator
report_generator = ExcelReportGenerator(results, intro_text)

# Generate the workbook with embedded images
output_path = "sales_report_with_charts.xlsx"
report_generator.generate_workbook(output_path, image_functions=image_functions)
```

## Backward Compatibility
All existing code will continue to work without modification. The new image features are optional and only activated when:
1. `image_functions` parameter is provided to `generate_workbook()`
2. Manual calls are made to `add_image_from_figure()` or `add_image_from_file()`

The package gracefully handles missing matplotlib dependency by printing warnings and skipping image operations.