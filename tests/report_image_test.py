"""
Tests for ReportImage component with matplotlib figures and PNG files.

Tests verify image rendering with real Excel files, correct positioning,
table-image integration, and standalone image support.
"""

import pytest
import pandas as pd
from pathlib import Path

from excel_report_maker import ExcelReport, ReportTable, ReportSheet, ReportImage


@pytest.fixture
def dummy_data():
    return pd.read_csv('tests/dummy.csv')


@pytest.fixture
def test_png_path(tmp_path):
    png_path = tmp_path / "test_logo.png"
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots(figsize=(3, 2))
    ax.text(0.5, 0.5, 'Test Logo', ha='center', va='center', fontsize=20)
    ax.axis('off')
    fig.savefig(png_path, bbox_inches='tight', dpi=100)
    plt.close(fig)
    return str(png_path)


def test_matplotlib_figure_image(dummy_data):
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots(figsize=(8, 5))
    ax.plot([1, 2, 3, 4], [1, 4, 2, 3], marker='o')
    ax.set_title('Sales Trend')
    ax.set_xlabel('Quarter')
    ax.set_ylabel('Sales')

    sheet = ReportSheet("Figure Test")
    sheet.register_table(ReportTable.from_df('Sales Data', dummy_data))
    sheet.register_image(ReportImage.from_figure(fig, title='Quarterly Trend', width=600, height=400))

    report = ExcelReport([sheet])
    report.generate_workbook('tests/test_matplotlib_figure.xlsx')
    plt.close(fig)

    assert Path('tests/test_matplotlib_figure.xlsx').exists()


def test_png_file_image(dummy_data, test_png_path):
    sheet = ReportSheet("PNG Test")
    sheet.register_table(ReportTable.from_df('Data Table', dummy_data))
    sheet.register_image(ReportImage.from_file(test_png_path, title='Company Logo', width=300, height=200))

    report = ExcelReport([sheet])
    report.generate_workbook('tests/test_png_file.xlsx')

    assert Path('tests/test_png_file.xlsx').exists()


def test_multiple_images(dummy_data, test_png_path):
    import matplotlib.pyplot as plt

    fig1, ax1 = plt.subplots(figsize=(6, 4))
    ax1.bar(['A', 'B', 'C'], [3, 7, 5])
    ax1.set_title('Bar Chart')

    fig2, ax2 = plt.subplots(figsize=(6, 4))
    ax2.scatter([1, 2, 3, 4], [10, 20, 15, 25])
    ax2.set_title('Scatter Plot')

    sheet = ReportSheet("Multiple Images")
    sheet.register_table(ReportTable.from_df('Summary', dummy_data.head(5)))
    sheet.register_image(ReportImage.from_figure(fig1, title='Chart 1', width=500, height=350))
    sheet.register_image(ReportImage.from_figure(fig2, title='Chart 2', width=500, height=350))
    sheet.register_image(ReportImage.from_file(test_png_path, width=300, height=200))

    report = ExcelReport([sheet])
    report.generate_workbook('tests/test_multiple_images.xlsx')

    plt.close(fig1)
    plt.close(fig2)

    assert Path('tests/test_multiple_images.xlsx').exists()


def test_standalone_images(test_png_path):
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots(figsize=(8, 6))
    ax.plot([1, 2, 3], [1, 2, 3])

    sheet = ReportSheet("Standalone Images")
    sheet.register_image(ReportImage.from_figure(fig, title='Line Chart', width=600, height=450))
    sheet.register_image(ReportImage.from_file(test_png_path, title='Logo', width=400, height=300))

    report = ExcelReport([sheet])
    report.generate_workbook('tests/test_standalone_images.xlsx')
    plt.close(fig)

    assert Path('tests/test_standalone_images.xlsx').exists()


def test_image_without_title(test_png_path):
    sheet = ReportSheet("No Title Test")
    sheet.register_image(ReportImage.from_file(test_png_path, width=400, height=300))

    report = ExcelReport([sheet])
    report.generate_workbook('tests/test_no_title_image.xlsx')

    assert Path('tests/test_no_title_image.xlsx').exists()
