import pytest

from excel_report_maker import ExcelReport, ReportTable, ReportSheet
import pandas as pd

@pytest.fixture
def dummy_data() -> pd.DataFrame:
    return pd.read_csv('tests/dummy.csv')

def test_excel_report(dummy_data):
    sheet_1 = ReportSheet(
        "Dummy Sheet 1",
        [ReportTable.from_df(f"Dummy {x}", dummy_data) for x in range(3)]
    )
    sheet_2 = ReportSheet(
        "Dummy Sheet 2",
        [ReportTable.from_df(f"Dummy {x}", dummy_data) for x in range(2)]
    )

    sheet_1.gridlines = False
    sheets = [sheet_1, sheet_2]

    ExcelReport(sheets, "Hello, world!")\
        .generate_workbook('tests/demo_report.xlsx')

def test_excel_workflow(dummy_data):
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
    second_topic_sheet.gridlines = False
    # Register a table
    rtable3 = ReportTable.from_df('Custom Table', data2)
    rtable3.set_table_style(name='TableStyleLight3', showFirstColumn=True)
    # Changing the text styling
    # See the openpyxl documentation for more details:
    # https://openpyxl.readthedocs.io/en/stable/styles.html
    from openpyxl.styles import Font
    rtable3.set_text_style('title', 'font', Font(size=36))
    second_topic_sheet.register_table(rtable3)

    # Build the report ----------------------------------------------------------------------------
    report.generate_workbook('tests/demo_workflow.xlsx')

def test_excel_dict(dummy_data):
    dummy_dict = {
        'sheet 1': {
            'Special Table': dummy_data,
            'Subset Table': dummy_data.iloc[:, 1:3]
        },
        'Sheet CRAZY TRAIN' :{
            'Secret table': dummy_data.iloc[1:4, :]
        }
    }

    report = ExcelReport.from_dict(dummy_dict)
    report.generate_workbook('tests/demo_dict.xlsx')