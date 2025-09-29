"""
Unit tests for existing Excel report generation functionality.

Tests the core ExcelReportGenerator methods to ensure backward compatibility
and proper functioning of table creation, worksheet generation, and workbook
assembly. Verifies that all existing functionality continues to work unchanged.

Example:
    pytest test_generation.py -v
"""

import pytest
import pandas as pd
import tempfile
import shutil
from pathlib import Path
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))
from excel_report_maker import ExcelReportGenerator


@pytest.fixture
def sample_data():
    df1 = pd.DataFrame({
        'Category': ['A', 'B', 'C'],
        'Values': [10, 20, 30],
        'Rate': [0.1, 0.2, 0.3]
    })
    
    df2 = pd.DataFrame({
        'Month': ['Jan', 'Feb', 'Mar'],
        'Sales': [100, 150, 200]
    })
    
    return df1, df2


@pytest.fixture
def results_dict(sample_data):
    df1, df2 = sample_data
    return {
        "Analysis": {
            "Category Data": df1,
            "Sales Data": df2
        },
        "Summary": {
            "Overview": df1
        }
    }


@pytest.fixture
def intro_text():
    return [
        "Test Report",
        "This is a test report.",
        "Generated automatically."
    ]


@pytest.fixture
def temp_dir():
    temp_dir = tempfile.mkdtemp()
    yield temp_dir
    shutil.rmtree(temp_dir)


class TestInitialization:
    
    def test_init_basic(self, results_dict, intro_text):
        generator = ExcelReportGenerator(results_dict, intro_text)
        
        assert generator.results == results_dict
        assert generator.intro_text == intro_text
        assert generator.wb is not None
        assert generator.global_table_counter == 1
        assert "Sheet" not in generator.wb.sheetnames
    
    def test_init_empty_results(self, intro_text):
        generator = ExcelReportGenerator({}, intro_text)
        
        assert generator.results == {}
        assert len(generator.wb.sheetnames) == 0


class TestIntroductionSheet:
    
    def test_create_introduction_sheet(self, results_dict, intro_text):
        generator = ExcelReportGenerator(results_dict, intro_text)
        generator.create_introduction_sheet()
        
        assert "Introduction" in generator.wb.sheetnames
        ws = generator.wb["Introduction"]
        assert ws["A1"].value == "Introduction"
        assert ws["A2"].value == intro_text[0]
        assert ws["A3"].value == intro_text[1]
    
    def test_introduction_sheet_formatting(self, results_dict, intro_text):
        generator = ExcelReportGenerator(results_dict, intro_text)
        generator.create_introduction_sheet()
        
        ws = generator.wb["Introduction"]
        assert ws["A1"].font.bold is True
        assert ws["A1"].font.size == 14
        assert ws.column_dimensions['A'].width == 100


class TestTableAppending:
    
    def test_append_df_as_table_basic(self, results_dict, intro_text, sample_data):
        generator = ExcelReportGenerator(results_dict, intro_text)
        generator.create_introduction_sheet()
        
        ws = generator.wb["Introduction"]
        df1, _ = sample_data
        
        next_row = generator.append_df_as_table(ws, df1, "Test Table", 5)
        
        assert ws.cell(row=5, column=1).value == "Test Table"
        assert ws.cell(row=6, column=1).value == "Category"
        assert ws.cell(row=7, column=1).value == "A"
        assert next_row > 5
    
    def test_append_df_rate_formatting(self, results_dict, intro_text, sample_data):
        generator = ExcelReportGenerator(results_dict, intro_text)
        generator.create_introduction_sheet()
        
        ws = generator.wb["Introduction"]
        df1, _ = sample_data
        
        generator.append_df_as_table(ws, df1, "Test Table", 1)
        
        rate_cell = ws.cell(row=3, column=3)
        assert rate_cell.number_format == '0%'
    
    def test_table_counter_increment(self, results_dict, intro_text, sample_data):
        generator = ExcelReportGenerator(results_dict, intro_text)
        generator.create_introduction_sheet()
        
        ws = generator.wb["Introduction"]
        df1, df2 = sample_data
        
        initial_counter = generator.global_table_counter
        generator.append_df_as_table(ws, df1, "Table 1", 1)
        generator.append_df_as_table(ws, df2, "Table 2", 10)
        
        assert generator.global_table_counter == initial_counter + 2


class TestResultsSheets:
    
    def test_create_results_sheets(self, results_dict, intro_text):
        generator = ExcelReportGenerator(results_dict, intro_text)
        generator.create_results_sheets()
        
        assert "Analysis" in generator.wb.sheetnames
        assert "Summary" in generator.wb.sheetnames
        assert len(generator.wb.sheetnames) == 2
    
    def test_results_sheet_content(self, results_dict, intro_text):
        generator = ExcelReportGenerator(results_dict, intro_text)
        generator.create_results_sheets()
        
        analysis_ws = generator.wb["Analysis"]
        assert analysis_ws.cell(row=1, column=1).value == "Category Data"
        assert analysis_ws.cell(row=2, column=1).value == "Category"
    
    def test_column_width_adjustment(self, results_dict, intro_text):
        generator = ExcelReportGenerator(results_dict, intro_text)
        generator.create_results_sheets()
        
        ws = generator.wb["Analysis"]
        assert ws.column_dimensions['A'].width > 0


class TestWorkbookGeneration:
    
    def test_generate_workbook_basic(self, results_dict, intro_text, temp_dir):
        generator = ExcelReportGenerator(results_dict, intro_text)
        output_path = Path(temp_dir) / "test.xlsx"
        
        generator.generate_workbook(str(output_path))
        
        assert output_path.exists()
        assert "Introduction" in generator.wb.sheetnames
        assert "Analysis" in generator.wb.sheetnames
        assert "Summary" in generator.wb.sheetnames
    
    def test_generate_workbook_without_image_functions(self, results_dict, intro_text, temp_dir):
        generator = ExcelReportGenerator(results_dict, intro_text)
        output_path = Path(temp_dir) / "test.xlsx"
        
        generator.generate_workbook(str(output_path), image_functions=None)
        
        assert output_path.exists()
    
    def test_generate_workbook_empty_image_functions(self, results_dict, intro_text, temp_dir):
        generator = ExcelReportGenerator(results_dict, intro_text)
        output_path = Path(temp_dir) / "test.xlsx"
        
        generator.generate_workbook(str(output_path), image_functions={})
        
        assert output_path.exists()


class TestMethodSignatures:
    
    def test_existing_method_signatures(self, results_dict, intro_text):
        generator = ExcelReportGenerator(results_dict, intro_text)
        
        assert hasattr(generator, 'create_introduction_sheet')
        assert hasattr(generator, 'append_df_as_table')
        assert hasattr(generator, 'create_results_sheets')
        assert hasattr(generator, 'generate_workbook')
    
    def test_generate_workbook_backward_compatibility(self, results_dict, intro_text, temp_dir):
        generator = ExcelReportGenerator(results_dict, intro_text)
        output_path = Path(temp_dir) / "test.xlsx"
        
        generator.generate_workbook(str(output_path))
        
        assert output_path.exists()


class TestEdgeCases:
    
    def test_empty_dataframe(self, intro_text, temp_dir):
        empty_df = pd.DataFrame(columns=['A', 'B'])
        results = {"Empty": {"Empty Table": empty_df}}
        
        generator = ExcelReportGenerator(results, intro_text)
        output_path = Path(temp_dir) / "empty.xlsx"
        
        generator.generate_workbook(str(output_path))
        
        assert output_path.exists()
    
    def test_single_row_dataframe(self, intro_text, temp_dir):
        single_df = pd.DataFrame({'A': [1], 'B': [2]})
        results = {"Single": {"Single Row": single_df}}
        
        generator = ExcelReportGenerator(results, intro_text)
        output_path = Path(temp_dir) / "single.xlsx"
        
        generator.generate_workbook(str(output_path))
        
        assert output_path.exists()
    
    def test_large_dataframe(self, intro_text, temp_dir):
        large_df = pd.DataFrame({
            'A': range(1000),
            'B': range(1000, 2000)
        })
        results = {"Large": {"Large Table": large_df}}
        
        generator = ExcelReportGenerator(results, intro_text)
        output_path = Path(temp_dir) / "large.xlsx"
        
        generator.generate_workbook(str(output_path))
        
        assert output_path.exists()