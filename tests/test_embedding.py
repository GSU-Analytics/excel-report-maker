"""
Unit tests for image embedding functionality in excel-report-maker.

Tests the new add_image_from_figure() and add_image_from_file() methods, 
as well as the create_results_sheets_with_images() workflow. Verifies
proper handling of matplotlib figures, PNG files, and graceful degradation
when matplotlib is not available.

Example:
    pytest test_embedding.py -v
"""

import pytest
import pandas as pd
import tempfile
import shutil
from pathlib import Path
from unittest.mock import patch, MagicMock
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))
from excel_report_maker import ExcelReportGenerator


@pytest.fixture
def generator():
    results = {"Sheet1": {"Table1": pd.DataFrame({"A": [1, 2], "B": [3, 4]})}}
    intro_text = ["Test report"]
    return ExcelReportGenerator(results, intro_text)


@pytest.fixture
def temp_dir():
    temp_dir = tempfile.mkdtemp()
    yield temp_dir
    shutil.rmtree(temp_dir)


class TestImageFromFigure:
    
    @patch('openpyxl.drawing.image.Image')
    def test_add_image_from_figure_success(self, mock_image_class, generator):
        generator.create_introduction_sheet()
        ws = generator.wb["Introduction"]
        
        mock_fig = MagicMock()
        mock_fig.savefig = MagicMock()
        
        mock_image_instance = MagicMock()
        mock_image_class.return_value = mock_image_instance
        
        result_row = generator.add_image_from_figure(ws, mock_fig, start_row=5)
        
        mock_fig.savefig.assert_called_once()
        mock_image_class.assert_called_once()
        assert result_row > 5
    
    def test_add_image_from_figure_no_matplotlib(self, generator, capsys):
        generator.create_introduction_sheet()
        ws = generator.wb["Introduction"]
        
        mock_fig = MagicMock()
        
        with patch('openpyxl.drawing.image.Image', side_effect=ImportError("No module")):
            result_row = generator.add_image_from_figure(ws, mock_fig, start_row=5)
        
        assert result_row == 6
        captured = capsys.readouterr()
        assert "Warning: Image embedding requires matplotlib" in captured.out
    
    @patch('openpyxl.drawing.image.Image')
    def test_add_image_custom_parameters(self, mock_image_class, generator):
        generator.create_introduction_sheet()
        ws = generator.wb["Introduction"]
        
        mock_fig = MagicMock()
        mock_image_instance = MagicMock()
        mock_image_class.return_value = mock_image_instance
        
        generator.add_image_from_figure(ws, mock_fig, start_row=3, start_col='D', width=800, height=600)
        
        assert mock_image_instance.width == 800
        assert mock_image_instance.height == 600
        assert mock_image_instance.anchor == 'D3'


class TestImageFromFile:
    
    @patch('openpyxl.drawing.image.Image')
    def test_add_image_from_file_success(self, mock_image_class, generator, temp_dir):
        generator.create_introduction_sheet()
        ws = generator.wb["Introduction"]
        
        png_path = Path(temp_dir) / "test.png"
        png_path.write_bytes(b"dummy png content")
        
        mock_image = MagicMock()
        mock_image_class.return_value = mock_image
        
        result_row = generator.add_image_from_file(ws, str(png_path), start_row=3)
        
        mock_image_class.assert_called_once_with(str(png_path))
        assert result_row > 3
    
    def test_add_image_from_file_missing(self, generator, capsys):
        generator.create_introduction_sheet()
        ws = generator.wb["Introduction"]
        
        result_row = generator.add_image_from_file(ws, "nonexistent.png", start_row=3)
        
        assert result_row == 4
        captured = capsys.readouterr()
        assert "Warning: Image file not found" in captured.out
    
    @patch('openpyxl.drawing.image.Image')
    def test_add_image_file_column_integer(self, mock_image_class, generator, temp_dir):
        generator.create_introduction_sheet()
        ws = generator.wb["Introduction"]
        
        png_path = Path(temp_dir) / "test.png"
        png_path.write_bytes(b"dummy")
        
        mock_image = MagicMock()
        mock_image_class.return_value = mock_image
        
        generator.add_image_from_file(ws, str(png_path), start_row=1, start_col=3)
        
        assert mock_image.anchor == 'C1'


class TestImageSheetCreation:
    
    def test_create_results_sheets_with_images(self, generator):
        def mock_image_func(df, table_title):
            mock_fig = MagicMock()
            mock_fig.savefig = MagicMock()
            return [mock_fig]
        
        image_functions = {"Sheet1": mock_image_func}
        
        with patch.object(generator, 'add_image_from_figure', return_value=10):
            generator.create_introduction_sheet()
            generator.create_results_sheets_with_images(image_functions)
            
            assert "Sheet1" in generator.wb.sheetnames
    
    def test_generate_workbook_with_image_functions(self, generator, temp_dir):
        output_path = Path(temp_dir) / "test.xlsx"
        
        def mock_image_func(df, table_title):
            return []
        
        image_functions = {"Sheet1": mock_image_func}
        
        generator.generate_workbook(str(output_path), image_functions=image_functions)
        
        assert output_path.exists()
    
    def test_image_function_single_figure(self, generator):
        def single_figure_func(df, table_title):
            mock_fig = MagicMock()
            return mock_fig
        
        image_functions = {"Sheet1": single_figure_func}
        
        with patch.object(generator, 'add_image_from_figure', return_value=10):
            generator.create_introduction_sheet()
            generator.create_results_sheets_with_images(image_functions)
    
    def test_image_function_exception_handling(self, generator, capsys):
        def failing_func(df, table_title):
            raise ValueError("Test error")
        
        image_functions = {"Sheet1": failing_func}
        
        generator.create_introduction_sheet()
        generator.create_results_sheets_with_images(image_functions)
        
        captured = capsys.readouterr()
        assert "Warning: Failed to add images" in captured.out
    
    def test_image_function_none_figures(self, generator):
        def none_func(df, table_title):
            return [None, MagicMock(), None]
        
        image_functions = {"Sheet1": none_func}
        
        with patch.object(generator, 'add_image_from_figure', return_value=10) as mock_add:
            generator.create_introduction_sheet()
            generator.create_results_sheets_with_images(image_functions)
            
            assert mock_add.call_count == 1