from dataclasses import dataclass, field

from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

from excel_report_maker.excel_reporter.excel_report_df import ExcelReportGenerator
from excel_report_maker.excel_components.report_sheet import ReportSheet
from excel_report_maker.excel_components.report_table import ReportTable

from typing import Optional, TYPE_CHECKING
if TYPE_CHECKING:
    import pandas as pd


@dataclass
class ExcelReport(ExcelReportGenerator):
    results: list["ReportSheet"] = field(default_factory=list)
    intro_text_str: Optional[str] = None
    wb: Workbook = field(default_factory=Workbook)

    def __post_init__(self):
        # Global counter for table names to ensure uniqueness across the workbook.
        self.global_table_counter: int = 1
        # Remove the default sheet if it exists.
        if "Sheet" in self.wb.sheetnames:
            del self.wb["Sheet"]
    
    def _create_introduction_sheet(self):
        if self.intro_text_str:
            self.intro_text = self.intro_text_str.splitlines()
            super().create_introduction_sheet()
    
    def _build_report_sheet(self, report_sheet: "ReportSheet"):
        report_sheet._create_report_sheet(self)

    def _build_all_report_sheets(self):
        for report_sheet in self.results:
            self._build_report_sheet(report_sheet)

    @staticmethod
    def from_dict(sheet_table_dict: dict[str, dict[str, "pd.DataFrame"]], intro_text: str = 'Hello world!'):
        results = [
            ReportSheet.from_table_list(sheet_name, [
                ReportTable.from_df(
                    table_name,
                    sheet_table_dict[sheet_name][table_name]
                )
                for table_name in sheet_table_dict[sheet_name]
            ])
            for sheet_name
            in sheet_table_dict
        ]
        return ExcelReport(results=results, intro_text_str=intro_text)

    def register_sheet(self, report_sheet: "ReportSheet"):
        self.results.append(report_sheet)
        return report_sheet

    def generate_workbook(self, output_path):
        self._create_introduction_sheet()
        self._build_all_report_sheets()
        self.wb.save(output_path)