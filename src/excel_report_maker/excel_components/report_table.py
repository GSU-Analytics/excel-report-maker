from dataclasses import dataclass, field

import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

from typing import Optional, TYPE_CHECKING
if TYPE_CHECKING:
    from excel_report_maker.excel_reporter.excel_report import ExcelReport
    from excel_report_maker.excel_components.report_sheet import ReportSheet


DEFAULT_TABLE_STYLE = TableStyleInfo(
    name="TableStyleDark9",
    showFirstColumn=False,
    showLastColumn=False,
    showRowStripes=True,
    showColumnStripes=False
)


@dataclass
class ReportTable:
    title: str
    df: pd.DataFrame
    table: Optional[Table] = None
    table_style: TableStyleInfo = DEFAULT_TABLE_STYLE

    @staticmethod
    def from_df(title, df):
        return ReportTable(title, df)

    def set_table_style(self, 
        name="TableStyleDark9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    ):
        '''The name will determine the base theme and style.
        The form should be "TableStyle<Mode><Num>" ...where Mode is
        one of Light, Medium, or Dark, and Num is an integer from 1 to approximately 21.

        To check, open Excel and mouse over the various table styles to see their names.
        '''
        self.table_style = TableStyleInfo(
            name=name,
            showFirstColumn=showFirstColumn,
            showLastColumn=showLastColumn,
            showRowStripes=showRowStripes,
            showColumnStripes=showColumnStripes
        )
        return self
    
    def _as_excel_table(self, displayName: str, ref: str, **table_kwargs):
        table = Table(displayName=displayName, ref=ref, **table_kwargs)
        table.tableStyleInfo = self.table_style
        self.table = table
        return table

    def _append_table_to_sheet(self, report_sheet: "ReportSheet", global_table_counter: int, start_row: int):
        """
        Appends a Pandas DataFrame to a worksheet as a formatted table.

        Args:
            ws (Worksheet): The worksheet object.
            df (DataFrame): The Pandas DataFrame to add.
            table_title (str): Title to display above the table.
            start_row (int): The starting row number in the worksheet.

        Returns:
            int: The row number after the inserted table.
        """
        # Get the Excel worksheet object
        ws: Optional[Worksheet] = report_sheet.ws
        assert ws is not None

        # Write the table title with formatting.
        title_cell = ws.cell(
            row=start_row,
            column=1,
            value=self.title
        )
        title_cell.font = Font(bold=True, size=12)

        # Append the DataFrame rows (header and data).
        start_row += 1
        initial_data_row = start_row
        i = 0
        for i, row in enumerate(dataframe_to_rows(self.df, index=False, header=True), start=0):
            ws.append(row)
        rows_appended = i + 1
        end_row = start_row + rows_appended - 1
        start_col = 1
        end_col = self.df.shape[1]

        # Create an Excel table with a unique name.
        table_ref = f"{get_column_letter(start_col)}{initial_data_row}:{get_column_letter(end_col)}{end_row}"
        excel_table = self._as_excel_table(displayName=f"Table{global_table_counter}", ref=table_ref)
        
        # Add the table
        ws.add_table(excel_table)

        # Format any column whose name contains "Rate" as a percentage.
        self._format_rates(
            ws,
            initial_data_row + 1,
            end_row + 1
        )

        # Return the next available row (with a couple of blank rows added).
        return end_row + 2

    def _format_rates(self, ws: Worksheet, data_top_row, data_bottom_row):
        rate_cols = [col for col in self.df.columns if "rate" in col.lower()]
        for col_idx, col_name in enumerate(self.df.columns, start=1):
            if col_name in rate_cols:
                for row_idx in range(data_top_row, data_bottom_row):
                    ws.cell(row=row_idx, column=col_idx).number_format = '0%'
