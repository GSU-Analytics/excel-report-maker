"""
Image component for Excel reports supporting matplotlib figures and PNG files.

Usage:
    from excel_report_maker import ReportImage

    # From matplotlib figure
    img = ReportImage.from_figure(fig, title='Sales Trend', width=600, height=400)
    sheet.register_image(img)

    # From PNG file
    img = ReportImage.from_file('logo.png', title='Company Logo', width=300, height=200)
    sheet.register_image(img)

Rendering logic converts matplotlib figures to PNG bytes via BytesIO, anchors images
to cells, calculates row consumption based on height, and returns next available row.
"""

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Optional, TYPE_CHECKING

from openpyxl.drawing.image import Image as OpenpyxlImage

if TYPE_CHECKING:
    from excel_report_maker.excel_components.report_sheet import ReportSheet


@dataclass
class ReportImageSettings:
    title: Optional[str] = None
    width: int = 600
    height: int = 400


class ReportImage:
    def __init__(self, source, title: Optional[str] = None, width: int = 600, height: int = 400):
        self.source = source
        self.settings = ReportImageSettings(title=title, width=width, height=height)

    @staticmethod
    def from_figure(figure, title: Optional[str] = None, width: int = 600, height: int = 400):
        return ReportImage(source=('figure', figure), title=title, width=width, height=height)

    @staticmethod
    def from_file(path: str | Path, title: Optional[str] = None, width: int = 600, height: int = 400):
        return ReportImage(source=('file', path), title=title, width=width, height=height)

    def _append_image_to_sheet(self, report_sheet: "ReportSheet", start_row: int):
        ws = report_sheet.ws
        assert ws is not None

        if self.settings.title:
            title_cell = ws.cell(row=start_row, column=1, value=self.settings.title)
            from openpyxl.styles import Font
            title_cell.font = Font(bold=True, size=14)
            start_row += 1

        source_type, source_data = self.source

        if source_type == 'figure':
            img_bytes = BytesIO()
            source_data.savefig(img_bytes, format='png', bbox_inches='tight')
            img_bytes.seek(0)
            img = OpenpyxlImage(img_bytes)
        else:
            img = OpenpyxlImage(source_data)

        img.width = self.settings.width
        img.height = self.settings.height
        img.anchor = f"A{start_row}"
        ws.add_image(img)

        rows_consumed = int(self.settings.height / 15) + 2
        return start_row + rows_consumed
