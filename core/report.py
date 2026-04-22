"""Produce a result Excel report summarizing each PDF's outcome."""

import os
from typing import List

from openpyxl import Workbook

from utils.logger import get_logger
from utils.models import FileResult

logger = get_logger(__name__)

REPORT_FILENAME = "bookmark_result.xlsx"

HEADERS = [
    "PDF File",
    "Excel FILENAME",
    "Bookmark Title",
    "Status",
    "Error",
]


def generate_report(results: List[FileResult], output_folder: str) -> str:
    """Write an .xlsx report to output_folder. Returns the absolute file path."""
    os.makedirs(output_folder, exist_ok=True)
    output_path = os.path.join(output_folder, REPORT_FILENAME)

    wb = Workbook()
    ws = wb.active
    ws.title = "Bookmark Result"

    ws.append(HEADERS)

    for r in results:
        ws.append(
            [
                r.pdf_name,
                r.excel_filename,
                r.title,
                r.status,
                r.error,
            ]
        )

    widths = [45, 45, 80, 15, 50]
    for idx, width in enumerate(widths, start=1):
        column_letter = ws.cell(row=1, column=idx).column_letter
        ws.column_dimensions[column_letter].width = width

    wb.save(output_path)
    logger.info("Report written to %s", output_path)
    return output_path
