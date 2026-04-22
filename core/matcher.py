"""Match Excel rows with PDF files.

Matching rule (confirmed with user): the Excel FILENAME differs from the real
PDF file name only by the missing ".pdf" suffix, so we append ".pdf" to the
Excel value and do an exact string comparison against the actual filenames.
"""

from typing import Dict, List

from utils.logger import get_logger
from utils.models import MatchedPair, MatchResult, PdfFileInfo

logger = get_logger(__name__)


def _to_pdf_name(excel_filename: str) -> str:
    """Append .pdf if not already present."""
    name = excel_filename.strip()
    if name.lower().endswith(".pdf"):
        return name
    return name + ".pdf"


def match(mapping: Dict[str, str], pdf_files: List[PdfFileInfo]) -> MatchResult:
    """Match Excel mapping against scanned PDF files."""
    result = MatchResult()

    expected_by_pdf_name: Dict[str, str] = {}
    duplicate_excel_keys: List[str] = []
    for excel_filename in mapping.keys():
        pdf_name = _to_pdf_name(excel_filename)
        if pdf_name in expected_by_pdf_name:
            if pdf_name not in duplicate_excel_keys:
                duplicate_excel_keys.append(pdf_name)
            continue
        expected_by_pdf_name[pdf_name] = excel_filename

    pdf_name_to_info: Dict[str, PdfFileInfo] = {}
    duplicate_pdf_names: List[str] = []
    for pdf in pdf_files:
        if pdf.name in pdf_name_to_info:
            if pdf.name not in duplicate_pdf_names:
                duplicate_pdf_names.append(pdf.name)
            continue
        pdf_name_to_info[pdf.name] = pdf

    for pdf_name, excel_filename in expected_by_pdf_name.items():
        pdf_info = pdf_name_to_info.get(pdf_name)
        if pdf_info is None:
            result.missing_pdfs.append(excel_filename)
            continue
        title = mapping[excel_filename]
        result.matched.append(
            MatchedPair(pdf=pdf_info, excel_filename=excel_filename, title=title)
        )

    matched_pdf_names = {pair.pdf.name for pair in result.matched}
    for pdf in pdf_files:
        if pdf.name not in matched_pdf_names and pdf.name not in duplicate_pdf_names:
            if pdf.name not in expected_by_pdf_name:
                result.unmatched_pdfs.append(pdf)

    result.conflicts = duplicate_excel_keys + duplicate_pdf_names

    logger.info(
        "Match done: matched=%d unmatched_pdfs=%d missing_pdfs=%d conflicts=%d",
        len(result.matched),
        len(result.unmatched_pdfs),
        len(result.missing_pdfs),
        len(result.conflicts),
    )
    return result
