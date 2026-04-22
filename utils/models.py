"""Shared data structures used across the tool."""

from dataclasses import dataclass, field
from typing import List, Optional


STATUS_PENDING = "pending"
STATUS_MATCHED = "matched"
STATUS_WRITTEN = "written"
STATUS_FAILED = "failed"
STATUS_SKIPPED = "skipped"
STATUS_NO_TITLE = "no_title"
STATUS_NO_PDF = "no_pdf"
STATUS_CONFLICT = "conflict"
STATUS_NOT_SINGLE_PAGE = "not_single_page"


@dataclass
class ExcelRow:
    row_number: int
    filename: str
    title: str


@dataclass
class ExcelMappingResult:
    mapping: dict = field(default_factory=dict)
    duplicate_filenames: List[str] = field(default_factory=list)
    empty_title_rows: List[ExcelRow] = field(default_factory=list)
    missing_columns: List[str] = field(default_factory=list)
    total_rows: int = 0

    @property
    def is_valid(self) -> bool:
        return not self.missing_columns


@dataclass
class PdfFileInfo:
    name: str
    path: str
    size: int = 0
    pages: int = 0
    status: str = STATUS_PENDING
    error: str = ""


@dataclass
class MatchedPair:
    pdf: PdfFileInfo
    excel_filename: str
    title: str


@dataclass
class MatchResult:
    matched: List[MatchedPair] = field(default_factory=list)
    unmatched_pdfs: List[PdfFileInfo] = field(default_factory=list)
    missing_pdfs: List[str] = field(default_factory=list)
    conflicts: List[str] = field(default_factory=list)


@dataclass
class FileResult:
    pdf_name: str
    excel_filename: str
    title: str
    status: str
    error: str = ""


@dataclass
class WriteResult:
    success: bool
    error: str = ""


@dataclass
class BatchReport:
    total_pdfs: int = 0
    total_excel_rows: int = 0
    matched: int = 0
    written: int = 0
    failed: int = 0
    unmatched_pdfs: int = 0
    missing_pdfs: int = 0
    conflicts: int = 0
    non_single_page: int = 0
    empty_titles: int = 0
    results: List[FileResult] = field(default_factory=list)
    report_path: Optional[str] = None
