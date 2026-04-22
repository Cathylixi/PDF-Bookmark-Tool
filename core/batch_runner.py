"""Top-level orchestrator. The UI layer only talks to this module."""

from typing import Callable, Optional

from core.bookmark_writer import write_bookmark
from core.excel_reader import read_excel_mapping
from core.matcher import match
from core.pdf_scanner import scan_pdf_folder
from core.report import generate_report
from utils.logger import get_logger
from utils.models import (
    BatchReport,
    FileResult,
    STATUS_CONFLICT,
    STATUS_FAILED,
    STATUS_NO_PDF,
    STATUS_NO_TITLE,
    STATUS_SKIPPED,
    STATUS_WRITTEN,
)

logger = get_logger(__name__)

ProgressCallback = Optional[Callable[[str], None]]


def _notify(cb: ProgressCallback, message: str) -> None:
    logger.info(message)
    if cb is not None:
        try:
            cb(message)
        except Exception:
            logger.exception("Progress callback raised; ignoring")


def _collect_precheck_results(excel_mapping, match_result) -> BatchReport:
    report = BatchReport()
    report.total_excel_rows = len(excel_mapping.mapping) + len(
        excel_mapping.empty_title_rows
    ) + len(excel_mapping.duplicate_filenames)
    report.total_pdfs = (
        len(match_result.matched)
        + len(match_result.unmatched_pdfs)
        + len(match_result.conflicts)
    )
    report.matched = len(match_result.matched)
    report.missing_pdfs = len(match_result.missing_pdfs)
    report.unmatched_pdfs = len(match_result.unmatched_pdfs)
    report.conflicts = len(match_result.conflicts)
    report.empty_titles = len(excel_mapping.empty_title_rows)

    for pair in match_result.matched:
        if pair.pdf.pages != 1:
            report.non_single_page += 1

    for pair in match_result.matched:
        report.results.append(
            FileResult(
                pdf_name=pair.pdf.name,
                excel_filename=pair.excel_filename,
                title=pair.title,
                status="ready",
                error=pair.pdf.error or "",
            )
        )

    for pdf in match_result.unmatched_pdfs:
        report.results.append(
            FileResult(
                pdf_name=pdf.name,
                excel_filename="",
                title="",
                status=STATUS_NO_TITLE,
                error="PDF has no matching row in Excel",
            )
        )

    for excel_filename in match_result.missing_pdfs:
        report.results.append(
            FileResult(
                pdf_name="",
                excel_filename=excel_filename,
                title=excel_mapping.mapping.get(excel_filename, ""),
                status=STATUS_NO_PDF,
                error="Excel row has no matching PDF file",
            )
        )

    for name in match_result.conflicts:
        report.results.append(
            FileResult(
                pdf_name=name,
                excel_filename=name,
                title="",
                status=STATUS_CONFLICT,
                error="Duplicate filename in Excel or PDF folder",
            )
        )

    for row in excel_mapping.empty_title_rows:
        report.results.append(
            FileResult(
                pdf_name="",
                excel_filename=row.filename,
                title="",
                status=STATUS_NO_TITLE,
                error=f"Empty title at Excel row {row.row_number}",
            )
        )

    return report


def run_precheck(
    excel_path: str,
    folder: str,
    on_progress: ProgressCallback = None,
) -> BatchReport:
    """Read Excel + scan PDFs + match, without modifying any files."""
    _notify(on_progress, f"[Precheck] Reading Excel: {excel_path}")
    excel_mapping = read_excel_mapping(excel_path)

    if not excel_mapping.is_valid:
        _notify(
            on_progress,
            f"[Precheck] Excel missing required columns: {excel_mapping.missing_columns}",
        )
        report = BatchReport()
        report.results.append(
            FileResult(
                pdf_name="",
                excel_filename="",
                title="",
                status=STATUS_FAILED,
                error=(
                    f"Excel missing required columns: "
                    f"{', '.join(excel_mapping.missing_columns)}"
                ),
            )
        )
        return report

    _notify(on_progress, f"[Precheck] Excel rows usable: {len(excel_mapping.mapping)}")

    _notify(on_progress, f"[Precheck] Scanning PDF folder: {folder}")
    pdf_files = scan_pdf_folder(folder)
    _notify(on_progress, f"[Precheck] Found {len(pdf_files)} PDF file(s)")

    match_result = match(excel_mapping.mapping, pdf_files)

    report = _collect_precheck_results(excel_mapping, match_result)

    _notify(
        on_progress,
        f"[Precheck] matched={report.matched}  "
        f"missing_pdfs={report.missing_pdfs}  "
        f"unmatched_pdfs={report.unmatched_pdfs}  "
        f"conflicts={report.conflicts}  "
        f"non_single_page={report.non_single_page}  "
        f"empty_titles={report.empty_titles}",
    )
    return report


def run_full(
    excel_path: str,
    folder: str,
    on_progress: ProgressCallback = None,
) -> BatchReport:
    """Run precheck, then actually write bookmarks into matched PDFs."""
    _notify(on_progress, f"[Run] Reading Excel: {excel_path}")
    excel_mapping = read_excel_mapping(excel_path)

    if not excel_mapping.is_valid:
        msg = f"Excel missing required columns: {', '.join(excel_mapping.missing_columns)}"
        _notify(on_progress, f"[Run] {msg}")
        report = BatchReport()
        report.results.append(
            FileResult(
                pdf_name="",
                excel_filename="",
                title="",
                status=STATUS_FAILED,
                error=msg,
            )
        )
        return report

    _notify(on_progress, f"[Run] Scanning folder: {folder}")
    pdf_files = scan_pdf_folder(folder)

    match_result = match(excel_mapping.mapping, pdf_files)

    report = BatchReport()
    report.total_excel_rows = len(excel_mapping.mapping) + len(
        excel_mapping.empty_title_rows
    )
    report.total_pdfs = len(pdf_files)
    report.matched = len(match_result.matched)
    report.missing_pdfs = len(match_result.missing_pdfs)
    report.unmatched_pdfs = len(match_result.unmatched_pdfs)
    report.conflicts = len(match_result.conflicts)
    report.empty_titles = len(excel_mapping.empty_title_rows)

    total_to_write = len(match_result.matched)
    for idx, pair in enumerate(match_result.matched, start=1):
        prefix = f"[Run] ({idx}/{total_to_write})"

        if pair.pdf.error:
            _notify(on_progress, f"{prefix} SKIP {pair.pdf.name}: {pair.pdf.error}")
            report.failed += 1
            report.results.append(
                FileResult(
                    pdf_name=pair.pdf.name,
                    excel_filename=pair.excel_filename,
                    title=pair.title,
                    status=STATUS_FAILED,
                    error=pair.pdf.error,
                )
            )
            continue

        if pair.pdf.pages != 1:
            report.non_single_page += 1
            _notify(
                on_progress,
                f"{prefix} NOTE {pair.pdf.name} has {pair.pdf.pages} pages; "
                f"bookmark will still target page 1",
            )

        _notify(on_progress, f"{prefix} WRITE {pair.pdf.name} -> {pair.title}")
        result = write_bookmark(pair.pdf.path, pair.title)
        if result.success:
            report.written += 1
            report.results.append(
                FileResult(
                    pdf_name=pair.pdf.name,
                    excel_filename=pair.excel_filename,
                    title=pair.title,
                    status=STATUS_WRITTEN,
                )
            )
        else:
            report.failed += 1
            _notify(
                on_progress,
                f"{prefix} FAIL {pair.pdf.name}: {result.error}",
            )
            report.results.append(
                FileResult(
                    pdf_name=pair.pdf.name,
                    excel_filename=pair.excel_filename,
                    title=pair.title,
                    status=STATUS_FAILED,
                    error=result.error,
                )
            )

    for pdf in match_result.unmatched_pdfs:
        report.results.append(
            FileResult(
                pdf_name=pdf.name,
                excel_filename="",
                title="",
                status=STATUS_SKIPPED,
                error="PDF has no matching row in Excel",
            )
        )

    for excel_filename in match_result.missing_pdfs:
        report.results.append(
            FileResult(
                pdf_name="",
                excel_filename=excel_filename,
                title=excel_mapping.mapping.get(excel_filename, ""),
                status=STATUS_NO_PDF,
                error="Excel row has no matching PDF file",
            )
        )

    for name in match_result.conflicts:
        report.results.append(
            FileResult(
                pdf_name=name,
                excel_filename=name,
                title="",
                status=STATUS_CONFLICT,
                error="Duplicate filename in Excel or PDF folder",
            )
        )

    for row in excel_mapping.empty_title_rows:
        report.results.append(
            FileResult(
                pdf_name="",
                excel_filename=row.filename,
                title="",
                status=STATUS_NO_TITLE,
                error=f"Empty title at Excel row {row.row_number}",
            )
        )

    report_path = generate_report(report.results, folder)
    report.report_path = report_path
    _notify(
        on_progress,
        f"[Run] Done. written={report.written} failed={report.failed} "
        f"non_single_page={report.non_single_page} report={report_path}",
    )
    return report
