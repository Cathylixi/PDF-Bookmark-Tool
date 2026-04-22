"""Write a single top-level bookmark into a PDF file, overwriting the original.

Safety strategy:
  1. Produce a new PDF next to the original with a .tmp suffix.
  2. Re-open the temp file to confirm the bookmark is present.
  3. Atomically replace the original using os.replace.
  4. If anything fails, the temp file is removed and the original is untouched.
"""

import os
from typing import Optional

from pypdf import PdfReader, PdfWriter
from pypdf.errors import PdfReadError

from utils.logger import get_logger
from utils.models import WriteResult

logger = get_logger(__name__)


def _cleanup(path: Optional[str]) -> None:
    if path and os.path.exists(path):
        try:
            os.remove(path)
        except OSError:
            pass


def write_bookmark(pdf_path: str, title: str) -> WriteResult:
    """Overwrite pdf_path so it contains exactly one top-level bookmark.

    The bookmark targets page 1 and uses ``title`` as its label. Any existing
    outline entries are dropped.
    """
    if not title or not title.strip():
        return WriteResult(success=False, error="Empty bookmark title")

    if not os.path.isfile(pdf_path):
        return WriteResult(success=False, error="PDF file not found")

    tmp_path = pdf_path + ".bookmark.tmp"

    try:
        reader = PdfReader(pdf_path)
        if reader.is_encrypted:
            return WriteResult(success=False, error="PDF is encrypted")

        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)

        if len(writer.pages) == 0:
            return WriteResult(success=False, error="PDF has no pages")

        writer.add_outline_item(title=title.strip(), page_number=0)

        with open(tmp_path, "wb") as fp:
            writer.write(fp)
    except PdfReadError as exc:
        _cleanup(tmp_path)
        return WriteResult(success=False, error=f"Invalid PDF: {exc}")
    except Exception as exc:
        _cleanup(tmp_path)
        return WriteResult(success=False, error=f"Write failed: {exc}")

    try:
        verify_reader = PdfReader(tmp_path)
        outline = verify_reader.outline
        if not outline:
            _cleanup(tmp_path)
            return WriteResult(success=False, error="Verification failed: no outline written")
    except Exception as exc:
        _cleanup(tmp_path)
        return WriteResult(success=False, error=f"Verification failed: {exc}")

    try:
        os.replace(tmp_path, pdf_path)
    except OSError as exc:
        _cleanup(tmp_path)
        return WriteResult(success=False, error=f"Replace failed: {exc}")

    logger.info("Wrote bookmark to %s: %s", pdf_path, title)
    return WriteResult(success=True)
