"""Scan a folder and collect metadata for each PDF file."""

import os
from typing import List

from pypdf import PdfReader
from pypdf.errors import PdfReadError

from utils.logger import get_logger
from utils.models import PdfFileInfo, STATUS_PENDING

logger = get_logger(__name__)


def scan_pdf_folder(folder: str) -> List[PdfFileInfo]:
    """Return a list of PdfFileInfo for every .pdf in the given folder (non-recursive)."""
    if not os.path.isdir(folder):
        raise FileNotFoundError(f"Folder does not exist: {folder}")

    files: List[PdfFileInfo] = []
    for name in sorted(os.listdir(folder)):
        if not name.lower().endswith(".pdf"):
            continue

        full_path = os.path.join(folder, name)
        if not os.path.isfile(full_path):
            continue

        info = PdfFileInfo(
            name=name,
            path=full_path,
            status=STATUS_PENDING,
        )

        try:
            info.size = os.path.getsize(full_path)
        except OSError as exc:
            info.error = f"Cannot read file size: {exc}"

        try:
            reader = PdfReader(full_path)
            if reader.is_encrypted:
                info.error = "PDF is encrypted"
            info.pages = len(reader.pages)
        except PdfReadError as exc:
            info.error = f"Invalid PDF: {exc}"
        except Exception as exc:
            info.error = f"Cannot open PDF: {exc}"

        files.append(info)

    logger.info("Scanned %d PDF file(s) in %s", len(files), folder)
    return files
