# PDF Bookmark Tool

A standalone local tool that reads `FILENAME` and `FINAL_TITLE` columns from an
Excel file and writes a one-line table bookmark into every matching PDF in a
local folder.

## For End Users (No Python Required)

1. Unzip the `PDF_Bookmark_Tool` folder that was shared with you.
2. Double-click `PDF_Bookmark_Tool.exe` to launch the app.
3. Click **Browse Excel...** and pick an `.xlsx` file that contains the columns
   `FILENAME` and `FINAL_TITLE`.
4. Click **Browse Folder...** and pick a folder that contains your PDF files.
5. Click **Precheck (no changes)** to validate the match without touching any
   files. The log area will show how many PDFs will be written, unmatched,
   missing, or in conflict.
6. When you are happy with the precheck results, click
   **Run (overwrite PDFs)**. The tool writes one top-level bookmark into each
   matched PDF, **overwriting the original file**.
7. When the run finishes, a `bookmark_result.xlsx` report is generated inside
   the PDF folder. It lists the outcome for every PDF.

## Matching Rule

- The Excel `FILENAME` value (for example `t-14-01-01.rtf`) has `.pdf` appended
  automatically, and must exactly match a real PDF file name in the folder.
- Matching is case-sensitive and does not do any fuzzy normalization.

## First-Version Limits

- Only scans the selected folder itself; subfolders are not recursed into.
- Every matched PDF receives exactly one top-level bookmark that jumps to
  page 1.
- Any existing bookmarks in the PDF are removed; only this new one remains.
- The tool overwrites the original PDF. Internally it writes to a temp file
  first and then atomically replaces the original, so a failure mid-run will
  not corrupt the original file.

## Result Report Columns

| Column | Description |
| --- | --- |
| PDF File | The PDF file name inside the folder |
| Excel FILENAME | The `FILENAME` value from Excel |
| Bookmark Title | The title that was written into the PDF |
| Status | `written` / `failed` / `no_title` / `no_pdf` / `conflict` / `skipped` |
| Error | Filled only when something actually fails; empty for successful rows |

---

## For Developers (Rebuilding the Executable)

Prerequisites: Python 3.10 or newer, `pip`.

1. Install dependencies:

   ```powershell
   cd "C:\Users\Xi.Li\Desktop\pdf bookmark\implementation"
   pip install -r requirements.txt
   ```

2. Run from source (useful while developing):

   ```powershell
   python main.py
   ```

3. Rebuild the standalone Windows executable for distribution:

   Double-click `build.bat`. The script cleans previous builds, installs
   dependencies, and calls PyInstaller. When it finishes, the new build is in
   `dist/PDF_Bookmark_Tool/`. Zip that whole folder and share it.

## Project Layout

```
implementation/
├── main.py                 # Entry point
├── build.bat               # One-click rebuild script
├── requirements.txt        # Python dependencies
├── README.md               # This file
├── PLAN.md                 # Feature plan
├── CODE_PLAN.md            # Code structure plan
├── ui/main_window.py       # Tkinter GUI (front end)
├── core/
│   ├── excel_reader.py     # Read Excel -> FILENAME -> FINAL_TITLE mapping
│   ├── pdf_scanner.py      # Scan a folder for .pdf files
│   ├── matcher.py          # Match Excel rows to PDFs
│   ├── bookmark_writer.py  # Write a bookmark into a single PDF
│   ├── batch_runner.py     # Orchestrator (UI only talks to this module)
│   └── report.py           # Produce the result .xlsx report
└── utils/
    ├── models.py           # Shared dataclasses
    └── logger.py           # Unified logger
```
