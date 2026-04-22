# PDF Bookmark Tool

A standalone local tool that reads `FILENAME` and `FINAL_TITLE` columns from an
Excel file and writes a one-line table bookmark into every matching PDF in a
local folder.

---

## For End Users (No Python Required)

You do not need to install Python or build anything to use this tool. 

1. Download the `PDF_Bookmark_Tool.zip` file provided by the developer and unzip it to your computer.
2. Open the unzipped `PDF_Bookmark_Tool` folder.
   *(Note: You can move this folder anywhere you like, but **never** move the `.exe` file out of the folder by itself, as it relies on the `_internal` folder next to it.)*
3. Double-click `PDF_Bookmark_Tool.exe` to launch the app.
4. Click **Browse Excel...** and pick an `.xlsx` file that contains the columns `FILENAME` and `FINAL_TITLE`.
5. Click **Browse Folder...** and pick a folder that contains your PDF files.
6. Click **Precheck (no changes)** to validate the match without touching any files. The log area will show how many PDFs will be written, unmatched, missing, or in conflict.
7. When you are happy with the precheck results, click **Run (overwrite PDFs)**. The tool writes one top-level bookmark into each matched PDF, **overwriting the original file**.
8. When the run finishes, a `bookmark_result.xlsx` report is generated inside the PDF folder. It lists the outcome for every PDF.

### Matching Rule
- The Excel `FILENAME` value (for example `t-14-01-01.rtf`) has `.pdf` appended automatically, and must exactly match a real PDF file name in the folder.
- Matching is case-sensitive and does not do any fuzzy normalization.

### Result Report Columns
| Column | Description |
| --- | --- |
| PDF File | The PDF file name inside the folder |
| Excel FILENAME | The `FILENAME` value from Excel |
| Bookmark Title | The title that was written into the PDF |
| Status | `written` / `failed` / `no_title` / `no_pdf` / `conflict` / `skipped` |
| Error | Filled only when something actually fails; empty for successful rows |

---

## For Developers

If you want to modify the code or run it directly from the source without packaging:

### Prerequisites
Python 3.10 or newer, and `pip`.

### 1. Install dependencies
Open your terminal/command prompt, navigate to the `implementation` folder, and run:
```powershell
pip install -r requirements.txt
```

### 2. Run from source
To launch the GUI directly using Python during development:
```powershell
python main.py
```

### 3. Rebuild the executable
If you make changes to the code and want to generate a new `.exe`, simply run:
```powershell
.\build.bat
```
The script cleans previous builds, installs dependencies, and calls PyInstaller to regenerate the `dist/PDF_Bookmark_Tool/` folder.

### Project Layout
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