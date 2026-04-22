"""Tkinter main window. The UI only talks to core.batch_runner."""

import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Optional

from core.batch_runner import run_full, run_precheck
from utils.logger import get_logger
from utils.models import BatchReport

logger = get_logger(__name__)


class MainWindow:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("PDF Bookmark Tool")
        self.root.geometry("860x620")
        self.root.minsize(760, 520)

        self.excel_path_var = tk.StringVar()
        self.folder_path_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")

        self._message_queue: "queue.Queue[str]" = queue.Queue()
        self._worker: Optional[threading.Thread] = None

        self._build_layout()
        self.root.after(100, self._drain_queue)

    def _build_layout(self) -> None:
        padding = {"padx": 10, "pady": 6}

        top_frame = ttk.LabelFrame(self.root, text="1. Select Inputs")
        top_frame.pack(fill="x", **padding)

        ttk.Label(top_frame, text="Excel File:").grid(
            row=0, column=0, sticky="w", padx=6, pady=6
        )
        self.excel_entry = ttk.Entry(top_frame, textvariable=self.excel_path_var, width=80)
        self.excel_entry.grid(row=0, column=1, sticky="ew", padx=6, pady=6)
        ttk.Button(top_frame, text="Browse Excel...", command=self._pick_excel).grid(
            row=0, column=2, padx=6, pady=6
        )

        ttk.Label(top_frame, text="PDF Folder:").grid(
            row=1, column=0, sticky="w", padx=6, pady=6
        )
        self.folder_entry = ttk.Entry(top_frame, textvariable=self.folder_path_var, width=80)
        self.folder_entry.grid(row=1, column=1, sticky="ew", padx=6, pady=6)
        ttk.Button(top_frame, text="Browse Folder...", command=self._pick_folder).grid(
            row=1, column=2, padx=6, pady=6
        )

        top_frame.columnconfigure(1, weight=1)

        action_frame = ttk.LabelFrame(self.root, text="2. Actions")
        action_frame.pack(fill="x", **padding)

        self.precheck_btn = ttk.Button(
            action_frame, text="Precheck (no changes)", command=self._on_precheck
        )
        self.precheck_btn.pack(side="left", padx=6, pady=8)

        self.run_btn = ttk.Button(
            action_frame, text="Run (overwrite PDFs)", command=self._on_run
        )
        self.run_btn.pack(side="left", padx=6, pady=8)

        ttk.Button(action_frame, text="Clear Log", command=self._clear_log).pack(
            side="left", padx=6, pady=8
        )

        log_frame = ttk.LabelFrame(self.root, text="3. Log / Progress")
        log_frame.pack(fill="both", expand=True, **padding)

        self.log_text = tk.Text(log_frame, wrap="none", height=18)
        self.log_text.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
        log_scroll_y = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        log_scroll_y.pack(side="right", fill="y", pady=6)
        self.log_text.configure(yscrollcommand=log_scroll_y.set, state="disabled")

        status_bar = ttk.Frame(self.root)
        status_bar.pack(fill="x", side="bottom")
        ttk.Label(status_bar, textvariable=self.status_var, anchor="w").pack(
            fill="x", padx=12, pady=4
        )

    def _pick_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            self.excel_path_var.set(path)

    def _pick_folder(self) -> None:
        path = filedialog.askdirectory(title="Select PDF Folder")
        if path:
            self.folder_path_var.set(path)

    def _validate_inputs(self) -> bool:
        excel_path = self.excel_path_var.get().strip()
        folder_path = self.folder_path_var.get().strip()

        if not excel_path:
            messagebox.showwarning("Notice", "Please select an Excel file first.")
            return False
        if not os.path.isfile(excel_path):
            messagebox.showerror("Error", f"Excel file not found:\n{excel_path}")
            return False
        if not folder_path:
            messagebox.showwarning("Notice", "Please select a PDF folder first.")
            return False
        if not os.path.isdir(folder_path):
            messagebox.showerror("Error", f"Folder does not exist:\n{folder_path}")
            return False
        return True

    def _on_precheck(self) -> None:
        if not self._validate_inputs():
            return
        self._start_worker("precheck")

    def _on_run(self) -> None:
        if not self._validate_inputs():
            return
        confirm = messagebox.askyesno(
            "Confirm",
            "This will write bookmarks into the matched PDFs and overwrite the "
            "original files.\n\nContinue?",
        )
        if not confirm:
            return
        self._start_worker("run")

    def _start_worker(self, mode: str) -> None:
        if self._worker and self._worker.is_alive():
            messagebox.showinfo(
                "Notice", "A task is already running. Please wait for it to finish."
            )
            return

        self._set_busy(True)
        title = "Precheck" if mode == "precheck" else "Run"
        self._log(f"=== {title} ===")
        excel_path = self.excel_path_var.get().strip()
        folder_path = self.folder_path_var.get().strip()

        def worker() -> None:
            try:
                if mode == "precheck":
                    report = run_precheck(
                        excel_path, folder_path, on_progress=self._enqueue
                    )
                else:
                    report = run_full(
                        excel_path, folder_path, on_progress=self._enqueue
                    )
                self._enqueue(_summary_line(report, mode))
                if mode == "run" and report.report_path:
                    self._enqueue(f"Result report: {report.report_path}")
            except Exception as exc:
                logger.exception("Worker failed")
                self._enqueue(f"[ERROR] {exc}")
            finally:
                self._enqueue("__DONE__")

        self._worker = threading.Thread(target=worker, daemon=True)
        self._worker.start()

    def _enqueue(self, message: str) -> None:
        self._message_queue.put(message)

    def _drain_queue(self) -> None:
        try:
            while True:
                msg = self._message_queue.get_nowait()
                if msg == "__DONE__":
                    self._set_busy(False)
                    continue
                self._log(msg)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._drain_queue)

    def _log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.precheck_btn.configure(state=state)
        self.run_btn.configure(state=state)
        self.status_var.set("Working..." if busy else "Ready")


def _summary_line(report: BatchReport, mode: str) -> str:
    if mode == "precheck":
        return (
            f"[Summary] matched={report.matched} missing_pdfs={report.missing_pdfs} "
            f"unmatched_pdfs={report.unmatched_pdfs} conflicts={report.conflicts} "
            f"non_single_page={report.non_single_page} empty_titles={report.empty_titles}"
        )
    return (
        f"[Summary] written={report.written} failed={report.failed} "
        f"non_single_page={report.non_single_page} missing_pdfs={report.missing_pdfs} "
        f"unmatched_pdfs={report.unmatched_pdfs} conflicts={report.conflicts}"
    )


def launch() -> None:
    root = tk.Tk()
    MainWindow(root)
    root.mainloop()
