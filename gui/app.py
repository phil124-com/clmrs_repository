"""
CLMRS Data Analyzer — desktop GUI (Tkinter, ships with Python so no extra
runtime is needed once packaged with PyInstaller).

Workflow: Select Programme -> Upload 4 files -> Validate -> Run Complete
Analysis -> Review -> Export Client Report.
"""
from __future__ import annotations
import os
import sys
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config.programmes import ProgrammeConfig, butter_default, ensure_builtin_programmes
from engine.pipeline import run_complete_analysis
from engine.excel_export import build_report

APP_TITLE = "CLMRS Data Analyzer"
FILE_SLOTS = ["Farmer List", "Household", "Inspection", "Follow Up"]


class CLMRSApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("880x640")
        self.minsize(760, 560)

        ensure_builtin_programmes()
        self.programme_cfg: ProgrammeConfig = butter_default()
        self.file_paths = {slot: None for slot in FILE_SLOTS}
        self.analysis_result = None

        self._build_ui()
        self._refresh_programme_list()

    # ---------------------------------------------------------------- UI

    def _build_ui(self):
        header = tk.Frame(self, bg="#1F4E5F", height=64)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text=APP_TITLE, bg="#1F4E5F", fg="white",
                 font=("Segoe UI", 16, "bold")).pack(side="left", padx=20, pady=14)

        body = tk.Frame(self, padx=20, pady=16)
        body.pack(fill="both", expand=True)

        # Programme selector
        prog_row = tk.Frame(body)
        prog_row.pack(fill="x", pady=(0, 12))
        tk.Label(prog_row, text="Programme:", font=("Segoe UI", 10, "bold")).pack(side="left")
        self.programme_var = tk.StringVar(value="Butter")
        self.programme_combo = ttk.Combobox(prog_row, textvariable=self.programme_var,
                                             state="readonly", width=30)
        self.programme_combo.pack(side="left", padx=8)
        self.programme_combo.bind("<<ComboboxSelected>>", self._on_programme_change)

        # File upload rows
        files_frame = tk.LabelFrame(body, text="Data Import", padx=12, pady=12, font=("Segoe UI", 10, "bold"))
        files_frame.pack(fill="x", pady=(0, 12))

        self.file_labels = {}
        for slot in FILE_SLOTS:
            row = tk.Frame(files_frame)
            row.pack(fill="x", pady=4)
            tk.Label(row, text=slot, width=14, anchor="w", font=("Segoe UI", 10)).pack(side="left")
            status = tk.Label(row, text="Not loaded", fg="#a33", width=46, anchor="w", font=("Segoe UI", 9))
            status.pack(side="left", padx=8)
            self.file_labels[slot] = status
            tk.Button(row, text="Browse...", command=lambda s=slot: self._browse(s)).pack(side="left")

        # Action buttons
        actions = tk.Frame(body)
        actions.pack(fill="x", pady=(0, 12))
        self.validate_btn = tk.Button(actions, text="Validate", command=self._validate,
                                       width=16, height=2)
        self.validate_btn.pack(side="left", padx=(0, 8))
        self.run_btn = tk.Button(actions, text="RUN COMPLETE ANALYSIS", command=self._run_analysis,
                                  width=24, height=2, bg="#1F4E5F", fg="white",
                                  font=("Segoe UI", 10, "bold"))
        self.run_btn.pack(side="left", padx=(0, 8))
        self.export_btn = tk.Button(actions, text="Export Client Report", command=self._export,
                                     width=20, height=2, state="disabled")
        self.export_btn.pack(side="left")

        # Progress / log
        log_frame = tk.LabelFrame(body, text="Progress", padx=8, pady=8, font=("Segoe UI", 10, "bold"))
        log_frame.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_frame, height=10, state="disabled", font=("Consolas", 9))
        self.log_text.pack(fill="both", expand=True)

        # Executive summary preview
        summary_frame = tk.LabelFrame(body, text="Executive Summary", padx=8, pady=8, font=("Segoe UI", 10, "bold"))
        summary_frame.pack(fill="both", expand=True, pady=(12, 0))
        self.summary_text = tk.Text(summary_frame, height=8, state="disabled", font=("Consolas", 9))
        self.summary_text.pack(fill="both", expand=True)

    def _refresh_programme_list(self):
        names = ProgrammeConfig.list_saved() or ["Butter"]
        self.programme_combo["values"] = names
        if "Butter" in names:
            self.programme_var.set("Butter")

    def _on_programme_change(self, event=None):
        cfg = ProgrammeConfig.load(self.programme_var.get())
        if cfg:
            self.programme_cfg = cfg
            self._log(f"Loaded programme configuration: {cfg.name}")

    # ---------------------------------------------------------------- Actions

    def _log(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self.update_idletasks()

    def _browse(self, slot):
        path = filedialog.askopenfilename(
            title=f"Select {slot} file",
            filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")],
        )
        if path:
            self.file_paths[slot] = path
            self.file_labels[slot].configure(text=f"\u2713 {os.path.basename(path)}", fg="#080")

    def _validate(self):
        missing = [s for s, p in self.file_paths.items() if not p]
        if missing:
            messagebox.showwarning(APP_TITLE, "Missing files:\n" + "\n".join(missing))
            return
        for slot, p in self.file_paths.items():
            if not os.path.exists(p):
                messagebox.showerror(APP_TITLE, f"{slot} file no longer exists:\n{p}")
                return
        messagebox.showinfo(APP_TITLE, "All four files are present. Ready to run analysis.")

    def _run_analysis(self):
        missing = [s for s, p in self.file_paths.items() if not p]
        if missing:
            messagebox.showwarning(APP_TITLE, "Please upload all four files first:\n" + "\n".join(missing))
            return

        self.run_btn.configure(state="disabled")
        self.export_btn.configure(state="disabled")
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        def worker():
            try:
                result = run_complete_analysis(
                    self.programme_cfg,
                    farmer_list_path=self.file_paths["Farmer List"],
                    household_path=self.file_paths["Household"],
                    inspection_path=self.file_paths["Inspection"],
                    follow_up_path=self.file_paths["Follow Up"],
                    progress_cb=lambda msg: self.after(0, self._log, msg),
                )
                self.after(0, self._on_analysis_done, result)
            except Exception as e:
                tb = traceback.format_exc()
                self.after(0, self._on_analysis_error, str(e), tb)

        threading.Thread(target=worker, daemon=True).start()

    def _on_analysis_done(self, result):
        self.analysis_result = result
        for w in result.warnings:
            self._log(f"WARNING: {w}")
        self.summary_text.configure(state="normal")
        self.summary_text.delete("1.0", "end")
        for k, v in result.exec_summary.items():
            self.summary_text.insert("end", f"{k}: {v}\n")
        self.summary_text.configure(state="disabled")
        self.run_btn.configure(state="normal")
        self.export_btn.configure(state="normal")
        messagebox.showinfo(APP_TITLE, "Analysis complete. You can now export the client report.")

    def _on_analysis_error(self, msg, tb):
        self.run_btn.configure(state="normal")
        self._log(f"ERROR: {msg}")
        messagebox.showerror(APP_TITLE, f"Analysis failed:\n{msg}\n\nSee log for details.")
        print(tb, file=sys.stderr)

    def _export(self):
        if not self.analysis_result:
            return
        default_name = f"{self.programme_cfg.name}_CLMRS_Client_Report.xlsx"
        path = filedialog.asksaveasfilename(
            title="Save Client Report", defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not path:
            return
        try:
            build_report(self.analysis_result.sheets, self.analysis_result.exec_summary,
                          self.analysis_result.warnings, self.programme_cfg.name,
                          self.analysis_result.meta, path)
            self._log(f"Client report saved: {path}")
            if messagebox.askyesno(APP_TITLE, f"Report saved to:\n{path}\n\nOpen it now?"):
                os.startfile(path) if os.name == "nt" else None
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Export failed:\n{e}")


def main():
    app = CLMRSApp()
    app.mainloop()


if __name__ == "__main__":
    main()
