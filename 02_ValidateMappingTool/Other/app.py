"""
app.py - ECU-TEST Project File Matrix Validator GUI.
A modern, responsive desktop application for validating and injecting K-Matrix mapping parameters.
"""

import os
import sys
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, Any, List

# Add parent directory to sys.path if running directly
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from matrix_parser import parse_matrix_excel
from prj_engine import audit_project, apply_fixes


class PrjMatrixValidatorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("ECU-TEST PRJ Matrix Validator")
        self.root.geometry("820x620")
        self.root.minsize(720, 500)

        # Communication Queue for background threads
        self.log_queue: queue.Queue = queue.Queue()

        # Cache parsed matrix data
        self.matrix_signals: Dict[str, Dict[str, Any]] = {}
        self.discrepancies: List[Dict[str, Any]] = []

        self._build_ui()
        self._check_default_paths()
        self.root.after(100, self._process_log_queue)

    def _build_ui(self):
        # Configure styles
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"))
        style.configure("Status.TLabel", font=("Segoe UI", 9, "italic"))
        style.configure("Accent.TButton", font=("Segoe UI", 9, "bold"))

        # Top Frame: File inputs
        input_frame = ttk.LabelFrame(self.root, text=" Target Files ", padding=12)
        input_frame.pack(fill=tk.X, padx=12, pady=8)

        # Excel Matrix path
        ttk.Label(input_frame, text="K-Matrix Excel (.xlsx):").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.excel_var = tk.StringVar()
        self.excel_entry = ttk.Entry(input_frame, textvariable=self.excel_var, width=65)
        self.excel_entry.grid(row=0, column=1, padx=6, pady=4, sticky=tk.EW)
        ttk.Button(input_frame, text="Browse...", command=self._browse_excel).grid(row=0, column=2, pady=4)

        # PRJ File path
        ttk.Label(input_frame, text="Target Project (.prj):").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.prj_var = tk.StringVar()
        self.prj_entry = ttk.Entry(input_frame, textvariable=self.prj_var, width=65)
        self.prj_entry.grid(row=1, column=1, padx=6, pady=4, sticky=tk.EW)
        ttk.Button(input_frame, text="Browse...", command=self._browse_prj).grid(row=1, column=2, pady=4)

        input_frame.columnconfigure(1, weight=1)

        # Button Action Bar
        action_frame = ttk.Frame(self.root, padding=6)
        action_frame.pack(fill=tk.X, padx=12, pady=4)

        self.btn_validate = ttk.Button(
            action_frame,
            text="🔍 Validate Matrix vs. PRJ",
            style="Accent.TButton",
            command=self._start_validation
        )
        self.btn_validate.pack(side=tk.LEFT, padx=4)

        self.btn_apply = ttk.Button(
            action_frame,
            text="⚡ Apply Fixes Atomically",
            command=self._start_apply_fixes,
            state=tk.NORMAL
        )
        self.btn_apply.pack(side=tk.LEFT, padx=4)

        self.btn_clear = ttk.Button(
            action_frame,
            text="Clear Log",
            command=self._clear_log
        )
        self.btn_clear.pack(side=tk.RIGHT, padx=4)

        # Progress / Status Bar
        status_frame = ttk.Frame(self.root, padding=4)
        status_frame.pack(fill=tk.X, padx=12)
        self.status_var = tk.StringVar(value="Ready. Select files and click Validate.")
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, style="Status.TLabel")
        self.status_label.pack(side=tk.LEFT)

        self.progress = ttk.Progressbar(status_frame, mode="indeterminate", length=180)

        # Log & Results Frame
        log_frame = ttk.LabelFrame(self.root, text=" Audit & Execution Log ", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

        self.log_text = tk.Text(log_frame, wrap=tk.WORD, font=("Consolas", 9), background="#1E1E1E", foreground="#D4D4D4")
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Configure log text tags for syntax coloring
        self.log_text.tag_configure("info", foreground="#569CD6")
        self.log_text.tag_configure("success", foreground="#4EC9B0")
        self.log_text.tag_configure("warning", foreground="#CE9178")
        self.log_text.tag_configure("error", foreground="#F44747")
        self.log_text.tag_configure("highlight", foreground="#DCDCAA")

    def _check_default_paths(self):
        """Pre-populate default paths if they exist locally or on network shares."""
        default_excel = r"C:\Workspace\23_Rex_Prototype(1015.1.0)\ComSCL_NetMtrx_RbaNds\comscl_netmtrx_rbands.xlsx"
        default_prj = r"\\fe00fx21.de.bosch.com\USD-WTS-MDG_714_USD$\EPC22\00_LAG8HC\00_Exchange\SiL_Rx_CheckSignals.prj"

        if os.path.exists(default_excel):
            self.excel_var.set(default_excel)
        if os.path.exists(default_prj):
            self.prj_var.set(default_prj)

    def _browse_excel(self):
        initial = os.path.dirname(self.excel_var.get()) if self.excel_var.get() else os.getcwd()
        f = filedialog.askopenfilename(
            title="Select K-Matrix Excel File",
            initialdir=initial,
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if f:
            self.excel_var.set(os.path.normpath(f))

    def _browse_prj(self):
        initial = os.path.dirname(self.prj_var.get()) if self.prj_var.get() else os.getcwd()
        f = filedialog.askopenfilename(
            title="Select Target ECU-TEST Project File",
            initialdir=initial,
            filetypes=[("ECU-TEST Project Files", "*.prj"), ("XML Files", "*.xml"), ("All Files", "*.*")]
        )
        if f:
            self.prj_var.set(os.path.normpath(f))

    def _log(self, text: str, tag: str = "info"):
        self.log_queue.put((text, tag))

    def _clear_log(self):
        self.log_text.delete("1.0", tk.END)

    def _process_log_queue(self):
        try:
            while True:
                msg, tag = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, msg + "\n", tag)
                self.log_text.see(tk.END)
        except queue.Empty:
            pass
        self.root.after(100, self._process_log_queue)

    def _set_busy(self, is_busy: bool, message: str = ""):
        if is_busy:
            self.btn_validate.configure(state=tk.DISABLED)
            self.btn_apply.configure(state=tk.DISABLED)
            self.progress.pack(side=tk.RIGHT, padx=4)
            self.progress.start(10)
            if message:
                self.status_var.set(message)
        else:
            self.btn_validate.configure(state=tk.NORMAL)
            self.btn_apply.configure(state=tk.NORMAL)
            self.progress.stop()
            self.progress.pack_forget()
            if message:
                self.status_var.set(message)

    def _start_validation(self):
        excel_path = self.excel_var.get().strip()
        prj_path = self.prj_var.get().strip()

        if not os.path.exists(excel_path):
            messagebox.showerror("Error", f"Excel file does not exist:\n{excel_path}")
            return
        if not os.path.exists(prj_path):
            messagebox.showerror("Error", f"Project file does not exist:\n{prj_path}")
            return

        self._set_busy(True, "Auditing project against K-Matrix...")
        self._log(f"=== Starting Audit ===", "highlight")
        self._log(f"Matrix:  {excel_path}")
        self._log(f"Project: {prj_path}")

        def worker():
            try:
                # 1. Parse Excel
                self._log("Parsing Excel SignalMapping sheet...", "info")
                self.matrix_signals = parse_matrix_excel(excel_path)
                self._log(f"Found {len(self.matrix_signals)} valid receive signals in Excel.", "info")

                # 2. Audit PRJ
                self.discrepancies = audit_project(
                    prj_path,
                    self.matrix_signals,
                    log_callback=lambda msg: self._log(msg, "info")
                )

                if not self.discrepancies:
                    self._log("SUCCESS: 0 Discrepancies found! Project is 100% in sync with Matrix.", "success")
                    self.root.after(0, lambda: self._set_busy(False, "Audit Passed: 0 Discrepancies."))
                else:
                    self._log(f"\nWARNING: Found {len(self.discrepancies)} Discrepancies:", "warning")
                    for d in self.discrepancies:
                        self._log(
                            f"  [{d['pdu']}] {d['subproject']} -> {d['package']} | {d['param']}: Expected='{d['expected']}', Actual='{d['actual']}'",
                            "error"
                        )
                    self._log(f"\nClick 'Apply Fixes Atomically' to inject missing parameters.", "highlight")
                    self.root.after(0, lambda: self._set_busy(False, f"Audit complete. {len(self.discrepancies)} discrepancies found."))

            except Exception as e:
                self._log(f"ERROR during audit: {str(e)}", "error")
                self.root.after(0, lambda: self._set_busy(False, "Audit encountered an error."))

        threading.Thread(target=worker, daemon=True).start()

    def _start_apply_fixes(self):
        excel_path = self.excel_var.get().strip()
        prj_path = self.prj_var.get().strip()

        if not os.path.exists(excel_path):
            messagebox.showerror("Error", f"Excel file does not exist:\n{excel_path}")
            return
        if not os.path.exists(prj_path):
            messagebox.showerror("Error", f"Project file does not exist:\n{prj_path}")
            return

        confirm = messagebox.askyesno(
            "Confirm Update",
            f"Are you sure you want to adapt the project file?\n\nA backup (.bak) will be created before modification.\n\nTarget:\n{prj_path}"
        )
        if not confirm:
            return

        self._set_busy(True, "Applying fixes atomically...")
        self._log(f"\n=== Starting Atomic Parameter Injection ===", "highlight")

        def worker():
            try:
                # Ensure matrix is parsed
                if not self.matrix_signals:
                    self._log("Parsing Excel SignalMapping sheet...", "info")
                    self.matrix_signals = parse_matrix_excel(excel_path)

                stats = apply_fixes(
                    prj_path,
                    self.matrix_signals,
                    log_callback=lambda msg: self._log(msg, "info")
                )

                self._log("\nUpdate Summary:", "highlight")
                self._log(f"  P_SignalPortDefaultValue: Injected {stats['injected_defaults']} new, Updated {stats['updated_defaults']} existing.", "success")
                self._log(f"  P_SignalPortBitPos:       Injected {stats['injected_bitpos']} new, Updated {stats['updated_bitpos']} existing.", "success")
                self._log("SUCCESS: Atomic project adaptation completed!", "success")

                self.root.after(0, lambda: self._set_busy(False, "Fixes applied successfully!"))
                self.root.after(0, lambda: messagebox.showinfo("Success", "Project file updated and validated successfully!"))

            except Exception as e:
                self._log(f"ERROR during update: {str(e)}", "error")
                self.root.after(0, lambda: self._set_busy(False, "Update failed."))
                self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to apply fixes:\n{str(e)}"))

        threading.Thread(target=worker, daemon=True).start()


def main():
    root = tk.Tk()
    app = PrjMatrixValidatorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
