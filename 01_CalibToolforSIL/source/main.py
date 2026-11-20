"""
SiL Calibration Tool (Tkinter Desktop Application)
Analyzes vECU.a2l and Modified_Calibrations.dcm, discovers communication buses dynamically,
extracts and cross-checks calibrations against existing datasets, and presents an interactive
panel for approval and automatic appending.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, List, Set

from a2l_parser import CalibrationItem, extract_needed_calibrations
from dcm_parser import parse_existing_dcm_variables
from dcm_writer import append_calibrations_to_dcm
from excel_exporter import export_calibrations_to_excel


class SilCalibApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("SiL Software Calibration Tool")
        self.root.geometry("1100x750")
        self.root.minsize(900, 600)

        # State
        self.discovered_buses: Set[str] = set()
        self.existing_dcm_vars: Dict[str, float] = {}
        self.all_needed_items: List[CalibrationItem] = []
        self.pending_items: List[CalibrationItem] = []
        self.item_checkbox_vars: Dict[str, tk.BooleanVar] = {}
        self.last_log_path: str = ""

        # Styles
        self._setup_styles()
        self._build_ui()

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        
        # Colors
        self.bg_color = "#f8fafc"
        self.card_bg = "#ffffff"
        self.header_bg = "#0f172a"
        self.header_fg = "#ffffff"
        self.accent_color = "#2563eb"
        self.success_color = "#16a34a"
        
        self.root.configure(bg=self.bg_color)
        
        style.configure("TFrame", background=self.bg_color)
        style.configure("Card.TFrame", background=self.card_bg, relief="groove")
        style.configure("TLabel", background=self.bg_color, font=("Segoe UI", 9))
        style.configure("Header.TLabel", background=self.header_bg, foreground=self.header_fg, font=("Segoe UI", 14, "bold"))
        style.configure("SubHeader.TLabel", background=self.header_bg, foreground="#94a3b8", font=("Segoe UI", 9))
        style.configure("Section.TLabel", font=("Segoe UI", 10, "bold"), background=self.bg_color, foreground="#1e293b")
        style.configure("StatNum.TLabel", font=("Segoe UI", 16, "bold"), foreground=self.accent_color, background=self.card_bg)
        style.configure("StatLabel.TLabel", font=("Segoe UI", 8), foreground="#64748b", background=self.card_bg)
        
        style.configure("Action.TButton", font=("Segoe UI", 9, "bold"), padding=6)
        style.configure("Success.TButton", font=("Segoe UI", 10, "bold"), padding=8, background=self.success_color, foreground="white")

    def _build_ui(self):
        # 1. Header Banner
        header_frame = tk.Frame(self.root, bg=self.header_bg, height=70)
        header_frame.pack(fill=tk.X, side=tk.TOP)
        
        tk.Label(
            header_frame, text="SiL Software Calibration Merger & Validator",
            bg=self.header_bg, fg="white", font=("Segoe UI", 14, "bold")
        ).pack(anchor="w", padx=20, pady=(12, 2))
        
        tk.Label(
            header_frame, text="Universal Channel Auto-Discovery  •  DCM Conflict Avoidance  •  Pre-Testing Flash Automation",
            bg=self.header_bg, fg="#94a3b8", font=("Segoe UI", 9)
        ).pack(anchor="w", padx=20, pady=(0, 12))

        # Main Container
        main_container = ttk.Frame(self.root, padding=15)
        main_container.pack(fill=tk.BOTH, expand=True)

        # 2. File Selection Card
        file_card = ttk.Frame(main_container, style="Card.TFrame", padding=12)
        file_card.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(file_card, text="File Configuration Paths", style="Section.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

        # DCM File Picker
        ttk.Label(file_card, text="Target Modified_Calibrations.dcm:").grid(row=1, column=0, sticky="w", pady=4)
        self.dcm_path_var = tk.StringVar(value=r"C:\Workspace\LAG8HC\02_1015_1_Rex_Prototype\1015.1_Prototype\Simulate\02_ECU\vECU\Config\DCM\Modified_Calibrations.dcm")
        self.dcm_entry = ttk.Entry(file_card, textvariable=self.dcm_path_var, width=75)
        self.dcm_entry.grid(row=1, column=1, sticky="ew", padx=6, pady=4)
        ttk.Button(file_card, text="Browse...", command=self._browse_dcm).grid(row=1, column=2, padx=4, pady=4)

        # A2L File Picker
        ttk.Label(file_card, text="Source vECU.a2l:").grid(row=2, column=0, sticky="w", pady=4)
        self.a2l_path_var = tk.StringVar(value=r"C:\Workspace\LAG8HC\02_1015_1_Rex_Prototype\1015.1_Prototype\Simulate\02_ECU\vECU\vECU.a2l")
        self.a2l_entry = ttk.Entry(file_card, textvariable=self.a2l_path_var, width=75)
        self.a2l_entry.grid(row=2, column=1, sticky="ew", padx=6, pady=4)
        ttk.Button(file_card, text="Browse...", command=self._browse_a2l).grid(row=2, column=2, padx=4, pady=4)

        file_card.columnconfigure(1, weight=1)

        # Analyze Button
        btn_bar = ttk.Frame(file_card)
        btn_bar.grid(row=3, column=0, columnspan=3, sticky="e", pady=(8, 0))
        self.analyze_btn = ttk.Button(btn_bar, text="🔍 Analyze & Cross-Check", style="Action.TButton", command=self._start_analysis)
        self.analyze_btn.pack(side=tk.RIGHT)

        # 3. Stats / Channel Summary Bar
        self.stats_frame = ttk.Frame(main_container, style="Card.TFrame", padding=10)
        self.stats_frame.pack(fill=tk.X, pady=(0, 10))

        # Cards inside stats frame
        self.bus_stat_label = tk.Label(self.stats_frame, text="None", font=("Segoe UI", 11, "bold"), fg="#2563eb", bg="white")
        self._create_stat_box(self.stats_frame, 0, "Discovered Channels", self.bus_stat_label)

        self.needed_stat_label = tk.Label(self.stats_frame, text="0", font=("Segoe UI", 14, "bold"), fg="#0f172a", bg="white")
        self._create_stat_box(self.stats_frame, 1, "Total Required in A2L", self.needed_stat_label)

        self.existing_stat_label = tk.Label(self.stats_frame, text="0", font=("Segoe UI", 14, "bold"), fg="#64748b", bg="white")
        self._create_stat_box(self.stats_frame, 2, "Already Present in DCM (Excluded)", self.existing_stat_label)

        self.pending_stat_label = tk.Label(self.stats_frame, text="0", font=("Segoe UI", 14, "bold"), fg="#16a34a", bg="white")
        self._create_stat_box(self.stats_frame, 3, "New Calibrations to Add", self.pending_stat_label)

        for col in range(4):
            self.stats_frame.columnconfigure(col, weight=1)

        # 4. Calibration Table / Validation Panel
        table_card = ttk.Frame(main_container, style="Card.TFrame", padding=10)
        table_card.pack(fill=tk.BOTH, expand=True)

        # Table Header & Filter Bar
        tbl_top = ttk.Frame(table_card)
        tbl_top.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(tbl_top, text="Proposed Calibration Overrides for Validation", style="Section.TLabel").pack(side=tk.LEFT)

        # Search / Filter
        ttk.Label(tbl_top, text="Filter: ").pack(side=tk.LEFT, padx=(25, 4))
        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", lambda *args: self._apply_filter())
        self.filter_entry = ttk.Entry(tbl_top, textvariable=self.filter_var, width=25)
        self.filter_entry.pack(side=tk.LEFT)

        # Select / Deselect All / Copy
        ttk.Button(tbl_top, text="Select All", command=self._select_all).pack(side=tk.RIGHT, padx=2)
        ttk.Button(tbl_top, text="Deselect All", command=self._deselect_all).pack(side=tk.RIGHT, padx=2)
        ttk.Button(tbl_top, text="📋 Copy Names", command=self._copy_selected_names).pack(side=tk.RIGHT, padx=4)
        ttk.Button(tbl_top, text="📊 Export to Excel (.xlsx)", command=self._export_excel).pack(side=tk.RIGHT, padx=4)

        # Treeview with Scrollbars
        tree_container = ttk.Frame(table_card)
        tree_container.pack(fill=tk.BOTH, expand=True)

        columns = ("include", "name", "value", "module", "rule")
        self.tree = ttk.Treeview(tree_container, columns=columns, show="headings", selectmode="extended")
        
        self.tree.heading("include", text="Add?", anchor="center")
        self.tree.heading("name", text="Variable Name", anchor="w")
        self.tree.heading("value", text="Target Value", anchor="center")
        self.tree.heading("module", text="Function Group", anchor="center")
        self.tree.heading("rule", text="Rule Description", anchor="w")

        self.tree.column("include", width=60, anchor="center")
        self.tree.column("name", width=360, anchor="w")
        self.tree.column("value", width=100, anchor="center")
        self.tree.column("module", width=140, anchor="center")
        self.tree.column("rule", width=380, anchor="w")

        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.bind("<Button-3>", self._show_context_menu)

        # Right-click context menu
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="📋 Copy Variable Name(s)", command=self._copy_selected_names)
        self.context_menu.add_command(label="📋 Copy as ECU-TEST Python Calls", command=self._copy_selected_python)
        self.context_menu.add_command(label="📋 Copy as DCM FESTWERT Blocks", command=self._copy_selected_dcm)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Toggle Selected (Add / Exclude)", command=self._toggle_selected_items)

        # 5. Bottom Action Bar
        action_bar = ttk.Frame(main_container, padding=(0, 10, 0, 0))
        action_bar.pack(fill=tk.X, side=tk.BOTTOM)

        self.status_var = tk.StringVar(value="Ready. Specify paths and click 'Analyze & Cross-Check'.")
        self.status_label = ttk.Label(action_bar, textvariable=self.status_var, foreground="#475569")
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.export_excel_btn = ttk.Button(action_bar, text="📊 Export to Excel (.xlsx)", command=self._export_excel, state=tk.DISABLED)
        self.export_excel_btn.pack(side=tk.RIGHT, padx=4)

        self.export_btn = ttk.Button(action_bar, text="📄 Export Document (.doc/.txt)", command=self._export_document, state=tk.DISABLED)
        self.export_btn.pack(side=tk.RIGHT, padx=4)

        self.apply_btn = ttk.Button(action_bar, text="✅ Approve & Append to DCM", style="Success.TButton", command=self._approve_and_append, state=tk.DISABLED)
        self.apply_btn.pack(side=tk.RIGHT, padx=6)

    def _create_stat_box(self, parent, col, title, num_label):
        box = tk.Frame(parent, bg="white", relief="solid", bd=1, padx=10, pady=8)
        box.grid(row=0, column=col, sticky="nsew", padx=4)
        tk.Label(box, text=title, font=("Segoe UI", 8), fg="#64748b", bg="white").pack(anchor="w")
        num_label.pack(anchor="w", in_=box)

    def _browse_dcm(self):
        f = filedialog.askopenfilename(title="Select Modified_Calibrations.dcm", filetypes=[("DCM Files", "*.dcm"), ("All Files", "*.*")])
        if f:
            self.dcm_path_var.set(f)

    def _browse_a2l(self):
        f = filedialog.askopenfilename(title="Select vECU.a2l", filetypes=[("A2L Files", "*.a2l"), ("All Files", "*.*")])
        if f:
            self.a2l_path_var.set(f)

    def _start_analysis(self):
        dcm_path = self.dcm_path_var.get().strip()
        a2l_path = self.a2l_path_var.get().strip()

        if not os.path.isfile(dcm_path):
            messagebox.showerror("File Error", f"DCM file not found:\n{dcm_path}")
            return
        if not os.path.isfile(a2l_path):
            messagebox.showerror("File Error", f"A2L file not found:\n{a2l_path}")
            return

        self.analyze_btn.config(state=tk.DISABLED)
        self.status_var.set("Scanning A2L and DCM files (streaming parse)... Please wait.")

        # Run in thread so GUI remains responsive
        threading.Thread(target=self._run_analysis, args=(dcm_path, a2l_path), daemon=True).start()

    def _run_analysis(self, dcm_path: str, a2l_path: str):
        try:
            # 1. Parse DCM
            self.existing_dcm_vars = parse_existing_dcm_variables(dcm_path)

            # 2. Parse A2L
            self.all_needed_items, self.discovered_buses = extract_needed_calibrations(a2l_path)

            # 3. Filter out existing items
            self.pending_items = []
            self.item_checkbox_vars.clear()

            for item in self.all_needed_items:
                if item.name not in self.existing_dcm_vars:
                    self.pending_items.append(item)
                    self.item_checkbox_vars[item.name] = tk.BooleanVar(value=True)

            self.root.after(0, self._analysis_complete)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Analysis Error", str(e)))
            self.root.after(0, lambda: self.status_var.set(f"Error during analysis: {e}"))
            self.root.after(0, lambda: self.analyze_btn.config(state=tk.NORMAL))

    def _analysis_complete(self):
        self.analyze_btn.config(state=tk.NORMAL)
        
        # Update Stats Cards
        bus_str = ", ".join(sorted(self.discovered_buses)) if self.discovered_buses else "None"
        self.bus_stat_label.config(text=bus_str)
        self.needed_stat_label.config(text=str(len(self.all_needed_items)))
        self.existing_stat_label.config(text=str(len(self.all_needed_items) - len(self.pending_items)))
        self.pending_stat_label.config(text=str(len(self.pending_items)))

        # Populate Table
        self._populate_tree()

        if self.pending_items:
            self.status_var.set(f"Analysis Complete: Found {len(self.pending_items)} missing calibrations. Review below and click Approve.")
            self.apply_btn.config(state=tk.NORMAL)
            self.export_btn.config(state=tk.NORMAL)
            self.export_excel_btn.config(state=tk.NORMAL)
        else:
            self.status_var.set("All required calibrations are already present in DCM. No changes needed.")
            self.apply_btn.config(state=tk.DISABLED)
            self.export_btn.config(state=tk.DISABLED)
            self.export_excel_btn.config(state=tk.DISABLED)

    def _populate_tree(self):
        self.tree.delete(*self.tree.get_children())
        filter_text = self.filter_var.get().strip().lower()

        for item in self.pending_items:
            if filter_text and filter_text not in item.name.lower() and filter_text not in item.rule_desc.lower():
                continue
            is_checked = self.item_checkbox_vars[item.name].get()
            check_mark = "☑ YES" if is_checked else "☐ NO"
            self.tree.insert("", tk.END, iid=item.name, values=(check_mark, item.name, f"{item.value:.1f}", item.function_group, item.rule_desc))

    def _apply_filter(self):
        self._populate_tree()

    def _on_tree_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if item_id and item_id in self.item_checkbox_vars:
            curr = self.item_checkbox_vars[item_id].get()
            self.item_checkbox_vars[item_id].set(not curr)
            self._populate_tree()

    def _show_context_menu(self, event):
        item_id = self.tree.identify_row(event.y)
        if item_id:
            # If clicked item is not in current selection, select it
            if item_id not in self.tree.selection():
                self.tree.selection_set(item_id)
            self.context_menu.post(event.x_root, event.y_root)

    def _toggle_selected_items(self):
        selected_ids = self.tree.selection()
        for item_id in selected_ids:
            if item_id in self.item_checkbox_vars:
                curr = self.item_checkbox_vars[item_id].get()
                self.item_checkbox_vars[item_id].set(not curr)
        self._populate_tree()

    def _get_target_items_for_copy(self) -> List[CalibrationItem]:
        """Returns items based on tree selection if any, else all checked items."""
        selected_ids = self.tree.selection()
        if selected_ids:
            id_set = set(selected_ids)
            return [it for it in self.pending_items if it.name in id_set]
        return [it for it in self.pending_items if self.item_checkbox_vars.get(it.name, tk.BooleanVar(value=False)).get()]

    def _copy_selected_names(self):
        """Mass-copies variable names to system clipboard."""
        target_items = self._get_target_items_for_copy()
        if not target_items:
            messagebox.showwarning("Warning", "No variables selected to copy.")
            return

        text = "\n".join(it.name for it in target_items)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status_var.set(f"Copied {len(target_items)} variable name(s) to clipboard!")

    def _copy_selected_python(self):
        """Mass-copies ECU-TEST Python snippet calls to system clipboard."""
        target_items = self._get_target_items_for_copy()
        if not target_items:
            messagebox.showwarning("Warning", "No variables selected to copy.")
            return

        lines = []
        for it in target_items:
            val_str = str(int(it.value)) if it.value.is_integer() else str(it.value)
            lines.append(f'BaseLIB_Sys_Apl_SetCalibValFromName("{it.name}", {val_str})')

        text = "\n".join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status_var.set(f"Copied {len(target_items)} ECU-TEST Python call(s) to clipboard!")

    def _copy_selected_dcm(self):
        """Mass-copies DCM FESTWERT blocks to system clipboard."""
        target_items = self._get_target_items_for_copy()
        if not target_items:
            messagebox.showwarning("Warning", "No variables selected to copy.")
            return

        blocks = []
        for it in target_items:
            blocks.append(
                f"FESTWERT {it.name}\n"
                f"   LANGNAME \"\"\n"
                f"   FUNKTION {it.function_group}\n"
                f"   EINHEIT_W \"-\"\n"
                f"   WERT {it.value:.16f}\n"
                f"END\n"
            )

        text = "\n".join(blocks)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status_var.set(f"Copied {len(target_items)} DCM FESTWERT block(s) to clipboard!")

    def _select_all(self):
        for v in self.item_checkbox_vars.values():
            v.set(True)
        self._populate_tree()

    def _deselect_all(self):
        for v in self.item_checkbox_vars.values():
            v.set(False)
        self._populate_tree()

    def _export_excel(self):
        """Exports the approved/validated calibration list to a formatted Excel workbook (.xlsx)."""
        selected_items = [it for it in self.pending_items if self.item_checkbox_vars[it.name].get()]
        if not selected_items:
            messagebox.showwarning("Warning", "No variables selected to export.")
            return

        export_path = filedialog.asksaveasfilename(
            title="Save Calibration Excel Workbook",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")],
            initialfile="SiL_Calibration_Variables_Export.xlsx"
        )
        if not export_path:
            return

        try:
            export_calibrations_to_excel(
                items=selected_items,
                export_path=export_path,
                target_dcm=self.dcm_path_var.get(),
                source_a2l=self.a2l_path_var.get(),
                discovered_buses=self.discovered_buses
            )
            messagebox.showinfo(
                "Excel Export Successful",
                f"Successfully exported {len(selected_items)} variables to Excel:\n{export_path}\n\n"
                "Sheets included:\n"
                "• Calibration_Table (full formatted table with auto-filter)\n"
                "• Mass_Copy_Variables (single column for 1-click mass copying)\n"
                "• ECU_TEST_Snippets (ready-to-paste Python calls)\n"
                "• Summary (metadata & category statistics)"
            )
            self.status_var.set(f"Excel workbook exported: {os.path.basename(export_path)}")
        except Exception as e:
            messagebox.showerror("Export Failed", str(e))

    def _export_document(self):
        """Exports the approved/validated calibration list to a .doc / .txt document."""
        selected_items = [it for it in self.pending_items if self.item_checkbox_vars[it.name].get()]
        if not selected_items:
            messagebox.showwarning("Warning", "No variables selected to export.")
            return

        export_path = filedialog.asksaveasfilename(
            title="Save Validation Report Document",
            defaultextension=".doc",
            filetypes=[("Word Document", "*.doc"), ("Text Report", "*.txt"), ("All Files", "*.*")],
            initialfile="SiL_Calibration_Validation_Report.doc"
        )
        if not export_path:
            return

        try:
            with open(export_path, "w", encoding="utf-8") as f:
                f.write("SiL SOFTWARE CALIBRATION VALIDATION REPORT\n")
                f.write("=" * 85 + "\n\n")
                f.write(f"Target DCM Path     : {self.dcm_path_var.get()}\n")
                f.write(f"Source A2L Path     : {self.a2l_path_var.get()}\n")
                f.write(f"Discovered Channels : {', '.join(sorted(self.discovered_buses))}\n")
                f.write(f"Selected Variables  : {len(selected_items)}\n")
                f.write("-" * 85 + "\n")
                f.write(f"{'No.':<4} | {'Variable Name':<45} | {'Value':<8} | {'Function':<16} | {'Rule Description'}\n")
                f.write("-" * 85 + "\n")
                for i, it in enumerate(selected_items, 1):
                    f.write(f"{i:<4} | {it.name:<45} | {it.value:<8.1f} | {it.function_group:<16} | {it.rule_desc}\n")
                f.write("=" * 85 + "\n")
                f.write("Status: APPROVED FOR DCM APPENDING\n")

            messagebox.showinfo("Export Successful", f"Validation report document saved to:\n{export_path}")
        except Exception as e:
            messagebox.showerror("Export Failed", str(e))

    def _approve_and_append(self):
        selected_items = [it for it in self.pending_items if self.item_checkbox_vars[it.name].get()]
        if not selected_items:
            messagebox.showwarning("Warning", "No variables selected to append.")
            return

        dcm_path = self.dcm_path_var.get().strip()
        a2l_path = self.a2l_path_var.get().strip()

        confirm = messagebox.askyesno(
            "Confirm Append",
            f"You are about to append {len(selected_items)} new calibrations to:\n{dcm_path}\n\n"
            "A safe backup (.bak) will be created automatically.\nProceed?"
        )
        if not confirm:
            return

        try:
            log_path = append_calibrations_to_dcm(
                dcm_path=dcm_path,
                items_to_add=selected_items,
                source_a2l=a2l_path,
                discovered_buses=self.discovered_buses
            )
            self.last_log_path = log_path

            messagebox.showinfo(
                "Merge Complete",
                f"Successfully appended {len(selected_items)} variables to:\n{dcm_path}\n\n"
                f"Audit Log written to:\n{log_path}"
            )
            self.status_var.set(f"Completed! Added {len(selected_items)} variables. Log: {os.path.basename(log_path)}")

            # Refresh
            self._start_analysis()

        except Exception as e:
            messagebox.showerror("Merge Failed", str(e))
            self.status_var.set(f"Failed to append: {e}")


def launch():
    root = tk.Tk()
    app = SilCalibApp(root)
    root.mainloop()


if __name__ == "__main__":
    launch()
