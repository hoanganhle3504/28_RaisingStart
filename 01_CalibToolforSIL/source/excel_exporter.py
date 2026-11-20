"""
Excel Exporter for SiL Calibration Tool.
Generates styled Excel workbooks (.xlsx) designed for easy single-line or mass copying
of calibration variables into INCA, CANape, ECU-TEST, or DCM datasets.
"""

import os
from typing import List, Set
from a2l_parser import CalibrationItem

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def export_calibrations_to_excel(
    items: List[CalibrationItem],
    export_path: str,
    target_dcm: str,
    source_a2l: str,
    discovered_buses: Set[str]
) -> str:
    """
    Exports calibration items to a multi-sheet formatted .xlsx workbook:
    - Sheet 1: 'Calibration_Table' (Full table with auto-filter for filtering and copying)
    - Sheet 2: 'Mass_Copy_Variables' (Single column of plain names for 1-click column selection)
    - Sheet 3: 'ECU_TEST_Snippets' (Ready-to-paste Python calls for ECU-TEST packages)
    - Sheet 4: 'Summary' (Metadata and category statistics)
    """
    wb = openpyxl.Workbook()

    # Styling definitions
    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=9)
    bold_data_font = Font(name="Segoe UI", size=9, bold=True)
    code_font = Font(name="Consolas", size=9)
    
    thin_border_side = Side(style="thin", color="CBD5E1")
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    # =========================================================================
    # SHEET 1: Calibration_Table (Full Detailed Table)
    # =========================================================================
    ws1 = wb.active
    ws1.title = "Calibration_Table"
    ws1.views.sheetView[0].showGridLines = True

    headers_1 = [
        "No.",
        "Variable Name",
        "Target Value",
        "Function Group",
        "Rule Category",
        "DCM Block Code"
    ]
    ws1.append(headers_1)

    for col_idx in range(1, len(headers_1) + 1):
        cell = ws1.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = cell_border
    ws1.row_dimensions[1].height = 26

    for idx, item in enumerate(items, 1):
        row_num = idx + 1
        dcm_block = f"FESTWERT {item.name} WERT {item.value:.1f} END"
        row_data = [
            idx,
            item.name,
            item.value,
            item.function_group,
            item.rule_desc,
            dcm_block
        ]
        ws1.append(row_data)

        is_even = (idx % 2 == 0)
        for col_idx in range(1, len(row_data) + 1):
            c = ws1.cell(row=row_num, column=col_idx)
            c.font = code_font if col_idx in (2, 6) else data_font
            c.border = cell_border
            if is_even:
                c.fill = zebra_fill
            if col_idx == 1 or col_idx == 3:
                c.alignment = Alignment(horizontal="center")
            elif col_idx == 4:
                c.alignment = Alignment(horizontal="center")

    # Auto-filter
    ws1.auto_filter.ref = ws1.dimensions

    # =========================================================================
    # SHEET 2: Mass_Copy_Variables (1-Click Column Copy)
    # =========================================================================
    ws2 = wb.create_sheet(title="Mass_Copy_Variables")
    ws2.views.sheetView[0].showGridLines = True
    
    ws2.append(["Variable Name (Click Column A to Mass Copy)"])
    header_cell_2 = ws2.cell(row=1, column=1)
    header_cell_2.fill = header_fill
    header_cell_2.font = header_font
    header_cell_2.alignment = Alignment(horizontal="left", vertical="center")
    ws2.row_dimensions[1].height = 24

    for idx, item in enumerate(items, 1):
        c = ws2.cell(row=idx + 1, column=1, value=item.name)
        c.font = code_font
        c.border = cell_border

    # =========================================================================
    # SHEET 3: ECU_TEST_Snippets (Ready-to-paste Python testcase calls)
    # =========================================================================
    ws3 = wb.create_sheet(title="ECU_TEST_Snippets")
    ws3.views.sheetView[0].showGridLines = True

    headers_3 = ["No.", "Variable Name", "ECU-TEST Python Call Snippet"]
    ws3.append(headers_3)
    for col_idx in range(1, len(headers_3) + 1):
        c = ws3.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = cell_border
    ws3.row_dimensions[1].height = 24

    for idx, item in enumerate(items, 1):
        val_str = str(int(item.value)) if item.value.is_integer() else str(item.value)
        snippet = f'BaseLIB_Sys_Apl_SetCalibValFromName("{item.name}", {val_str})'
        row_data = [idx, item.name, snippet]
        ws3.append(row_data)

        for col_idx in range(1, len(row_data) + 1):
            c = ws3.cell(row=idx + 1, column=col_idx)
            c.font = code_font if col_idx in (2, 3) else data_font
            c.border = cell_border
            if col_idx == 1:
                c.alignment = Alignment(horizontal="center")

    ws3.auto_filter.ref = ws3.dimensions

    # =========================================================================
    # SHEET 4: Summary_Metadata
    # =========================================================================
    ws4 = wb.create_sheet(title="Summary")
    ws4.views.sheetView[0].showGridLines = True

    ws4.append(["Property", "Value"])
    ws4.cell(1, 1).fill = header_fill
    ws4.cell(1, 1).font = header_font
    ws4.cell(1, 2).fill = header_fill
    ws4.cell(1, 2).font = header_font

    meta_rows = [
        ("Target DCM File", target_dcm),
        ("Source A2L File", source_a2l),
        ("Discovered Channels / Buses", ", ".join(sorted(discovered_buses))),
        ("Total New Calibrations", len(items)),
        ("PDU Coding Switches (*swtPduCodActv = 0)", sum(1 for it in items if "Com_swtPduCodActv" in it.name)),
        ("Network Diagnostic Indices (Com_idxDiagDi*_C = 2)", sum(1 for it in items if it.rule_desc.startswith("Network Diagnostic"))),
        ("PDU Diagnostic Indices (Com_idxDiagDi<BUS><PDU>_C = 6)", sum(1 for it in items if it.rule_desc.startswith("PDU Diagnostic"))),
        ("PDU Rx/Tx Gates (Com_swt*xena*_C = 1)", sum(1 for it in items if "Com_swt" in it.name and "Ena" in it.name)),
        ("DFC Control Masks (DFC_CtlMsk2.DFC_Com*_C = 1)", sum(1 for it in items if "DFC_CtlMsk2" in it.name)),
        ("DFC Disable Masks (DFC_DisblMsk2.DFC_Com*_C = 0)", sum(1 for it in items if "DFC_DisblMsk2" in it.name)),
    ]

    for p_name, p_val in meta_rows:
        ws4.append([p_name, p_val])
        row_idx = ws4.max_row
        ws4.cell(row_idx, 1).font = bold_data_font
        ws4.cell(row_idx, 1).border = cell_border
        ws4.cell(row_idx, 2).font = data_font
        ws4.cell(row_idx, 2).border = cell_border

    # Auto-adjust column widths across all sheets
    for ws in [ws1, ws2, ws3, ws4]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = str(cell.value or "")
                # Only consider first line of multiline strings
                first_line = val.split("\n")[0]
                max_len = max(max_len, len(first_line))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(export_path)
    return export_path
