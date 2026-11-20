import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from pathlib import Path
import os
import glob
from datetime import datetime


# ==========================================
# COLOR CONSTANTS (matching original file)
# ==========================================
HEADER_GRAY = 'A5A5A5'         # Header row background
ROW_LIGHT_GRAY = 'EDEDED'      # Alternating row stripe (80% tint of accent3 #A5A5A5)
ROW_WHITE = 'FFFFFF'            # Other rows

RESULT_COLORS = {
    'NONE': 'CDC5B9',
    'SUCCESS': '77FF5B',
    'INCONCLUSIVE': 'FFF15C',
    'FAILED': 'F25757',
    'ERROR': 'B40000'
}

RESULT_HIERARCHY = {
    'NONE': 5,
    'SUCCESS': 4,
    'INCONCLUSIVE': 3,
    'FAILED': 2,
    'ERROR': 1,
    '': 0,
    None: 0
}


def auto_fit_columns(ws, min_width=10, max_width=50, start_row=1):
    """Auto-fit column widths based on content"""
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row >= start_row and cell.value is not None:
                cell_len = len(str(cell.value))
                if cell_len > max_length:
                    max_length = cell_len
        adjusted = max(min_width, min(max_length + 3, max_width))
        ws.column_dimensions[col_letter].width = adjusted


def preprocess_window_precond(df, col_C, col_E):
    df = df.copy()
    processed_indices = set()

    i = 0
    while i < len(df):
        if i in processed_indices:
            i += 1
            continue

        if df.loc[i, col_E] == "Disable_WindowPreCond":
            disable_idx = i
            test_count = 0
            enable_idx = None
            tests_after_disable = []

            j = i + 1
            while j < len(df):
                if df.loc[j, col_E] == "Enable_WindowPreCond":
                    enable_idx = j
                    break
                tests_after_disable.append(j)
                test_count += 1
                j += 1

            if enable_idx is not None:
                tests_after_enable = []
                for k in range(enable_idx + 1, min(enable_idx + 1 + test_count, len(df))):
                    tests_after_enable.append(k)

                if len(tests_after_enable) == test_count:
                    for test_idx in tests_after_disable:
                        old_value = df.loc[test_idx, col_E]
                        df.loc[test_idx, col_E] = f"{old_value}_Dis_Precond"
                        processed_indices.add(test_idx)
                    for test_idx in tests_after_enable:
                        old_value = df.loc[test_idx, col_E]
                        df.loc[test_idx, col_E] = f"{old_value}_En_Precond"
                        processed_indices.add(test_idx)
                else:
                    print(f"  ⚠ Warning: Mismatch at row {disable_idx + 14}, skipping...")
            else:
                print(f"  ⚠ Warning: No matching Enable for Disable at row {disable_idx + 14}")

        i += 1

    return df


def write_summary_area(ws, autofill_col_letter, data_start_row, last_data_row):
    """Write rows 1-12 summary area with SUBTOTAL formulas that respect filters"""

    bold_font = Font(name='Calibri', size=11, bold=True)
    gray_font = Font(name='Calibri', size=11, color='808080')

    # Data range for AutoFilled column
    af_range = f"{autofill_col_letter}{data_start_row}:{autofill_col_letter}{last_data_row}"

    # Row 2: Title
    ws.cell(row=2, column=2, value="COM-Test Report (Merged)").font = Font(name='Calibri', size=28, bold=True)
    ws.row_dimensions[2].height = 36

    # Row 4: Overview + TOTAL (SUBTOTAL 103 = COUNTA that ignores hidden/filtered rows)
    ws.cell(row=4, column=2, value="Overview").font = Font(name='Calibri', size=14, bold=True)
    ws.row_dimensions[4].height = 18.75
    ws.cell(row=4, column=5, value="TOTAL").font = Font(name='Calibri', size=11)
    ws.cell(row=4, column=6).value = f"=SUBTOTAL(103,{af_range})"
    ws.cell(row=4, column=6).font = Font(name='Calibri', size=11)

    # Row 5: Evaluation (worst result in visible rows)
    ws.cell(row=5, column=2, value="Evaluation:").font = bold_font
    # Use SUMPRODUCT + SUBTOTAL to check filtered results
    eval_formula = (
        f'=IF(SUMPRODUCT(SUBTOTAL(103,OFFSET({autofill_col_letter}{data_start_row},ROW({af_range})-ROW({autofill_col_letter}{data_start_row}),0))*({af_range}="ERROR"))>0,"ERROR",'
        f'IF(SUMPRODUCT(SUBTOTAL(103,OFFSET({autofill_col_letter}{data_start_row},ROW({af_range})-ROW({autofill_col_letter}{data_start_row}),0))*({af_range}="FAILED"))>0,"FAILED",'
        f'IF(SUMPRODUCT(SUBTOTAL(103,OFFSET({autofill_col_letter}{data_start_row},ROW({af_range})-ROW({autofill_col_letter}{data_start_row}),0))*({af_range}="INCONCLUSIVE"))>0,"INCONCLUSIVE",'
        f'IF(SUMPRODUCT(SUBTOTAL(103,OFFSET({autofill_col_letter}{data_start_row},ROW({af_range})-ROW({autofill_col_letter}{data_start_row}),0))*({af_range}="SUCCESS"))>0,"SUCCESS","NONE"))))'
    )
    ws.cell(row=5, column=3).value = eval_formula
    ws.cell(row=5, column=3).font = bold_font

    # Metadata
    ws.cell(row=6, column=2, value="Testbench:").font = bold_font
    ws.cell(row=6, column=3, value="Merged from multiple files").font = gray_font
    ws.cell(row=7, column=2, value="Software:").font = bold_font
    ws.cell(row=7, column=3, value="-").font = gray_font
    ws.cell(row=8, column=2, value="Tester:").font = bold_font
    ws.cell(row=8, column=3, value="-").font = gray_font
    ws.cell(row=9, column=2, value="Executed:").font = bold_font
    ws.cell(row=9, column=3, value=datetime.now().strftime("%Y-%m-%d %H:%M:%S")).font = gray_font
    ws.cell(row=10, column=2, value="Duration in h:").font = bold_font
    ws.cell(row=10, column=3, value="-").font = gray_font

    # Result stats in E5-G9
    # SUMPRODUCT + SUBTOTAL trick to count specific values in visible rows only
    results_order = ['NONE', 'SUCCESS', 'INCONCLUSIVE', 'FAILED', 'ERROR']
    for idx, result in enumerate(results_order):
        row = 5 + idx

        # E column: result label with color
        e_cell = ws.cell(row=row, column=5, value=result)
        e_cell.font = Font(name='Calibri', size=11, bold=True)
        if result in RESULT_COLORS:
            e_cell.fill = PatternFill(start_color=RESULT_COLORS[result],
                                      end_color=RESULT_COLORS[result],
                                      fill_type='solid')

        # F column: filter-aware COUNT formula
        # SUMPRODUCT(SUBTOTAL(103,OFFSET(first_cell,ROW(range)-ROW(first_cell),0))*(range="VALUE"))
        count_formula = (
            f'=SUMPRODUCT(SUBTOTAL(103,OFFSET({autofill_col_letter}{data_start_row},'
            f'ROW({af_range})-ROW({autofill_col_letter}{data_start_row}),0))'
            f'*({af_range}="{result}"))'
        )
        ws.cell(row=row, column=6).value = count_formula
        ws.cell(row=row, column=6).font = Font(name='Calibri', size=11)

        # G column: PERCENTAGE formula
        ws.cell(row=row, column=7).value = f'=IF(F4>0,F{row}/F4,0)'
        ws.cell(row=row, column=7).font = Font(name='Calibri', size=11)
        ws.cell(row=row, column=7).number_format = '0.00%'


def apply_manual_table_formatting(ws, header_row, data_start_row, last_data_row, first_col, last_col):
    """
    Apply manual gray formatting to look like the original file.
    This avoids theme-dependent table style colors.
    """
    header_fill = PatternFill(start_color=HEADER_GRAY, end_color=HEADER_GRAY, fill_type='solid')
    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    stripe_fill = PatternFill(start_color=ROW_LIGHT_GRAY, end_color=ROW_LIGHT_GRAY, fill_type='solid')

    thin_border = Border(
        left=Side(style='thin', color='E0E0E0'),
        right=Side(style='thin', color='E0E0E0'),
        top=Side(style='thin', color='E0E0E0'),
        bottom=Side(style='thin', color='E0E0E0')
    )

    # Header row
    for col in range(first_col, last_col + 1):
        cell = ws.cell(row=header_row, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    # Data rows: alternating stripes
    for row in range(data_start_row, last_data_row + 1):
        row_offset = row - data_start_row
        for col in range(first_col, last_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border
            # Every other row gets light gray stripe
            if row_offset % 2 == 1:
                cell.fill = stripe_fill


def merge_excel_sheets(file_paths, output_path):
    """
    Merge sheets from one or multiple Excel files.
    Output matches reference file: gray table + filters + conditional formatting + dropdown.
    """
    all_data = []

    for file_path in file_paths:
        print(f"\n📂 Processing file: {os.path.basename(file_path)}")
        print("-" * 40)

        excel_file = pd.ExcelFile(file_path)

        for sheet_name in excel_file.sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=12)

                if len(df) > 0:
                    df['_source_sheet'] = sheet_name
                    df['_source_file'] = os.path.basename(file_path)
                    all_data.append(df)
                    print(f"  ✓ Added {len(df)} rows from '{sheet_name}'")
                else:
                    print(f"  ⚠ Warning: Sheet '{sheet_name}' has no data rows, skipping...")
            except Exception as e:
                print(f"  ✗ Error reading sheet '{sheet_name}': {str(e)}")

    if not all_data:
        print("\n✗ No data found in any file!")
        return

    print("\n" + "=" * 60)

    merged_df = pd.concat(all_data, ignore_index=True)
    print(f"✓ Total rows collected: {len(merged_df)}")

    # Remove 'Unnamed' columns
    unnamed_cols = [c for c in merged_df.columns if str(c).startswith('Unnamed')]
    if unnamed_cols:
        merged_df = merged_df.drop(columns=unnamed_cols)
        print(f"  Removed {len(unnamed_cols)} empty column(s)")

    # Identify columns
    col_names = list(merged_df.columns)
    col_Bus = col_names[0]
    col_Frame = col_names[1]
    col_Signal = col_names[2]
    col_TestName = col_names[3]
    col_Result = col_names[5]

    merged_df[col_Signal] = merged_df[col_Signal].fillna('').astype(str)
    merged_df[col_TestName] = merged_df[col_TestName].fillna('').astype(str)
    merged_df[col_Frame] = merged_df[col_Frame].fillna('').astype(str)

    # WindowPreCond preprocessing
    merged_df = preprocess_window_precond(merged_df, col_Frame, col_TestName)

    # Deduplicate with cross-file SUCCESS detection
    merged_df['_identifier'] = (
        merged_df[col_Bus].astype(str) + '|||' +
        merged_df[col_Frame] + '|||' +
        merged_df[col_Signal] + '|||' +
        merged_df[col_TestName]
    )
    merged_df['_result_level'] = merged_df[col_Result].map(
        lambda x: RESULT_HIERARCHY.get(x, 0) if pd.notna(x) else 0
    )

    # Find Reclassification column name
    reclass_col = None
    for c in merged_df.columns:
        if 'reclassification' in str(c).lower():
            reclass_col = c
            break

    # =============================================
    # CROSS-FILE LOGIC:
    # For each duplicate group with different results:
    #   - Keep the WORST result as the main Result row
    #   - If another file has SUCCESS, set Reclassification = SUCCESS
    #   - Add NOTE = filename that had SUCCESS
    # =============================================
    print("\n🔄 Analyzing cross-file results...")

    # Build a map: identifier -> list of (result, source_file) 
    id_results = merged_df.groupby('_identifier').apply(
        lambda g: list(zip(g[col_Result].tolist(), g['_source_file'].tolist()))
    ).to_dict()

    # For each identifier, determine reclassification and note
    reclass_map = {}  # identifier -> SUCCESS
    note_map = {}     # identifier -> filename(s) that had SUCCESS
    cross_file_count = 0

    for ident, result_file_pairs in id_results.items():
        # Get unique results
        unique_results = set(r for r, f in result_file_pairs if pd.notna(r))

        if len(unique_results) > 1 and 'SUCCESS' in unique_results:
            # There are different results AND one is SUCCESS
            # Find which file(s) had SUCCESS
            success_files = [f for r, f in result_file_pairs if r == 'SUCCESS']
            non_success = [r for r in unique_results if r != 'SUCCESS']

            if non_success:  # There are non-SUCCESS results too
                reclass_map[ident] = 'SUCCESS'
                note_map[ident] = ', '.join(sorted(set(success_files)))
                cross_file_count += 1

    if cross_file_count > 0:
        print(f"  ✓ Found {cross_file_count} test(s) with SUCCESS in another file")
    else:
        print(f"  ℹ No cross-file SUCCESS differences found")

    # Sort: keep WORST result (ascending level = worst first)
    merged_df = merged_df.sort_values(
        by=['_identifier', '_result_level'],
        ascending=[True, True]  # ascending=True → lowest level (worst) first
    )
    final_df = merged_df.drop_duplicates(subset='_identifier', keep='first')

    # Apply reclassification and notes from cross-file analysis
    final_df = final_df.copy()
    final_df['_identifier_temp'] = (
        final_df[col_Bus].astype(str) + '|||' +
        final_df[col_Frame] + '|||' +
        final_df[col_Signal] + '|||' +
        final_df[col_TestName]
    )

    # Set Reclassification for cross-file SUCCESS cases
    if reclass_col:
        # Ensure column is string type (may be float64 if all NaN)
        final_df[reclass_col] = final_df[reclass_col].astype(object)
        for idx in final_df.index:
            ident = final_df.loc[idx, '_identifier_temp']
            if ident in reclass_map:
                final_df.loc[idx, reclass_col] = reclass_map[ident]

    # Add NOTE column for cross-file info
    note_col = None
    for c in final_df.columns:
        if str(c).strip().upper() == 'NOTE':
            note_col = c
            break

    if note_col is None:
        note_col = 'NOTE'
        final_df[note_col] = None

    # Ensure NOTE column is string type
    final_df[note_col] = final_df[note_col].astype(object)
    for idx in final_df.index:
        ident = final_df.loc[idx, '_identifier_temp']
        if ident in note_map:
            final_df.loc[idx, note_col] = f"SUCCESS from: {note_map[ident]}"

    # Clean up temp column
    final_df = final_df.drop(columns=['_identifier_temp'])

    columns_to_drop = ['_identifier', '_result_level', '_source_sheet', '_source_file']
    final_df = final_df.drop(columns=[c for c in columns_to_drop if c in final_df.columns])
    final_df.reset_index(drop=True, inplace=True)

    # ==========================================
    # CLEANUP: Clear junk data from specific columns
    # ==========================================
    valid_results = {'NONE', 'SUCCESS', 'INCONCLUSIVE', 'FAILED', 'ERROR'}

    # Columns to always clear completely
    for col in final_df.columns:
        if any(k.lower() in str(col).lower() for k in ['issafety', 'l1 req', 'l4 req']):
            final_df[col] = None
            print(f"  🧹 Cleared column: {col}")

    # Reclassification: keep only valid result values, clear everything else
    for col in final_df.columns:
        if 'reclassification' in str(col).lower():
            mask = final_df[col].notna() & ~final_df[col].isin(valid_results)
            cleared = mask.sum()
            final_df.loc[mask, col] = None
            if cleared > 0:
                print(f"  🧹 Cleared {cleared} invalid values from: {col}")
            break

    print(f"✓ Final unique rows: {len(final_df)}")

    result_counts_dict = final_df[col_Result].value_counts().to_dict()
    total_tests = len(final_df)

    # ==========================================
    # BUILD OUTPUT WORKBOOK
    # ==========================================
    wb_out = openpyxl.Workbook()
    ws = wb_out.active
    ws.title = "Results"

    # --- Header at row 13 (starting column B) ---
    header_row = 13
    headers = list(final_df.columns)
    first_col = 2  # Column B
    last_col = len(headers) + 1

    for col_idx, header in enumerate(headers):
        cell = ws.cell(row=header_row, column=col_idx + first_col)
        cell.value = header
        cell.font = Font(name='Calibri', size=11)

    # --- Data starting at row 14 ---
    data_start_row = 14
    for row_idx in range(len(final_df)):
        for col_idx, col_name in enumerate(headers):
            value = final_df.iloc[row_idx][col_name]
            if pd.isna(value):
                value = None
            cell = ws.cell(row=data_start_row + row_idx, column=col_idx + first_col)
            cell.value = value
            cell.font = Font(name='Calibri', size=11)

    last_data_row = data_start_row + len(final_df) - 1
    if last_data_row < data_start_row:
        last_data_row = data_start_row

    # --- Find column positions ---
    result_col_excel = headers.index(col_Result) + first_col
    result_col_letter = get_column_letter(result_col_excel)

    reclass_col_name = None
    reclass_col_excel = None
    reclass_col_letter = None
    for h in headers:
        if 'reclassification' in str(h).lower():
            reclass_col_name = h
            reclass_col_excel = headers.index(h) + first_col
            reclass_col_letter = get_column_letter(reclass_col_excel)
            break

    version_col_excel = None
    version_col_letter = None
    for h in headers:
        if 'version' in str(h).lower():
            version_col_excel = headers.index(h) + first_col
            version_col_letter = get_column_letter(version_col_excel)
            break

    autofill_col_excel = None
    autofill_col_letter = None
    for h in headers:
        if 'autofill' in str(h).lower():
            autofill_col_excel = headers.index(h) + first_col
            autofill_col_letter = get_column_letter(autofill_col_excel)
            break

    # --- Summary area (rows 1-12) with FORMULAS referencing AutoFilled column ---
    if autofill_col_letter:
        write_summary_area(ws, autofill_col_letter, data_start_row, last_data_row)
    else:
        # Fallback: use Result column if AutoFilled not found
        write_summary_area(ws, result_col_letter, data_start_row, last_data_row)

    # --- Bold the Result column ---
    for row_idx in range(len(final_df)):
        cell = ws.cell(row=data_start_row + row_idx, column=result_col_excel)
        cell.font = Font(name='Calibri', size=11, bold=True)

    # =============================================
    # TABLE (for filter dropdowns only, no coloring)
    # =============================================
    last_col_letter = get_column_letter(last_col)
    table_ref = f"B{header_row}:{last_col_letter}{last_data_row}"

    table = Table(displayName="MergedResults", ref=table_ref)
    # Use "None" style so table doesn't override our manual colors
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleLight1",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=False,       # We handle stripes manually
        showColumnStripes=False
    )
    ws.add_table(table)

    # =============================================
    # MANUAL GRAY TABLE FORMATTING (theme-safe)
    # =============================================
    print("\n🎨 Applying gray table formatting...")
    apply_manual_table_formatting(ws, header_row, data_start_row, last_data_row, first_col, last_col)

    # =============================================
    # CONDITIONAL FORMATTING (colors + bold font)
    # =============================================
    print("🎨 Applying conditional formatting...")

    # Result + Version columns (G:H)
    if version_col_letter:
        gh_range = f"{result_col_letter}{data_start_row}:{version_col_letter}{last_data_row}"
    else:
        gh_range = f"{result_col_letter}{data_start_row}:{result_col_letter}{last_data_row}"

    cf_rules_gh = [
        ('ERROR',        RESULT_COLORS['ERROR'],        7),
        ('FAILED',       RESULT_COLORS['FAILED'],       9),
        ('INCONCLUSIVE', RESULT_COLORS['INCONCLUSIVE'], 11),
        ('SUCCESS',      RESULT_COLORS['SUCCESS'],      13),
        ('NONE',         RESULT_COLORS['NONE'],         15),
    ]
    for value, color, priority in cf_rules_gh:
        fill = PatternFill(bgColor=color)
        bold = Font(bold=True)
        rule = CellIsRule(operator='equal', formula=[f'"{value}"'], fill=fill, font=bold)
        rule.priority = priority
        ws.conditional_formatting.add(gh_range, rule)

    # Reclassification column (J) + Evaluation cell (C5) — SAME format as Result
    if reclass_col_letter:
        reclass_range = f"C5 {reclass_col_letter}{data_start_row}:{reclass_col_letter}{last_data_row}"
        cf_rules_j = [
            ('ERROR',        RESULT_COLORS['ERROR'],        1),
            ('FAILED',       RESULT_COLORS['FAILED'],       2),
            ('INCONCLUSIVE', RESULT_COLORS['INCONCLUSIVE'], 3),
            ('SUCCESS',      RESULT_COLORS['SUCCESS'],      4),
            ('NONE',         RESULT_COLORS['NONE'],         5),
        ]
        for value, color, priority in cf_rules_j:
            fill = PatternFill(bgColor=color)
            bold = Font(bold=True)
            rule = CellIsRule(operator='equal', formula=[f'"{value}"'], fill=fill, font=bold)
            rule.priority = priority
            ws.conditional_formatting.add(reclass_range, rule)

    # AutoFilled column — same formatting as Result
    if autofill_col_letter:
        af_range_cf = f"{autofill_col_letter}{data_start_row}:{autofill_col_letter}{last_data_row}"
        cf_rules_af = [
            ('ERROR',        RESULT_COLORS['ERROR'],        16),
            ('FAILED',       RESULT_COLORS['FAILED'],       17),
            ('INCONCLUSIVE', RESULT_COLORS['INCONCLUSIVE'], 18),
            ('SUCCESS',      RESULT_COLORS['SUCCESS'],      19),
            ('NONE',         RESULT_COLORS['NONE'],         20),
        ]
        for value, color, priority in cf_rules_af:
            fill = PatternFill(bgColor=color)
            bold = Font(bold=True)
            rule = CellIsRule(operator='equal', formula=[f'"{value}"'], fill=fill, font=bold)
            rule.priority = priority
            ws.conditional_formatting.add(af_range_cf, rule)

    # =============================================
    # DATA VALIDATION (dropdown on Reclassification)
    # =============================================
    if reclass_col_letter:
        print("📋 Adding dropdown to Reclassification column...")
        dv = DataValidation(
            type="list",
            formula1="$E$5:$E$9",
            allow_blank=True,
            showDropDown=False  # False = SHOW dropdown (Excel convention inverted)
        )
        dv.prompt = "Select a result"
        dv.promptTitle = "Reclassification"
        dv.sqref = f"{reclass_col_letter}{data_start_row}:{reclass_col_letter}{last_data_row}"
        ws.add_data_validation(dv)

    # --- Formula in AutoFilled column ---
    if autofill_col_excel and reclass_col_letter:
        for row_idx in range(len(final_df)):
            excel_row = data_start_row + row_idx
            formula = f"=IF(ISBLANK({reclass_col_letter}{excel_row}),{result_col_letter}{excel_row},{reclass_col_letter}{excel_row})"
            ws.cell(row=excel_row, column=autofill_col_excel).value = formula

    # --- Auto-fit columns ---
    print("📐 Auto-fitting columns...")
    auto_fit_columns(ws, min_width=13, max_width=50, start_row=header_row)
    ws.column_dimensions['B'].width = max(ws.column_dimensions['B'].width, 20)

    # Row heights
    ws.row_dimensions[header_row].height = 15
    for r in range(data_start_row, last_data_row + 1):
        ws.row_dimensions[r].height = 15

    # --- Save ---
    wb_out.save(output_path)
    wb_out.close()

    print("  ✓ Done!")

    # Result distribution
    print(f"\nResult distribution:")
    print("-" * 60)
    for result, count in final_df[col_Result].value_counts().items():
        pct = (count / len(final_df)) * 100
        print(f"  {str(result):15s}: {count:5d} ({pct:5.1f}%)")
    print("-" * 60)

    return final_df


if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))

    input_patterns = [
        os.path.join(script_dir, "*.xlsm"),
        os.path.join(script_dir, "*.xlsx"),
    ]

    output_file = os.path.join(script_dir, "Merged.xlsx")

    input_files = []
    for pattern in input_patterns:
        for f in glob.glob(pattern):
            if os.path.abspath(f) != os.path.abspath(output_file):
                input_files.append(f)

    input_files = sorted(set(input_files))

    print(f"Script location: {script_dir}")
    print(f"Output file:     Merged.xlsx")
    print("=" * 60)

    if not input_files:
        print(f"\n✗ No Excel files (.xlsm, .xlsx) found!")
        print(f"\nPlace your Excel files in the same folder as this script:")
        print(f"  {script_dir}")
        print("\n  your_folder/")
        print("  ├── mergerp.py")
        print("  ├── File1.xlsm")
        print("  ├── File2.xlsx")
        print("  └── Merged.xlsx  (auto-created)")
    else:
        print(f"\n📋 Found {len(input_files)} input file(s):")
        for f in input_files:
            print(f"  • {os.path.basename(f)}")
        print()

        try:
            merged_data = merge_excel_sheets(input_files, output_file)
            print("\n" + "=" * 60)
            print("         PROCESS COMPLETED SUCCESSFULLY!")
            print("=" * 60)
        except Exception as e:
            print(f"\n✗ Error: {str(e)}")
            import traceback
            traceback.print_exc()
