# SiL Software Calibration Merger & Validator Tool

A dedicated desktop utility for Software-in-the-Loop (SiL) engineers to prepare, validate, and bake permanent communication and diagnostic calibration datasets into Synopsys Silver virtual ECUs (`Modified_Calibrations.dcm`).

---

## 1. Features & Architecture

- **Universal Channel Auto-Discovery**: Dynamically detects any communication cluster name from the A2L file (e.g. `HVCAN`, `ACANFD`, `BCAN`, `LIN1`, `MotorSubcan`, `SES_LIN1`) by parsing `Com_swtPduCodActv<BUS>_C`. No hardcoding of channel names.
- **Conflict & Overlap Avoidance**: Reads the target `Modified_Calibrations.dcm` and automatically excludes variables that are already calibrated, ensuring no duplicate entries or conflicting values.
- **Interactive Validation Panel**: Displays all proposed calibrations in a categorized table with checkboxes (`Add?`), target values, and rule descriptions.
- **Safety First**: Automatically generates a timestamped backup (`Modified_Calibrations.dcm.<timestamp>.bak`) before modifying the DCM file.
- **Docs-as-Code & Audit Logging**:
  - Export validation review reports directly to `.doc` or `.txt`.
  - Generates detailed execution audit logs (`calibration_merge_log_<timestamp>.txt`) detailing all added variables, values, and rules.

---

## 2. Implemented Calibration Rules

| Rule Pattern | Discovered Scope | Value | Target BSW Module | Purpose |
| :--- | :--- | :---: | :---: | :--- |
| `*swtPduCodActv = 0` | All channels (`Com_swtPduCodActv<BUS>_C`) | `0.0` | `comcil_diag_auto` | Deactivate variant coding restrictions; accept all frames |
| `Com_idxDiagDi*_C = 2` | Global bus index (`Com_idxDiagDi<BUS>_C`) | `2.0` | `comcil_diag_auto` | Enable global network bus-off and timeout diagnostics |
| `Com_idxDiagDi<BUS><PDU>_C = 6` | Every PDU on all clusters | `6.0` | `comcil_diag_auto` | Enable full PDU timeout, checksum, and counter diagnostics |
| `Com_swt*xena*_C = 1` | `Com_swtRxEna*` and `Com_swtTxEna*` | `1.0` | `comcil_diag_auto` | Gate on all incoming Rx and outgoing Tx communication frames |
| `DFC_CtlMsk2.DFC_Com*_C = 1` | All communication DFCs | `1.0` | `DFC` | Activate background diagnostic fault monitoring |
| `DFC_DisblMsk2.DFC_Com*_C = 0` | All communication DFCs | `0.0` | `DFC` | Clear disable mask; permit defect qualification into memory |

---

## 3. Directory Layout & How to Launch

### Folder Layout
```text
C:\Workspace\25_CalibToolforSIL\
├── SiL_CalibTool.exe         <-- Standalone Windows executable (No Python required!)
├── README.md                 <-- User manual
└── source\                   <-- Full Python source code
    ├── main.py               <-- Tkinter GUI frontend
    ├── a2l_parser.py         <-- Dynamic channel & A2L parser
    ├── dcm_parser.py         <-- DCM dataset parser
    ├── dcm_writer.py         <-- DCM formatter, backup & logger
    ├── excel_exporter.py     <-- Multi-sheet Excel workbook generator (.xlsx)
    ├── test_tool.py          <-- Automated test suite
    ├── run_tool.bat          <-- Batch launcher for Python source
    └── build_exe.bat         <-- One-click builder to recompile SiL_CalibTool.exe
```

### Quick Launch Options
- **Option 1 (Standalone Executable - Recommended)**:
  Simply double-click **`SiL_CalibTool.exe`** in the root folder. It runs independently without needing Python or external libraries.
- **Option 2 (From Source)**:
  Navigate to `source\` and double-click `run_tool.bat`, or run `python main.py`.

### Step-by-Step Workflow
1. **Specify Paths**:
   - `Target Modified_Calibrations.dcm`: Path to your simulation overlay DCM file.
   - `Source vECU.a2l`: Path to your compiled virtual ECU A2L file.
2. **Click "🔍 Analyze & Cross-Check"**:
   - The tool scans `vECU.a2l` and discovers all communication channels.
   - It extracts all matching calibrations (Rule 1 to 6).
   - It cross-checks against `Modified_Calibrations.dcm` and displays only the **missing** calibrations.
3. **Review & Validate**:
   - Double-click any row to toggle its `Add?` state (or use `Select All` / `Deselect All`).
   - Use the `Filter` box to search for specific PDUs or DFCs.
4. **Export / Mass-Copy / Approve**:
   - **`📊 Export to Excel (.xlsx)`**: Generates a professional multi-sheet Excel workbook:
     - `Calibration_Table`: Full table with AutoFilter enabled for fast filtering and row copying.
     - `Mass_Copy_Variables`: Single clean column of plain variable names; click column header A to mass-copy all names with Ctrl+C.
     - `ECU_TEST_Snippets`: Pre-formatted Python testcase calls (`BaseLIB_Sys_Apl_SetCalibValFromName(...)`).
     - `Summary`: Audit metadata and category counts.
   - **`📋 Copy Names` & Right-Click Menu**: Mass-copy variable names, Python calls, or DCM blocks directly to your clipboard.
   - **`📄 Export Document (.doc/.txt)`**: Generates a formal validation document for team review.
   - **`✅ Approve & Append to DCM`**: Safely appends the calibrations and generates the audit log.
