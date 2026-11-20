# ECU-TEST Project File Matrix Validator — Architecture & Extension Guide

## 1. Executive Summary & Purpose

The **ECU-TEST PRJ Matrix Validator** is a specialized desktop automation and verification tool designed for automotive Software-in-the-Loop (SiL) test engineering.

### The Problem It Solves
When generating ECU-TEST project files (`.prj`) from automotive network matrices (such as VW ComSCL / K-Matrix Excel sheets like `comscl_netmtrx_rbands.xlsx`), automated generator toolchains frequently suffer from change propagation gaps:
1. **Omitted Bit-Position Parameters**: Bit-packed receive signals mapped to multi-bit ECU ports (e.g., `AB_Stellgliedtest` mapped to bit 5 of `Com_stAirBg`) are generated with `P_SignalPortConversion = 'Value_Direct'`, but `P_SignalPortBitPos` is omitted (defaulting to Python `None`). During test execution, ECU-TEST reads the full unmasked integer port value (`32.0`) and evaluates `1.0 == 32.0`, causing test failure.
2. **Omitted Default Replacement Values**: During PDU timeout and diagnosis fadeout evaluation (`PortRxErrHndlgPduTout`), the ECU substitutes lost signals with predefined replacement error patterns (e.g., `3` for High-Voltage Interlock errors, `15` for Operation Mode). The test package asserts `MsgVal_ref == Replacement_ref`, where `Replacement_ref` is dynamically bound to `P_SignalPortDefaultValue`. When this parameter is omitted from the `.prj`, it defaults to `0`, causing valid ECU replacement behavior to fail (`3.0 != 0`).
3. **SMB / Network Share Latency Trap**: Large ECU-TEST project files (8–10+ MB XML trees) hosted on enterprise UNC network shares (`\\fe00fx21...`) cannot be safely edited via iterative in-place XML writers; small-buffer network latency leads to process deadlocks or file truncation.

This tool provides a graphical and scriptable audit engine that verifies 100% of signals against the Excel matrix ground truth and applies atomic parameter injections safely.

---

## 2. High-Level Architecture & Layered Diagram

```text
+-----------------------------------------------------------------------------------+
|                              GUI Layer (app.py)                                   |
|  - Tkinter / TTK Desktop Window with Clam Theme                                   |
|  - File Dialog Pickers (Excel Matrix .xlsx, Target Project .prj)                  |
|  - Thread-Safe Queue (queue.Queue) & Event Loop Polling (root.after)              |
|  - Syntax-Highlighted Consolas Terminal View & Progress Indicators                |
+-----------------------------------------------------------------------------------+
                                  |
                                  | Dispatches async background tasks
                                  v
+-----------------------------------------------------------------------------------+
|                      Matrix Parser Layer (matrix_parser.py)                       |
|  - openpyxl streaming read of 'SignalMapping' sheet                               |
|  - Dynamic header index discovery (BusName, PduName, SignalName, Port, etc.)      |
|  - Macro Resolution: MAXUINT8 (255), MAXUINT16 (65535), MAXSINT16 (32767), etc.   |
|  - Returns: Dict[<SignalName>__<PortName>, SignalMetadata]                        |
+-----------------------------------------------------------------------------------+
                                  |
                                  | Feeds ground-truth matrix data
                                  v
+-----------------------------------------------------------------------------------+
|                        PRJ Engine Layer (prj_engine.py)                           |
|                                                                                   |
|  [Audit Mode]:                                                                    |
|    - Streams XML DOM via xml.etree.ElementTree                                    |
|    - Maps <subproject> elements to matrix keys                                    |
|    - Audits <packageTest> parameter sets for target test suites                   |
|    - Returns Discrepancy List: Missing / Mismatched parameters                     |
|                                                                                   |
|  [Atomic Injection Mode]:                                                         |
|    1. Network Backup: Creates `<prj_path>.bak_<timestamp>` on network share       |
|    2. Local SSD Staging: Copies network file to %TEMP%\staged_prj_<filename>      |
|    3. XML DOM Manipulation: Injects <PARAMETER> and updates <VALUE> nodes        |
|    4. Local Schema Validation: Parses modified file with ET.parse to guarantee 100%|
|       well-formed XML without declaration header (format-rev="10" / "11")          |
|    5. Atomic Deployment: Overwrites target UNC file in a single fast block copy   |
+-----------------------------------------------------------------------------------+
```

---

## 3. Directory Structure

```text
C:\Workspace\14_Analyze_RP_Trellis_AI\
├── prj_matrix_validator\
│   ├── __init__.py               # Package marker
│   ├── app.py                    # Tkinter GUI entrypoint
│   ├── matrix_parser.py          # Excel K-Matrix parsing and macro resolution
│   ├── prj_engine.py             # XML auditing and atomic SSD injection engine
│   └── ARCHITECTURE.md           # Architectural documentation (this file)
├── tests\
│   └── test_validator.py         # Unit tests for parser and PRJ engine
└── run_validator_gui.bat         # 1-click Windows desktop launcher
```

---

## 4. Key Component Responsibilities

### 4.1 `matrix_parser.py`
- **Data Source**: Excel file (`.xlsx`), specifically the sheet named `SignalMapping`.
- **Dynamic Header Discovery**: Scans initial rows for `SignalName`, `SignalPortName`, `SignalPortBitPos`, `SignalPortConversion`, and `SignalPortDefaultValue` columns. This ensures backwards and forwards compatibility if matrix column order changes.
- **Direction Filtering**: Filters strictly for `SignalDirection == 'receive'`.
- **Macro Translation Table**:
  ```python
  MACRO_TABLE = {
      "MAXUINT8": 255,
      "MAXUINT16": 65535,
      "MAXSINT16": 32767,
      "MAXSINT32": 2147483647,
      "MAXUINT32": 4294967295
  }
  ```
- **Output Contract**: Returns a dictionary keyed by `f"{SignalName}__{SignalPortName}"`.

### 4.2 `prj_engine.py`
- **Namespace Handling**: Registers `xsi = "http://www.w3.org/2001/XMLSchema-instance"` to ensure clean attribute output (`xsi:type="parameterEntry"`) without spurious XML prefixes (`ns0:`).
- **Target Test Suites**:
  - `PACKAGES_FOR_BITPOS`: Packages that assert port bit positions against single-bit bus stimulation (`PortRxConvRaw`, `PortRxErrHndlgPduTout_*`, `PortRxErrHndlgDlc_*`, `PortRxErrHndlgPduDi`, `PortRxErrHndlgAlvCntr_*`, `PortRxErrHndlgChks_*`, `PortRxErrHndlgLinResp_*`).
  - `PACKAGES_FOR_DEFVAL`: Packages that evaluate port replacement values during error injection (`PortRxErrHndlgPduTout_*`, `PortRxErrHndlgDlc_*`, `PortRxErrHndlgPduDi`, `PortRxStSigPat*`).
- **Atomic Local Staging**:
  ```python
  # Staged to %TEMP% on local fast SSD
  local_staged_prj = os.path.join(tempfile.gettempdir(), f"staged_prj_{basename}")
  shutil.copy2(prj_path, local_staged_prj)
  ```
  Prevents network latency bottlenecks and guarantees that if a process is cancelled, the network file remains untouched.

### 4.3 `app.py`
- **Thread Safety**: Long operations (parsing large Excel files, reading 9MB XML trees, copying over SMB) execute in background `threading.Thread(daemon=True)` workers.
- **Queue-Based Logging**: Worker threads post formatted messages to a `queue.Queue`. The GUI main thread polls the queue every 100ms (`root.after(100, ...)`), preventing UI lockups.
- **Color-Coded Status Output**:
  - `highlight`: Yellow section headers.
  - `info`: Cyan progress indicators.
  - `warning`: Orange alert messages.
  - `error`: Red discrepancy line items.
  - `success`: Green completion and verification notifications.

---

## 5. How to Extend the Tool

### 5.1 Adding a New Package to Audit/Injection Lists
To add a new ECU-TEST test package to receive `P_SignalPortBitPos` or `P_SignalPortDefaultValue`:
1. Open `prj_matrix_validator/prj_engine.py`.
2. Locate `PACKAGES_FOR_BITPOS` or `PACKAGES_FOR_DEFVAL`.
3. Add the package name:
   ```python
   PACKAGES_FOR_BITPOS.add("MyNewCustomPackageTest")
   ```

### 5.2 Adding a New Parameter (e.g. `P_SignalPortConversionFactor`)
1. **In `matrix_parser.py`**:
   - Add column index discovery for `SignalPortConversionFactor`.
   - Add `"factor": factor_val` to the returned dictionary.
2. **In `prj_engine.py`**:
   - In `audit_project()`: Add a check comparing `params.get("P_SignalPortConversionFactor")` against `matrix_info["factor"]`.
   - In `apply_fixes()`: Add injection block for `P_SignalPortConversionFactor`.

### 5.3 Headless / CI Pipeline Execution
The engine can be invoked programmatically without the GUI:
```python
from prj_matrix_validator.matrix_parser import parse_matrix_excel
from prj_matrix_validator.prj_engine import audit_project, apply_fixes

# 1. Parse Excel Matrix
matrix = parse_matrix_excel("path/to/matrix.xlsx")

# 2. Audit
discrepancies = audit_project("path/to/project.prj", matrix)
print(f"Found {len(discrepancies)} discrepancies")

# 3. Apply fixes if needed
if discrepancies:
    stats = apply_fixes("path/to/project.prj", matrix)
    print("Fixes applied:", stats)
```

---

## 6. Verification and Testing

Run the automated test suite from the repository root:
```bash
python -m unittest tests/test_validator.py
```
Test suite coverage:
- Macro translation dictionary accuracy.
- Mock project XML auditing (verifying missing parameter detection).
- Mock project parameter injection (verifying clean XML generation and discrepancy resolution).
