"""
matrix_parser.py - Parser for ComSCL NetMatrix Excel files (e.g., comscl_netmtrx_rbands.xlsx)
Extracts signal-to-port mappings, bit positions, conversion types, and default values.
"""

from typing import Dict, Any, Optional
import openpyxl

MACRO_TABLE: Dict[str, int] = {
    "MAXUINT8": 255,
    "MAXUINT16": 65535,
    "MAXSINT16": 32767,
    "MAXSINT32": 2147483647,
    "MAXUINT32": 4294967295
}


def parse_matrix_excel(excel_path: str) -> Dict[str, Dict[str, Any]]:
    """
    Parses the SignalMapping sheet from the given Excel file.
    
    Returns a dictionary mapping '<SignalName>__<SignalPortName>' to:
    {
        'bus': str,
        'pdu': str,
        'sig': str,
        'direction': str,
        'port': str,
        'bitpos': Optional[int],
        'conv': Optional[str],
        'defval': Optional[int]
    }
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    if "SignalMapping" not in wb.sheetnames:
        raise ValueError(f"Sheet 'SignalMapping' not found in {excel_path}. Available sheets: {wb.sheetnames}")
    
    ws = wb["SignalMapping"]
    signals: Dict[str, Dict[str, Any]] = {}
    
    # We find headers dynamically or assume standard layout (Row 3 has column names)
    col_bus = 0
    col_pdu = 1
    col_sig = 2
    col_dir = 3
    col_port = 4
    col_bitpos = 5
    col_conv = 7
    col_defval = 14

    for r_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
        if r_idx < 4:
            # Header search fallback
            row_strs = [str(c or "").strip().lower() for c in row]
            if "signalname" in row_strs:
                for c_i, name in enumerate(row_strs):
                    if name == "busname": col_bus = c_i
                    elif name == "pduname": col_pdu = c_i
                    elif name == "signalname": col_sig = c_i
                    elif name == "signaldirection": col_dir = c_i
                    elif name == "signalportname": col_port = c_i
                    elif name == "signalportbitpos": col_bitpos = c_i
                    elif name == "signalportconversion": col_conv = c_i
                    elif name == "signalportdefaultvalue": col_defval = c_i
            continue
            
        direction = str(row[col_dir] or "").strip()
        if direction.lower() != "receive":
            continue
            
        sig = str(row[col_sig] or "").strip()
        port = str(row[col_port] or "").strip()
        if not sig or not port:
            continue
            
        bus = str(row[col_bus] or "").strip()
        pdu = str(row[col_pdu] or "").strip()
        
        # BitPos
        bitpos_val: Optional[int] = None
        raw_bp = row[col_bitpos] if col_bitpos < len(row) else None
        if raw_bp is not None and str(raw_bp).strip() != "":
            try:
                bitpos_val = int(raw_bp)
            except (ValueError, TypeError):
                pass
                
        # Conversion
        raw_conv = row[col_conv] if col_conv < len(row) else None
        conv_val = str(raw_conv or "").strip() or None
        
        # DefaultValue
        defval_val: Optional[int] = None
        raw_dv = row[col_defval] if col_defval < len(row) else None
        if raw_dv is not None and str(raw_dv).strip() != "":
            dv_str = str(raw_dv).strip()
            if dv_str in MACRO_TABLE:
                defval_val = MACRO_TABLE[dv_str]
            else:
                try:
                    defval_val = int(float(dv_str))
                except (ValueError, TypeError):
                    pass
                    
        key = f"{sig}__{port}"
        signals[key] = {
            "bus": bus,
            "pdu": pdu,
            "sig": sig,
            "direction": direction,
            "port": port,
            "bitpos": bitpos_val,
            "conv": conv_val,
            "defval": defval_val
        }
        
    return signals
