"""
DCM Parser for SiL Calibration Tool.
Reads existing DCM files to extract already-calibrated variables and their values.
"""

import re
from typing import Dict, Set


def parse_existing_dcm_variables(dcm_path: str) -> Dict[str, float]:
    """
    Parses a DCM file and returns a dictionary of {variable_name: calibrated_value}.
    Supports FESTWERT and general DCM parameter declarations.
    """
    existing_vars = {}
    current_var = None
    
    var_start_pattern = re.compile(r"^\s*(?:FESTWERT|GRUPPENWERT|KENNFELD|KENNLINIE)\s+([A-Za-z0-9_\.]+)", re.IGNORECASE)
    wert_pattern = re.compile(r"^\s*WERT\s+([+-]?[0-9]*\.?[0-9]+(?:[eE][+-]?[0-9]+)?)", re.IGNORECASE)
    
    with open(dcm_path, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            start_match = var_start_pattern.match(line)
            if start_match:
                current_var = start_match.group(1).strip()
                continue
                
            if current_var:
                wert_match = wert_pattern.match(line)
                if wert_match:
                    try:
                        existing_vars[current_var] = float(wert_match.group(1))
                    except ValueError:
                        existing_vars[current_var] = 0.0
                    current_var = None
                    continue
                    
            if line.strip().upper() == "END":
                current_var = None

    return existing_vars
