"""
A2L Parser for SiL Calibration Tool.
Extracts calibration variables matching predefined communication and DFC rules,
using dynamic bus discovery to support any channel (HVCAN, ACANFD, LIN, etc.).
"""

import re
from typing import Dict, List, Set, Tuple


class CalibrationItem:
    def __init__(self, name: str, value: float, rule_desc: str, function_group: str = "comcil_diag_auto"):
        self.name = name
        self.value = value
        self.rule_desc = rule_desc
        self.function_group = function_group

    def __repr__(self):
        return f"CalibrationItem({self.name}, {self.value}, {self.rule_desc})"


def extract_characteristics_from_a2l(a2l_path: str) -> Set[str]:
    """
    Stream-reads an A2L file and extracts all CHARACTERISTIC names.
    Fast and memory-efficient for large A2L files (50MB+).
    """
    characteristics = set()
    char_pattern = re.compile(r"^\s*/begin\s+CHARACTERISTIC\s+([A-Za-z0-9_\.]+)", re.IGNORECASE)
    
    with open(a2l_path, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            match = char_pattern.match(line)
            if match:
                characteristics.add(match.group(1).strip())
                continue
            
            # Also catch standard layout where name is on the next line after /begin CHARACTERISTIC
            if "/begin CHARACTERISTIC" in line:
                for next_line in f:
                    stripped = next_line.strip()
                    if stripped and not stripped.startswith("/*"):
                        characteristics.add(stripped.split()[0])
                        break
                        
    return characteristics


def discover_buses(characteristics: Set[str]) -> Set[str]:
    """
    Discovers all bus/cluster names by finding Com_swtPduCodActv<BUS>_C.
    Supports any channel name dynamically (HVCAN, ACANFD, LIN, MotorSubcan, etc.).
    """
    buses = set()
    pattern = re.compile(r"^Com_swtPduCodActv([A-Za-z0-9_]+)_C$")
    for name in characteristics:
        m = pattern.match(name)
        if m:
            buses.add(m.group(1))
    return buses


def extract_needed_calibrations(a2l_path: str) -> Tuple[List[CalibrationItem], Set[str]]:
    """
    Extracts all variables matching the 6 calibration rules from vECU.a2l:
    1. *swtPduCodActv = 0
    2. Com_idxDiagDi<BUS>_C = 2 (Network level)
    3. Com_idxDiagDi<BUS><PDU>_C = 6 (PDU level)
    4. Com_swt*xena*_C = 1 (Rx and Tx enables)
    5. DFC_CtlMsk2.DFC_Com*_C = 1
    6. DFC_DisblMsk2.DFC_Com*_C = 0
    """
    characteristics = extract_characteristics_from_a2l(a2l_path)
    buses = discover_buses(characteristics)
    
    items = []
    
    # Sort for deterministic output
    for name in sorted(characteristics):
        # Rule 1: *swtPduCodActv = 0
        if re.search(r"Com_swtPduCodActv.*_C$", name):
            items.append(CalibrationItem(name, 0.0, "PDU Coding Deactivation (*swtPduCodActv=0)", "comcil_diag_auto"))
            continue

        # Rule 2 & 3: Com_idxDiagDi*_C
        idx_match = re.match(r"^Com_idxDiagDi([A-Za-z0-9_]+)_C$", name)
        if idx_match:
            suffix = idx_match.group(1)
            # If suffix exactly matches a bus name -> Network Level (Value = 2)
            if suffix in buses:
                items.append(CalibrationItem(name, 2.0, "Network Diagnostic Index (Com_idxDiagDi*_C=2)", "comcil_diag_auto"))
            # Otherwise -> PDU Level (Value = 6)
            else:
                items.append(CalibrationItem(name, 6.0, "PDU Diagnostic Index (Com_idxDiagDi<BUS><PDU>_C=6)", "comcil_diag_auto"))
            continue

        # Rule 4: Com_swt*xena*_C = 1 (Com_swtRxEna / Com_swtTxEna)
        if re.match(r"^Com_swt[RT]xEna.*_C$", name):
            items.append(CalibrationItem(name, 1.0, "PDU Rx/Tx Reception/Transmission Enable (Com_swt*xena*_C=1)", "comcil_diag_auto"))
            continue

        # Rule 5: DFC_CtlMsk2.DFC_Com*_C = 1
        if re.match(r"^DFC_CtlMsk2\.DFC_Com.*_C$", name):
            items.append(CalibrationItem(name, 1.0, "DFC Control Mask (DFC_CtlMsk2.DFC_Com*_C=1)", "DFC"))
            continue

        # Rule 6: DFC_DisblMsk2.DFC_Com*_C = 0
        if re.match(r"^DFC_DisblMsk2\.DFC_Com.*_C$", name):
            items.append(CalibrationItem(name, 0.0, "DFC Disable Mask Unmask (DFC_DisblMsk2.DFC_Com*_C=0)", "DFC"))
            continue

    return items, buses
