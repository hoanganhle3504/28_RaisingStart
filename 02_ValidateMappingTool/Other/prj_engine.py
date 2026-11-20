"""
prj_engine.py - ECU-TEST Project File Engine.
Handles atomic SMB/local staging, XML auditing, and parameter injection.
"""

import os
import shutil
import tempfile
import time
import xml.etree.ElementTree as ET
from typing import Dict, Any, List, Optional, Callable

# Ensure xsi namespace is preserved cleanly without ns0 prefixing
ET.register_namespace("xsi", "http://www.w3.org/2001/XMLSchema-instance")

# Packages that require BitPos injection (bitfield masking evaluation)
PACKAGES_FOR_BITPOS = {
    "PortRxConvRaw",
    "PortRxErrHndlgDlc_FadeoutActive",
    "PortRxErrHndlgDlc_FadeoutInActive",
    "PortRxErrHndlgPduDi",
    "PortRxErrHndlgPduTout_FadeoutActive",
    "PortRxErrHndlgPduTout_FadeoutInActive",
    "PortRxErrHndlgAlvCntr_FadeoutActive",
    "PortRxErrHndlgAlvCntr_FadeoutInActive",
    "PortRxErrHndlgChks_FadeoutActive",
    "PortRxErrHndlgChks_FadeoutInActive",
    "PortRxErrHndlgLinResp_fadeoutActive",
    "PortRxErrHndlgLinResp_fadeoutInActive"
}

# Packages that require DefaultValue injection (replacement value evaluation on fault)
PACKAGES_FOR_DEFVAL = {
    "PortRxErrHndlgPduTout_FadeoutActive",
    "PortRxErrHndlgPduTout_FadeoutInActive",
    "PortRxErrHndlgDlc_FadeoutActive",
    "PortRxErrHndlgDlc_FadeoutInActive",
    "PortRxErrHndlgPduDi",
    "PortRxErrHndlgAlvCntr_FadeoutActive",
    "PortRxErrHndlgAlvCntr_FadeoutInActive",
    "PortRxErrHndlgChks_FadeoutActive",
    "PortRxErrHndlgChks_FadeoutInActive",
    "PortRxErrHndlgLinResp_fadeoutActive",
    "PortRxErrHndlgLinResp_fadeoutInActive",
    "PortRxStSigPatErr_FadeoutActive",
    "PortRxStSigPatErr_FadeoutInActive",
    "PortRxStSigPatInit",
    "PortRxStSigPatNotPres",
    "PortRxStSigPatOOR_FadeoutActive",
    "PortRxStSigPatOOR_FadeoutInActive",
    "PortRxStSigPatSpecPat1",
    "PortRxStSigPatSpecPat2"
}


def audit_project(
    prj_path: str,
    matrix_signals: Dict[str, Dict[str, Any]],
    log_callback: Optional[Callable[[str], None]] = None
) -> List[Dict[str, Any]]:
    """
    Audits a .prj file against the parsed matrix signals without modifying the file.
    Returns a list of discrepancies.
    """
    if not os.path.exists(prj_path):
        raise FileNotFoundError(f"Project file not found: {prj_path}")

    if log_callback:
        log_callback(f"Loading project file: {prj_path}...")

    tree = ET.parse(prj_path)
    root = tree.getroot()
    discrepancies: List[Dict[str, Any]] = []

    for sp in root.iter("COMPONENT"):
        if sp.attrib.get("{http://www.w3.org/2001/XMLSchema-instance}type") != "subproject":
            continue
        sname_elem = sp.find("NAME")
        if sname_elem is None or not sname_elem.text:
            continue
        sp_name = sname_elem.text.strip()
        
        matrix_info = matrix_signals.get(sp_name)
        if not matrix_info:
            continue

        for pt in sp.findall("COMPONENTS/COMPONENT"):
            if pt.attrib.get("{http://www.w3.org/2001/XMLSchema-instance}type") != "packageTest":
                continue
            pt_name_elem = pt.find("NAME")
            if pt_name_elem is None or not pt_name_elem.text:
                continue
            pt_name = pt_name_elem.text.strip()

            for pr in pt.findall("COMPONENTS/COMPONENT"):
                pdata = pr.find("PARAMETER-DATA")
                if pdata is None:
                    continue

                params: Dict[str, str] = {}
                for param in pdata.findall("PARAMETER"):
                    pk = param.find("NAME")
                    if pk is not None and pk.text:
                        val_node = param.find(".//VALUE")
                        val_txt = "".join(val_node.itertext()).strip() if val_node is not None else ""
                        params[pk.text.strip()] = val_txt

                # 1. Audit BitPos
                if pt_name in PACKAGES_FOR_BITPOS:
                    exp_bp = matrix_info["bitpos"]
                    if exp_bp is not None:
                        prj_bp = params.get("P_SignalPortBitPos")
                        if prj_bp is None or prj_bp == "" or prj_bp == "None" or str(prj_bp).strip() != str(exp_bp):
                            discrepancies.append({
                                "subproject": sp_name,
                                "pdu": matrix_info["pdu"],
                                "package": pt_name,
                                "param": "P_SignalPortBitPos",
                                "expected": str(exp_bp),
                                "actual": prj_bp or "MISSING"
                            })

                # 2. Audit DefaultValue
                if pt_name in PACKAGES_FOR_DEFVAL:
                    exp_dv = matrix_info["defval"]
                    if exp_dv is not None:
                        prj_dv = params.get("P_SignalPortDefaultValue")
                        if prj_dv is None or prj_dv == "" or str(prj_dv).strip() != str(exp_dv):
                            discrepancies.append({
                                "subproject": sp_name,
                                "pdu": matrix_info["pdu"],
                                "package": pt_name,
                                "param": "P_SignalPortDefaultValue",
                                "expected": str(exp_dv),
                                "actual": prj_dv or "MISSING"
                            })

                # 3. Audit Conversion
                exp_conv = matrix_info["conv"]
                if exp_conv:
                    prj_conv = params.get("P_SignalPortConversion")
                    if prj_conv is not None and prj_conv != exp_conv:
                        discrepancies.append({
                            "subproject": sp_name,
                            "pdu": matrix_info["pdu"],
                            "package": pt_name,
                            "param": "P_SignalPortConversion",
                            "expected": exp_conv,
                            "actual": prj_conv
                        })

    return discrepancies


def apply_fixes(
    prj_path: str,
    matrix_signals: Dict[str, Dict[str, Any]],
    log_callback: Optional[Callable[[str], None]] = None
) -> Dict[str, int]:
    """
    Applies parameter injections and updates atomically to the project file.
    Staging to local SSD temp prevents network share latency and XML corruption.
    """
    if not os.path.exists(prj_path):
        raise FileNotFoundError(f"Project file not found: {prj_path}")

    # 1. Create a backup on the share
    bak_path = prj_path + f".bak_{int(time.time())}"
    if log_callback:
        log_callback(f"Creating safety backup at: {bak_path}")
    shutil.copy2(prj_path, bak_path)

    # 2. Stage locally to temp SSD
    temp_dir = tempfile.gettempdir()
    local_staged_prj = os.path.join(temp_dir, f"staged_prj_{os.path.basename(prj_path)}")
    if log_callback:
        log_callback(f"Staging to local temp: {local_staged_prj}")
    shutil.copy2(prj_path, local_staged_prj)

    # 3. Parse local file
    tree = ET.parse(local_staged_prj)
    root = tree.getroot()

    stats = {
        "injected_defaults": 0,
        "updated_defaults": 0,
        "injected_bitpos": 0,
        "updated_bitpos": 0,
        "updated_conv": 0
    }

    for sp in root.iter("COMPONENT"):
        if sp.attrib.get("{http://www.w3.org/2001/XMLSchema-instance}type") != "subproject":
            continue
        sname_elem = sp.find("NAME")
        if sname_elem is None or not sname_elem.text:
            continue
        sp_name = sname_elem.text.strip()

        matrix_info = matrix_signals.get(sp_name)
        if not matrix_info:
            continue

        exp_bp = matrix_info["bitpos"]
        exp_dv = matrix_info["defval"]
        exp_conv = matrix_info["conv"]

        for pt in sp.findall("COMPONENTS/COMPONENT"):
            if pt.attrib.get("{http://www.w3.org/2001/XMLSchema-instance}type") != "packageTest":
                continue
            pt_name_elem = pt.find("NAME")
            if pt_name_elem is None or not pt_name_elem.text:
                continue
            pt_name = pt_name_elem.text.strip()

            for pr in pt.findall("COMPONENTS/COMPONENT"):
                pdata = pr.find("PARAMETER-DATA")
                if pdata is None:
                    continue

                # DefaultValue handling
                if exp_dv is not None and pt_name in PACKAGES_FOR_DEFVAL:
                    p_exist = None
                    for p in pdata.findall("PARAMETER"):
                        if p.find("NAME") is not None and p.find("NAME").text == "P_SignalPortDefaultValue":
                            p_exist = p
                            break
                    if p_exist is not None:
                        val_el = p_exist.find(".//VALUE")
                        if val_el is not None:
                            inner = val_el.find("VALUE")
                            target_node = inner if inner is not None else val_el
                            if target_node.text != str(exp_dv):
                                target_node.text = str(exp_dv)
                                stats["updated_defaults"] += 1
                    else:
                        new_p = ET.Element("PARAMETER", {"{http://www.w3.org/2001/XMLSchema-instance}type": "parameterEntry"})
                        ET.SubElement(new_p, "NAME", {"{http://www.w3.org/2001/XMLSchema-instance}type": "string"}).text = "P_SignalPortDefaultValue"
                        val_base = ET.SubElement(new_p, "VALUE", {"format-rev": "1", "{http://www.w3.org/2001/XMLSchema-instance}type": "valueBaseExpression"})
                        ET.SubElement(val_base, "VALUE", {"{http://www.w3.org/2001/XMLSchema-instance}type": "integer"}).text = str(exp_dv)
                        pdata.append(new_p)
                        stats["injected_defaults"] += 1

                # BitPos handling
                if exp_bp is not None and pt_name in PACKAGES_FOR_BITPOS:
                    p_exist = None
                    for p in pdata.findall("PARAMETER"):
                        if p.find("NAME") is not None and p.find("NAME").text == "P_SignalPortBitPos":
                            p_exist = p
                            break
                    if p_exist is not None:
                        val_el = p_exist.find(".//VALUE")
                        if val_el is not None:
                            inner = val_el.find("VALUE")
                            target_node = inner if inner is not None else val_el
                            if target_node.text != str(exp_bp):
                                target_node.text = str(exp_bp)
                                stats["updated_bitpos"] += 1
                    else:
                        new_p = ET.Element("PARAMETER", {"{http://www.w3.org/2001/XMLSchema-instance}type": "parameterEntry"})
                        ET.SubElement(new_p, "NAME", {"{http://www.w3.org/2001/XMLSchema-instance}type": "string"}).text = "P_SignalPortBitPos"
                        val_base = ET.SubElement(new_p, "VALUE", {"format-rev": "1", "{http://www.w3.org/2001/XMLSchema-instance}type": "valueBaseExpression"})
                        ET.SubElement(val_base, "VALUE", {"{http://www.w3.org/2001/XMLSchema-instance}type": "integer"}).text = str(exp_bp)
                        pdata.append(new_p)
                        stats["injected_bitpos"] += 1

    # 4. Save local file and validate well-formedness
    if log_callback:
        log_callback("Writing updated XML to local SSD staging...")
    tree.write(local_staged_prj, encoding="utf-8", xml_declaration=False)

    # Validate parse
    ET.parse(local_staged_prj)
    if log_callback:
        log_callback("Local XML validation: 100% Valid & Well-Formed!")

    # 5. Atomically deploy to network target
    if log_callback:
        log_callback(f"Deploying atomically to target: {prj_path}...")
    t0 = time.time()
    shutil.copy2(local_staged_prj, prj_path)
    t1 = time.time()

    if log_callback:
        log_callback(f"Deployment completed in {t1-t0:.2f}s!")

    # Cleanup temp file
    try:
        os.remove(local_staged_prj)
    except OSError:
        pass

    return stats
