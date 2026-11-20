"""
DCM Writer & Logger for SiL Calibration Tool.
Formats FESTWERT blocks, creates backups, safely appends calibrations to DCM,
and generates detailed audit logs.
"""

import datetime
import os
import shutil
from typing import List, Set
from a2l_parser import CalibrationItem


def format_festwert_block(item: CalibrationItem) -> str:
    """Formats a CalibrationItem into an ASAM-compliant FESTWERT DCM block."""
    return (
        f"FESTWERT {item.name}\n"
        f"   LANGNAME \"\"\n"
        f"   FUNKTION {item.function_group}\n"
        f"   EINHEIT_W \"-\"\n"
        f"   WERT {item.value:.16f}\n"
        f"END\n"
    )


def append_calibrations_to_dcm(
    dcm_path: str,
    items_to_add: List[CalibrationItem],
    source_a2l: str,
    discovered_buses: Set[str],
    log_dir: str = None
) -> str:
    """
    Creates a backup, appends approved items to the DCM file,
    and writes a detailed log file. Returns the path of the log file.
    """
    if not items_to_add:
        return "No items to append."

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    file_timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # 1. Create a safe backup of DCM
    backup_path = f"{dcm_path}.{file_timestamp}.bak"
    shutil.copy2(dcm_path, backup_path)

    # 2. Append entries to DCM
    with open(dcm_path, "a", encoding="utf-8") as f:
        f.write("\n* ====================================================================\n")
        f.write(f"* SiL Auto-Calibrations Added: {timestamp}\n")
        f.write(f"* Source A2L: {os.path.basename(source_a2l)}\n")
        f.write(f"* Channels: {', '.join(sorted(discovered_buses))}\n")
        f.write(f"* Total Items: {len(items_to_add)}\n")
        f.write("* ====================================================================\n\n")
        for item in items_to_add:
            f.write(format_festwert_block(item))
            f.write("\n")

    # 3. Write detailed log file
    if log_dir is None:
        log_dir = os.path.dirname(dcm_path)
        
    log_filename = f"calibration_merge_log_{file_timestamp}.txt"
    log_path = os.path.join(log_dir, log_filename)
    
    with open(log_path, "w", encoding="utf-8") as log_file:
        log_file.write("=" * 80 + "\n")
        log_file.write("           SiL CALIBRATION MERGE AUDIT LOG\n")
        log_file.write("=" * 80 + "\n")
        log_file.write(f"Date / Time       : {timestamp}\n")
        log_file.write(f"Target DCM File   : {dcm_path}\n")
        log_file.write(f"DCM Backup File   : {backup_path}\n")
        log_file.write(f"Source A2L File   : {source_a2l}\n")
        log_file.write(f"Discovered Buses  : {', '.join(sorted(discovered_buses))}\n")
        log_file.write(f"Total Added Vars  : {len(items_to_add)}\n")
        log_file.write("-" * 80 + "\n")
        log_file.write(f"{'Variable Name':50} | {'Value':10} | {'Rule / Description'}\n")
        log_file.write("-" * 80 + "\n")
        for item in items_to_add:
            log_file.write(f"{item.name:50} | {item.value:<10} | {item.rule_desc}\n")
        log_file.write("=" * 80 + "\n")
        log_file.write("STATUS: SUCCESS - All calibrations successfully appended to DCM.\n")

    return log_path
