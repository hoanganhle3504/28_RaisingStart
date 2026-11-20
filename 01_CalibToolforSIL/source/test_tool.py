"""
Automated Verification Suite for SiL Calibration Tool.
Tests A2L parsing, DCM parsing, diffing logic, backup creation, and log generation.
"""

import os
import shutil
import tempfile
import unittest

from a2l_parser import CalibrationItem, extract_needed_calibrations
from dcm_parser import parse_existing_dcm_variables
from dcm_writer import append_calibrations_to_dcm, format_festwert_block
from excel_exporter import export_calibrations_to_excel


class TestSilCalibTool(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.prototype_a2l = r"C:\Workspace\LAG8HC\02_1015_1_Rex_Prototype\1015.1_Prototype\Simulate\02_ECU\vECU\vECU.a2l"
        cls.prototype_dcm = r"C:\Workspace\LAG8HC\02_1015_1_Rex_Prototype\1015.1_Prototype\Simulate\02_ECU\vECU\Config\DCM\Modified_Calibrations.dcm"

    def test_a2l_extraction_and_bus_discovery(self):
        items, buses = extract_needed_calibrations(self.prototype_a2l)
        self.assertIn("HVCAN", buses)
        self.assertIn("LIN1", buses)
        self.assertGreater(len(items), 100)

        # Verify Rule 1 (*swtPduCodActv = 0)
        pdu_cod = [it for it in items if "Com_swtPduCodActv" in it.name]
        self.assertTrue(len(pdu_cod) >= 2)
        for it in pdu_cod:
            self.assertEqual(it.value, 0.0)

        # Verify Rule 2 (Com_idxDiagDi<BUS>_C = 2)
        net_idx = [it for it in items if it.name in ("Com_idxDiagDiHVCAN_C", "Com_idxDiagDiLIN1_C")]
        self.assertEqual(len(net_idx), 2)
        for it in net_idx:
            self.assertEqual(it.value, 2.0)

        # Verify Rule 3 (Com_idxDiagDi<BUS><PDU>_C = 6)
        pdu_idx = [it for it in items if it.name == "Com_idxDiagDiHVCANESP_10_C"]
        self.assertEqual(len(pdu_idx), 1)
        self.assertEqual(pdu_idx[0].value, 6.0)

        # Verify Rule 4 (Com_swtRxEna / TxEna = 1)
        rx_ena = [it for it in items if it.name == "Com_swtRxEnaHVCANAirbag_10_C"]
        self.assertEqual(len(rx_ena), 1)
        self.assertEqual(rx_ena[0].value, 1.0)

        # Verify Rule 5 & 6 (DFC Ctl=1, Disbl=0)
        dfc_ctl = [it for it in items if it.name == "DFC_CtlMsk2.DFC_ComTOutESP_C"]
        dfc_disbl = [it for it in items if it.name == "DFC_DisblMsk2.DFC_ComTOutESP_C"]
        self.assertEqual(len(dfc_ctl), 1)
        self.assertEqual(dfc_ctl[0].value, 1.0)
        self.assertEqual(dfc_ctl[0].function_group, "DFC")
        self.assertEqual(len(dfc_disbl), 1)
        self.assertEqual(dfc_disbl[0].value, 0.0)

    def test_dcm_parsing_and_exclusion(self):
        existing = parse_existing_dcm_variables(self.prototype_dcm)
        self.assertGreater(len(existing), 50)

    def test_dcm_writer_and_backup(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_dcm = os.path.join(temp_dir, "test_modified.dcm")
            shutil.copy2(self.prototype_dcm, temp_dcm)

            test_item = CalibrationItem("Com_swtRxEnaHVCANTestPdu_C", 1.0, "Test Rule", "comcil_diag_auto")
            log_path = append_calibrations_to_dcm(
                dcm_path=temp_dcm,
                items_to_add=[test_item],
                source_a2l=self.prototype_a2l,
                discovered_buses={"HVCAN", "LIN1"},
                log_dir=temp_dir
            )

            self.assertTrue(os.path.isfile(log_path))
            with open(log_path, "r", encoding="utf-8") as f:
                log_content = f.read()
                self.assertIn("Com_swtRxEnaHVCANTestPdu_C", log_content)
                self.assertIn("STATUS: SUCCESS", log_content)

            with open(temp_dcm, "r", encoding="utf-8") as f:
                dcm_content = f.read()
                self.assertIn("FESTWERT Com_swtRxEnaHVCANTestPdu_C", dcm_content)
                self.assertIn("WERT 1.0000000000000000", dcm_content)

    def test_excel_export(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_xlsx = os.path.join(temp_dir, "test_calibrations.xlsx")
            items = [
                CalibrationItem("Com_swtPduCodActvHVCAN_C", 0.0, "Coding switch", "comcil_diag_auto"),
                CalibrationItem("Com_idxDiagDiHVCAN_C", 2.0, "Network diag index", "comcil_diag_auto"),
                CalibrationItem("Com_swtRxEnaHVCANESP_10_C", 1.0, "Rx enable", "comcil_diag_auto"),
                CalibrationItem("DFC_CtlMsk2.DFC_ComTOutESP_C", 1.0, "DFC control mask", "DFC")
            ]
            export_calibrations_to_excel(
                items=items,
                export_path=temp_xlsx,
                target_dcm=self.prototype_dcm,
                source_a2l=self.prototype_a2l,
                discovered_buses={"HVCAN", "LIN1"}
            )
            self.assertTrue(os.path.isfile(temp_xlsx))
            self.assertGreater(os.path.getsize(temp_xlsx), 1000)


if __name__ == "__main__":
    unittest.main()
