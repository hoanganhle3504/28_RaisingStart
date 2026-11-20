"""
test_validator.py - Unit tests for Matrix Parser and PRJ Engine.
"""

import os
import sys
import unittest
import tempfile
import xml.etree.ElementTree as ET

# Ensure prj_matrix_validator is importable
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "prj_matrix_validator")))

from matrix_parser import parse_matrix_excel, MACRO_TABLE
from prj_engine import audit_project, apply_fixes, PACKAGES_FOR_BITPOS, PACKAGES_FOR_DEFVAL


class TestMatrixValidator(unittest.TestCase):
    def test_macro_table(self):
        self.assertEqual(MACRO_TABLE["MAXUINT8"], 255)
        self.assertEqual(MACRO_TABLE["MAXUINT16"], 65535)
        self.assertEqual(MACRO_TABLE["MAXSINT16"], 32767)
        self.assertEqual(MACRO_TABLE["MAXSINT32"], 2147483647)
        self.assertEqual(MACRO_TABLE["MAXUINT32"], 4294967295)

    def test_mock_prj_audit_and_fix(self):
        # Create a mock XML project file
        mock_xml = """<PROJECT xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" format-rev="10">
    <COMPONENTS>
        <COMPONENT xsi:type="subproject">
            <NAME>TestSig__TestPort</NAME>
            <COMPONENTS>
                <COMPONENT xsi:type="packageTest">
                    <NAME>PortRxErrHndlgPduTout_FadeoutInActive</NAME>
                    <COMPONENTS>
                        <COMPONENT xsi:type="paramRecord">
                            <PARAMETER-DATA xsi:type="packageParameterData">
                                <PARAMETER xsi:type="parameterEntry">
                                    <NAME>P_PduName</NAME>
                                    <VALUE format-rev="1" xsi:type="valueBaseExpression">
                                        <VALUE xsi:type="string">TestPdu</VALUE>
                                    </VALUE>
                                </PARAMETER>
                            </PARAMETER-DATA>
                        </COMPONENT>
                    </COMPONENTS>
                </COMPONENT>
            </COMPONENTS>
        </COMPONENT>
    </COMPONENTS>
</PROJECT>"""

        with tempfile.TemporaryDirectory() as td:
            prj_path = os.path.join(td, "mock_test.prj")
            with open(prj_path, "w", encoding="utf-8") as f:
                f.write(mock_xml)

            mock_matrix = {
                "TestSig__TestPort": {
                    "bus": "HVCANFD",
                    "pdu": "TestPdu",
                    "sig": "TestSig",
                    "port": "TestPort",
                    "bitpos": 5,
                    "conv": "Value_Direct",
                    "defval": 15
                }
            }

            # 1. Audit should detect missing BitPos and DefaultValue
            disc = audit_project(prj_path, mock_matrix)
            self.assertEqual(len(disc), 2)
            params_found = {d["param"] for d in disc}
            self.assertIn("P_SignalPortBitPos", params_found)
            self.assertIn("P_SignalPortDefaultValue", params_found)

            # 2. Apply fixes
            stats = apply_fixes(prj_path, mock_matrix)
            self.assertEqual(stats["injected_bitpos"], 1)
            self.assertEqual(stats["injected_defaults"], 1)

            # 3. Re-audit should find 0 discrepancies
            disc_after = audit_project(prj_path, mock_matrix)
            self.assertEqual(len(disc_after), 0)


if __name__ == "__main__":
    unittest.main()
