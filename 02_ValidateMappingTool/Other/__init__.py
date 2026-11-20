"""
ECU-TEST PRJ Matrix Validator package.
"""

from .matrix_parser import parse_matrix_excel
from .prj_engine import audit_project, apply_fixes

__all__ = ["parse_matrix_excel", "audit_project", "apply_fixes"]
