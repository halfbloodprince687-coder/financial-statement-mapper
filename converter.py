"""Converts a legacy .xls workbook to .xlsx using headless LibreOffice.

Why LibreOffice and not a pure-Python library: xlrd (the only pure-Python
.xls reader left) can only return each formula cell's last CACHED result,
not the formula text itself - so a naive Python-only "conversion" would
silently turn every live =SUM(...) total in the template into a dead static
number. LibreOffice's own file-format engine reconstructs the real formulas,
so a template converted this way still recalculates normally in Excel.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile


class ConversionUnavailable(RuntimeError):
    pass


def _find_soffice() -> str:
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    raise ConversionUnavailable(
        "LibreOffice isn't installed here, so a .xls file can't be auto-converted. "
        "Either install LibreOffice (see README) or save the file as .xlsx yourself "
        "(Excel/LibreOffice: File -> Save As -> .xlsx) and upload that instead."
    )


def convert_xls_to_xlsx(xls_path: str, timeout: int = 60) -> str:
    """Returns the path to a converted .xlsx copy of xls_path."""
    soffice = _find_soffice()
    out_dir = tempfile.mkdtemp(prefix="fsm_convert_")
    result = subprocess.run(
        [soffice, "--headless", "--norestore", "--convert-to", "xlsx", "--outdir", out_dir, xls_path],
        capture_output=True, text=True, timeout=timeout,
    )
    stem = os.path.splitext(os.path.basename(xls_path))[0]
    out_path = os.path.join(out_dir, f"{stem}.xlsx")
    if result.returncode != 0 or not os.path.exists(out_path):
        raise ConversionUnavailable(
            f"LibreOffice could not convert this .xls file (exit code {result.returncode}). "
            f"{result.stderr.strip()[:300]}\n\nTry saving it as .xlsx manually instead."
        )
    return out_path
