"""Excel/CSV loading without python-calamine.

Some Corning rate cards have invalid stylesheet colors. openpyxl refuses those
files, so we repair xl/styles.xml inside a temp copy, then read with openpyxl.
"""

from __future__ import annotations

import re
import tempfile
import zipfile
from pathlib import Path

import pandas as pd

# openpyxl requires RGB like "FFRRGGBB". Broken files often have "RGB", theme refs, etc.
_RGB_ATTR = re.compile(r'\brgb="([^"]*)"', re.IGNORECASE)


def _sanitize_rgb(match: re.Match[str]) -> str:
    raw = match.group(1).strip()
    hex_part = re.sub(r"[^0-9A-Fa-f]", "", raw)
    if len(hex_part) == 8:
        return f'rgb="{hex_part.upper()}"'
    if len(hex_part) == 6:
        return f'rgb="FF{hex_part.upper()}"'
    # Fallback: opaque black — keeps workbook readable; styles are not needed for values
    return 'rgb="FF000000"'


def _repair_xlsx_styles(src: Path, dest: Path) -> None:
    """Copy xlsx and fix invalid rgb= attributes in styles.xml."""
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(
        dest, "w", compression=zipfile.ZIP_DEFLATED
    ) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.lower() in {"xl/styles.xml", "xl/styles.xml"}:
                text = data.decode("utf-8", errors="replace")
                text = _RGB_ATTR.sub(_sanitize_rgb, text)
                data = text.encode("utf-8")
            zout.writestr(info, data)


def _read_xlsx_openpyxl(path: Path) -> pd.DataFrame:
    try:
        return pd.read_excel(path, engine="openpyxl")
    except ValueError as first_exc:
        # Typical: "Colors must be aRGB hex values" / unable to read stylesheet
        msg = str(first_exc).lower()
        if "stylesheet" not in msg and "argb" not in msg and "color" not in msg:
            raise
        with tempfile.TemporaryDirectory() as tmp:
            repaired = Path(tmp) / f"repaired_{path.name}"
            _repair_xlsx_styles(path, repaired)
            return pd.read_excel(repaired, engine="openpyxl")


def load_to_dataframe(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path)
    if suffix == ".tsv":
        return pd.read_csv(path, sep="\t")
    if suffix == ".xls":
        return pd.read_excel(path)
    if suffix != ".xlsx":
        raise ValueError(f"Unsupported file type: {path.suffix}")
    return _read_xlsx_openpyxl(path)
