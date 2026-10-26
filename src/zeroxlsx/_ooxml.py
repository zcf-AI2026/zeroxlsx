"""Low-level OOXML plumbing: references, escaping and shared namespaces.

Everything here is pure data shaping — no file I/O, no state.
"""

from __future__ import annotations

import re
from xml.sax.saxutils import escape as _xml_escape

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_DOC_PROPS = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_TYPE = {
    "office_document": f"{NS_REL}/officeDocument",
    "worksheet": f"{NS_REL}/worksheet",
    "styles": f"{NS_REL}/styles",
    "drawing": f"{NS_REL}/drawing",
    "image": f"{NS_REL}/image",
    "core": "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
    "app": f"{NS_REL}/extended-properties",
}

_CONTROL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_REF = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


def xml_text(value: str) -> str:
    """Escape text for XML and drop characters that are illegal in XML 1.0."""
    return _xml_escape(_CONTROL.sub("", value))


def column_index(letter: str) -> int:
    """``"A" -> 1``, ``"AA" -> 27``."""
    if not letter or not letter.isalpha():
        raise ValueError(f"bad column: {letter!r}")
    index = 0
    for char in letter.upper():
        index = index * 26 + (ord(char) - 64)
    return index


def column_letter(index: int) -> str:
    """``1 -> "A"``, ``27 -> "AA"``."""
    if index < 1:
        raise ValueError(f"column index must be >= 1, got {index}")
    letters = ""
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def cell_ref(row: int, column: int) -> str:
    """``(7, 2) -> "B7"``."""
    if row < 1 or column < 1:
        raise ValueError("row and column are 1-based")
    return f"{column_letter(column)}{row}"


def parse_ref(ref: str) -> tuple[int, int]:
    """``"B7" -> (7, 2)``."""
    match = _REF.match(ref.strip().upper())
    if not match:
        raise ValueError(f"bad cell reference: {ref!r}")
    return int(match.group(2)), column_index(match.group(1))


def parse_range(ref: str) -> tuple[int, int, int, int]:
    """``"A1:C3" -> (1, 1, 3, 3)`` (top, left, bottom, right)."""
    parts = ref.split(":")
    if len(parts) == 1:
        row, col = parse_ref(parts[0])
        return row, col, row, col
    if len(parts) != 2:
        raise ValueError(f"bad range: {ref!r}")
    top, left = parse_ref(parts[0])
    bottom, right = parse_ref(parts[1])
    return min(top, bottom), min(left, right), max(top, bottom), max(left, right)


def color(value: str | None) -> str | None:
    """Normalise ``#rrggbb`` / ``rrggbb`` to the ``FFRRGGBB`` form OOXML wants."""
    if value is None:
        return None
    hexdigits = value.lstrip("#")
    if len(hexdigits) == 6:
        hexdigits = "FF" + hexdigits
    if len(hexdigits) != 8:
        raise ValueError(f"bad colour: {value!r} (use #RRGGBB or FFRRGGBB)")
    return hexdigits.upper()


def xml_declaration() -> str:
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
