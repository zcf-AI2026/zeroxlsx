"""The workbook model and the .xlsx writer."""

from __future__ import annotations

import datetime as _dt
import io
import os
import zipfile
from pathlib import Path

from . import _images, _ooxml
from ._ooxml import (
    NS_A,
    NS_DRAWING,
    NS_MAIN,
    NS_REL,
    REL_TYPE,
    cell_ref,
    column_index,
    parse_range,
    xml_declaration,
    xml_text,
)
from ._styles import Style, StyleRegistry

EXCEL_EPOCH = _dt.datetime(1899, 12, 30)
MAX_SHEET_NAME = 31
_FORBIDDEN = set('[]:*?/\\')

# Fixed timestamp so two identical workbooks produce identical bytes.
_ZIP_DATE = (1980, 1, 1, 0, 0, 0)


def _as_serial(value: _dt.date | _dt.datetime) -> float:
    if isinstance(value, _dt.datetime):
        delta = value - EXCEL_EPOCH
        return delta.days + delta.seconds / 86400 + delta.microseconds / 86400e6
    delta = value - EXCEL_EPOCH.date()
    return float(delta.days)


class Sheet:
    """One worksheet. Created through :meth:`Workbook.add_sheet`."""

    def __init__(self, title: str, book: Workbook) -> None:
        self.title = title
        self._book = book
        self._cells: dict[int, dict[int, tuple[object, Style | None]]] = {}
        self._widths: dict[int, float] = {}
        self._merges: list[str] = []
        self._images: list[dict] = []
        self._max_row = 0
        self._max_col = 0

    # -- writing ------------------------------------------------------------

    def write(self, row: int, column: int, value: object, style: Style | None = None) -> None:
        """Put ``value`` in a 1-based cell. ``column`` may be ``3`` or ``"C"``."""
        if isinstance(column, str):
            column = column_index(column)
        if row < 1 or column < 1:
            raise ValueError("row and column are 1-based")
        line = self._cells.setdefault(row, {})
        if value is None and style is None:
            line.pop(column, None)
            return
        line[column] = (value, style)
        self._max_row = max(self._max_row, row)
        self._max_col = max(self._max_col, column)

    def append(self, values, style: Style | None = None) -> int:
        """Write ``values`` on the next free row and return that row number."""
        row = self._max_row + 1
        for offset, value in enumerate(values, start=1):
            if value is not None or style is not None:
                self.write(row, offset, value, style)
        return row

    def set_column_width(self, column: int | str, width: float) -> None:
        """Column width in Excel character units (8.43 ≈ default)."""
        if isinstance(column, str):
            column = column_index(column)
        if width <= 0:
            raise ValueError("width must be positive")
        self._widths[column] = float(width)

    def merge(self, cells: str, value: object | None = None,
              style: Style | None = None) -> None:
        """Merge a range such as ``"A1:C1"``; optionally write ``value`` into the
        top-left cell (that is where Excel keeps the content)."""
        top, left, bottom, right = parse_range(cells)
        if (top, left) == (bottom, right):
            raise ValueError("a merge needs at least two cells")
        self._merges.append(f"{cell_ref(top, left)}:{cell_ref(bottom, right)}")
        if value is not None or style is not None:
            self.write(top, left, value, style)
        self._max_row = max(self._max_row, bottom)
        self._max_col = max(self._max_col, right)

    def add_image(self, source, anchor: str = "A1", width: float | None = None,
                  height: float | None = None, name: str | None = None) -> None:
        """Embed a PNG/JPEG as a floating picture anchored at ``anchor``.

        ``width`` / ``height`` are pixels; omit both to use the image's own size,
        pass only one to keep the aspect ratio.
        """
        data = _read_image(source)
        extension, px_w, px_h = _images.probe(data)
        if width is None and height is None:
            width, height = px_w, px_h
        elif width is None:
            width = height * px_w / px_h
        elif height is None:
            height = width * px_h / px_w
        row, column = _ooxml.parse_ref(anchor)
        self._images.append({
            "data": data,
            "extension": extension,
            "row": row - 1,
            "column": column - 1,
            "cx": _images.to_emu(width),
            "cy": _images.to_emu(height),
            "name": name or f"Picture {len(self._images) + 1}",
        })

    # -- serialisation ------------------------------------------------------

    def to_xml(self) -> str:
        parts = [xml_declaration(), f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_REL}">']
        last = cell_ref(max(self._max_row, 1), max(self._max_col, 1))
        parts.append(f'<dimension ref="A1:{last}"/>')
        parts.append('<sheetViews><sheetView workbookViewId="0"/></sheetViews>')
        parts.append('<sheetFormatPr defaultRowHeight="15"/>')

        if self._widths:
            parts.append("<cols>")
            for column in sorted(self._widths):
                parts.append(
                    f'<col min="{column}" max="{column}" width="{self._widths[column]}" '
                    'customWidth="1"/>'
                )
            parts.append("</cols>")

        parts.append("<sheetData>")
        for row in sorted(self._cells):
            cells = self._cells[row]
            if not any(cells.values()):
                continue
            parts.append(f'<row r="{row}">')
            for column in sorted(cells):
                parts.append(self._cell_xml(row, column, *cells[column]))
            parts.append("</row>")
        parts.append("</sheetData>")

        if self._merges:
            parts.append(f'<mergeCells count="{len(self._merges)}">')
            for ref in self._merges:
                parts.append(f'<mergeCell ref="{ref}"/>')
            parts.append("</mergeCells>")

        if self._images:
            parts.append('<drawing r:id="rId1"/>')

        parts.append("</worksheet>")
        return "".join(parts)

    def _cell_xml(self, row: int, column: int, value: object,
                  style: Style | None) -> str:
        ref = cell_ref(row, column)
        index = self._book.styles.index_of(style)
        style_attr = f' s="{index}"' if index else ""

        if value is None:
            return f'<c r="{ref}"{style_attr}/>'

        if isinstance(value, bool):
            return f'<c r="{ref}" t="b"{style_attr}><v>{int(value)}</v></c>'

        if isinstance(value, (_dt.date, _dt.datetime)):
            serial = _as_serial(value)
            if style is None:
                fmt = "yyyy-mm-dd" if not isinstance(value, _dt.datetime) else \
                    "yyyy-mm-dd hh:mm:ss"
                index = self._book.styles.index_of(Style(number_format=fmt))
                style_attr = f' s="{index}"'
            return f'<c r="{ref}"{style_attr}><v>{serial:g}</v></c>'

        if isinstance(value, (int, float)):
            return f'<c r="{ref}"{style_attr}><v>{value!r}</v></c>'

        text = str(value)
        if text.startswith("=") and len(text) > 1:
            return f'<c r="{ref}"{style_attr}><f>{xml_text(text[1:])}</f></c>'
        space = ' xml:space="preserve"' if text != text.strip() else ""
        return (f'<c r="{ref}" t="inlineStr"{style_attr}>'
                f"<is><t{space}>{xml_text(text)}</t></is></c>")


class Workbook:
    """A minimal .xlsx workbook. Call :meth:`save` to write the file."""

    def __init__(self, title: str = "Sheet") -> None:
        self.styles = StyleRegistry()
        self._sheets: list[Sheet] = []
        self._props = {"title": title, "creator": "zeroxlsx"}
        self._default_title = title

    # -- structure ----------------------------------------------------------

    def add_sheet(self, title: str | None = None) -> Sheet:
        """Append a worksheet. Names follow Excel's rules: 1-31 chars, no
        ``[]:*?/\\``, and unique (case-insensitively)."""
        if title is None:
            taken = {sheet.title.lower() for sheet in self._sheets}
            number = 1
            while f"{self._default_title}{number}".lower() in taken:
                number += 1
            name = f"{self._default_title}{number}"
        else:
            name = title
        self._validate_title(name)
        sheet = Sheet(name, self)
        self._sheets.append(sheet)
        return sheet

    @property
    def sheets(self) -> list[Sheet]:
        if not self._sheets:
            self.add_sheet()
        return self._sheets

    def _validate_title(self, name: str) -> None:
        if not name:
            raise ValueError("sheet title cannot be empty")
        if len(name) > MAX_SHEET_NAME:
            raise ValueError(f"sheet title is longer than {MAX_SHEET_NAME} characters")
        if _FORBIDDEN & set(name):
            raise ValueError('sheet title cannot contain any of: [ ] : * ? / \\')
        if name.startswith("'") or name.endswith("'"):
            raise ValueError("sheet title cannot start or end with an apostrophe")
        if any(existing.title.lower() == name.lower() for existing in self._sheets):
            raise ValueError(f"a sheet named {name!r} already exists")

    @property
    def sheet_count(self) -> int:
        return len(self._sheets)

    # -- writing ------------------------------------------------------------

    def save(self, target) -> None:
        """Write the workbook to a path (``str`` / ``Path``) or a binary file object."""
        data = self.to_bytes()
        if hasattr(target, "write"):
            target.write(data)
        else:
            Path(target).write_bytes(data)

    def to_bytes(self) -> bytes:
        """Serialise the whole package in memory."""
        sheets = self.sheets
        # Sheet XML must be built first: it is what registers cell styles, and
        # styles.xml has to be written with the final cellXfs table.
        sheet_xml = [(sheet, sheet.to_xml()) for sheet in sheets]
        styles_xml = self.styles.to_xml()

        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
            self._write_part(archive, "[Content_Types].xml", self._content_types(sheets))
            self._write_part(archive, "_rels/.rels", self._package_rels())
            self._write_part(archive, "docProps/core.xml", self._core_props())
            self._write_part(archive, "docProps/app.xml", self._app_props(sheets))
            self._write_part(archive, "xl/workbook.xml", self._workbook_xml(sheets))
            self._write_part(archive, "xl/_rels/workbook.xml.rels",
                             self._workbook_rels(sheets))
            self._write_part(archive, "xl/styles.xml", styles_xml)

            image_number = 0
            for number, (sheet, xml) in enumerate(sheet_xml, start=1):
                self._write_part(archive, f"xl/worksheets/sheet{number}.xml", xml)
                if not sheet._images:
                    continue
                self._write_part(
                    archive,
                    f"xl/worksheets/_rels/sheet{number}.xml.rels",
                    self._sheet_rels(number),
                )
                self._write_part(archive, f"xl/drawings/drawing{number}.xml",
                                 self._drawing_xml(sheet, number))
                self._write_part(archive, f"xl/drawings/_rels/drawing{number}.xml.rels",
                                 self._drawing_rels(sheet))
                for image in sheet._images:
                    image_number += 1
                    ext = image["extension"]
                    self._write_part(archive, f"xl/media/image{image_number}.{ext}",
                                     image["data"])
        return buffer.getvalue()

    @staticmethod
    def _write_part(archive: zipfile.ZipFile, name: str, data) -> None:
        if isinstance(data, str):
            data = data.encode("utf-8")
        info = zipfile.ZipInfo(name, date_time=_ZIP_DATE)
        info.compress_type = zipfile.ZIP_DEFLATED
        info.external_attr = 0o600 << 16
        archive.writestr(info, data)

    # -- package parts ------------------------------------------------------

    def _content_types(self, sheets: list[Sheet]) -> str:
        parts = [xml_declaration(), f'<Types xmlns="{_ooxml.NS_CT}">',
                 '<Default Extension="rels" ContentType="application/'
                 'vnd.openxmlformats-package.relationships+xml"/>',
                 '<Default Extension="xml" ContentType="application/xml"/>',
                 '<Default Extension="png" ContentType="image/png"/>',
                 '<Default Extension="jpeg" ContentType="image/jpeg"/>',
                 '<Override PartName="/xl/workbook.xml" ContentType="application/'
                 'vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
                 '<Override PartName="/xl/styles.xml" ContentType="application/'
                 'vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
                 '<Override PartName="/docProps/core.xml" ContentType="application/'
                 'vnd.openxmlformats-package.core-properties+xml"/>',
                 '<Override PartName="/docProps/app.xml" ContentType="application/'
                 'vnd.openxmlformats-officedocument.extended-properties+xml"/>']
        for number, sheet in enumerate(sheets, start=1):
            parts.append(f'<Override PartName="/xl/worksheets/sheet{number}.xml" '
                         'ContentType="application/vnd.openxmlformats-officedocument.'
                         'spreadsheetml.worksheet+xml"/>')
            if sheet._images:
                parts.append(f'<Override PartName="/xl/drawings/drawing{number}.xml" '
                             'ContentType="application/vnd.openxmlformats-officedocument.'
                             'drawing+xml"/>')
        parts.append("</Types>")
        return "".join(parts)

    @staticmethod
    def _package_rels() -> str:
        return "".join([
            xml_declaration(), f'<Relationships xmlns="{_ooxml.NS_PKG_REL}">',
            f'<Relationship Id="rId1" Type="{REL_TYPE["office_document"]}" '
            'Target="xl/workbook.xml"/>',
            f'<Relationship Id="rId2" Type="{REL_TYPE["core"]}" '
            'Target="docProps/core.xml"/>',
            f'<Relationship Id="rId3" Type="{REL_TYPE["app"]}" '
            'Target="docProps/app.xml"/>',
            "</Relationships>",
        ])

    def _workbook_xml(self, sheets: list[Sheet]) -> str:
        parts = [xml_declaration(), f'<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_REL}">',
                 "<workbookPr/>", "<sheets>"]
        for number, sheet in enumerate(sheets, start=1):
            parts.append(f'<sheet name="{xml_text(sheet.title)}" sheetId="{number}" '
                         f'r:id="rId{number}"/>')
        parts.append("</sheets></workbook>")
        return "".join(parts)

    def _workbook_rels(self, sheets: list[Sheet]) -> str:
        parts = [xml_declaration(), f'<Relationships xmlns="{_ooxml.NS_PKG_REL}">']
        for number in range(1, len(sheets) + 1):
            parts.append(f'<Relationship Id="rId{number}" Type="{REL_TYPE["worksheet"]}" '
                         f'Target="worksheets/sheet{number}.xml"/>')
        parts.append(f'<Relationship Id="rId{len(sheets) + 1}" Type="{REL_TYPE["styles"]}" '
                     'Target="styles.xml"/>')
        parts.append("</Relationships>")
        return "".join(parts)

    @staticmethod
    def _sheet_rels(number: int) -> str:
        return "".join([
            xml_declaration(), f'<Relationships xmlns="{_ooxml.NS_PKG_REL}">',
            f'<Relationship Id="rId1" Type="{REL_TYPE["drawing"]}" '
            f'Target="../drawings/drawing{number}.xml"/>',
            "</Relationships>",
        ])

    @staticmethod
    def _drawing_xml(sheet: Sheet, number: int) -> str:
        parts = [xml_declaration(),
                 f'<xdr:wsDr xmlns:xdr="{NS_DRAWING}" xmlns:a="{NS_A}" xmlns:r="{NS_REL}">']
        for index, image in enumerate(sheet._images, start=1):
            cx, cy = image["cx"], image["cy"]
            parts.append(
                '<xdr:oneCellAnchor>'
                f'<xdr:from><xdr:col>{image["column"]}</xdr:col><xdr:colOff>0</xdr:colOff>'
                f'<xdr:row>{image["row"]}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
                f'<xdr:ext cx="{cx}" cy="{cy}"/>'
                "<xdr:pic><xdr:nvPicPr>"
                f'<xdr:cNvPr id="{index + 1}" name="{xml_text(image["name"])}"/>'
                '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>'
                "</xdr:nvPicPr>"
                f'<xdr:blipFill><a:blip r:embed="rId{index}"/>'
                "<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>"
                '<xdr:spPr><a:xfrm><a:off x="0" y="0"/>'
                f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>'
                "</xdr:pic><xdr:clientData/></xdr:oneCellAnchor>"
            )
        parts.append("</xdr:wsDr>")
        return "".join(parts)

    @staticmethod
    def _drawing_rels(sheet: Sheet) -> str:
        parts = [xml_declaration(), f'<Relationships xmlns="{_ooxml.NS_PKG_REL}">']
        for index, image in enumerate(sheet._images, start=1):
            parts.append(
                f'<Relationship Id="rId{index}" Type="{REL_TYPE["image"]}" '
                f'Target="../media/image{index}.{image["extension"]}"/>'
            )
        parts.append("</Relationships>")
        return "".join(parts)

    def _core_props(self) -> str:
        now = _dt.datetime.now(_dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        return "".join([
            xml_declaration(),
            '<cp:coreProperties '
            'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/'
            'core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" '
            'xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
            f"<dc:title>{xml_text(self._props['title'])}</dc:title>",
            f"<dc:creator>{xml_text(self._props['creator'])}</dc:creator>",
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>',
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>',
            "</cp:coreProperties>",
        ])

    @staticmethod
    def _app_props(sheets: list[Sheet]) -> str:
        return "".join([
            xml_declaration(),
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/'
            '2006/extended-properties">',
            "<Application>zeroxlsx</Application>",
            f"<Sheets>{len(sheets)}</Sheets>",
            "</Properties>",
        ])


def _read_image(source) -> bytes:
    """Accept a path, ``bytes`` or an already-open binary file object."""
    if isinstance(source, (bytes, bytearray, memoryview)):
        return bytes(source)
    if hasattr(source, "read"):
        return source.read()
    return Path(os.fspath(source)).read_bytes()
