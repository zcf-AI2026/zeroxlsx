"""Cell styling: a small ``Style`` value object plus the styles.xml builder."""

from __future__ import annotations

from dataclasses import dataclass

from ._ooxml import color, xml_text

BUILTIN_NUMBER_FORMATS = {
    "general": 0,
    "0": 1,
    "0.00": 2,
    "#,##0": 3,
    "#,##0.00": 4,
    "0%": 9,
    "0.00%": 10,
}

FIRST_CUSTOM_FORMAT_ID = 164


@dataclass(frozen=True)
class Style:
    """A cell format. Instances are hashable, so they de-duplicate naturally."""

    bold: bool = False
    italic: bool = False
    size: float = 11
    font_color: str | None = None
    fill: str | None = None
    number_format: str | int | None = None
    wrap: bool = False
    h_align: str | None = None
    v_align: str | None = None

    def merged(self, **overrides) -> Style:
        """Return a copy with some fields replaced."""
        data = self.__dict__ | overrides
        return Style(**data)


class StyleRegistry:
    """Assigns indices to fonts / fills / number formats / cell formats."""

    def __init__(self) -> None:
        self._fonts: list[Style] = []
        self._fills: list[str | None] = [None, "__gray125__"]
        self._formats: dict[str, int] = {}
        self._xfs: list[Style] = []
        self._next_format_id = FIRST_CUSTOM_FORMAT_ID
        self._register(Style())  # index 0 = the default "Normal" format

    # -- registration -------------------------------------------------------

    def _register(self, style: Style) -> int:
        if style in self._xfs:
            return self._xfs.index(style)
        if style not in self._fonts:
            self._fonts.append(style)
        if style.fill is not None and style.fill not in self._fills:
            self._fills.append(style.fill)
        self._format_id(style.number_format)
        self._xfs.append(style)
        return len(self._xfs) - 1

    def index_of(self, style: Style | None) -> int:
        if style is None:
            return 0
        return self._register(style)

    def _format_id(self, number_format: str | int | None) -> int:
        if number_format is None:
            return 0
        if isinstance(number_format, int):
            return number_format
        key = number_format.strip()
        builtin = BUILTIN_NUMBER_FORMATS.get(key.lower())
        if builtin is not None:
            return builtin
        if key not in self._formats:
            self._formats[key] = self._next_format_id
            self._next_format_id += 1
        return self._formats[key]

    # -- serialisation ------------------------------------------------------

    def to_xml(self) -> str:
        parts = [
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        ]
        if self._formats:
            parts.append(f'<numFmts count="{len(self._formats)}">')
            for code, fmt_id in self._formats.items():
                parts.append(
                    f'<numFmt numFmtId="{fmt_id}" formatCode="{xml_text(code)}"/>'
                )
            parts.append("</numFmts>")

        parts.append(f'<fonts count="{len(self._fonts)}">')
        for style in self._fonts:
            parts.append(self._font_xml(style))
        parts.append("</fonts>")

        parts.append(f'<fills count="{len(self._fills)}">')
        for fill in self._fills:
            parts.append(self._fill_xml(fill))
        parts.append("</fills>")

        parts.append('<borders count="1"><border><left/><right/><top/><bottom/><diagonal/>'
                     "</border></borders>")
        parts.append('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" '
                     'borderId="0"/></cellStyleXfs>')

        parts.append(f'<cellXfs count="{len(self._xfs)}">')
        for style in self._xfs:
            parts.append(self._xf_xml(style))
        parts.append("</cellXfs>")
        parts.append('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/>'
                     "</cellStyles>")
        parts.append("</styleSheet>")
        return "".join(parts)

    def _font_xml(self, style: Style) -> str:
        bits = ["<font>"]
        if style.bold:
            bits.append("<b/>")
        if style.italic:
            bits.append("<i/>")
        bits.append(f'<sz val="{style.size}"/>')
        rgb = color(style.font_color)
        if rgb:
            bits.append(f'<color rgb="{rgb}"/>')
        bits.append('<name val="Calibri"/><family val="2"/></font>')
        return "".join(bits)

    @staticmethod
    def _fill_xml(fill: str | None) -> str:
        if fill is None:
            return '<fill><patternFill patternType="none"/></fill>'
        if fill == "__gray125__":
            return '<fill><patternFill patternType="gray125"/></fill>'
        rgb = color(fill)
        return ('<fill><patternFill patternType="solid">'
                f'<fgColor rgb="{rgb}"/><bgColor indexed="64"/>'
                "</patternFill></fill>")

    def _xf_xml(self, style: Style) -> str:
        font_id = self._fonts.index(style)
        fill_id = self._fills.index(style.fill) if style.fill in self._fills else 0
        fmt_id = self._format_id(style.number_format)
        attrs = [
            f'numFmtId="{fmt_id}"',
            f'fontId="{font_id}"',
            f'fillId="{fill_id}"',
            'borderId="0"',
            'xfId="0"',
        ]
        if fmt_id:
            attrs.append('applyNumberFormat="1"')
        if font_id:
            attrs.append('applyFont="1"')
        if fill_id > 1:
            attrs.append('applyFill="1"')

        align_bits = []
        if style.h_align:
            align_bits.append(f'horizontal="{style.h_align}"')
        if style.v_align:
            align_bits.append(f'vertical="{style.v_align}"')
        if style.wrap:
            align_bits.append('wrapText="1"')
        if align_bits:
            attrs.append('applyAlignment="1"')
            return f"<xf {' '.join(attrs)}><alignment {' '.join(align_bits)}/></xf>"
        return f"<xf {' '.join(attrs)}/>"
