# zeroxlsx

**Write real `.xlsx` files with zero dependencies — styles, formulas, merged cells, column widths and embedded floating images, all from the standard library.**

[![CI](https://github.com/zcf-AI2026/zeroxlsx/actions/workflows/ci.yml/badge.svg)](https://github.com/zcf-AI2026/zeroxlsx/actions/workflows/ci.yml)
[![PyPI](https://img.shields.io/pypi/v/zeroxlsx.svg)](https://pypi.org/project/zeroxlsx/)
[![Python](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://pypi.org/project/zeroxlsx/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

`pip install openpyxl` pulls in four packages. `XlsxWriter` brings its own tree.
Sometimes you just need to *emit a spreadsheet* — inside a Lambda, a CI job, an
air-gapped box, or a one-file script — and you would rather not add a dependency
tree for it.

`zeroxlsx` is a single small package that writes correct OOXML by hand:
`zipfile` + `xml`. Nothing else. **It can even embed pictures without Pillow** —
the image size is read straight from the PNG/JPEG header.

```python
from zeroxlsx import Workbook, Style

book = Workbook()
sheet = book.add_sheet("Report")

sheet.merge("A1:D1", "Quarterly summary", Style(bold=True, size=14))
for column, label in enumerate(["Region", "Orders", "Revenue", "Margin"], start=1):
    sheet.write(2, column, label, Style(bold=True, fill="#F2E8D5", h_align="center"))

sheet.write(3, 1, "North")
sheet.write(3, 2, 128)
sheet.write(3, 3, 48230.50, Style(number_format="#,##0.00"))
sheet.write(3, 4, 0.312, Style(number_format="0.0%"))

sheet.add_image("logo.png", anchor="F2", width=72)   # PNG or JPEG, no Pillow
book.save("report.xlsx")
```

## What you get

| | openpyxl | XlsxWriter | **zeroxlsx** |
|---|---|---|---|
| Write .xlsx | ✅ | ✅ | ✅ |
| Read .xlsx | ✅ | ❌ | ❌ *(by design)* |
| Charts / pivot tables | ✅ | ✅ | ❌ |
| **Runtime dependencies** | 4 | 2 | **0** |
| Embed images without Pillow | ❌ | ✅ | ✅ |
| Styles, merges, widths, formulas | ✅ | ✅ | ✅ |
| Deterministic output (same input → same bytes) | ❌ | — | ✅ |

If you need to *read* workbooks or build charts, use openpyxl / XlsxWriter — they
are excellent at that. `zeroxlsx` covers the "just write the file" case with nothing
to install.

## Install

```bash
pip install zeroxlsx
```

Python 3.9+. No dependencies. The installed wheel is ~20 KB.

## Features

- **Values**: strings (inline, no shared-string table), numbers, booleans, dates
  and datetimes (as real Excel serials), formulas (`"=SUM(A1:A3)"`).
- **Styles**: bold / italic / font size / font colour / solid fill / wrap / horizontal
  and vertical alignment / number formats (built-in ids or any custom code).
  Identical styles collapse into one `cellXfs` entry automatically.
- **Layout**: merged ranges, per-column widths, multiple sheets with Excel's naming
  rules enforced (31 chars, no `[]:*?/\`, unique).
- **Images**: PNG and JPEG, anchored to any cell, scaled with or without keeping the
  aspect ratio. Dimensions are parsed from the file header — no Pillow.
- **Safety**: XML escaping, control characters stripped, sheet names validated,
  zip entries written deterministically.

## API at a glance

```python
book = Workbook()                      # optional: Workbook("Report")
sheet = book.add_sheet("Data")         # add_sheet() -> auto-named "Sheet1", "Sheet2", …

sheet.write(row, column, value, style=None)     # 1-based; column may be "C" or 3
sheet.append([...], style=None)                 # writes the next row, returns its number
sheet.merge("A1:C1", "title", style=None)       # optionally fill the top-left cell
sheet.set_column_width("A", 24)                 # or set_column_width(1, 24)
sheet.add_image(png_or_path, anchor="B2", width=None, height=None)

book.save("out.xlsx")                  # or book.save(buffer), or bytes = book.to_bytes()
```

`Style` is a frozen dataclass: `bold`, `italic`, `size`, `font_color`, `fill`,
`number_format`, `wrap`, `h_align`, `v_align`. Use `style.merged(italic=True)` to
derive a variant.

## How the file is built

An `.xlsx` is a ZIP of XML parts. `zeroxlsx` writes exactly the parts Excel needs:

```
[Content_Types].xml
_rels/.rels
docProps/{core,app}.xml
xl/workbook.xml  +  xl/_rels/workbook.xml.rels
xl/styles.xml
xl/worksheets/sheetN.xml  (+ _rels when the sheet has a drawing)
xl/drawings/drawingN.xml  (+ _rels)     ← only when images are embedded
xl/media/imageN.png|jpeg
```

Strings are written inline (`t="inlineStr"`) instead of in a shared-string table —
one part less to build, and Excel accepts it fine. Pictures use `oneCellAnchor`
with explicit EMU extents (1 px = 9525 EMU at 96 dpi).

## Tests

The library has no dependencies, but the tests deliberately use **openpyxl + Pillow**
to prove the files are readable by a mainstream implementation — every value, style,
merge, width, formula and image anchor is read back and asserted.

```bash
pip install -e ".[dev]"
pytest
```

## Limitations

- Write-only. No reader, no charts, no pivot tables, no images *inside* cells
  (use `add_image` for a floating picture), no conditional formatting, no data
  validation, no autofilter, no freeze panes yet.
- No streaming writer for millions of rows (each sheet is built in memory).

Roadmap: freeze panes, autofilter, a streaming row writer, in-cell images, and
`.xlsx` → JSON reading in a separate `zeroxlsx.read` module.

## License

MIT — see [LICENSE](LICENSE).
