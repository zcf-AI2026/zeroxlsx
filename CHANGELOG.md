# Changelog

All notable changes to this project are documented here.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and this
project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.0] — 2026-10-26

First public release.

### Added

- `Workbook` / `Sheet` / `Style` with a small, explicit API:
  `write`, `append`, `merge`, `set_column_width`, `add_image`, `save`, `to_bytes`.
- Cell values: strings (inline, no shared-string table), numbers, booleans, dates and
  datetimes as real Excel serials, and formulas written as `"=SUM(A1:A3)"`.
- Styles: bold, italic, font size and colour, solid fill, wrap, horizontal/vertical
  alignment, built-in and custom number formats. Identical styles are de-duplicated
  into a single `cellXfs` entry.
- Layout: merged ranges, per-column widths, multiple sheets, and Excel's sheet-name
  rules enforced (31 characters, no `[]:*?/\`, unique case-insensitively).
- Embedded floating pictures: PNG and JPEG, anchored to any cell, scaled with or
  without keeping the aspect ratio. **Image dimensions are parsed from the file
  header, so Pillow is not required.**
- Deterministic output: fixed zip timestamps, so the same input produces the same
  bytes (apart from the `docProps` timestamp).
- XML escaping and control-character stripping on every string that reaches a part.

### Notes

- Zero runtime dependencies; Python 3.9+.
- Tests read the generated files back with **openpyxl + Pillow** to prove that every
  value, style, merge, width, formula and image anchor survives a round trip through
  a mainstream implementation. Those two packages are test-only.

[Unreleased]: https://github.com/zcf-AI2026/zeroxlsx/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/zcf-AI2026/zeroxlsx/releases/tag/v0.1.0
