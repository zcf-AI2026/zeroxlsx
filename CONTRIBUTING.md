# Contributing

Thanks for taking a look. Small, focused pull requests are easiest to review.

## Development setup

```bash
git clone https://github.com/zcf-AI2026/zeroxlsx.git
cd zeroxlsx
python -m venv .venv && source .venv/bin/activate
pip install -e ".[dev]"
```

`openpyxl` and `Pillow` are installed **only** for the tests: they read our output
back and prove it is valid. Nothing in `src/` may import them.

## Before opening a PR

```bash
ruff check src tests examples
pytest
```

Both must be clean. CI runs them on Python 3.9 through 3.13.

## House rules

1. **`src/zeroxlsx/` must stay dependency-free.** Standard library only. A PR that
   adds an import outside stdlib will be closed — that is the point of this package.
2. **Every new feature needs a round-trip test.** Write the workbook, read it back
   with openpyxl, assert the value/style/anchor. “It looks right in Excel” is not a
   test we can run in CI.
3. **Keep the OOXML minimal.** Only add a part or an attribute when Excel actually
   needs it. Extra parts mean more to keep valid.
4. Line length 100, enforced by ruff. Public methods get a one-line docstring.

## Where things live

| File | Responsibility |
|---|---|
| `src/zeroxlsx/workbook.py` | `Workbook`, `Sheet`, and the ZIP/XML writer |
| `src/zeroxlsx/_styles.py` | `Style` and the `styles.xml` builder |
| `src/zeroxlsx/_images.py` | PNG/JPEG header probing, pixels → EMU |
| `src/zeroxlsx/_ooxml.py` | Cell references, escaping, namespaces |

## Commits and releases

- Present-tense subject lines (“add freeze panes”, not “added”).
- Update `CHANGELOG.md` under `[Unreleased]` in the same PR.
- Maintainers tag releases as `vX.Y.Z`.

## Reporting a bug

Please include: what you wrote, what Excel/openpyxl/LibreOffice shows, Python version,
and — if you can — the smallest script that reproduces it.

**Do not attach a spreadsheet containing real business data.** Reduce the problem to
a few fabricated rows, or paste the generated XML part.
