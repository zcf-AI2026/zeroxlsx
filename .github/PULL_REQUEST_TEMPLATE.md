## What this changes

<!-- One or two sentences. Link the issue if there is one: Fixes #123 -->

## Checklist

- [ ] `ruff check src tests examples` is clean
- [ ] `pytest` passes
- [ ] No new import outside the standard library in `src/zeroxlsx/`
- [ ] New behaviour is covered by a round-trip test (write → read back with openpyxl)
- [ ] `CHANGELOG.md` updated under `[Unreleased]`
- [ ] Generated files were opened in a real viewer at least once (Excel / LibreOffice),
      and the result is described below

## What I verified by opening the file

<!-- e.g. "Excel 365 on Windows: merged title spans A1:D1, the logo sits at F2, margins show 31.2%" -->

## Notes for reviewers
