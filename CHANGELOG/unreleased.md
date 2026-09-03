# Unreleased

These are the changes that will go out in the next release.

## Added

First release of pycelerate — build Excel formulas with Python operators, then
hand them straight to openpyxl.

- Expression tree rendered with Excel's precedence rules, including the two places
  Python and Excel genuinely disagree (`^` is left-associative in Excel; unary `-`
  binds tighter than `^`).
- `Sheet.put()` writes a value and hands back a `CellRef`, so no coordinate is
  typed twice.
- `CellRef` / `RangeRef` / `Ref`: relative by default, sheet qualification never
  inferred, `.abs()` / `.offset()` / `.shift()` / `.qualified()`.
- `F.ANYNAME(...)` builds a call for any function name, with the `_xlfn.` and
  `_xlfn._xlws.` prefixes the xlsx format needs applied automatically.
- Three layers of validation: arity checked at construction for known names, a
  did-you-mean warning for near-miss spellings, and type stubs (`functions.pyi`)
  that move both to edit time. Unknown names stay unchecked, because the set of
  real Excel functions is open.
- Zero runtime dependencies. Nothing in `src/` imports openpyxl; worksheets are
  duck-typed.

### Project infrastructure

- Requires Python 3.11+, tested on 3.11, 3.12, 3.13 and 3.14.
- Ruff formats and lints, both gating CI. Formatting is Python-only: Ruff's
  Markdown code-fence formatting is disabled, because the docs use aligned
  trailing comments to show a call beside the formula Excel receives.
- Both pyright and ty gate CI. pyright is authoritative, because the package ships
  `py.typed` and its own type errors would surface in consumers' editors.
- `tests/test_recalc.py` runs in CI against a real LibreOffice install — the only
  check that tests what a formula *means* rather than how it is shaped.
- Security scanning: CodeQL, Dependabot, and a Zizmor workflow audit.
- Docs site with Zensical + mkdocstrings, deployed to GitHub Pages.
- Trusted publishing to PyPI with OIDC and build provenance attestation.

### Contributors

[@dunlopWill](https://github.com/dunlopWill) (Will Dunlop) created pycelerate.
