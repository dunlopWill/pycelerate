# pycelerate

[![CI](https://github.com/dunlopWill/pycelerate/actions/workflows/ci.yml/badge.svg)](https://github.com/dunlopWill/pycelerate/actions/workflows/ci.yml)
[![PyPI](https://img.shields.io/pypi/v/pycelerate.svg)](https://pypi.org/project/pycelerate/)
[![Python versions](https://img.shields.io/pypi/pyversions/pycelerate.svg)](https://pypi.org/project/pycelerate/)

Build Excel formulas with Python operators, then hand them straight to openpyxl.

```python
from pycelerate import Sheet, F

s = Sheet(ws)
rev  = s.put("B2", 1000)                 # writes the value, hands back the ref
cost = s.put("B3", 600)

s.put("B4", rev - cost)                  # =B2-B3
s.put("B5", (rev - cost) / rev)          # =(B2-B3)/B2
s.put("B6", F.IF(rev.gt(cost), "profit", "loss"))
```

No tracking which cell a value went to, no `f"={a}-{b}"` string assembly, and no
guessing about parentheses.

## Why

Four problems, in the order they actually bite.

1. **Coordinates drift.** Hand-written formulas name each cell twice — once where
   the value is written, once where it is referenced — and nothing connects the
   two. `s.put()` returns a reference to where the value actually landed, so the
   coordinate is written once and never typed again.
2. **Functions added in Excel 2010 or later need a prefix you have never heard
   of.** The file format stores them as `_xlfn.XLOOKUP`, `_xlfn._xlws.FILTER`.
   Write `=XLOOKUP(...)` bare and the workbook opens showing `#NAME?`, with no
   error at write time. `F.XLOOKUP(...)` emits the prefixed form automatically.
3. **Precedence goes wrong silently.** Naive concatenation turns
   `(revenue - cost) / revenue` into `=B3-B4/B3` — a wrong number, not an error.
   pycelerate renders an expression tree with Excel's own precedence rules.
4. **You have to open Excel to remember the arguments.** Does `SUMIFS` take the
   sum range first or last? A formula written as a string gives your editor
   nothing to work with. `F.` completes 263 function names, and the 85 you reach
   for most carry full signatures — so the argument order shows up as you type,
   and the wrong count is flagged before you run anything.

If you are writing computed **values** into a spreadsheet — `df.to_excel(...)` and
friends — you need none of this. It is for workbooks that must contain **live
formulas**, because a human will open the file and change the inputs.

## Install

```bash
uv add pycelerate      # or: pip install pycelerate
```

Python 3.11+. Zero runtime dependencies — nothing here imports openpyxl.

## Documentation

**<https://dunlopWill.github.io/pycelerate/>**

- [Installation](https://dunlopWill.github.io/pycelerate/installation/)
- [Usage](https://dunlopWill.github.io/pycelerate/usage/) — writing values, references, operators, functions
- [Validation](https://dunlopWill.github.io/pycelerate/validation/) — what is checked, when, and what nothing can check
- [API Reference](https://dunlopWill.github.io/pycelerate/api/)

## Development

```bash
uv sync
just check                # ruff, pyright, ty, pytest
just fast                 # skip the LibreOffice recalculation tests
just recalc               # only those (needs `soffice` on PATH)
just testall              # every supported Python version
```

`just check` is the gate CI runs. See [CONTRIBUTING.md](CONTRIBUTING.md) and
[CLAUDE.md](CLAUDE.md).

## License

MIT — see [LICENSE](LICENSE).
