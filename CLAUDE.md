# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
uv sync                                   # install (dev group: openpyxl, pytest)
uv run pytest                             # full suite (103 tests, <1s)
uv run pytest tests/test_precedence.py    # one file
uv run pytest -k "parens"                 # one pattern
uv run pytest tests/test_refs.py::test_name
```

Requires Python >=3.14. No linter or formatter is configured.

## Architecture

A zero-runtime-dependency library that builds Excel formula strings from Python
operators. Four modules under `src/pycelerate/`, layered bottom-up:

`expr.py` → `refs.py` → `functions.py` → `sheet.py`

**The one invariant holding it together:** every node is a `str` subclass whose
string value is the *complete* `=`-prefixed formula, while `.text` is the body
without the `=`. Composition always reads `.text`; only the final assignment to a
cell uses the string value. That is what makes `ws["B5"] = rev - cost` work with no
conversion step and no openpyxl integration code.

**`expr.py`** — `Expr` base plus `BinOp`, `UnaryOp`, `Lit`, `Raw`. Each node
carries `_prec` (the `P_*` constants, loosest→tightest). `Expr._operand()` is the
whole parenthesisation policy: parenthesise when the child binds looser than the
parent, plus when an equal-precedence *right* operand belongs to a
non-associative operator (`_RIGHT_NEEDS_PARENS` = `- / ^`). `+ * &` are
deliberately excluded so output stays readable. Two places where Excel and Python
disagree are handled here, not by the caller: `^` is left-associative in Excel, and
unary `-` binds *tighter* than `^`.

**`refs.py`** — `CellRef` / `RangeRef` / the `Ref` sheet factory. References are
relative by default (so Excel's own fix-up works when a human later edits the
file) and sheet qualification is never inferred. Each ref carries both `sheet` (the
prefix actually rendered) and `home` (the sheet it was written to, never rendered);
`home` exists only so `.qualified()` can supply a sheet name without retyping it.

**`functions.py`** — `F.ANYNAME(...)` builds a `Func` for any name; there is no
allowlist (`_COMMON` exists purely for REPL tab-completion). `stored_name()`
applies the `_xlfn.` / `_xlfn._xlws.` prefixes that the xlsx format requires for
Excel-2010-and-later functions — omit them and the workbook shows `#NAME?`. `__`
in an attribute name becomes `.` (`F.STDEV__S` → `STDEV.S`).

**`sheet.py`** — `Sheet.put()` writes a value through openpyxl and hands back a
`CellRef`, so no coordinate is typed twice. Holds no state beyond the worksheet.

**openpyxl is never imported** anywhere in `src/`, including by `Sheet`. Worksheets
are duck-typed on `.title` and `.cell(row=, column=, value=)`. Keep it that way —
it is the reason `dependencies = []`. openpyxl is a dev-only dep used by
`tests/test_openpyxl_roundtrip.py`, which guards it with `importorskip`.

## Conventions

- Comparisons are methods (`.eq() .lt() .gt()`…), not operator overloads, so `==`
  and `hash()` keep native `str` behaviour and refs stay usable as dict keys.
- `&` is Excel string concatenation, not boolean AND. `%` raises `TypeError`
  (it would otherwise be Python string formatting); use `F.MOD` or `.pct()`.
- `shift()` is implemented on every node type and must recurse — it re-points
  relative refs and leaves `$`-locked ones alone.
- Nodes are constructed in `__new__` (they are `str` subclasses), attributes
  assigned on the object after `cls._make(text, prec)`.
- Tests are organised by concern, not by module; `test_precedence.py` is the
  load-bearing one — wrong parens give silently wrong numbers rather than errors.
