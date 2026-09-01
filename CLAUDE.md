# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
uv sync                                   # install (dev group: openpyxl, pytest)
uv run pytest                             # full suite (147 tests, ~6s)
uv run pytest -m "not slow"               # skip LibreOffice recalc (~0.5s)
uv run pytest -m slow                     # only the recalc tests
uv run pytest tests/test_precedence.py    # one file
uv run pytest -k "parens"                 # one pattern
uvx pyright src tests                     # type stubs (see functions.pyi)
```

Requires Python >=3.14. No linter or formatter is configured. `pyright src tests`
is clean and should stay that way — the package ships `py.typed`, so its own type
errors would surface in consumers' editors.

## Architecture

A zero-runtime-dependency library that builds Excel formula strings from Python
operators. Four modules under `src/pycelerate/`, layered bottom-up:

`expr.py` → `refs.py` → `functions.py` → `sheet.py`

`Expr._make()` returns `Self`, not `Expr`. That is load-bearing for typing: it is
what lets each subclass assign its own attributes to the result and return it as
its own type. Changing it back to `Expr` reintroduces ~26 type errors across the
package.

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
allowlist. `stored_name()` applies the `_xlfn.` / `_xlfn._xlws.` prefixes that the
xlsx format requires for Excel-2010-and-later functions — omit them and the
workbook shows `#NAME?`. `__` in an attribute name becomes `.` (`F.STDEV__S` →
`STDEV.S`).

`_check_name()` runs two checks at construction, both governed by one rule: **a
name in a table is checked, a name outside every table is left alone**, because the
set of real Excel functions is open (add-ins, newer releases). Arity mismatches on
names in `_ARITY` raise `TypeError`; an unknown name close to a known one emits
`UnknownFunctionWarning` and still renders. `check=False` bypasses both, and
`Func.shift()` passes it because shifting cannot change arity.

When adding to `_ARITY`, **widen bounds when unsure** — a false rejection blocks
correct code, a miss only restores the status quo. Optional trailing arguments are
counted, so the ranges track Excel's own `[...]` notation. Dotted names must be
spelled as `_XLFN` spells them or the lookup silently misses;
`test_every_arity_entry_is_actually_wired_up` catches that by exercising every
entry.

**`functions.pyi`** — stubs moving the same arity check to edit time. A `.pyi`
*replaces* the `.py` for type checkers, so anything importable must be declared
here too (`_XLFN`, `_ARITY`, `stored_name` are imported by tests). Two rules:
reference positions are typed `Expr` (so a bare `"A:A"` string is rejected —
it would render as a text literal); **everything else stays `_Value`**, because
`None` is a legal omitted argument anywhere and values may arrive as computed
expressions. Narrowing a parameter to `int`/`bool`/`str` breaks both.

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
- `CellRef._replace()` uses a `_KEEP` sentinel rather than `None` for "unchanged",
  because `sheet=None` is a real value meaning "no prefix" (what `.bare()` asks
  for). Keep the parameters explicit — the old `**kw` version type-checked nothing.
- `RangeRef.count()` deliberately shadows `str.count`, so that one subclass raises
  on `expr.count(substring)`. It carries a pyright ignore; renaming it would be a
  breaking API change.
- Tests are organised by concern, not by module; `test_precedence.py` is the
  load-bearing one — wrong parens give silently wrong numbers rather than errors.
- `test_recalc.py` evaluates generated workbooks in headless LibreOffice — the only
  check that catches formulas which are well-formed but *mean* the wrong thing.
  It carries a negative control (`test_the_harness_can_actually_detect_a_bad_formula`);
  keep it, or the whole file could silently stop testing anything.
