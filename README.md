# pycelerate

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

Three problems, in the order they actually bite.

**1. Coordinates drift.** Hand-written formulas name cells twice — once where the
value is written, once where it is referenced — and nothing connects the two:

```python
ws["B2"] = 1000
ws["B3"] = 600
ws["B4"] = "=B2-B3"
ws["B5"] = "=(B2-B3)/B2"
```

Insert one header row at the top of the generating code and every string below is
quietly pointing at the wrong cell. `s.put()` returns a reference to where the value
actually landed, so the coordinate is written once and never typed again. This is the
one that costs you time on an ordinary Tuesday.

**2. Functions added in Excel 2010 or later need a prefix you have never heard of.** The
file format stores them as `_xlfn.XLOOKUP`, `_xlfn.TEXTJOIN`, `_xlfn._xlws.FILTER`.
Write `=XLOOKUP(...)` bare through openpyxl and the workbook opens showing `#NAME?`,
with no error at write time and no hint as to why. `F.XLOOKUP(...)` emits the
prefixed form automatically — you write the plain name.

**3. Precedence goes wrong silently.** The obvious version of "just concatenate the
coordinates" turns `(revenue - cost) / revenue` into `=B3-B4/B3` — a wrong number,
not an error. pycelerate builds a small expression tree and renders it with Excel's
precedence rules, so the parentheses are right and there are no redundant ones.

It handles two places where Python and Excel genuinely disagree:

| You write | Python means | Excel would read | pycelerate emits |
|---|---|---|---|
| `a ** b ** c` | `a ** (b ** c)` | `^` is left-associative | `a^(b^c)` |
| `-a ** 2` | `-(a ** 2)` | unary `-` binds tighter than `^` | `-(a^2)` |

## Is this for you?

If you are writing computed **values** into a spreadsheet — `df.to_excel(...)` and
friends — you need none of this. No formula ever appears, so none of the above can
go wrong.

This is for workbooks that must contain **live formulas**, because a human will open
the file and change the inputs: financial models, budget templates, audit
deliverables, anything where "the numbers recalculate" is the point.

## Install

```bash
uv add pycelerate      # or: pip install pycelerate
```

Zero runtime dependencies — nothing here imports openpyxl. It works with openpyxl
because expressions *are* strings (see below), not because of any integration code.
That also means it works anywhere else a formula string is accepted.

## Writing values

`Sheet` wraps an openpyxl worksheet so a write hands back a reference:

```python
s = Sheet(ws)
rev = s.put("B2", 1000)          # or s.put(2, 2, 1000)
s.put("B4", rev * 1.2)           # =B2*1.2
```

It holds no state beyond the worksheet, and `s.ws` is the plain openpyxl object
whenever you want it. Formatting passes straight through — `number_format` plus any
keyword you could set on a cell:

```python
s.put("B5", margin, number_format="0.0%", font=Font(bold=True))
```

References come back **unqualified**, because most formulas stay on one sheet. When
one crosses sheets, `.qualified()` fills in the sheet name so you don't retype it:

```python
rev = Sheet(source).put("B2", 1000)
summary["A1"] = rev.qualified() * 2      # ='Source Data'!B2*2
```

`s["B2"]`, `s["A1:A9"]` and `s.cell(2, 2)` give references without writing anything.
`s.span(top, bottom)` and `s.column(rev)` build ranges from references you already
have.

## References

```python
cell("B3")                       # B3
cell(3, 2)                       # B3           -- (row, col), 1-based
cell("B3").abs()                 # $B$3         -- also .abs_r() / .abs_c() / .rel()
cell("B3").offset(rows=2)        # B5
cell("B3").to(cell("D5"))        # B3:D5

rng("B3:B10")                    # B3:B10
rng("A:A")                       # whole column -- "3:3" for a whole row
rng("B3:B10").sum()              # SUM(B3:B10)  -- .average() .count() .counta() .min() .max()
```

References are **relative by default**, so Excel's own reference fix-up behaves
normally when someone later edits the workbook by hand.

Sheet qualification is always explicit — never inferred:

```python
data = Ref(ws_source)            # an openpyxl worksheet, or just a name string
data["B3"]                       # 'Source Data'!B3
data["A:A"]                      # 'Source Data'!A:A
cell("B3", sheet="Data")         # Data!B3
```

## Operators

| Python | Excel | |
|---|---|---|
| `+ - * /` | `+ - * /` | |
| `**` | `^` | |
| `&` | `&` | string concatenation, **not** boolean AND — use `F.AND(...)` |
| `-x` | `-x` | |
| `x.pct()` | `x%` | |

Comparisons are **methods**, not operators: `.eq() .ne() .lt() .le() .gt() .ge()`.
That keeps `==` and `hash()` behaving like normal Python, so references stay usable
as dict keys and in sets. It is a real ergonomic cost in condition-heavy models,
paid deliberately. Reflected operators work too, so `1000 - cost` is fine.

## Functions

`F` builds a call for any name — nothing limits you to a listed set:

```python
F.SUM(rng("B3:B10"))
F.VLOOKUP(key, rng("A:D"), 3, False)
F.SOME_ADDIN_FUNCTION(a, b)
F.STDEV__S(rng("A1:A9"))         # double underscore becomes a dot: STDEV.S
```

The `_xlfn.` rewriting described above happens here:

```python
F.XLOOKUP(key, lookup, result)   # =_xlfn.XLOOKUP(...)
F.FILTER(rng("A:A"), cond)       # =_xlfn._xlws.FILTER(...)
```

`None` renders as an omitted argument: `F.VLOOKUP(a, b, 3, None)` → `VLOOKUP(a,b,3,)`.

`Raw("...")` is the escape hatch for formula text pycelerate does not model. It is
parenthesised whenever nested, since its structure is unknown — pass `atomic=True`
if it is a single self-contained term.

## Two things worth knowing

**Expressions are `str` subclasses.** That is what lets you assign one to a cell with
no conversion step. The string value is the finished formula:

```python
str(rev - cost)          # "=B3-B4"
(rev - cost).text        # "B3-B4"   -- the body, for building strings yourself
```

The trade-off: `f"Total: {ref}"` gives `"Total: =B3"`, and `isinstance(expr, str)` is
true. Use `.text` when you want the body.

**openpyxl does not update formulas when you insert or delete rows.** From its
[docs](https://openpyxl.readthedocs.io/en/stable/editing_worksheets.html): *"Openpyxl
does not manage dependencies, such as formulae, tables, charts, etc., when rows or
columns are inserted or deleted."* Excel fixes references only when a **human** edits
the workbook; a file written by `ws.delete_rows(3)` just contains stale strings.

pycelerate does not change this. Settle your data in Python, then write. If you must
move things afterwards, `.shift()` re-points every relative reference in an
expression and leaves `$`-locked ones alone:

```python
growth = F.ROUND((cell("B3") - cell("B4")) / cell("$A$1"), 2)
growth.shift(rows=-1)            # =ROUND((B2-B3)/$A$1,2)
```

## Development

```bash
uv sync
uv run pytest
uv run pytest tests/test_precedence.py    # one file
```
