# Installation

## From PyPI

```bash
uv add pycelerate      # or: pip install pycelerate
```

Zero runtime dependencies — nothing here imports openpyxl. It works with openpyxl
because [expressions *are* strings](usage.md#two-things-worth-knowing), not because
of any integration code. That also means it works anywhere else a formula string is
accepted.

## Supported Python versions

pycelerate runs on **Python 3.11 and later**, and is tested on 3.11, 3.12, 3.13 and
3.14 on every commit.

openpyxl is not required to install or import pycelerate. You need it only to write
the workbook that the formulas go into.

## From source

```bash
git clone https://github.com/dunlopWill/pycelerate
cd pycelerate
uv sync
```

That installs the dev group (ruff, pyright, ty, pytest, coverage, openpyxl). See
[Contributing](https://github.com/dunlopWill/pycelerate/blob/main/CONTRIBUTING.md)
for the local workflow.
