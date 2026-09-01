"""Excel worksheet functions.

``F`` builds a call for any function name, so nothing here limits which functions
you can use::

    F.SUM(rng("B3:B10"))
    F.XLOOKUP(key, lookup_range, return_range)
    F.SOME_ADDIN_FUNCTION(a, b)

Functions added in Excel 2010 or later are stored in the file with an ``_xlfn.``
prefix; written bare they show as ``#NAME?`` when the workbook is opened.  That
rewriting happens automatically, so you always write the plain name.
"""

from __future__ import annotations

from .expr import Expr, P_ATOM, lit

__all__ = ["Func", "F"]

# Functions added in Excel 2010 and later, which the file format stores prefixed.
# Taken from XlsxWriter's table, which tracks the Microsoft list.
_XLFN = {
    "ACOT", "ACOTH", "AGGREGATE", "ANCHORARRAY", "ARABIC", "ARRAYTOTEXT", "BASE",
    "BETA.DIST", "BETA.INV", "BINOM.DIST", "BINOM.DIST.RANGE", "BINOM.INV", "BITAND",
    "BITLSHIFT", "BITOR", "BITRSHIFT", "BITXOR", "BYCOL", "BYROW", "CEILING.MATH",
    "CEILING.PRECISE", "CHISQ.DIST", "CHISQ.DIST.RT", "CHISQ.INV", "CHISQ.INV.RT",
    "CHISQ.TEST", "CHOOSECOLS", "CHOOSEROWS", "COMBINA", "CONCAT", "CONFIDENCE.NORM",
    "CONFIDENCE.T", "COT", "COTH", "COVARIANCE.P", "COVARIANCE.S", "CSC", "CSCH",
    "DAYS", "DECIMAL", "DROP", "ERF.PRECISE", "ERFC.PRECISE", "EXPAND", "EXPON.DIST",
    "F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST", "FILTERXML", "FLOOR.MATH",
    "FLOOR.PRECISE", "FORECAST.ETS", "FORECAST.ETS.CONFINT", "FORECAST.ETS.SEASONALITY",
    "FORECAST.ETS.STAT", "FORECAST.LINEAR", "FORMULATEXT", "GAMMA", "GAMMA.DIST",
    "GAMMA.INV", "GAMMALN.PRECISE", "GAUSS", "HSTACK", "HYPGEOM.DIST", "IFNA", "IFS",
    "IMAGE", "IMCOSH", "IMCOT", "IMCSC", "IMCSCH", "IMSEC", "IMSECH", "IMSINH", "IMTAN",
    "ISFORMULA", "ISOMITTED", "ISOWEEKNUM", "LAMBDA", "LET", "LOGNORM.DIST",
    "LOGNORM.INV", "MAKEARRAY", "MAP", "MAXIFS", "MINIFS", "MODE.MULT", "MODE.SNGL",
    "MUNIT", "NEGBINOM.DIST", "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV",
    "NUMBERVALUE", "PDURATION", "PERCENTILE.EXC", "PERCENTILE.INC", "PERCENTRANK.EXC",
    "PERCENTRANK.INC", "PERMUTATIONA", "PHI", "POISSON.DIST", "QUARTILE.EXC",
    "QUARTILE.INC", "QUERYSTRING", "RANDARRAY", "RANK.AVG", "RANK.EQ", "REDUCE", "RRI",
    "SCAN", "SEC", "SECH", "SEQUENCE", "SHEET", "SHEETS", "SINGLE", "SKEW.P", "SORTBY",
    "STDEV.P", "STDEV.S", "SWITCH", "T.DIST", "T.DIST.2T", "T.DIST.RT", "T.INV",
    "T.INV.2T", "T.TEST", "TAKE", "TEXTAFTER", "TEXTBEFORE", "TEXTJOIN", "TEXTSPLIT",
    "TOCOL", "TOROW", "UNICHAR", "UNICODE", "UNIQUE", "VALUETOTEXT", "VAR.P", "VAR.S",
    "VSTACK", "WEBSERVICE", "WEIBULL.DIST", "WRAPCOLS", "WRAPROWS", "XLOOKUP", "XMATCH",
    "XOR", "Z.TEST",
}

# Two dynamic-array functions take a further worksheet-class prefix.
_XLWS = {
    "FILTER", "SORT",
}


def stored_name(name: str) -> str:
    """The name as it must appear inside the xlsx file."""
    upper = name.upper()
    if upper in _XLWS:
        return "_xlfn._xlws." + upper
    if upper in _XLFN:
        return "_xlfn." + upper
    return name


class Func(Expr):
    """A worksheet function call.

    ``None`` renders as an omitted argument, so ``Func("VLOOKUP", a, b, c, None)``
    gives ``VLOOKUP(a,b,c,)``.
    """

    name: str
    args: tuple

    def __new__(cls, name: str, *args) -> "Func":
        parts = tuple(lit(a) for a in args)
        text = f"{stored_name(name)}({','.join(p.text for p in parts)})"
        obj = cls._make(text, P_ATOM)
        obj.name, obj.args = name, parts
        return obj

    def shift(self, rows: int = 0, cols: int = 0) -> "Func":
        return Func(self.name, *(a.shift(rows, cols) for a in self.args))


class _Functions:
    """Attribute access builds a call: ``F.SUM(a, b)`` -> ``SUM(a,b)``."""

    __slots__ = ()

    def __getattr__(self, name: str):
        if name.startswith("_"):
            raise AttributeError(name)

        def call(*args) -> Func:
            return Func(name.replace("__", "."), *args)

        call.__name__ = name
        call.__qualname__ = f"F.{name}"
        call.__doc__ = f"Build an Excel {name.replace('__', '.')}(...) call."
        return call

    def __dir__(self):
        return sorted(_XLFN | _XLWS | _COMMON)

    def __repr__(self) -> str:
        return "<pycelerate.F: Excel worksheet functions>"


# Only for tab-completion in a REPL -- any name works whether it is listed or not.
_COMMON = {
    "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "AVERAGE", "AVERAGEIF", "AVERAGEIFS",
    "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS", "MIN", "MAX", "ROUND",
    "ROUNDUP", "ROUNDDOWN", "ABS", "MOD", "INT", "IF", "IFERROR", "IFNA", "AND",
    "OR", "NOT", "VLOOKUP", "HLOOKUP", "INDEX", "MATCH", "OFFSET", "INDIRECT",
    "TEXT", "LEN", "LEFT", "RIGHT", "MID", "TRIM", "UPPER", "LOWER", "SUBSTITUTE",
    "FIND", "SEARCH", "TODAY", "NOW", "DATE", "YEAR", "MONTH", "DAY", "EOMONTH",
    "EDATE", "DATEDIF", "NPV", "IRR", "XNPV", "XIRR", "PMT", "PV", "FV", "RATE",
    "NPER", "ISBLANK", "ISNUMBER", "ISTEXT", "ISERROR", "NA", "ROW", "COLUMN",
}

F = _Functions()
