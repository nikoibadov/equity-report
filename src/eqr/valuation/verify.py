"""Recalculate the written workbook with the `formulas` engine (no Excel/LibreOffice needed)
and check the live formulas reproduce the Python engine."""
from __future__ import annotations

from pathlib import Path


def _val(sol: dict, book: str, sheet: str, cell: str):
    key = f"'[{book}]{sheet.upper()}'!{cell}"
    for k, v in sol.items():
        if k.upper() == key.upper():
            x = v.value[0][0] if hasattr(v, "value") else v
            return float(x)
    raise KeyError(key)


def check_workbook(path: Path, res: dict, tol: float = 0.01) -> tuple[bool, str]:
    import formulas
    xl = formulas.ExcelModel().loads(str(path)).finish()
    sol = xl.calculate()
    book = path.name
    d = res["dcf"]
    checks = {
        "WACC": (_val(sol, book, "WACC", "B15"), d["wacc"], 1e-6),
        "EV": (_val(sol, book, "DCF", "B23"), d["ev"], 0.5),
        "Equity": (_val(sol, book, "DCF", "B30"), d["equity"], 0.5),
        "DCF price": (_val(sol, book, "DCF", "B34"), d["price"], tol),
    }
    lines, ok = [], True
    for k, (xl_v, py_v, t) in checks.items():
        good = abs(xl_v - py_v) <= t
        ok &= good
        lines.append(f"  {'OK ' if good else 'BAD'} {k:10s} excel={xl_v:,.4f} python={py_v:,.4f}")
    return ok, f"Excel recalculation check ({book}):\n" + "\n".join(lines)
