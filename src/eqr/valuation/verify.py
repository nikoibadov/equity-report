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
        "Equity": (_val(sol, book, "DCF", "B32"), d["equity"], 0.5),
        "DCF price": (_val(sol, book, "DCF", "B36"), d["price"], tol),
    }
    if res.get("prec"):
        pm, wt = _label_cells(path)
        checks["Precedent"] = (_val(sol, book, "Precedents", pm), res["prec"]["prices"]["core"]["median"], tol)
        checks["Weighted"] = (_val(sol, book, "Precedents", wt), res["prec"]["weighted"], tol)
    lines, ok = [], True
    try:
        chk = [k for k in sol if "CHECKS" in k.upper() and k.upper().endswith("!B" + str(_last_check_row(path)))]
        allpass = bool(sol[chk[0]].value[0][0]) if chk else None
    except Exception:  # noqa: BLE001
        allpass = None
    if allpass is not None:
        ok &= allpass
        lines.append(f"  {'OK ' if allpass else 'BAD'} Checks sheet: all integrity checks {'pass' if allpass else 'do NOT pass'}")
    for k, (xl_v, py_v, t) in checks.items():
        good = abs(xl_v - py_v) <= t
        ok &= good
        lines.append(f"  {'OK ' if good else 'BAD'} {k:10s} excel={xl_v:,.4f} python={py_v:,.4f}")
    return ok, f"Excel recalculation check ({book}):\n" + "\n".join(lines)


def _last_check_row(path: Path) -> int:
    import openpyxl
    ws = openpyxl.load_workbook(path)["Checks"]
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=1).value == "All checks pass":
            return r
    raise KeyError("All checks pass")


def scenario_price(path: Path, selector: int, tmp: Path) -> float:
    """Set the scenario selector in a copy of the workbook, recalculate and return the DCF value per share."""
    import formulas
    import openpyxl
    wb = openpyxl.load_workbook(path)
    ws = wb["Assumptions"]
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(row=r, column=2).value or "").startswith("Active scenario"):
            ws.cell(row=r, column=4).value = selector
    out = tmp / f"scenario_{selector}.xlsx"
    wb.save(out)
    sol = formulas.ExcelModel().loads(str(out)).finish().calculate()
    return _val(sol, out.name, "DCF", "B36")


def _label_cells(path: Path) -> tuple[str, str]:
    """Cells of the precedent median price and the weighted value, found by their row labels."""
    import openpyxl
    ws = openpyxl.load_workbook(path)["Precedents"]
    med = wt = None
    for r in range(1, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        if a == "Core set: median":
            med = f"E{r}"
        elif a == "Weighted value per share":
            wt = f"D{r}"
    return med, wt
