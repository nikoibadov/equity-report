"""Formula-driven valuation workbook for any listed company, built on the SAP case-study template.

Same layout and conventions as the SAP workbook (excel_writer.py): Cover, Summary, Assumptions with sources and a live
Bear/Base/Bull selector, WACC, Forecast, DCF, DDM, Checks, Scenarios, Sensitivity, Monte Carlo, Comps, Historicals,
FCFF bridge, Sources & notes. Live formulas run Assumptions -> WACC -> Forecast -> DCF -> DDM -> Comps -> Summary;
the engine values (grids, scenarios, Monte Carlo runs) are labelled as such. Sheets whose data a ticker does not have
(no dividend, no peers, no statements) are left out rather than filled with placeholders. SAP-only material
(segment build, precedent transactions, intangibles, sum of the parts) has no public equivalent for an arbitrary
ticker and is not part of this template.

`verify()` recalculates the saved workbook with the `formulas` engine and fails if it does not reproduce generic.py."""
from __future__ import annotations

import math
import tempfile
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as L

from . import generic as G
from .excel_writer import (BOX, EUR, EUR1, F_B, F_H, F_IN, F_LINK, FILL_H, FILL_K, FILL_S, MULT, PCT, PCT2, TEAL, _hdr,
                           _title, _widths)

M = 1e6
GREY = Font(italic=True, color="8091A5")
SMALL = Font(color="8091A5", size=9)
NOTE = Font(italic=True, color="5B6B7F")
YC = {t: L(3 + t) for t in range(0, G.N_YEARS + 1)}          # FY0 -> C, FY1 -> D ... FY10 -> M
SYM = {"EUR": "€", "USD": "$", "GBP": "£", "JPY": "¥", "CHF": "CHF ", "CNY": "CN¥", "KRW": "₩", "INR": "₹", "CAD": "C$",
       "AUD": "A$", "HKD": "HK$", "SEK": "SEK ", "DKK": "DKK ", "NOK": "NOK ", "TWD": "NT$", "BRL": "R$", "SGD": "S$"}
DEFAULT_WEIGHTS = {"DCF": 0.50, "Comparable companies": 0.25, "Precedent transactions": 0.25}


def _ok(v) -> bool:
    return v is not None and not (isinstance(v, float) and math.isnan(v))


class Ctx:
    def __init__(self, res: dict, weights: dict | None):
        self.res, self.p, self.d = res, res["inputs"], res["dcf"]
        p = self.p
        self.sym = SYM.get(p["q_ccy"], p["q_ccy"] + " ")
        self.PX = f'"{self.sym}"#,##0.00'
        self.PX0 = f'"{self.sym}"#,##0'
        self.m = f"{p['fin_ccy']}m"
        self.fy0 = int(p.get("fy0") or 0)
        self.vd = res.get("meta", {}).get("valuation_date") or p.get("valuation_date") or res["market"].get("price_date", "")
        w = dict(weights or DEFAULT_WEIGHTS)
        self.has_comps = bool(res["comps"].get("stats"))
        avail = {"DCF": w.get("DCF", 0.0), **({"Comparable companies": w.get("Comparable companies", 0.0)} if self.has_comps else {})}
        self.weights = avail
        self.ref: dict = {}

    def fy(self, t: int, suffix: str = "") -> str:
        return f"FY{self.fy0 + t}{suffix}" if self.fy0 else f"Y{t}{suffix}"


def _put(ws, cell, v, fmt=None, font=None):
    c = ws[cell]
    c.value = (float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) and _ok(v) else (v if _ok(v) else None))
    if fmt:
        c.number_format = fmt
    if font:
        c.font = font
    return c


# ======================================================================= Assumptions
def _assumptions(wb, c: Ctx):
    ws = wb.create_sheet("Assumptions")
    p, S, mk = c.p, c.p.get("sources", {}), c.res["market"]
    _title(ws, "Assumptions (blue = input; edit here and the model recalculates)",
           f"{p['name']} ({p['ticker']}) | valuation date {c.vd} | {c.m} unless stated; per-share values in {p['q_ccy']}")
    _hdr(ws, 4, ["Key", "Assumption", "Unit"] + [c.fy(t, "E") for t in range(1, 6)] + ["Terminal", "Source", "Rationale"])
    bsd, m = p.get("bs_date") or c.vd, c.m
    sections = [
        ("Market & capital structure", [
            ("price", "Share price (valuation date)", p["q_ccy"], p["price"], c.PX, "Yahoo Finance close", f"Close on {mk.get('price_date', c.vd)}"),
            ("shares", "Shares outstanding, all classes (m)", "m", p["shares"] / M, EUR1, "Yahoo Finance", S.get("shares", "")),
            ("fx", f"FX: {p['q_ccy']} per {p['fin_ccy']}", "x", p["fx"], "0.0000", "ECB reference rates / Yahoo Finance",
             "Converts equity value in the reporting currency into the quote currency (1 if they are the same)"),
            ("e_mv", "Market value of equity", m, "={price}*{shares}/{fx}", EUR, "Formula", "Price x shares / FX"),
            ("d_mv", "Debt incl. leases for WACC weights", m, "={debt}", EUR, "Link to the bridge", "Book value of debt as a proxy for its market value"),
        ]),
        ("Cost of capital inputs", [
            ("rf", "Risk-free rate (10y government bond)", "%", p["rf"], PCT2, S.get("rf", ""), "Long government yield in the currency of the cash flows, on the valuation date"),
            ("beta", "Levered beta", "x", p["beta"], "0.00", S.get("beta", ""), "Regression beta shrunk towards 1 (Blume)"),
            ("erp", "Equity risk premium", "%", p["erp"], PCT2, S.get("erp", ""), "Same house view as the SAP model"),
            ("kd", "Pre-tax cost of debt", "%", p["kd"], PCT2, S.get("kd", ""), "Risk-free rate plus a default spread for the implied rating"),
            ("tax_wacc", "Tax rate for the debt tax shield", "%", p["tax_wacc"], PCT2, S.get("tax_terminal", ""), "Marginal (statutory) rate"),
        ]),
        (f"Base year {c.fy(0)} (actual)", [
            ("rev0", "Revenue", m, p["rev0"] / M, EUR, "Yahoo Finance annual income statement", f"Fiscal year ended {p.get('fy0_date') or 'n/a'}"),
        ]),
        ("Revenue growth (FY6-FY10 fade linearly to terminal growth)", [
            ("growth", "Revenue growth", "%", list(p["growth"]), PCT, "Consensus / Yahoo Finance", S.get("growth", "")),
        ]),
        ("Profitability", [
            ("margin0", f"EBIT margin {c.fy(1)}", "%", p["margin0"], PCT, "Yahoo Finance", S.get("margin", "")),
            ("margin_target", f"Target EBIT margin (reached in {c.fy(5)}, held after)", "%", p["margin_target"], PCT, "Analyst judgement", "Default = FY1 margin; edit for a margin story"),
            ("tax", "Tax rate on EBIT", "%", [p["tax"]] * 5, PCT, S.get("tax", ""), "Effective rate while the current tax structure holds", ("tax_terminal", p["tax_terminal"])),
        ]),
        ("Reinvestment & SBC", [
            ("da_pct", "D&A % of revenue", "%", [p["da_pct"]] * 5, PCT, S.get("reinvestment", ""), "Recent average", ("da_terminal", p["da_terminal"])),
            ("capex_pct", "Capex % of revenue", "%", [p["capex_pct"]] * 5, PCT, S.get("reinvestment", ""), "Recent average; terminal capex should cover D&A plus growth", ("capex_terminal", p["capex_terminal"])),
            ("nwc_pct", "Working-capital investment % of revenue increase", "%", p["nwc_pct"], PCT, S.get("nwc", ""), "Negative = working capital releases cash as revenue grows"),
            ("sbc_pct", "Stock-based compensation % of revenue", "%", p["sbc_pct"], PCT, "Yahoo Finance cash-flow statement", "Memo; used only if the add-back switch is 1"),
            ("sbc_addback", "Add back SBC as non-cash? (0/1)", "0/1", p["sbc_addback"], "0", "Model choice", S.get("sbc", "")),
        ]),
        ("Terminal value & timing", [
            ("g", "Terminal growth g", "%", p["g"], PCT2, "House view", "Long-run nominal growth, below the risk-free rate"),
            ("stub", f"Part of {c.fy(1)} still to come after {bsd}", "years", p["stub"], "0.000", "Fiscal calendar", S.get("timing", "")),
            ("roll_days", f"Days from balance sheet ({bsd}) to valuation date", "days", p["roll_days"], "0", "Calendar", "Value rolled forward at the cost of equity"),
        ]),
        (f"EV-to-equity bridge (balance sheet {bsd})", [
            ("cash", "Cash and short-term investments", m, p["cash"] / M, EUR, "Yahoo Finance quarterly balance sheet", S.get("bridge", "")),
            ("lt_inv", "Long-term financial investments", m, p["lt_inv"] / M, EUR, "Yahoo Finance quarterly balance sheet", "Marketable securities and investments outside operations"),
            ("other_nonop", "Other non-operating assets", m, p["other_nonop"] / M, EUR, "Filings (manual)", "E.g. tax-loss carryforwards, non-core stakes; 0 unless added"),
            ("debt", "Financial debt incl. leases", m, p["debt"] / M, EUR, "Yahoo Finance quarterly balance sheet", "Total debt as reported"),
            ("pension", "Pension deficit", m, p["pension"] / M, EUR, "Filings (manual)", "Not in Yahoo's data; add from the annual report"),
            ("nci", "Non-controlling interests", m, p["nci"] / M, EUR, "Yahoo Finance quarterly balance sheet", "Book value"),
            ("pref", "Preferred stock", m, p["pref"] / M, EUR, "Yahoo Finance quarterly balance sheet", "Book value"),
        ]),
    ]
    if c.res.get("ddm"):
        sections.append(("Dividend discount model (cross-check)", [
            ("dps0", "Indicated annual dividend per share", p["q_ccy"], p["dps0"], c.PX, "Yahoo Finance", S.get("dividend", "")),
            ("g_div", f"Dividend growth {c.fy(1)}-{c.fy(5)}", "%", p["g_div"], PCT, "Consensus EPS growth", "Fades to terminal growth in FY6-FY10"),
        ]))
    ref, formulas_ = c.ref, []
    r = 5
    for title, items in sections:
        ws.cell(row=r, column=1, value=title).font = F_B
        for col in range(1, 12):
            ws.cell(row=r, column=col).fill = FILL_S
        r += 1
        for it in items:
            k, lab, unit, v, fmt, src, rat = it[:7]
            term = it[7] if len(it) > 7 else None
            ws.cell(row=r, column=1, value=k).font = SMALL
            ws.cell(row=r, column=2, value=lab)
            ws.cell(row=r, column=3, value=unit)
            if isinstance(v, list):
                ref[k] = {}
                for t, x in enumerate(v, start=1):
                    _put(ws, f"{L(3 + t)}{r}", x, fmt, F_IN)
                    ref[k][t] = f"Assumptions!${L(3 + t)}${r}"
                if term:
                    _put(ws, f"I{r}", term[1], fmt, F_IN)
                    ref[term[0]] = f"Assumptions!$I${r}"
            elif isinstance(v, str) and v.startswith("="):
                formulas_.append((f"D{r}", v, fmt))
                ref[k] = f"Assumptions!$D${r}"
            else:
                _put(ws, f"D{r}", v, fmt, F_IN)
                ref[k] = f"Assumptions!$D${r}"
            ws.cell(row=r, column=10, value=src)
            ws.cell(row=r, column=11, value=rat).alignment = Alignment(wrap_text=True, vertical="top")
            r += 1
        r += 1
    for cell, f, fmt in formulas_:
        _put(ws, cell, f.format(**{k: v for k, v in ref.items() if isinstance(v, str)}), fmt, F_LINK)
    # ---- live scenario selector
    ws.cell(row=r, column=1, value="Scenario selector (drives Forecast, WACC, DCF and DDM)").font = F_B
    for col in range(1, 12):
        ws.cell(row=r, column=col).fill = FILL_S
    r += 1
    sel = f"Assumptions!$D${r}"
    ws.cell(row=r, column=2, value="Active scenario: 1 = Bear, 2 = Base, 3 = Bull")
    cc = ws[f"D{r}"]; cc.value = 2; cc.font = Font(bold=True, color="0000FF", size=12); cc.fill = FILL_K
    ws.cell(row=r, column=5, value=f'=CHOOSE($D${r},"Bear","Base","Bull")').font = F_B
    ws.cell(row=r, column=11, value="Type 1, 2 or 3. Base (2) reproduces the Python engine; the Scenarios sheet lists each case's value.")
    r += 1
    _hdr(ws, r, ["", "Scenario settings", "", "Bear", "Base", "Bull", "Active"])
    r += 1
    sc = G.SCENARIOS
    settings = {}
    for key, lab, fmt in (("d_growth", f"Revenue growth shift {c.fy(2)}-{c.fy(5)} (per year)", PCT), ("d_margin", "Target EBIT margin shift", PCT),
                          ("d_g", "Terminal growth shift", PCT2), ("d_wacc", "WACC shift", PCT2)):
        ws.cell(row=r, column=2, value=lab)
        for col, n in zip("DEF", ("Bear", "Base", "Bull")):
            _put(ws, f"{col}{r}", sc[n][key], fmt, F_IN)
        _put(ws, f"G{r}", f"=CHOOSE({sel},D{r},E{r},F{r})", fmt, F_B)
        ws[f"H{r}"] = "<- active"
        settings[key] = f"Assumptions!$G${r}"
        r += 1
    _hdr(ws, r, ["", "Active inputs used by the model", ""] + [c.fy(t) for t in range(1, 6)] + ["Terminal"])
    r += 1
    ws.cell(row=r, column=2, value="Revenue growth (active)")
    act_g = {}
    for t in range(1, 6):
        col = L(3 + t)
        _put(ws, f"{col}{r}", f"={ref['growth'][t]}" + ("" if t == 1 else f"+{settings['d_growth']}"), PCT)
        act_g[t] = f"Assumptions!${col}${r}"
    r += 1
    ws.cell(row=r, column=2, value="Target EBIT margin (active)")
    _put(ws, f"I{r}", f"={ref['margin_target']}+{settings['d_margin']}", PCT)
    act_mt = f"Assumptions!$I${r}"
    r += 1
    ws.cell(row=r, column=2, value="Terminal growth (active)")
    _put(ws, f"I{r}", f"={ref['g']}+{settings['d_g']}", PCT2)
    act_gt = f"Assumptions!$I${r}"
    ref.update(growth_act=act_g, mt_act=act_mt, g_act=act_gt, scen_wacc=settings["d_wacc"], scen_sel=sel)
    _widths(ws, {"A": 14, "B": 46, "C": 9, "D": 11, "E": 11, "F": 11, "G": 11, "H": 11, "I": 11, "J": 44, "K": 100})
    ws.freeze_panes = "D5"


# ======================================================================= WACC
def _wacc(wb, c: Ctx):
    ws = wb.create_sheet("WACC")
    ref, d, p, s = c.ref, c.d, c.p, c.res.get("snapshot", {})
    _title(ws, "Cost of capital (CAPM / market-value weights)")
    rows = [
        ("Risk-free rate", f"={ref['rf']}", PCT2), ("Levered beta", f"={ref['beta']}", "0.00"),
        ("Equity risk premium", f"={ref['erp']}", PCT2), ("Cost of equity  Ke = Rf + β·ERP", "=B4+B5*B6", PCT2),
        ("Pre-tax cost of debt", f"={ref['kd']}", PCT2), ("Tax rate (marginal)", f"={ref['tax_wacc']}", PCT2),
        ("After-tax cost of debt", "=B8*(1-B9)", PCT2), (f"Equity (market value, {c.m})", f"={ref['e_mv']}", EUR),
        (f"Debt incl. leases ({c.m})", f"={ref['d_mv']}", EUR), ("Equity weight", "=B11/(B11+B12)", PCT),
        ("Debt weight", "=B12/(B11+B12)", PCT), ("WACC", "=B13*B7+B14*B10", PCT2),
        ("Scenario shift (Assumptions selector)", f"={ref['scen_wacc']}", PCT2), ("WACC used in the DCF", "=B15+B16", PCT2),
    ]
    _hdr(ws, 3, ["Component", "Value", "Python check"])
    checks = {7: d["ke"], 10: d["kd_at"], 13: d["we"], 14: d["wd"], 15: d["wacc"]}
    for i, (lab, f, fmt) in enumerate(rows, start=4):
        ws.cell(row=i, column=1, value=lab)
        cc = ws.cell(row=i, column=2, value=f); cc.number_format = fmt; cc.font = F_LINK if "Assump" in f else Font()
        if i in checks:
            _put(ws, f"C{i}", checks[i], fmt, GREY)
    for a in ("A15", "B15", "A17", "B17"):
        ws[a].font = F_B
    ws["B15"].fill = FILL_K; ws["B17"].fill = FILL_K
    r = 20
    ws.cell(row=r, column=1, value="Reference points").font = F_B
    _hdr(ws, r + 1, ["Item", "Value", "Comment"])
    refs = [("Beta regression (raw)", s.get("beta_raw"), f"5y weekly returns vs {s.get('beta_index', '?')}, n={s.get('beta_n', 0)}"
             + (f", R² {s['beta_r2']:.2f}" if s.get("beta_r2") else "") + "; Blume adjustment 0.67 x raw + 0.33"),
            ("Effective tax rate (explicit years)", p["tax"], p.get("sources", {}).get("tax", "")),
            ("Statutory tax rate (terminal and tax shield)", p["tax_terminal"], p.get("sources", {}).get("tax_terminal", "")),
            ("Equity risk premium range", "4.2% - 6.5%", "Damodaran implied ERP (lower end) to the IDW/FAUB range (upper end); base 5.0% as in the SAP model"),
            ("Cost of debt build", p["kd"], p.get("sources", {}).get("kd", ""))]
    for i, (a, v, cmt) in enumerate(refs, start=r + 2):
        ws.cell(row=i, column=1, value=a)
        cc = ws.cell(row=i, column=2, value=float(v) if isinstance(v, (int, float)) else v)
        if isinstance(v, float):
            cc.number_format = PCT2 if v < 0.5 else "0.00"
        ws.cell(row=i, column=3, value=cmt)
    _widths(ws, {"A": 44, "B": 16, "C": 120})
    c.ref.update(ke="WACC!$B$7", wacc="WACC!$B$17", wacc_capm="WACC!$B$15", we="WACC!$B$13", wd="WACC!$B$14")


# ======================================================================= Forecast
def _forecast(wb, c: Ctx):
    ws = wb.create_sheet("Forecast")
    ref, p, f = c.ref, c.p, c.d["forecast"]
    _title(ws, f"Revenue and free cash flow to the firm ({c.m})",
           f"{c.fy(0)} = reported; {c.fy(1)}-{c.fy(5)} explicit; {c.fy(6)}-{c.fy(10)} linear fade to the terminal values")
    _hdr(ws, 4, ["", "Line item"] + [c.fy(t, "A" if t == 0 else "E") for t in range(0, G.N_YEARS + 1)])
    ws["A5"] = "year"
    for t in range(0, G.N_YEARS + 1):
        _put(ws, f"{YC[t]}5", c.fy0 + t if c.fy0 else t, None, SMALL)
    labels = [("revenue", "Revenue"), ("growth", "  growth y/y"), ("margin", "EBIT margin"), ("ebit", "EBIT (operating profit)"),
              ("tax_rate", "Tax rate on EBIT"), ("nopat", "NOPAT"), ("sbc", "+ SBC add-back (switch)"), ("da_pct", "D&A % revenue"),
              ("da", "+ D&A"), ("capex_pct", "Capex % revenue"), ("capex", "- Capex"), ("dnwc", "- Investment in working capital"),
              ("fcff", "Free cash flow to firm (FCFF)"), ("fcff_m", "  FCFF margin"), ("py_fcff", "Python engine check (FCFF)")]
    R = {k: i for i, (k, _) in enumerate(labels, start=6)}
    for k, lab in labels:
        ws.cell(row=R[k], column=2, value=lab)
    # FY0 actuals (memo)
    h = c.res.get("hist")
    if h is not None and len(h) and p.get("fy0_date") in h.index:
        a = h.loc[p["fy0_date"]]
        _put(ws, f"C{R['revenue']}", f"={ref['rev0']}", None, F_LINK)
        if _ok(a.get("ebit")):
            _put(ws, f"C{R['ebit']}", a["ebit"] / M, None, F_IN)
            _put(ws, f"C{R['margin']}", f"=C{R['ebit']}/C{R['revenue']}")
        if _ok(a.get("tax_rate")):
            _put(ws, f"C{R['tax_rate']}", a["tax_rate"], None, F_IN)
        if _ok(a.get("da")):
            _put(ws, f"C{R['da']}", a["da"] / M, None, F_IN); _put(ws, f"C{R['da_pct']}", f"=C{R['da']}/C{R['revenue']}")
        if _ok(a.get("capex")):
            _put(ws, f"C{R['capex']}", -a["capex"] / M, None, F_IN); _put(ws, f"C{R['capex_pct']}", f"=C{R['capex']}/C{R['revenue']}")
        prev = [d_ for d_ in h.index if d_ < p["fy0_date"]]
        if prev and _ok(h.loc[prev[-1], "revenue"]):
            _put(ws, f"C{R['growth']}", f"=C{R['revenue']}/{h.loc[prev[-1], 'revenue'] / M}-1")
    else:
        _put(ws, f"C{R['revenue']}", f"={ref['rev0']}", None, F_LINK)
    H = YC[5]
    for t in range(1, G.N_YEARS + 1):
        col, pc = YC[t], YC[t - 1]
        k = t - 5
        if t <= 5:
            _put(ws, f"{col}{R['growth']}", f"={ref['growth_act'][t]}", None, F_LINK)
            _put(ws, f"{col}{R['margin']}", f"={ref['margin0']}" if t == 1 else f"={ref['margin0']}+({ref['mt_act']}-{ref['margin0']})*{t - 1}/4", None, F_LINK)
            for row, key in (("tax_rate", "tax"), ("da_pct", "da_pct"), ("capex_pct", "capex_pct")):
                _put(ws, f"{col}{R[row]}", f"={ref[key][t]}", None, F_LINK)
        else:
            _put(ws, f"{col}{R['growth']}", f"=${H}{R['growth']}+({ref['g_act']}-${H}{R['growth']})*{k}/5")
            _put(ws, f"{col}{R['margin']}", f"={ref['mt_act']}", None, F_LINK)
            for row, key in (("tax_rate", "tax_terminal"), ("da_pct", "da_terminal"), ("capex_pct", "capex_terminal")):
                _put(ws, f"{col}{R[row]}", f"=${H}{R[row]}+({ref[key]}-${H}{R[row]})*{k}/5")
        _put(ws, f"{col}{R['revenue']}", f"={pc}{R['revenue']}*(1+{col}{R['growth']})")
        _put(ws, f"{col}{R['ebit']}", f"={col}{R['revenue']}*{col}{R['margin']}")
        _put(ws, f"{col}{R['nopat']}", f"={col}{R['ebit']}*(1-{col}{R['tax_rate']})")
        _put(ws, f"{col}{R['sbc']}", f"={col}{R['revenue']}*{ref['sbc_pct']}*{ref['sbc_addback']}")
        _put(ws, f"{col}{R['da']}", f"={col}{R['revenue']}*{col}{R['da_pct']}")
        _put(ws, f"{col}{R['capex']}", f"={col}{R['revenue']}*{col}{R['capex_pct']}")
        _put(ws, f"{col}{R['dnwc']}", f"={ref['nwc_pct']}*({col}{R['revenue']}-{pc}{R['revenue']})")
        _put(ws, f"{col}{R['fcff']}", f"={col}{R['nopat']}+{col}{R['sbc']}+{col}{R['da']}-{col}{R['capex']}-{col}{R['dnwc']}")
        _put(ws, f"{col}{R['fcff_m']}", f"={col}{R['fcff']}/{col}{R['revenue']}")
        _put(ws, f"{col}{R['py_fcff']}", float(f.iloc[t - 1]["fcff"]) / M, None, GREY)
    pct_rows = {R[k] for k in ("growth", "margin", "tax_rate", "da_pct", "capex_pct", "fcff_m")}
    for row in range(6, R["py_fcff"] + 1):
        for t in range(0, G.N_YEARS + 1):
            ws[f"{YC[t]}{row}"].number_format = PCT if row in pct_rows else EUR
    for k in ("revenue", "ebit", "fcff"):
        for col in ["B"] + [YC[t] for t in range(0, G.N_YEARS + 1)]:
            ws[f"{col}{R[k]}"].font = F_B
    for t in range(1, G.N_YEARS + 1):
        ws[f"{YC[t]}{R['fcff']}"].fill = FILL_K
    _widths(ws, {"A": 3, "B": 40, **{YC[t]: 11 for t in range(0, G.N_YEARS + 1)}})
    ws.freeze_panes = "C6"
    c.R = R


# ======================================================================= DCF
def _dcf(wb, c: Ctx):
    ws = wb.create_sheet("DCF")
    ref, R, d, p = c.ref, c.R, c.d, c.p
    bsd = p.get("bs_date") or c.vd
    _title(ws, f"DCF valuation (FCFF, mid-period convention, anchored at the {bsd} balance sheet)")
    _hdr(ws, 4, ["", "Item"] + [c.fy(t, "E") for t in range(1, G.N_YEARS + 1)])
    ws["A5"] = "year"
    for t in range(1, G.N_YEARS + 1):
        _put(ws, f"{YC[t]}5", c.fy0 + t if c.fy0 else t, None, SMALL)
    ws["B6"], ws["B7"], ws["B8"], ws["B9"], ws["B10"] = ("FCFF (Forecast)", f"Cash flow valued ({c.fy(1)}: part still to come only)",
                                                         f"Discount period t (years from {bsd}, mid-period)", "Discount factor", "Present value")
    D = YC[1]
    for t in range(1, G.N_YEARS + 1):
        col = YC[t]
        _put(ws, f"{col}6", f"=Forecast!{col}{R['fcff']}", EUR, F_LINK)
        _put(ws, f"{col}7", f"={col}6*{ref['stub']}" if t == 1 else f"={col}6", EUR)
        _put(ws, f"{col}8", f"={ref['stub']}/2" if t == 1 else f"={ref['stub']}+({col}5-${D}$5-1)+0.5", "0.000")
        _put(ws, f"{col}9", f"=1/(1+{ref['wacc']})^{col}8", "0.0000")
        _put(ws, f"{col}10", f"={col}7*{col}9", EUR)
    Mc = YC[G.N_YEARS]
    PX = c.PX
    block = [
        ("WACC", f"={ref['wacc']}", PCT2, d["wacc"], False),
        ("Terminal growth g", f"={ref['g_act']}", PCT2, p["g"], False),
        (f"Revenue T+1 ({c.fy(G.N_YEARS + 1)})", f"=Forecast!{Mc}{R['revenue']}*(1+B15)", EUR, None, True),
        ("EBIT T+1 at target margin", f"=B16*{ref['mt_act']}", EUR, d["ebit_t1"], True),
        ("FCFF T+1", f"=B17*(1-{ref['tax_terminal']})+B16*{ref['sbc_pct']}*{ref['sbc_addback']}+B16*({ref['da_terminal']}-{ref['capex_terminal']})"
                     f"-{ref['nwc_pct']}*(B16-Forecast!{Mc}{R['revenue']})", EUR, d["fcff_t1"], True),
        (f"Terminal value at end-{c.fy(G.N_YEARS)} (Gordon)", "=B18/(B14-B15)", EUR, d["tv"], True),
        ("Discount period of TV", f"={ref['stub']}+{G.N_YEARS - 1}", "0.00", d["t_tv"], False),
        ("PV of terminal value", "=B19/(1+B14)^B20", EUR, d["pv_tv"], True),
        (f"Sum of PV of FCFF {c.fy(1)}-{c.fy(G.N_YEARS)}", f"=SUM({D}10:{Mc}10)", EUR, d["sum_pv"], True),
        (f"Enterprise value ({bsd})", "=B21+B22", EUR, d["ev"], True),
        ("+ Cash and short-term investments", f"={ref['cash']}", EUR, None, True),
        ("+ Long-term financial investments", f"={ref['lt_inv']}", EUR, None, True),
        ("+ Other non-operating assets", f"={ref['other_nonop']}", EUR, None, True),
        ("- Financial debt incl. leases", f"=-{ref['debt']}", EUR, None, True),
        ("- Pension deficit", f"=-{ref['pension']}", EUR, None, True),
        ("- Non-controlling interests", f"=-{ref['nci']}", EUR, None, True),
        ("- Preferred stock", f"=-{ref['pref']}", EUR, None, True),
        (f"Equity value ({bsd})", "=SUM(B23:B30)", EUR, d["equity"], True),
        ("Shares outstanding (m)", f"={ref['shares']}", EUR1, None, False),
        (f"FX ({p['q_ccy']} per {p['fin_ccy']})", f"={ref['fx']}", "0.0000", None, False),
        (f"Value per share at {bsd}", "=B31*B33/B32", PX, d["ps_anchor"], False),
        (f"Roll-forward to {c.vd}: (1+Ke)^(days/365)", f"=(1+{ref['ke']})^({ref['roll_days']}/365)", "0.0000", d["roll"], False),
        (f"DCF value per share at {c.vd}", "=B34*B35", PX, d["price"], False),
        ("Share price", f"={ref['price']}", PX, None, False),
        ("Upside / (downside)", "=B36/B37-1", PCT, d["upside"], False),
        ("TV share of EV", "=B21/B23", PCT, d["tv_share"], False),
        ("Implied exit multiple EV/EBIT (TV / EBIT T+1)", "=B19/B17", MULT, d["exit_ev_ebit"], False),
        ("Implied EV / NTM revenue", f"=B23/({ref['stub']}*Forecast!{YC[1]}{R['revenue']}+(1-{ref['stub']})*Forecast!{YC[2]}{R['revenue']})", MULT, d["implied_ev_sales_ntm"], False),
        ("Implied EV / NTM EBIT", f"=B23/({ref['stub']}*Forecast!{YC[1]}{R['ebit']}+(1-{ref['stub']})*Forecast!{YC[2]}{R['ebit']})", MULT, d["implied_ev_ebit_ntm"], False),
    ]
    _hdr(ws, 13, ["Terminal value & equity bridge", "Value", "Python check"])
    for i, (lab, f, fmt, chk, money) in enumerate(block, start=14):
        ws.cell(row=i, column=1, value=lab)
        _put(ws, f"B{i}", f, fmt)
        if _ok(chk):
            _put(ws, f"C{i}", float(chk) / (M if money else 1), fmt, GREY)
    for rr in (23, 31, 36):
        ws[f"A{rr}"].font = F_B; ws[f"B{rr}"].font = F_B; ws[f"B{rr}"].fill = FILL_K
    ws["A44"] = (f"{c.fy(1)} is valued only for the {p['stub']:.2f} years after the balance-sheet date; the cash it generated before that date "
                 "is already in the balance sheet. Same convention as the SAP model (Q4 stub from 30 Sep).")
    ws["A44"].font = NOTE
    _widths(ws, {"A": 52, "B": 16, "C": 14, **{YC[t]: 11 for t in range(1, G.N_YEARS + 1)}})
    c.ref.update(price_out="DCF!$B$36", ev_out="DCF!$B$23", equity_out="DCF!$B$31", upside_out="DCF!$B$38", tv_share_out="DCF!$B$39")


# ======================================================================= DDM
def _ddm(wb, c: Ctx):
    dd = c.res.get("ddm")
    if not dd:
        return
    ws = wb.create_sheet("DDM")
    ref, p = c.ref, c.p
    _title(ws, "Dividend discount model (two stages, discounted at the cost of equity)",
           "Cross-check on the DCF. Dividends are paid through the year (mid-period); buybacks are not included.")
    rows = [("Indicated annual dividend per share", f"={ref['dps0']}", c.PX), (f"Dividend growth {c.fy(1)}-{c.fy(5)}", f"={ref['g_div']}", PCT),
            ("Terminal growth (active)", f"={ref['g_act']}", PCT2), ("Cost of equity", f"={ref['ke']}", PCT2), ("First discount period (years)", 0.5, "0.00")]
    for i, (lab, f, fmt) in enumerate(rows, start=4):
        ws.cell(row=i, column=1, value=lab)
        _put(ws, f"B{i}", f, fmt, F_LINK if isinstance(f, str) else F_IN)
    _hdr(ws, 10, ["", "Item"] + [c.fy(t, "E") for t in range(1, G.N_YEARS + 1)])
    labs = {11: "year index", 12: "Dividend growth", 13: "Dividend per share", 14: "Discount period (years)", 15: "Present value",
            16: "Python check (dividend per share)"}
    for rr, lab in labs.items():
        ws.cell(row=rr, column=2, value=lab)
    tab = dd["table"]
    for t in range(1, G.N_YEARS + 1):
        col, pc = YC[t], YC[t - 1]
        _put(ws, f"{col}11", t, None, SMALL)
        _put(ws, f"{col}12", "=$B$5" if t <= 5 else f"=$B$5+($B$6-$B$5)*{t - 5}/5", PCT)
        _put(ws, f"{col}13", f"={'$B$4' if t == 1 else f'{pc}13'}*(1+{col}12)", "0.000")
        _put(ws, f"{col}14", f"=$B$8+({col}11-1)", "0.00")
        _put(ws, f"{col}15", f"={col}13/(1+$B$7)^{col}14", "0.000")
        _put(ws, f"{col}16", float(tab.loc[t - 1, "dps"]), "0.000", GREY)
    Mc = YC[G.N_YEARS]
    block = [(f"Dividend per share T+1 ({c.fy(G.N_YEARS + 1)})", f"={Mc}13*(1+B6)", "0.000", dd["dps_t1"]),
             ("Terminal value per share (Gordon growth at Ke)", "=B20/(B7-B6)", c.PX, dd["tv"]),
             ("PV of terminal value", f"=B21/(1+B7)^({Mc}14+0.5)", c.PX, dd["pv_tv"]),
             (f"PV of dividends {c.fy(1)}-{c.fy(G.N_YEARS)}", f"=SUM({YC[1]}15:{Mc}15)", c.PX, dd["pv_div"]),
             (f"DDM value per share at {c.vd}", "=B22+B23", c.PX, dd["price"]),
             ("DCF value per share (for comparison)", f"={ref['price_out']}", c.PX, None)]
    _hdr(ws, 19, ["Valuation", "Value", "Python check"])
    for i, (lab, f, fmt, chk) in enumerate(block, start=20):
        ws.cell(row=i, column=1, value=lab); _put(ws, f"B{i}", f, fmt)
        if chk is not None:
            _put(ws, f"C{i}", chk, fmt, GREY)
    ws["A24"].font = F_B; ws["B24"].font = F_B; ws["B24"].fill = FILL_K
    h = c.res.get("hist")
    note = "Reading: a DDM values only the cash paid as dividends."
    if h is not None and len(h) and {"dividends", "buybacks", "fcf"} <= set(h.columns):
        last = h.iloc[-1]
        if _ok(last["fcf"]) and last["fcf"]:
            dv, bb = -(last["dividends"] or 0), -(last["buybacks"] or 0)
            note += (f" In {c.fy(0)} dividends were {dv / last['fcf'] * 100:.0f}% and buybacks {bb / last['fcf'] * 100:.0f}% of free cash flow, "
                     "so for a company that returns cash mainly through buybacks the DDM sits well below the DCF.")
    ws["A27"] = note; ws["A27"].font = NOTE
    _widths(ws, {"A": 50, "B": 24, "C": 14, **{YC[t]: 10 for t in range(1, G.N_YEARS + 1)}})
    ws.column_dimensions["B"].width = 16
    c.ref["ddm_out"] = "DDM!$B$24"


# ======================================================================= Comps (live statistics)
def _comps(wb, c: Ctx):
    peers, cp = c.res.get("peers"), c.res["comps"]
    if peers is None or not len(peers):
        return
    ws = wb.create_sheet("Comps")
    ref, p = c.ref, c.p
    _title(ws, "Relative valuation: peer multiples (NTM = next twelve months, consensus from Yahoo Finance)",
           "Yahoo symbol = the code used to fetch data; local code and exchange shown for readability. Statistics are live formulas; blank = not meaningful (negative or missing).")
    cols = ["Company", "Yahoo symbol", "Local code", "Exchange", "Country", "Industry", "Mkt cap €m", "NTM growth %", "EBIT margin %",
            "Rule of 40", "EV/NTM sales", "EV/NTM EBIT", "P/NTM EPS"]
    _hdr(ws, 4, cols)
    r0 = r = 5
    for _, pr in (peers.sort_values("mcap_eur", ascending=False) if "mcap_eur" in peers else peers).iterrows():
        code, exch = G.exchange_of(pr["ticker"])
        vals = [pr.get("name"), pr["ticker"], code, exch, pr.get("country"), pr.get("industry"), (pr.get("mcap_eur") or 0) / M,
                pr.get("g_ntm"), pr.get("ebit_margin"), pr.get("rule40")]
        for j, v in enumerate(vals, start=1):
            _put(ws, f"{L(j)}{r}", v)
        for j, k in zip((11, 12, 13), ("ev_sales_ntm", "ev_ebit_ntm", "pe_ntm")):
            v = pr.get(k)
            _put(ws, f"{L(j)}{r}", v if _ok(v) and v > 0 else None)
        r += 1
    r1 = r - 1
    for row in ws.iter_rows(min_row=r0, max_row=r1, min_col=7, max_col=13):
        for cc in row:
            cc.number_format = EUR if cc.column == 7 else ("0.0" if cc.column <= 10 else '0.0"x"')
    st_r = r + 1
    for i, (lab, fn) in enumerate([("Lower quartile", "QUARTILE.INC({rng},1)"), ("Median", "MEDIAN({rng})"),
                                    ("Upper quartile", "QUARTILE.INC({rng},3)"), ("Mean", "AVERAGE({rng})"), ("Count", "COUNT({rng})")]):
        ws.cell(row=st_r + i, column=1, value=lab).font = F_B
        for j in range(8, 14):
            col = L(j)
            cc = ws.cell(row=st_r + i, column=j, value="=" + fn.format(rng=f"{col}{r0}:{col}{r1}"))
            cc.number_format = "0" if lab == "Count" else ("0.0" if j <= 10 else '0.0"x"')
    r = st_r + 7
    ws.cell(row=r, column=1, value=f"Implied value per share for {p['name']} (peer quartiles applied to its NTM consensus)").font = F_B
    r += 1
    _hdr(ws, r, ["Metric", f"{p['ticker']} NTM metric", "Peer Q1", "Peer median", "Peer Q3", "Price @Q1", "Price @median", "Price @Q3", "Python check (median)"])
    r += 1
    nonop = f"({ref['cash']}+{ref['lt_inv']}+{ref['other_nonop']}-{ref['debt']}-{ref['pension']}-{ref['nci']}-{ref['pref']})"
    colmap = {"ev_sales_ntm": "K", "ev_ebit_ntm": "L", "pe_ntm": "M"}
    impl = []
    for key, st in cp["stats"].items():
        ev_based = key != "pe_ntm"
        ws.cell(row=r, column=1, value=st["label"])
        _put(ws, f"B{r}", st["metric"] / M if ev_based else st["metric"], EUR1 if ev_based else "0.00", F_IN)
        for j, srow in zip("CDE", (st_r, st_r + 1, st_r + 2)):
            _put(ws, f"{j}{r}", f"={colmap[key]}{srow}", '0.0"x"')
        for j, mc in zip("FGH", "CDE"):
            f = f"=({mc}{r}*$B{r}+{nonop})*{ref['fx']}/{ref['shares']}" if ev_based else f"={mc}{r}*$B{r}"
            _put(ws, f"{j}{r}", f, c.PX)
        _put(ws, f"I{r}", st["p_med"], c.PX, GREY)
        impl.append(f"G{r}")
        r += 1
    ws.cell(row=r, column=1, value="Comparable-companies value (average of the medians)").font = F_B
    _put(ws, f"G{r}", f"=AVERAGE({','.join(impl)})", c.PX, F_B).fill = FILL_K
    _put(ws, f"I{r}", G.comps_value(cp), c.PX, GREY)
    c.ref["comps_out"] = f"Comps!$G${r}"
    c.ref["comps_rows"] = [(st["label"], f"Comps!$G${impl[i][1:]}", st["p_med"]) for i, st in enumerate(cp["stats"].values())]
    r += 2
    if "regression" in cp:
        reg = cp["regression"]
        ws.cell(row=r, column=1, value=f"Growth-adjusted: EV/NTM sales = a + b x Rule of 40 (n={reg['n']}, engine values)").font = F_B
        r += 1
        for lab, v, fmt in (("Slope b", reg["slope"], "0.0000"), ("Intercept a", reg["intercept"], "0.000"), ("Std error", reg["se"], "0.000"),
                            (f"{p['ticker']} Rule of 40", reg["rule40"], "0.0"), ("Fitted EV/NTM sales", reg["fit"], '0.00"x"'),
                            ("Implied price (fit)", reg["p_fit"], c.PX), ("Implied price (fit ± 1 SE)", f"{c.sym}{reg['p_low']:.2f} - {c.sym}{reg['p_high']:.2f}", None)):
            ws.cell(row=r, column=1, value=lab); _put(ws, f"B{r}", v, fmt)
            r += 1
        r += 1
    for n in ("Peer multiples use each peer's own fiscal-year-weighted consensus (Yahoo Finance) and its GAAP/IFRS operating margin, so EV/EBIT includes SBC on both sides.",
              f"EV-based prices use the same EV-to-equity bridge, shares and FX as the DCF (balance sheet {p.get('bs_date')}).",
              "Peers in another currency: EV is rebuilt in the reporting currency from market cap, debt and cash, so multiples are currency-neutral.",
              "Market cap shown in € for comparison only (ECB reference rates)."):
        ws.cell(row=r, column=1, value="• " + n); r += 1
    _widths(ws, {"A": 34, "B": 14, "C": 11, "D": 18, "E": 14, "F": 26, **{L(j): 12 for j in range(7, 14)}})
    ws.freeze_panes = "C5"


# ======================================================================= Engine-value sheets
def _grid(ws, r, title, grid, row_lab, col_lab, row_fmt, col_fmt, px_fmt, base_r=None, base_c=None):
    ws.cell(row=r, column=1, value=title).font = F_B
    ws.cell(row=r + 1, column=1, value=f"{row_lab} ↓ / {col_lab} →").font = NOTE
    for j, col in enumerate(grid.columns, start=2):
        cc = ws.cell(row=r + 1, column=j, value=float(col)); cc.number_format = col_fmt; cc.font = F_H; cc.fill = FILL_H
    for i, idx in enumerate(grid.index, start=r + 2):
        cc = ws.cell(row=i, column=1, value=float(idx)); cc.number_format = row_fmt; cc.font = F_H; cc.fill = FILL_H
        for j, col in enumerate(grid.columns, start=2):
            v = grid.loc[idx, col]
            cell = ws.cell(row=i, column=j, value=float(v) if _ok(v) else "n/m")
            cell.number_format = px_fmt; cell.border = BOX
            if base_r is not None and abs(float(idx) - base_r) < 1e-9 and abs(float(col) - base_c) < 1e-9:
                cell.font = F_B
    rng = f"B{r+2}:{L(1+len(grid.columns))}{r+1+len(grid.index)}"
    ws.conditional_formatting.add(rng, ColorScaleRule(start_type="min", start_color="F8D7CF", mid_type="percentile", mid_value=50,
                                                      mid_color="FFFFFF", end_type="max", end_color="BFE6DF"))
    return r + 3 + len(grid.index)


def _sensitivity(wb, c: Ctx, out_dir: Path):
    res, p, d = c.res, c.p, c.d
    if "grid_wacc_g" not in res:
        return
    ws = wb.create_sheet("Sensitivity")
    _title(ws, f"Sensitivity analysis (DCF value per share, {p['q_ccy']}, {c.vd})",
           "Engine values (same formulas as the DCF sheet); bold = base case; n/m = WACC too close to growth")
    r = 4
    r = _grid(ws, r, "1) WACC vs terminal growth", res["grid_wacc_g"], "WACC", "g", PCT2, PCT2, c.PX0, round(d["wacc"], 4), round(p["g"], 4))
    if "grid_wacc_margin" in res:
        r = _grid(ws, r + 1, "2) WACC vs target EBIT margin", res["grid_wacc_margin"], "WACC", "target margin", PCT2, PCT, c.PX0,
                  round(d["wacc"], 4), round(p["margin_target"], 4))
    if "grid_growth_margin" in res:
        r = _grid(ws, r + 1, f"3) Revenue growth shift {c.fy(2)}-{c.fy(5)} vs target margin shift", res["grid_growth_margin"], "Δ growth", "Δ margin",
                  PCT, PCT, c.PX0, 0.0, 0.0)
    if "grid_beta_erp" in res:
        r = _grid(ws, r + 1, "4) Beta vs equity risk premium", res["grid_beta_erp"], "beta", "ERP", "0.00", PCT2, c.PX0,
                  round(p["beta"], 4), round(p["erp"], 4))
    if res.get("tornado") is not None:
        ws.cell(row=r + 1, column=1, value="5) Tornado: one input at a time").font = F_B
        _hdr(ws, r + 2, ["Driver", "Low", "High", "Range"])
        for i, (_, t) in enumerate(res["tornado"].iterrows(), start=r + 3):
            ws.cell(row=i, column=1, value=t["driver"])
            for j, k in zip((2, 3, 4), ("low", "high", "range")):
                _put(ws, f"{L(j)}{i}", t[k], f'"{c.sym}"#,##0.0')
        png = out_dir / f"{_tag(p)}_tornado.png"
        _tornado_png(res, png, c.sym)
        img = XLImage(str(png)); img.width, img.height = 760, 388
        ws.add_image(img, f"J{r + 1}")
    _widths(ws, {"A": 36, **{L(j): 11 for j in range(2, 10)}})


def _scenarios(wb, c: Ctx):
    ws = wb.create_sheet("Scenarios")
    p = c.p
    _title(ws, "Scenario analysis (engine values)",
           "To see a scenario flow through the whole model, set the selector on the Assumptions sheet to 1, 2 or 3; the DCF value then equals the figure below.")
    _hdr(ws, 3, ["Scenario", "Weight", f"Value/share {p['q_ccy']}", "WACC", "Terminal g", "Target margin", f"Revenue {c.fy(5)}E {c.m}", "Narrative"])
    for i, n in enumerate(("Bear", "Base", "Bull"), start=4):
        s = c.res["scenarios"][n]
        vals = [n, s["weight"], s["price"], s.get("wacc"), s.get("g"), s.get("m_target"), (s.get("rev5") or 0) / M, s.get("story")]
        for j, (v, fm) in enumerate(zip(vals, [None, PCT, c.PX, PCT2, PCT2, PCT, EUR, None]), start=1):
            _put(ws, f"{L(j)}{i}", v, fm)
    ws["A8"] = "Probability-weighted value"; ws["A8"].font = F_B
    _put(ws, "C8", "=SUMPRODUCT(B4:B6,C4:C6)", c.PX, F_B).fill = FILL_K
    _widths(ws, {"A": 28, "B": 9, "C": 15, "D": 9, "E": 11, "F": 14, "G": 18, "H": 110})


def _monte_carlo(wb, c: Ctx, out_dir: Path):
    mc = c.res.get("mc")
    if not mc:
        return
    ws = wb.create_sheet("Monte Carlo")
    n, p = mc["n"], c.p
    _title(ws, f"Monte Carlo simulation of the DCF value ({n:,} runs, fixed seed {mc['seed']})",
           "Each run shocks four drivers independently (normal, cut at ±2.5 standard deviations; terminal growth capped 1pp below WACC) and reruns the DCF. Statistics are live formulas over the run table.")
    _hdr(ws, 4, ["Driver", "Standard deviation", "Correlation with value"])
    keys = [k for k, _, _ in mc["drivers"]]
    for i, (k, lab, sd) in enumerate(mc["drivers"], start=5):
        ws.cell(row=i, column=1, value=lab); _put(ws, f"B{i}", sd, PCT2)
        col = L(7 + keys.index(k))
        _put(ws, f"C{i}", f"=CORREL({col}5:{col}{n + 4},K5:K{n + 4})", "0.00")
    rng = f"K5:K{n + 4}"
    _hdr(ws, 12, ["Statistic", "Value per share", "Python check"])
    stats = [("5th percentile", f"=PERCENTILE({rng},0.05)", mc["pct"][5]), ("25th percentile", f"=PERCENTILE({rng},0.25)", mc["pct"][25]),
             ("Median", f"=MEDIAN({rng})", mc["pct"][50]), ("Mean", f"=AVERAGE({rng})", mc["mean"]),
             ("75th percentile", f"=PERCENTILE({rng},0.75)", mc["pct"][75]), ("95th percentile", f"=PERCENTILE({rng},0.95)", mc["pct"][95]),
             ("Standard deviation", f"=STDEV({rng})", mc["sd"])]
    for i, (lab, f, chk) in enumerate(stats, start=13):
        ws.cell(row=i, column=1, value=lab); _put(ws, f"B{i}", f, c.PX); _put(ws, f"C{i}", chk, c.PX, GREY)
    ws["A20"] = "Share of runs above the share price"
    _put(ws, "B20", f'=COUNTIF({rng},">"&{c.ref["price"]})/COUNT({rng})', PCT, F_B).fill = FILL_K
    ws["A20"].font = F_B
    ws["A21"] = "Base-case DCF value"; _put(ws, "B21", mc["base"], c.PX)
    ws["A23"] = "Shocks are independent; in reality growth and margins move together, so the true range is somewhat wider."; ws["A23"].font = NOTE
    _hdr(ws, 4, ["Run"] + [lab for _, lab, _ in mc["drivers"]] + ["Value per share"], start=6)
    dr = mc["draws"]
    for i in range(n):
        ws.cell(row=i + 5, column=6, value=i + 1)
        for j, k in enumerate(keys):
            ws.cell(row=i + 5, column=7 + j, value=float(dr[k].iloc[i])).number_format = PCT2
        ws.cell(row=i + 5, column=11, value=float(dr["price"].iloc[i])).number_format = c.PX
    png = out_dir / f"{_tag(p)}_monte_carlo.png"
    _mc_png(c.res, png, c.sym)
    img = XLImage(str(png)); img.width, img.height = 640, 300
    ws.add_image(img, "A25")
    _widths(ws, {"A": 40, "B": 18, "C": 22, "D": 3, "E": 3, "F": 7, "G": 12, "H": 16, "I": 16, "J": 20, "K": 15})


def _reverse(wb, c: Ctx):
    rv = c.res.get("reverse")
    if not rv:
        return
    ws = wb.create_sheet("Reverse DCF")
    p = c.p
    _title(ws, f"Reverse DCF: what the share price of {c.sym}{rv['price']:,.2f} implies",
           "Each row changes one driver and holds everything else at the base case until the DCF value equals the share price (engine values).")
    _hdr(ws, 4, ["Driver", "Base case", "Implied by the price", "Change", "For comparison", "How to read it"])
    ctx = rv.get("context", {})
    mh = ctx.get("margin_hist")
    rows = [
        ("Target EBIT margin", rv["margin_target"], PCT, (f"Reported range {mh[0] * 100:.1f}% - {mh[1] * 100:.1f}%" if mh else ""),
         "The long-run operating margin the market would need at the base growth and WACC"),
        (f"Revenue growth {c.fy(2)}-{c.fy(5)} (average)", rv["growth"], PCT,
         (f"Reported CAGR {ctx['rev_cagr'] * 100:.1f}% over {ctx['rev_years']} years" if "rev_cagr" in ctx else ""),
         "The growth needed in the explicit years at the base margin and WACC"),
        ("WACC", rv["wacc"], PCT2, f"Risk-free rate {p['rf'] * 100:.2f}%", "The discount rate at which the base forecast is worth the share price"),
        ("Terminal growth", rv["g"], PCT2, f"Risk-free rate {p['rf'] * 100:.2f}%", "The perpetual growth needed at the base forecast and WACC"),
    ]
    for i, (lab, d_, fmt, cmp_, read) in enumerate(rows, start=5):
        ws.cell(row=i, column=1, value=lab)
        _put(ws, f"B{i}", d_["base"], fmt)
        if d_["implied"] is None:
            ws.cell(row=i, column=3, value="no solution in range")
        else:
            _put(ws, f"C{i}", d_["implied"], fmt, F_B).fill = FILL_K
            _put(ws, f"D{i}", f"=C{i}-B{i}", fmt)
        ws.cell(row=i, column=5, value=cmp_); ws.cell(row=i, column=6, value=read)
    ws["A10"] = "Base-case DCF value per share"; _put(ws, "B10", f"={c.ref['price_out']}", c.PX, F_LINK)
    ws["A11"] = "Share price"; _put(ws, "B11", f"={c.ref['price']}", c.PX, F_LINK)
    notes = ["To check a row: type the implied value into its blue cell on the Assumptions sheet (WACC: adjust beta or the equity risk premium); the DCF value then equals the share price.",
             "'No solution in range' means that driver alone cannot explain the price (ranges: margin -20% to 95%, growth shift -30 to +60 points, WACC up to 30%, terminal growth below WACC).",
             "Reading: an implied value far outside the company's own history says either the market expects a different business from the one in the forecast, or the discount rate is too high or too low. It does not say which."]
    for i, n in enumerate(notes, start=13):
        ws.cell(row=i, column=1, value="• " + n).font = NOTE
    _widths(ws, {"A": 40, "B": 14, "C": 20, "D": 12, "E": 34, "F": 80})


# ======================================================================= Historicals and FCFF bridge
HIST_ROWS = [("revenue", "Revenue", EUR), ("growth", "  growth y/y", PCT), ("gross_profit", "Gross profit", EUR), ("gm", "  gross margin", PCT),
             ("rnd", "R&D expense", EUR), ("ebit", "Operating profit (EBIT)", EUR), ("em", "  EBIT margin", PCT), ("pretax", "Pre-tax income", EUR),
             ("tax_provision", "Tax expense", EUR), ("tax_rate", "Effective tax rate", PCT), ("net_income", "Net income to shareholders", EUR),
             ("eps_dil", "Diluted EPS (reporting currency)", "0.00"), ("shares_dil", "Diluted shares (m)", EUR1), (None, "", None),
             ("da", "Depreciation & amortisation", EUR), ("sbc", "Stock-based compensation", EUR), ("dwc", "Change in working capital (CF)", EUR),
             ("ocf", "Operating cash flow", EUR), ("capex", "Capital expenditure", EUR), ("fcf", "Free cash flow (reported)", EUR),
             ("cash_taxes", "Cash taxes paid", EUR), ("dividends", "Dividends paid", EUR), ("buybacks", "Share buybacks", EUR), (None, "", None),
             ("cash_sti", "Cash and short-term investments", EUR), ("lt_inv", "Long-term financial investments", EUR), ("debt", "Total debt", EUR),
             ("receivables", "Receivables", EUR), ("inventory", "Inventory", EUR), ("payables", "Payables", EUR),
             ("twc", "  trade working capital % revenue", PCT), ("equity", "Shareholders' equity", EUR)]


def _historicals(wb, c: Ctx):
    h = c.res.get("hist")
    if h is None or not len(h):
        return
    ws = wb.create_sheet("Historicals")
    p = c.p
    _title(ws, f"Historical financials ({c.m}) and consensus", "Yahoo Finance annual statements (fiscal years as reported); ratios are formulas")
    dates = list(h.index)
    _hdr(ws, 4, ["Line item"] + [f"FY{d[:4]}A" for d in dates])
    for j, d in enumerate(dates, start=2):
        _put(ws, f"{L(j)}5", d, None, SMALL)
    rows = {}
    for i, (k, lab, fmt) in enumerate(HIST_ROWS, start=6):
        ws.cell(row=i, column=1, value=lab)
        if k:
            rows[k] = i
    for j, d in enumerate(dates, start=2):
        col, pc = L(j), L(j - 1)
        a = h.loc[d]
        for k, lab, fmt in HIST_ROWS:
            if not k:
                continue
            i = rows[k]
            if k == "growth":
                f = f"={col}{rows['revenue']}/{pc}{rows['revenue']}-1" if j > 2 else None
            elif k == "gm":
                f = f"={col}{rows['gross_profit']}/{col}{rows['revenue']}" if _ok(a.get("gross_profit")) else None
            elif k == "em":
                f = f"={col}{rows['ebit']}/{col}{rows['revenue']}" if _ok(a.get("ebit")) else None
            elif k == "twc":
                f = (f"=({col}{rows['receivables']}+{col}{rows['inventory']}-{col}{rows['payables']})/{col}{rows['revenue']}"
                     if any(_ok(a.get(x)) for x in ("receivables", "payables")) else None)
            else:
                v = a.get(k)
                f = (v if k in ("tax_rate", "eps_dil") else v / M) if _ok(v) else None
            _put(ws, f"{col}{i}", f, fmt, F_IN if isinstance(f, float) else None)
    r = 6 + len(HIST_ROWS) + 1
    s = c.res.get("snapshot", {})
    re_, ee = s.get("rev_est", {}), s.get("eps_est", {})
    if re_ or ee:
        ws.cell(row=r, column=1, value="Analyst consensus (Yahoo Finance, run date)").font = F_B
        _hdr(ws, r + 1, ["Item", c.fy(1, "E") + " (0y)", c.fy(2, "E") + " (+1y)"])
        for i, (lab, src, sc_, fmt) in enumerate((("Revenue", re_, M, EUR), ("EPS", ee, 1, "0.00")), start=r + 2):
            ws.cell(row=i, column=1, value=lab)
            for j, k in ((2, "0y"), (3, "+1y")):
                if src.get(k):
                    _put(ws, f"{L(j)}{i}", src[k] / sc_, fmt, F_IN)
    _widths(ws, {"A": 38, **{L(j): 13 for j in range(2, 2 + len(dates))}})
    ws.freeze_panes = "B6"
    c.hist_rows, c.hist_col = rows, L(1 + len(dates))


def _recon(wb, c: Ctx):
    h = c.res.get("hist")
    if h is None or not len(h) or not hasattr(c, "hist_rows"):
        return
    a = h.iloc[-1]
    if not (_ok(a.get("ebit")) and _ok(a.get("fcf"))):
        return
    ws = wb.create_sheet("FCFF bridge")
    R, col = c.hist_rows, c.hist_col
    _title(ws, f"{c.fy(0) if c.fy0 else 'Last fiscal year'}: FCFF on the model definition vs the company's reported free cash flow ({c.m})",
           "Shows how the forecast's cash-flow definition relates to what the company reports; all lines link to Historicals")
    hl = lambda k: f"Historicals!{col}{R[k]}" if _ok(a.get(k)) else "0"  # noqa: E731
    rows = [("Operating profit (EBIT)", f"={hl('ebit')}"),
            ("- Taxes on EBIT at the effective rate", f"=-B4*{hl('tax_rate')}" if _ok(a.get("tax_rate")) else f"=-B4*{c.ref['tax'][1]}"),
            ("NOPAT", "=B4+B5"), ("+ D&A", f"={hl('da')}"), ("- Capex", f"={hl('capex')}"),
            ("+/- Change in working capital (cash-flow statement)", f"={hl('dwc')}"), ("FCFF on the model definition", "=SUM(B6:B9)"),
            ("+ Stock-based compensation (non-cash)", f"={hl('sbc')}"),
            ("+ Tax on EBIT minus cash taxes paid", f"=-B5-{hl('cash_taxes')}" if _ok(a.get("cash_taxes")) else "=0"),
            ("+/- Interest, leases and other items (balance)", "=B14-B10-B11-B12"), ("Reported free cash flow", f"={hl('fcf')}")]
    for i, (lab, f) in enumerate(rows, start=4):
        ws.cell(row=i, column=1, value=lab)
        _put(ws, f"B{i}", f, EUR, F_LINK if "Historicals" in f else None)
    for rr in (6, 10, 14):
        ws[f"A{rr}"].font = F_B; ws[f"B{rr}"].font = F_B
    ws["B10"].fill = FILL_K
    ws["A16"] = ("Reading: the model's FCFF excludes SBC add-backs (SBC is treated as a cost) and taxes EBIT at the effective rate, so it sits "
                 "below reported free cash flow when SBC is large." + ("" if _ok(a.get("cash_taxes")) else " Cash taxes paid are not reported by Yahoo for this year, so the tax difference is in the balance line."))
    ws["A16"].font = NOTE
    _widths(ws, {"A": 56, "B": 16})


# ======================================================================= Checks
def _checks(wb, c: Ctx):
    ws = wb.create_sheet("Checks")
    ref, R = c.ref, c.R
    _title(ws, "Model integrity checks (live formulas)", "Every row must read TRUE. The Python-engine comparisons apply to the Base scenario (selector = 2).")
    _hdr(ws, 4, ["Check", "Result", "Left side", "Right side"])
    D, Mc = YC[1], YC[G.N_YEARS]
    base_only = f"{ref['scen_sel']}<>2"
    F = lambda k: f"Forecast!{D}{R[k]}:{Mc}{R[k]}"  # noqa: E731
    checks = [
        ("FCFF = NOPAT + SBC add-back + D&A - capex - working capital (all years)",
         f"SUMPRODUCT(ABS({F('nopat')}+{F('sbc')}+{F('da')}-{F('capex')}-{F('dnwc')}-{F('fcff')}))", "0", "abs"),
        (f"Growth in {c.fy(G.N_YEARS)} equals terminal growth", f"Forecast!{Mc}{R['growth']}", ref["g_act"], "eq"),
        (f"Margin in {c.fy(G.N_YEARS)} equals the target margin", f"Forecast!{Mc}{R['margin']}", ref["mt_act"], "eq"),
        ("Capital-structure weights add up to 100%", f"{ref['we']}+{ref['wd']}", "1", "eq"),
        ("Enterprise value = PV of FCFF + PV of terminal value", "DCF!B23", "DCF!B21+DCF!B22", "eq"),
        ("Equity value = EV + non-operating assets - claims", "DCF!B31",
         f"DCF!B23+{ref['cash']}+{ref['lt_inv']}+{ref['other_nonop']}-{ref['debt']}-{ref['pension']}-{ref['nci']}-{ref['pref']}", "eq"),
        ("WACC is above terminal growth", ref["wacc"], ref["g_act"], "gt"),
        ("FCFF matches the Python engine in every year (Base)", f"SUMPRODUCT(ABS({F('fcff')}-{F('py_fcff')}))", "0.01", "le_base"),
        ("WACC matches the Python engine (Base)", f"ABS({ref['wacc']}-WACC!C15)", "0.000001", "le_base"),
        ("DCF value per share matches the Python engine (Base)", "ABS(DCF!B36-DCF!C36)", "0.01", "le_base"),
    ]
    if "ddm_out" in ref:
        checks.append(("DDM value per share matches the Python engine (Base)", "ABS(DDM!B24-DDM!C24)", "0.01", "le_base"))
    for lab, cell, py in ref.get("comps_rows", []):
        checks.append((f"Comps {lab}: median price matches the Python engine", f"ABS({cell}-{py})", "0.01", "le_any"))
    if "blend_out" in ref:
        checks.append(("Weighted value matches the Python engine (Base)", f"ABS({ref['blend_out']}-{ref['blend_py']})", "0.01", "le_base"))
    for i, (lab, left, right, kind) in enumerate(checks, start=5):
        ws.cell(row=i, column=1, value=lab)
        ws.cell(row=i, column=3, value=f"={left}"); ws.cell(row=i, column=4, value=f"={right}")
        f = {"abs": f"=ABS(C{i}-D{i})<0.001", "eq": f"=ABS(C{i}-D{i})<0.000001", "gt": f"=C{i}>D{i}",
             "le_base": f"=OR({base_only},C{i}<=D{i})", "le_any": f"=C{i}<=D{i}"}[kind]
        ws.cell(row=i, column=2, value=f).font = F_B
        for col in (3, 4):
            ws.cell(row=i, column=col).number_format = "#,##0.0000"
    last = 4 + len(checks)
    ws.cell(row=last + 2, column=1, value="All checks pass").font = F_B
    cc = ws.cell(row=last + 2, column=2, value=f"=AND(B5:B{last})"); cc.font = Font(bold=True, size=12, color=TEAL); cc.fill = FILL_K
    # warnings: judgement flags, not integrity failures
    r = last + 5
    ws.cell(row=r, column=1, value="Sanity flags (judgement; FALSE means look again, not that the model is broken)").font = F_B
    _hdr(ws, r + 1, ["Flag", "OK?", "Value", "Threshold"])
    n_peers = min((st["n"] for st in c.res["comps"]["stats"].values()), default=0)
    flags = [("WACC exceeds terminal growth by at least 2 points", f"{ref['wacc']}-{ref['g_act']}", "0.02", "ge"),
             ("Terminal growth not above the risk-free rate", ref["rf"], ref["g_act"], "ge"),
             ("Terminal value below 80% of enterprise value", "0.8", "DCF!B39", "ge"),
             ("DCF within ±40% of the share price", "0.4", "ABS(DCF!B38)", "ge"),
             ("At least 5 peers in every multiple", str(n_peers), "5", "ge")]
    for i, (lab, left, right, kind) in enumerate(flags, start=r + 2):
        ws.cell(row=i, column=1, value=lab)
        ws.cell(row=i, column=3, value=f"={left}"); ws.cell(row=i, column=4, value=f"={right}")
        ws.cell(row=i, column=2, value=f"=C{i}>=D{i}").font = F_B
        for col in (3, 4):
            ws.cell(row=i, column=col).number_format = "#,##0.0000"
    _widths(ws, {"A": 70, "B": 12, "C": 18, "D": 18})
    c.ref.update(checks_all=f"Checks!$B${last + 2}", n_checks=len(checks))


# ======================================================================= Summary, Cover, Sources
def _summary(wb, c: Ctx, out_dir: Path):
    ws = wb["Summary"]
    res, p, d, sc, mk, ref = c.res, c.p, c.d, c.res["scenarios"], c.res["market"], c.ref
    _title(ws, f"{p['name']} ({p['ticker']}): valuation as of {c.vd}",
           f"FCFF DCF + relative valuation | {p['q_ccy']} per share | figures link to the model sheets")
    _hdr(ws, 4, ["Key outputs", "Value"])
    rows = [("DCF value per share (base)", f"={ref['price_out']}", c.PX), ("Share price", f"={ref['price']}", c.PX),
            ("Upside / (downside) vs price", f"={ref['upside_out']}", PCT), ("Probability-weighted scenario value", "=Scenarios!C8", c.PX),
            ("WACC", f"={ref['wacc']}", PCT2), ("Terminal growth", f"={ref['g_act']}", PCT2),
            (f"Enterprise value ({c.m})", f"={ref['ev_out']}", EUR), (f"Equity value ({c.m})", f"={ref['equity_out']}", EUR),
            ("Terminal value share of EV", f"={ref['tv_share_out']}", PCT)]
    if res.get("mc"):
        rows.append(("Monte Carlo median (P5-P95 on the sheet)", "='Monte Carlo'!B15", c.PX))
    rv = res.get("reverse")
    if rv:
        if rv["margin_target"]["implied"] is not None:
            rows.append(("Price implies a target EBIT margin of", float(rv["margin_target"]["implied"]), PCT))
        if rv["wacc"]["implied"] is not None:
            rows.append(("Price implies a WACC of", float(rv["wacc"]["implied"]), PCT2))
    if "ddm_out" in ref:
        rows.append(("DDM value per share (dividends only)", f"={ref['ddm_out']}", c.PX))
    for lab, cell, _ in ref.get("comps_rows", []):
        rows.append((f"Comps: {lab} (median)", f"={cell}", c.PX))
    if mk.get("tgt_mean"):
        rows.append((f"Analyst target mean (n={mk['tgt_n']})", float(mk["tgt_mean"]), c.PX))
    rows.append(("52-week range", f"{c.sym}{mk['lo52']:.2f} - {c.sym}{mk['hi52']:.2f}", None))
    r = 5
    for lab, f, fm in rows:
        ws.cell(row=r, column=1, value=lab); _put(ws, f"B{r}", f, fm)
        r += 1
    ws["A5"].font = F_B; ws["B5"].font = Font(bold=True, size=12, color=TEAL); ws["B5"].fill = FILL_K
    # blended value (weights are inputs)
    r += 1
    ws.cell(row=r, column=1, value="Football-field weighting (blue = input; renormalised over the methods available)").font = F_B
    r += 1
    wrows = {}
    for k, w in c.weights.items():
        ws.cell(row=r, column=1, value=f"Weight: {k}"); _put(ws, f"B{r}", w, PCT, F_IN)
        wrows[k] = f"B{r}"
        r += 1
    from .generic import comps_value
    cv = comps_value(res["comps"])
    if "Comparable companies" in wrows and "comps_out" in ref:
        f = f"=({wrows['DCF']}*{ref['price_out']}+{wrows['Comparable companies']}*{ref['comps_out']})/({wrows['DCF']}+{wrows['Comparable companies']})"
        tot = c.weights["DCF"] + c.weights["Comparable companies"]
        py = (c.weights["DCF"] * d["price"] + c.weights["Comparable companies"] * cv) / tot if tot else d["price"]
    else:
        f, py = f"={ref['price_out']}", d["price"]
    ws.cell(row=r, column=1, value="Weighted value per share").font = F_B
    _put(ws, f"B{r}", f, c.PX, F_B).fill = FILL_K
    _put(ws, f"C{r}", py, c.PX, GREY)
    ref.update(blend_out=f"Summary!$B${r}", blend_py=f"Summary!$C${r}")
    if "Precedent transactions" not in c.weights:
        ws.cell(row=r + 1, column=1, value="Precedent transactions: not available for an arbitrary ticker (SAP model only); its weight is redistributed.").font = NOTE
    r += 3
    ws.cell(row=r, column=1, value="Key assumptions").font = F_B
    f5 = d["forecast"].iloc[4]
    ka = [f"Revenue {c.fy(0)} {c.sym if p['fin_ccy'] == p['q_ccy'] else p['fin_ccy'] + ' '}{p['rev0'] / 1e9:,.1f}bn; growth "
          + ", ".join(f"{x * 100:.1f}%" for x in p["growth"]) + f" ({c.fy(1)}-{c.fy(5)}), fading to {p['g'] * 100:.2f}% by {c.fy(G.N_YEARS)}",
          f"EBIT margin {p['margin0'] * 100:.1f}% ({c.fy(1)}) -> {p['margin_target'] * 100:.1f}% target; tax {p['tax'] * 100:.1f}% effective, "
          f"fading to the {p['tax_terminal'] * 100:.2f}% statutory rate",
          f"D&A {p['da_pct'] * 100:.1f}% and capex {p['capex_pct'] * 100:.1f}% of revenue; working capital {p['nwc_pct'] * 100:.1f}% of incremental revenue; "
          f"FCFF margin {c.fy(5)} {f5['fcff'] / f5['revenue'] * 100:.1f}%",
          f"WACC {d['wacc'] * 100:.2f}% = Ke {d['ke'] * 100:.2f}% (Rf {p['rf'] * 100:.2f}%, β {p['beta']:.2f}, ERP {p['erp'] * 100:.1f}%) x {d['we'] * 100:.0f}% "
          f"+ after-tax Kd {d['kd_at'] * 100:.2f}% x {d['wd'] * 100:.0f}%",
          f"Bridge ({p.get('bs_date')}): + cash {p['cash'] / 1e9:,.1f}bn + investments {p['lt_inv'] / 1e9:,.1f}bn - debt {p['debt'] / 1e9:,.1f}bn "
          f"- NCI/preferred {(p['nci'] + p['pref']) / 1e9:,.1f}bn ({p['fin_ccy']}); {p['shares'] / M:,.0f}m shares",
          f"Peers: {len(res['peers'])}" + (f" ({', '.join(res['peers']['ticker'])})" if len(res["peers"]) else " (comps skipped)")]
    for i, t in enumerate(ka, start=r + 1):
        ws.cell(row=i, column=1, value="• " + t)
    if res.get("meta", {}).get("lookahead"):
        ws.cell(row=r + len(ka) + 2, column=1, value=(
            "Look-ahead warning: the valuation date is earlier than today. Price, risk-free rate, beta window and balance sheet are as of that date; "
            "consensus, peer multiples and analyst targets are today's.")).font = Font(bold=True, color="C0392B")
    for j, wmsg in enumerate(res.get("meta", {}).get("warnings", []), start=r + len(ka) + 3):
        ws.cell(row=j, column=1, value="Check: " + wmsg).font = Font(bold=True, color="C0392B")
    png = out_dir / f"{_tag(p)}_football.png"
    from . import charts
    charts.football(res, png, f"{p['name']}: valuation football field ({p['q_ccy']} per share, {c.vd})",
                    items=G.football_items(res), currency=c.sym)
    img = XLImage(str(png)); img.width, img.height = 880, 640
    ws.add_image(img, "D4")
    _widths(ws, {"A": 52, "B": 18, "C": 14})


def _cover(wb, c: Ctx):
    ws = wb.create_sheet("Cover", 0)
    p, ref = c.p, c.ref
    _title(ws, f"{p['name']} ({p['ticker']}): company valuation model",
           f"Valuation date {c.vd} · {c.m} unless stated · per-share values in {p['q_ccy']} · public data (Yahoo Finance, ECB, government bond yields)")
    rows = [("DCF value per share", f"={ref['price_out']}", c.PX), ("Share price", f"={ref['price']}", c.PX),
            ("Weighted value per share", f"={ref['blend_out']}", c.PX),
            ("Active scenario", f'=CHOOSE({ref["scen_sel"]},"Bear","Base","Bull")', None), ("All integrity checks pass", f"={ref['checks_all']}", None)]
    for i, (lab, f, fm) in enumerate(rows, start=4):
        ws.cell(row=i, column=1, value=lab).font = F_B
        _put(ws, f"B{i}", f, fm, F_B).fill = FILL_K
    r = 10
    ws.cell(row=r, column=1, value="How to use").font = F_B
    for i, t in enumerate(("Change blue cells on the Assumptions sheet; every live sheet recalculates.",
                           "Switch the scenario with the selector at the bottom of the Assumptions sheet (1 = Bear, 2 = Base, 3 = Bull).",
                           "Check the Checks sheet after any change: every integrity row must read TRUE; sanity flags ask for judgement.",
                           "Sheets marked 'engine values' come from the Python engine and do not react to input changes; rerun the model to refresh them."), start=r + 1):
        ws.cell(row=i, column=1, value=f"{i - r}. {t}")
    r = 16
    ws.cell(row=r, column=1, value="Colour legend").font = F_B
    for i, (lab, font) in enumerate((("Blue: input you can change", F_IN), ("Black: formula", Font()), ("Green: link to another sheet", F_LINK),
                                     ("Grey italics: Python engine check value", GREY)), start=r + 1):
        ws.cell(row=i, column=1, value=lab).font = font
    ws.cell(row=r + 5, column=1, value="Highlighted: key result").fill = FILL_K
    r = 23
    _hdr(ws, r, ["Sheet", "Type", "Content"])
    desc = {"Summary": ("Live", "Key outputs, football field, weighting, key assumptions"),
            "Assumptions": ("Inputs", "Every input with source and rationale; scenario selector"),
            "WACC": ("Live", "Cost of capital and reference points"),
            "Forecast": ("Live", f"Revenue, EBIT and free cash flow, {c.fy(1)}-{c.fy(G.N_YEARS)}"),
            "DCF": ("Live", "Discounting, terminal value, equity bridge, value per share"),
            "DDM": ("Live", "Dividend discount model as a cross-check"), "Checks": ("Live", f"{ref['n_checks']} integrity checks and sanity flags"),
            "Scenarios": ("Engine values", "Bear, base and bull values and weights"),
            "Sensitivity": ("Engine values", "WACC vs growth and margin, growth vs margin, beta vs ERP; tornado"),
            "Reverse DCF": ("Engine values", "The margin, growth, WACC or terminal growth the share price implies"),
            "Monte Carlo": ("Engine values, live statistics", "Distribution of the DCF value"),
            "Comps": ("Live statistics", "Peer multiples, quartiles, implied values"),
            "Historicals": ("Data", "Annual statements and consensus"), "FCFF bridge": ("Live", f"{c.fy(0)} FCFF vs reported free cash flow"),
            "Sources & notes": ("Text", "Data sources, method notes and what this template leaves out")}
    names = [ws_.title for ws_ in wb.worksheets if ws_.title != "Cover"]
    order = [n for n in ORDER if n in names] + [n for n in names if n not in ORDER]
    for i, n in enumerate(order, start=r + 1):
        t, dsc = desc.get(n, ("", ""))
        for j, v in enumerate((n, t, dsc), start=1):
            ws.cell(row=i, column=j, value=v)
        ws.cell(row=i, column=1).font = F_B
    _widths(ws, {"A": 34, "B": 30, "C": 90})


def _sources(wb, c: Ctx):
    ws = wb.create_sheet("Sources & notes")
    p, res = c.p, c.res
    s = res.get("snapshot", {})
    _title(ws, "Data sources, method notes and what this template leaves out")
    _hdr(ws, 3, ["Item", "Source / note"])
    r = 4
    data = [("Share price", f"Yahoo Finance close {res['market'].get('price_date')}"),
            ("Financial statements", f"Yahoo Finance annual statements (base year ended {p.get('fy0_date')}) and quarterly balance sheet ({p.get('bs_date')})"),
            ("Consensus", "Yahoo Finance analyst revenue and EPS estimates (current and next fiscal year), as of the run date"),
            ("FX", "ECB reference rates (Yahoo Finance fallback)"), ("Company profile", f"{s.get('info', {}).get('sector') or ''} / {s.get('info', {}).get('industry') or ''}, {s.get('info', {}).get('country') or ''}")]
    data += [(k, v) for k, v in p.get("sources", {}).items()]
    for k, v in data:
        ws.cell(row=r, column=1, value=k); ws.cell(row=r, column=2, value=v); r += 1
    r += 1
    ws.cell(row=r, column=1, value="Same method as the SAP case model").font = F_B; r += 1
    for t in ("FCFF DCF with 5 explicit + 5 fade years, mid-period discounting, Gordon terminal value on a normalised year T+1",
              "Value anchored at the latest balance sheet, current year valued for its remaining part (stub), rolled forward to the valuation date at Ke",
              "Effective tax in the explicit years fading to the statutory rate; statutory rate for the debt tax shield",
              "Live Bear/Base/Bull selector, sensitivity grids, tornado, Monte Carlo, DDM cross-check, peer quartiles, integrity checks"):
        ws.cell(row=r, column=1, value="• " + t); r += 1
    r += 1
    ws.cell(row=r, column=1, value="Left out (no public data for an arbitrary ticker)").font = F_B; r += 1
    for k, v in (("Segment revenue build", "Yahoo has no segment data; revenue is forecast as one line from consensus"),
                 ("Precedent transactions", "Needs a curated deal list (SAP model: 11 software deals); weight redistributed to DCF and comps"),
                 ("Intangibles / R&D capitalisation, sum of the parts, bottom-up beta", "Need segment data and a curated peer set; not built"),
                 ("Bloomberg / Capital IQ reference points", "Licensed data; the public template uses Yahoo, ECB and government bond yields only")):
        ws.cell(row=r, column=1, value=k); ws.cell(row=r, column=2, value=v); r += 1
    if res.get("meta", {}).get("lookahead"):
        r += 1
        ws.cell(row=r, column=1, value="Look-ahead").font = F_B
        ws.cell(row=r, column=2, value="Valuation date earlier than today: consensus, peer multiples and analyst targets are today's and include later information.")
        r += 1
    r += 1
    ws.cell(row=r, column=1, value="Colour code: blue = input, green = link, black = formula, grey italics = Python engine check").font = NOTE
    _widths(ws, {"A": 34, "B": 150})


# ======================================================================= charts
def _tornado_png(res, path, sym):
    from . import charts as ch
    plt = ch.plt
    t = res["tornado"].iloc[::-1]
    base = res["dcf"]["price"]
    fig, ax = plt.subplots(figsize=(9, 4.6), dpi=150)
    span = (t["high"].max() - t["low"].min()) * 0.01
    for i, (_, r) in enumerate(t.iterrows()):
        ax.barh(i, r["high"] - base, left=base, color=ch.TEAL)
        ax.barh(i, r["low"] - base, left=base, color=ch.CORAL)
        ax.text(r["high"] + span, i, f"{r['high']:.0f}", va="center", fontsize=7)
        ax.text(r["low"] - span, i, f"{r['low']:.0f}", va="center", ha="right", fontsize=7)
    ax.axvline(base, color=ch.INK, lw=1)
    ax.set_yticks(range(len(t))); ax.set_yticklabels(t["driver"], fontsize=8)
    ax.set_xlabel(f"DCF value per share ({sym.strip()})")
    ax.set_title(f"DCF sensitivity to single inputs (base {sym}{base:.0f})", fontsize=11, color=ch.INK, loc="left")
    for s_ in ("top", "right"):
        ax.spines[s_].set_visible(False)
    fig.tight_layout(); fig.savefig(path); plt.close(fig)


def _mc_png(res, path, sym):
    from . import charts as ch
    plt = ch.plt
    mc, price = res["mc"], res["market"]["price"]
    fig, ax = plt.subplots(figsize=(9, 4.2), dpi=150)
    ax.hist(mc["draws"]["price"], bins=50, color=ch.TEAL, alpha=0.85)
    for q, ls in ((5, ":"), (50, "-"), (95, ":")):
        ax.axvline(mc["pct"][q], color=ch.INK, ls=ls, lw=1.2)
        ax.text(mc["pct"][q], ax.get_ylim()[1] * 0.96, f" P{q} {sym}{mc['pct'][q]:.0f}", fontsize=8, color=ch.INK, va="top")
    ax.axvline(price, color=ch.CORAL, ls="--", lw=1.4)
    ax.text(price, ax.get_ylim()[1] * 0.84, f" Price {sym}{price:.0f}", fontsize=8, color=ch.CORAL, va="top")
    ax.set_xlabel(f"DCF value per share ({sym.strip()})"); ax.set_ylabel("Simulations")
    ax.set_title(f"Monte Carlo: {mc['n']:,} simulations, {mc['p_above_price'] * 100:.0f}% above the share price", fontsize=11, color=ch.INK, loc="left")
    for s_ in ("top", "right"):
        ax.spines[s_].set_visible(False)
    fig.tight_layout(); fig.savefig(path); plt.close(fig)


# ======================================================================= entry points
ORDER = ["Cover", "Summary", "Assumptions", "WACC", "Forecast", "DCF", "DDM", "Checks", "Scenarios", "Sensitivity", "Reverse DCF", "Monte Carlo",
         "Comps", "Historicals", "FCFF bridge", "Sources & notes"]


def _tag(p) -> str:
    return str(p.get("ticker", "company")).replace(".", "_").replace("^", "")


def write(res: dict, path: Path, out_dir: Path | None = None, weights: dict | None = None) -> Path:
    """Write the workbook. `res` comes from generic.run (inputs, dcf, scenarios, comps, peers, market, ...)."""
    path = Path(path)
    out_dir = Path(out_dir or path.parent)
    out_dir.mkdir(parents=True, exist_ok=True)
    res = dict(res)
    res["inputs"] = G.norm(res["inputs"])
    c = Ctx(res, weights)
    wb = Workbook()
    wb.active.title = "Summary"
    _assumptions(wb, c)
    _wacc(wb, c)
    _forecast(wb, c)
    _dcf(wb, c)
    _ddm(wb, c)
    _scenarios(wb, c)
    _sensitivity(wb, c, out_dir)
    _monte_carlo(wb, c, out_dir)
    _reverse(wb, c)
    _comps(wb, c)
    _historicals(wb, c)
    _recon(wb, c)
    _summary(wb, c, out_dir)
    _checks(wb, c)
    _sources(wb, c)
    _cover(wb, c)
    names = [w.title for w in wb.worksheets]
    wb._sheets = [wb[n] for n in ORDER if n in names] + [w for w in wb._sheets if w.title not in ORDER]
    wb.save(path)
    return path


def verify(path: Path, res: dict, scenarios: bool = True, tol: float = 0.01) -> tuple[bool, str]:
    """Recalculate the saved workbook (no Excel needed) and compare it with the Python engine."""
    import formulas
    from .verify import _last_check_row, _val, scenario_price
    path = Path(path)
    sol = formulas.ExcelModel().loads(str(path)).finish().calculate()
    book, d = path.name, res["dcf"]
    checks = {"WACC": (_val(sol, book, "WACC", "B15"), d["wacc"], 1e-6),
              "EV (m)": (_val(sol, book, "DCF", "B23"), d["ev"] / M, max(abs(d["ev"]) / M * 1e-9, 0.01)),
              "Equity (m)": (_val(sol, book, "DCF", "B31"), d["equity"] / M, max(abs(d["equity"]) / M * 1e-9, 0.01)),
              "DCF price": (_val(sol, book, "DCF", "B36"), d["price"], tol)}
    if res.get("ddm"):
        checks["DDM price"] = (_val(sol, book, "DDM", "B24"), res["ddm"]["price"], tol)
    lines, ok = [], True
    row = _last_check_row(path)
    key = next((k for k in sol if "CHECKS" in k.upper() and k.upper().endswith(f"!B{row}")), None)
    allpass = bool(sol[key].value[0][0]) if key else False
    ok &= allpass
    lines.append(f"  {'OK ' if allpass else 'BAD'} Checks sheet: all integrity checks {'pass' if allpass else 'do NOT pass'}")
    for k, (xv, pv, t) in checks.items():
        good = abs(xv - pv) <= t
        ok &= good
        lines.append(f"  {'OK ' if good else 'BAD'} {k:10s} excel={xv:,.4f} python={pv:,.4f}")
    if scenarios:
        with tempfile.TemporaryDirectory() as tmp:
            for sel, n in ((1, "Bear"), (3, "Bull")):
                xv, pv = scenario_price(path, sel, Path(tmp)), res["scenarios"][n]["price"]
                good = abs(xv - pv) <= tol
                ok &= good
                lines.append(f"  {'OK ' if good else 'BAD'} {n + ' (sel)':10s} excel={xv:,.4f} python={pv:,.4f}")
    rv = res.get("reverse") or {}
    im = (rv.get("margin_target") or {}).get("implied")
    if scenarios and im is not None:
        import openpyxl
        wb = openpyxl.load_workbook(path)
        ws = wb["Assumptions"]
        row = next(r for r in range(1, ws.max_row + 1) if ws.cell(row=r, column=1).value == "margin_target")
        ws.cell(row=row, column=4).value = float(im)
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "reverse.xlsx"
            wb.save(out)
            xv = _val(formulas.ExcelModel().loads(str(out)).finish().calculate(), out.name, "DCF", "B36")
        good = abs(xv - rv["price"]) <= max(tol, 1e-4 * rv["price"])
        ok &= good
        lines.append(f"  {'OK ' if good else 'BAD'} {'Reverse':10s} excel at implied margin={xv:,.4f} share price={rv['price']:,.4f}")
    return ok, f"Excel recalculation check ({book}):\n" + "\n".join(lines)
