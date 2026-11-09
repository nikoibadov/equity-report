"""Writes the valuation package as a formula-driven Excel workbook.

Live formulas: Assumptions -> WACC -> Forecast -> DCF -> Summary, and the comps statistics.
Static values (written by the Python engine, labelled as such): sensitivity grids, scenarios,
peer data, historicals. A 'Python check' column next to key outputs proves both agree."""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .assumptions import EXPLICIT, FADE, YEARS

NAVY, TEAL, LIGHT, AMBERL = "0E1B2C", "0F8B8D", "F2F7F7", "FFF4E5"
F_IN = Font(color="1F4FD1")               # blue = hard-coded input
F_LINK = Font(color="0B7A3B")             # green = link to another sheet
F_B = Font(bold=True)
F_H = Font(bold=True, color="FFFFFF")
F_T = Font(bold=True, size=14, color=NAVY)
FILL_H = PatternFill("solid", fgColor=NAVY)
FILL_S = PatternFill("solid", fgColor=LIGHT)
FILL_K = PatternFill("solid", fgColor=AMBERL)
THIN = Side(style="thin", color="C9D3DD")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
EUR = '#,##0;[Red]-#,##0'
EUR1 = '#,##0.0;[Red]-#,##0.0'
PCT = '0.0%;[Red]-0.0%'
PCT2 = '0.00%'
MULT = '0.0"x"'
PX = '"€"#,##0.00'

YCOL = {2025: "C", **{y: get_column_letter(4 + i) for i, y in enumerate(YEARS)}}   # 2026->D ... 2035->M
PAIRED = {"m_nonifrs": "m_terminal", "acq_pct": "acq_terminal", "tax": "tax_terminal",
          "da_pct": "da_terminal", "capex_pct": "capex_terminal"}
PCT_KEYS = {"q4g_cloud", "q4g_licenses", "q4g_support", "q4g_services", "g_cloud", "g_licenses", "g_support", "g_services",
            "m_nonifrs", "acq_pct", "tax", "da_pct", "capex_pct", "nwc_pct", "sbc_pct", "sbc_noncash", "g",
            "rf", "erp", "kd", "tax_wacc", "m_terminal", "acq_terminal", "tax_terminal", "da_terminal", "capex_terminal"}


def _hdr(ws, row, values, start=1):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=start + i, value=v)
        c.font, c.fill, c.alignment = F_H, FILL_H, Alignment(horizontal="center", vertical="center", wrap_text=True)


def _title(ws, text, sub=None):
    ws["A1"] = text; ws["A1"].font = F_T
    if sub:
        ws["A2"] = sub; ws["A2"].font = Font(italic=True, color="5B6B7F")


def _widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


# ======================================================================= Assumptions
def _assumptions(wb, a, res):
    ws = wb.create_sheet("Assumptions")
    _title(ws, "Assumptions (blue = input; edit here and the model recalculates)",
           f"{res['meta']['name']} | edition: {res['edition'].upper()} | valuation date 13 Nov 2026 | € millions unless stated")
    _hdr(ws, 4, ["Key", "Assumption", "Unit", "2026", "2027", "2028", "2029", "2030", "Terminal", "Source", "Rationale"])
    sections = [
        ("Market & capital structure", ["price", "e_mv", "d_mv", "shares"]),
        ("Cost of capital inputs", ["rf", "beta", "erp", "kd", "tax_wacc"]),
        ("Base year FY2025 (actual)", ["base_cloud", "base_licenses", "base_support", "base_services"]),
        ("FY2026 anchor = 9M-2026 actual + Q4-2025 x (1 + Q4 growth)", ["ytd_cloud", "ytd_licenses", "ytd_support", "ytd_services",
                                                                     "ytdp_cloud", "ytdp_licenses", "ytdp_support", "ytdp_services",
                                                                     "q4g_cloud", "q4g_licenses", "q4g_support", "q4g_services"]),
        ("Segment revenue growth 2027-2030", ["g_cloud", "g_licenses", "g_support", "g_services"]),
        ("Profitability (fade 2031-35 linear to terminal)", ["m_nonifrs", "acq_pct", "oneoff_2026", "tax"]),
        ("Reinvestment & SBC", ["da_pct", "capex_pct", "nwc_pct", "sbc_pct", "sbc_noncash", "sbc_addback"]),
        ("Terminal value & timing", ["g", "stub", "roll_days"]),
        ("EV-to-equity bridge (30 Sep 2026)", ["cash", "st_inv", "lt_inv", "debt", "pension", "nci"]),
    ]
    ref = {}
    r = 5
    for title, keys in sections:
        ws.cell(row=r, column=1, value=title).font = F_B
        for c in range(1, 12):
            ws.cell(row=r, column=c).fill = FILL_S
        r += 1
        for k in keys:
            it = a.items[k]
            ws.cell(row=r, column=1, value=k).font = Font(color="8091A5", size=9)
            ws.cell(row=r, column=2, value=it.rationale.split(";")[0][:60] if False else _label(k))
            ws.cell(row=r, column=3, value=it.unit)
            fmt = PCT if k in PCT_KEYS else ('0.00' if k == "beta" else (PX if k == "price" else EUR))
            v = it.value
            if isinstance(v, list):
                cols = ["D", "E", "F", "G", "H"] if len(v) == 5 else ["E", "F", "G", "H"]
                ref[k] = {}
                for col, x in zip(cols, v):
                    c = ws[f"{col}{r}"]; c.value = x; c.font = F_IN; c.number_format = fmt
                    ref[k][2026 + "DEFGH".index(col)] = f"Assumptions!${col}${r}"
                if k in PAIRED:
                    tk = PAIRED[k]
                    c = ws[f"I{r}"]; c.value = a[tk]; c.font = F_IN; c.number_format = PCT
                    ref[tk] = f"Assumptions!$I${r}"
            else:
                c = ws[f"D{r}"]; c.value = v; c.font = F_IN; c.number_format = fmt
                ref[k] = f"Assumptions!$D${r}"
            src = it.source + (f" | terminal: {a.items[PAIRED[k]].source}" if k in PAIRED else "")
            rat = it.rationale + (f" | terminal: {a.items[PAIRED[k]].rationale}" if k in PAIRED else "")
            ws.cell(row=r, column=10, value=src)
            ws.cell(row=r, column=11, value=rat).alignment = Alignment(wrap_text=True, vertical="top")
            r += 1
        r += 1
    _widths(ws, {"A": 14, "B": 38, "C": 9, "D": 11, "E": 11, "F": 11, "G": 11, "H": 11, "I": 11, "J": 42, "K": 90})
    ws.freeze_panes = "D5"
    return ref


def _label(k):
    L = {
        "price": "Reference share price", "e_mv": "Market value of equity", "d_mv": "Debt (incl. leases) for WACC weights",
        "shares": "Diluted shares (m)", "rf": "Risk-free rate (10y Bund)", "beta": "Levered beta", "erp": "Equity risk premium",
        "kd": "Pre-tax cost of debt", "tax_wacc": "Tax rate (WACC)", "base_cloud": "Cloud revenue FY25",
        "base_licenses": "Software licences FY25", "base_support": "Software support FY25", "base_services": "Services FY25",
        "ytd_cloud": "Cloud 9M-26", "ytd_licenses": "Licences 9M-26", "ytd_support": "Support 9M-26", "ytd_services": "Services 9M-26",
        "ytdp_cloud": "Cloud 9M-25", "ytdp_licenses": "Licences 9M-25", "ytdp_support": "Support 9M-25", "ytdp_services": "Services 9M-25",
        "q4g_cloud": "Q4-26 growth: cloud", "q4g_licenses": "Q4-26 growth: licences", "q4g_support": "Q4-26 growth: support",
        "q4g_services": "Q4-26 growth: services", "g_cloud": "Growth: cloud", "g_licenses": "Growth: licences",
        "g_support": "Growth: support", "g_services": "Growth: services", "m_nonifrs": "Non-IFRS operating margin (incl. SBC)",
        "acq_pct": "Acquisition-related charges % rev", "oneoff_2026": "2026 one-offs (restructuring, litigation)",
        "tax": "Tax rate on EBIT", "da_pct": "D&A % revenue", "capex_pct": "Capex % revenue", "nwc_pct": "NWC release % of Δrevenue",
        "sbc_pct": "SBC % revenue", "sbc_noncash": "Non-cash share of SBC", "sbc_addback": "Add back non-cash SBC? (0/1)",
        "g": "Terminal growth g", "stub": "Stub (Q4-26) year fraction", "roll_days": "Roll-forward days 30 Sep -> 13 Nov",
        "cash": "Cash & equivalents", "st_inv": "Current financial assets", "lt_inv": "Non-current financial investments",
        "debt": "Financial debt incl. leases", "pension": "Pension deficit", "nci": "Non-controlling interests",
    }
    return L.get(k, k)


# ======================================================================= WACC
def _wacc(wb, ref, res):
    ws = wb.create_sheet("WACC")
    _title(ws, "Cost of capital (CAPM / market-value weights)")
    rows = [
        ("Risk-free rate", f"={ref['rf']}", PCT2), ("Levered beta", f"={ref['beta']}", "0.00"),
        ("Equity risk premium", f"={ref['erp']}", PCT2), ("Cost of equity  Ke = Rf + β·ERP", "=B4+B5*B6", PCT2),
        ("Pre-tax cost of debt", f"={ref['kd']}", PCT2), ("Tax rate", f"={ref['tax_wacc']}", PCT2),
        ("After-tax cost of debt", "=B8*(1-B9)", PCT2), ("Equity (market value, €m)", f"={ref['e_mv']}", EUR),
        ("Debt incl. leases (€m)", f"={ref['d_mv']}", EUR), ("Equity weight", "=B11/(B11+B12)", PCT),
        ("Debt weight", "=B12/(B11+B12)", PCT), ("WACC", "=B13*B7+B14*B10", PCT2),
    ]
    _hdr(ws, 3, ["Component", "Value", "Python check"])
    d = res["dcf"]
    checks = {7: d["ke"], 10: d["kd_at"], 13: d["we"], 14: d["wd"], 15: d["wacc"]}
    for i, (lab, f, fmt) in enumerate(rows, start=4):
        ws.cell(row=i, column=1, value=lab)
        c = ws.cell(row=i, column=2, value=f); c.number_format = fmt; c.font = F_LINK if "Assump" in f else Font()
        if i in checks:
            ws.cell(row=i, column=3, value=checks[i]).number_format = fmt
    ws["A15"].font = F_B; ws["B15"].font = F_B; ws["B15"].fill = FILL_K
    # Reference points
    ds = res["ds"]
    r = 18
    ws.cell(row=r, column=1, value="Reference points and why they are not used as base").font = F_B
    _hdr(ws, r + 1, ["Source", "WACC / input", "Comment"])
    if res["edition"] == "data":
        refs = [
            ("Bloomberg WACC function", ds["reference"]["bbg_wacc"], "Uses 10.32% 'country premium' (= Bloomberg expected market return 13.59% - Rf), i.e. Ke 13.1%: implausible for an A+/A1 company; rejected"),
            ("S&P Capital IQ DCF template", ds["reference"]["ciq_wacc"], "Rf 3.77%, ERP 4.23%, beta 0.92 -> Ke 7.65%: ERP at low end; our ERP 5.0% (IDW range) is more conservative"),
            ("Bloomberg beta 2y weekly (adj.)", 1.000, "Higher standard error (0.173) than 5y (0.087); 5y used"),
            ("Damodaran implied ERP 1 Nov 2026", 0.0463, "4.41% + 0.22% default-spread adj. (US-based); sensitivity"),
            ("IDW/FAUB market risk premium range", "5.0% - 6.5%", "German valuation standard (pre personal taxes); base 5.0% at its lower end"),
            ("Bund curve 20y / 30y", "3.60% / 3.63%", "Long end used in sensitivity (Rf +50bp row of tornado)"),
        ]
    else:
        wm = ds["web_meta"]
        refs = [
            ("Damodaran implied ERP (" + wm["damodaran"]["erp_date"] + ")", ds["reference"]["damodaran_erp"], "Implied US ERP incl. default-spread adjustment; below our 5.0% house view"),
            ("IDW/FAUB market risk premium range", "5.0% - 6.5%", "German valuation standard (pre personal taxes); base 5.0% at its lower end"),
            ("Beta regression (raw)", wm["beta"]["raw"], f"5y weekly vs DAX, n={wm['beta']['n']}, R² {wm['beta']['r2']:.2f}, SE {wm['beta']['se']:.3f}; Blume adjustment 0.67·raw + 0.33"),
            ("Damodaran A1/A+ default spread", wm["damodaran"]["a_plus_spread"], "Added to the Bund yield for the pre-tax cost of debt"),
            ("Bundesbank 10y Bund", wm["bund"]["yield"] / 100, "Date " + wm["bund"]["date"]),
        ]
    for i, (s, v, c) in enumerate(refs, start=r + 2):
        ws.cell(row=i, column=1, value=s)
        cc = ws.cell(row=i, column=2, value=v)
        if isinstance(v, float):
            cc.number_format = PCT2 if v < 0.5 else "0.000"
        ws.cell(row=i, column=3, value=c)
    _widths(ws, {"A": 40, "B": 16, "C": 110})
    return {"ke": "WACC!$B$7", "wacc": "WACC!$B$15"}


# ======================================================================= Forecast
def _forecast(wb, ref, res):
    ws = wb.create_sheet("Forecast")
    _title(ws, "Revenue build and FCFF (€m)", "2026 = 9M actual + Q4 estimate; 2027-30 segment-driven; 2031-35 linear fade to terminal")
    years = [2025] + YEARS
    _hdr(ws, 4, ["", "Line item"] + [f"{y}{'A' if y == 2025 else 'E'}" for y in years])
    ws["A5"] = "year"
    for y in years:
        ws[f"{YCOL[y]}5"] = y
        ws[f"{YCOL[y]}5"].font = Font(color="8091A5", size=9)
    R = {}
    labels = [
        ("cloud", "Cloud revenue"), ("licenses", "Software licences"), ("support", "Software support"), ("services", "Services"),
        ("revenue", "Total revenue"), ("growth", "  growth y/y"), ("cloud_share", "  cloud share of revenue"),
        ("margin", "Non-IFRS operating margin"), ("op", "Non-IFRS operating profit"), ("acq_pct", "Acquisition-related % rev"),
        ("acq", "Acquisition-related charges"), ("oneoff", "One-offs"), ("ebit", "EBIT (IFRS basis)"), ("ebit_m", "  EBIT margin"),
        ("tax_rate", "Tax rate"), ("nopat", "NOPAT"), ("sbc", "+ Non-cash SBC add-back (toggle)"), ("da_pct", "D&A % rev"),
        ("da", "+ D&A"), ("capex_pct", "Capex % rev"), ("capex", "- Capex"), ("dnwc", "- Δ Net working capital (neg. = release)"),
        ("fcff", "Free cash flow to firm (FCFF)"), ("fcff_m", "  FCFF margin"), ("py_fcff", "Python engine check (FCFF)"),
    ]
    for i, (k, lab) in enumerate(labels, start=6):
        R[k] = i
        ws.cell(row=i, column=2, value=lab)
    f = res["dcf"]["forecast"]
    h25 = res["hist"][2025]
    segs = ("cloud", "licenses", "support", "services")
    for s in segs:
        ws[f"C{R[s]}"] = f"={ref['base_' + s]}"; ws[f"C{R[s]}"].font = F_LINK
        ws[f"D{R[s]}"] = f"={ref['ytd_' + s]}+({ref['base_' + s]}-{ref['ytdp_' + s]})*(1+{ref['q4g_' + s]})"
        for y in EXPLICIT:
            c, p = YCOL[y], YCOL[y - 1]
            ws[f"{c}{R[s]}"] = f"={p}{R[s]}*(1+{ref['g_' + s][y]})"
    ws[f"C{R['revenue']}"] = f"=SUM(C{R['cloud']}:C{R['services']})"
    for y in [2026] + EXPLICIT:
        c = YCOL[y]
        ws[f"{c}{R['revenue']}"] = f"=SUM({c}{R['cloud']}:{c}{R['services']})"
        ws[f"{c}{R['growth']}"] = f"={c}{R['revenue']}/{YCOL[y-1]}{R['revenue']}-1"
        ws[f"{c}{R['cloud_share']}"] = f"={c}{R['cloud']}/{c}{R['revenue']}"
    ws[f"C{R['cloud_share']}"] = f"=C{R['cloud']}/C{R['revenue']}"
    for k, y in enumerate(FADE, start=1):
        c, p = YCOL[y], YCOL[y - 1]
        ws[f"{c}{R['growth']}"] = f"=$H{R['growth']}+({ref['g']}-$H{R['growth']})*{k}/5"
        ws[f"{c}{R['revenue']}"] = f"={p}{R['revenue']}*(1+{c}{R['growth']})"
    # paths
    def path(row, key):
        for y in [2026] + EXPLICIT:
            ws[f"{YCOL[y]}{row}"] = f"={ref[key][y]}"
            ws[f"{YCOL[y]}{row}"].font = F_LINK
        for k, y in enumerate(FADE, start=1):
            ws[f"{YCOL[y]}{row}"] = f"=$H{row}+({ref[PAIRED[key]]}-$H{row})*{k}/5"
    path(R["margin"], "m_nonifrs"); path(R["acq_pct"], "acq_pct"); path(R["tax_rate"], "tax")
    path(R["da_pct"], "da_pct"); path(R["capex_pct"], "capex_pct")
    # 2025A memo values (as reported)
    ws[f"C{R['op']}"] = res["ds"]["ir"]["fy2025_statement"]["fy2025"]["op_nonifrs"]
    ws[f"C{R['ebit']}"] = res["ds"]["ir"]["fy2025_statement"]["fy2025"]["op_ifrs"]
    ws[f"C{R['da']}"] = h25["da"]; ws[f"C{R['capex']}"] = h25["capex"]
    for cell in (f"C{R['op']}", f"C{R['ebit']}", f"C{R['da']}", f"C{R['capex']}"):
        ws[cell].font = F_IN
    for y in YEARS:
        c, p = YCOL[y], YCOL[y - 1]
        ws[f"{c}{R['op']}"] = f"={c}{R['revenue']}*{c}{R['margin']}"
        ws[f"{c}{R['acq']}"] = f"={c}{R['revenue']}*{c}{R['acq_pct']}"
        ws[f"{c}{R['oneoff']}"] = f"={ref['oneoff_2026']}" if y == 2026 else 0
        ws[f"{c}{R['ebit']}"] = f"={c}{R['op']}-{c}{R['acq']}-{c}{R['oneoff']}"
        ws[f"{c}{R['ebit_m']}"] = f"={c}{R['ebit']}/{c}{R['revenue']}"
        ws[f"{c}{R['nopat']}"] = f"={c}{R['ebit']}*(1-{c}{R['tax_rate']})"
        ws[f"{c}{R['sbc']}"] = f"={c}{R['revenue']}*{ref['sbc_pct']}*{ref['sbc_noncash']}*{ref['sbc_addback']}"
        ws[f"{c}{R['da']}"] = f"={c}{R['revenue']}*{c}{R['da_pct']}"
        ws[f"{c}{R['capex']}"] = f"={c}{R['revenue']}*{c}{R['capex_pct']}"
        ws[f"{c}{R['dnwc']}"] = f"=-{ref['nwc_pct']}*({c}{R['revenue']}-{p}{R['revenue']})"
        ws[f"{c}{R['fcff']}"] = f"={c}{R['nopat']}+{c}{R['sbc']}+{c}{R['da']}-{c}{R['capex']}-{c}{R['dnwc']}"
        ws[f"{c}{R['fcff_m']}"] = f"={c}{R['fcff']}/{c}{R['revenue']}"
        ws[f"{c}{R['py_fcff']}"] = float(f.loc[y, "fcff"])
    pct_rows = {R[k] for k in ("growth", "cloud_share", "margin", "acq_pct", "ebit_m", "tax_rate", "da_pct", "capex_pct", "fcff_m")}
    for row in range(6, R["py_fcff"] + 1):
        for y in years:
            ws[f"{YCOL[y]}{row}"].number_format = PCT if row in pct_rows else EUR
    for k in ("revenue", "ebit", "fcff"):
        for col in ["B"] + [YCOL[y] for y in years]:
            ws[f"{col}{R[k]}"].font = F_B
    for col in [YCOL[y] for y in years]:
        ws[f"{col}{R['fcff']}"].fill = FILL_K
        ws[f"{col}{R['py_fcff']}"].font = Font(italic=True, color="8091A5")
    _widths(ws, {"A": 3, "B": 42, **{YCOL[y]: 11 for y in years}})
    ws.freeze_panes = "C6"
    return R


# ======================================================================= DCF
def _dcf(wb, ref, wref, R, res):
    ws = wb.create_sheet("DCF")
    _title(ws, "DCF valuation (FCFF, mid-year convention, anchored 30 Sep 2026)")
    _hdr(ws, 4, ["", "Item"] + [f"{y}E" for y in YEARS])
    ws["A5"] = "year"
    for y in YEARS:
        ws[f"{YCOL[y]}5"] = y; ws[f"{YCOL[y]}5"].font = Font(color="8091A5", size=9)
    rows = {"fcff": 6, "cf": 7, "t": 8, "df": 9, "pv": 10}
    ws["B6"], ws["B7"], ws["B8"], ws["B9"], ws["B10"] = "FCFF (Forecast)", "Cash flow valued (2026: Q4 stub only)", "Discount period t (years from 30 Sep 2026, mid-period)", "Discount factor", "Present value"
    for y in YEARS:
        c = YCOL[y]
        ws[f"{c}6"] = f"=Forecast!{c}{R['fcff']}"; ws[f"{c}6"].font = F_LINK
        ws[f"{c}7"] = f"={c}6*{ref['stub']}" if y == 2026 else f"={c}6"
        ws[f"{c}8"] = f"={ref['stub']}/2" if y == 2026 else f"={ref['stub']}+({c}5-2027)+0.5"
        ws[f"{c}9"] = f"=1/(1+{wref['wacc']})^{c}8"
        ws[f"{c}10"] = f"={c}7*{c}9"
        for rr in (6, 7, 10):
            ws[f"{c}{rr}"].number_format = EUR
        ws[f"{c}8"].number_format = "0.000"; ws[f"{c}9"].number_format = "0.0000"
    d = res["dcf"]
    M = YCOL[2035]
    block = [
        ("WACC", f"={wref['wacc']}", PCT2, d["wacc"]),
        ("Terminal growth g", f"={ref['g']}", PCT2, None),
        ("Revenue T+1 (2036)", f"=Forecast!{M}{R['revenue']}*(1+B15)", EUR, None),
        ("EBIT T+1 at terminal margin", f"=B16*({ref['m_terminal']}-{ref['acq_terminal']})", EUR, d["ebit_t1"]),
        ("FCFF T+1", f"=B17*(1-{ref['tax_terminal']})+B16*{ref['sbc_pct']}*{ref['sbc_noncash']}*{ref['sbc_addback']}+B16*({ref['da_terminal']}-{ref['capex_terminal']})+{ref['nwc_pct']}*(B16-Forecast!{M}{R['revenue']})", EUR, d["fcff_t1"]),
        ("Terminal value at end-2035 (Gordon)", "=B18/(B14-B15)", EUR, d["tv"]),
        ("Discount period of TV", f"={ref['stub']}+(2035-2027)+1", "0.00", d["t_tv"]),
        ("PV of terminal value", "=B19/(1+B14)^B20", EUR, d["pv_tv"]),
        ("Sum of PV of FCFF 2026-35", f"=SUM(D10:{M}10)", EUR, d["sum_pv"]),
        ("Enterprise value (30 Sep 2026)", "=B21+B22", EUR, d["ev"]),
        ("+ Cash & equivalents", f"={ref['cash']}", EUR, None),
        ("+ Current financial assets", f"={ref['st_inv']}", EUR, None),
        ("+ Non-current financial investments", f"={ref['lt_inv']}", EUR, None),
        ("- Financial debt incl. leases", f"=-{ref['debt']}", EUR, None),
        ("- Pension deficit", f"=-{ref['pension']}", EUR, None),
        ("- Non-controlling interests", f"=-{ref['nci']}", EUR, None),
        ("Equity value (30 Sep 2026)", "=SUM(B23:B29)", EUR, d["equity"]),
        ("Diluted shares (m)", f"={ref['shares']}", EUR1, None),
        ("Value per share at 30 Sep 2026", "=B30/B31", PX, d["ps_anchor"]),
        ("Roll-forward factor to 13 Nov 2026 (1+Ke)^(days/365)", f"=(1+{wref['ke']})^({ref['roll_days']}/365)", "0.0000", d["roll"]),
        ("DCF value per share at 13 Nov 2026", "=B32*B33", PX, d["price"]),
        ("Reference share price", f"={ref['price']}", PX, None),
        ("Upside / (downside)", "=B34/B35-1", PCT, d["upside"]),
        ("TV share of EV", "=B21/B23", PCT, d["tv_share"]),
        ("Implied exit multiple EV/EBIT (TV / EBIT T+1)", "=B19/B17", MULT, d["exit_ev_ebit"]),
        ("Implied EV / NTM revenue", f"=B23/(0.25*Forecast!D{R['revenue']}+0.75*Forecast!E{R['revenue']})", MULT, d["implied_ev_sales_ntm"]),
        ("Implied EV / NTM non-IFRS EBIT", f"=B23/(0.25*Forecast!D{R['op']}+0.75*Forecast!E{R['op']})", MULT, d["implied_ev_ebit_ntm"]),
    ]
    _hdr(ws, 13, ["Terminal value & equity bridge", "Value", "Python check"])
    for i, (lab, f, fmt, chk) in enumerate(block, start=14):
        ws.cell(row=i, column=1, value=lab)
        c = ws.cell(row=i, column=2, value=f); c.number_format = fmt
        if chk is not None:
            cc = ws.cell(row=i, column=3, value=float(chk)); cc.number_format = fmt; cc.font = Font(italic=True, color="8091A5")
    for rr in (23, 30, 34):
        ws[f"A{rr}"].font = F_B; ws[f"B{rr}"].font = F_B; ws[f"B{rr}"].fill = FILL_K
    _widths(ws, {"A": 52, "B": 16, "C": 14, **{YCOL[y]: 11 for y in YEARS}})
    return {"price": "DCF!$B$34", "ev": "DCF!$B$23", "equity": "DCF!$B$30", "upside": "DCF!$B$36", "tv_share": "DCF!$B$37",
            "ps_anchor": "DCF!$B$32"}


# ======================================================================= Comps
def _comps(wb, ref, res):
    ws = wb.create_sheet("Comps")
    cp, peers = res["comps"], res["peers"]
    _title(ws, "Relative valuation: peer multiples (EUR, NTM = next twelve months)",
           "Core peers = enterprise application/platform software, market cap > €5bn, positive EBIT; statistics below are live formulas over the core block")
    cols = ["Company", "Ticker", "Group / screen result", "Mkt cap €m", "EV €m", "LTM revenue", "LTM EBIT margin %",
            "NTM growth %", "Rule of 40", "EV/NTM sales", "EV/NTM EBITDA", "EV/NTM EBIT", "P/NTM EPS", "EV/LTM sales", "P/E LTM"]
    keys = ["name", "ticker", "group", "mcap", "tev", "rev_ltm", "ebit_margin", "g_ntm", "rule40", "ev_sales_ntm",
            "ev_ebitda_ntm", "ev_ebit_ntm", "pe_ntm", "ev_sales_ltm", "pe_ltm"]
    _hdr(ws, 4, cols)
    core = peers[peers["core"]].sort_values("mcap", ascending=False)
    rest = peers[~peers["core"]].sort_values("mcap", ascending=False)
    r = 5
    ws.cell(row=r, column=1, value="CORE PEER SET").font = F_B; r += 1
    c0 = r
    for _, p in core.iterrows():
        for j, k in enumerate(keys, start=1):
            v = p[k]
            ws.cell(row=r, column=j, value=(None if v != v else (float(v) if isinstance(v, (int, float)) else v)))
        r += 1
    c1 = r - 1
    stats_r = r + 1
    ws.cell(row=stats_r - 1, column=1, value="")
    for i, (lab, fn) in enumerate([("Lower quartile", "QUARTILE.INC({rng},1)"), ("Median", "MEDIAN({rng})"),
                                    ("Upper quartile", "QUARTILE.INC({rng},3)"), ("Mean", "AVERAGE({rng})")]):
        ws.cell(row=stats_r + i, column=1, value=lab).font = F_B
        for j in range(8, 16):
            col = get_column_letter(j)
            ws.cell(row=stats_r + i, column=j, value="=" + fn.format(rng=f"{col}{c0}:{col}{c1}"))
    r = stats_r + 6
    ws.cell(row=r, column=1, value="SAP SE (target)").font = F_B
    sap = peers[peers["ticker"] == "SAP"]
    if len(sap):
        for j, k in enumerate(keys, start=1):
            v = sap.iloc[0][k]
            ws.cell(row=r, column=j, value=(None if v != v else (float(v) if isinstance(v, (int, float)) else v)))
    sap_row = r
    r += 2
    ws.cell(row=r, column=1, value="EXCLUDED / SECONDARY (shown for transparency)").font = F_B; r += 1
    for _, p in rest[rest["ticker"] != "SAP"].iterrows():
        for j, k in enumerate(keys, start=1):
            v = p[k]
            ws.cell(row=r, column=j, value=(None if v != v else (float(v) if isinstance(v, (int, float)) else v)))
            ws.cell(row=r, column=j).font = Font(color="8091A5")
        r += 1
    for row in ws.iter_rows(min_row=5, max_row=r, min_col=4, max_col=15):
        for c in row:
            c.number_format = EUR if c.column <= 6 else ("0.0" if c.column <= 9 else '0.0"x"')
    # Implied value block (live formulas)
    r += 2
    s = cp["sap"]
    ws.cell(row=r, column=1, value="Implied SAP value per share (core peer quartiles applied to SAP NTM consensus)").font = F_B
    r += 1
    _hdr(ws, r, ["Metric", "SAP NTM metric", "Peer Q1", "Peer median", "Peer Q3", "Price @Q1", "Price @median", "Price @Q3", "Python check (median)"])
    r += 1
    nonop = f"({ref['cash']}+{ref['st_inv']}+{ref['lt_inv']}-{ref['debt']}-{ref['pension']}-{ref['nci']})"
    metric_rows = [
        ("EV / NTM sales", s["rev_ntm"], "J", "ev", cp["stats"]["ev_sales_ntm"]["p_med"]),
        ("EV / NTM EBIT", s["ebit_ntm"], "L", "ev", cp["stats"]["ev_ebit_ntm"]["p_med"]),
        ("P / NTM EPS", s["eps_ntm"], "M", "px", cp["stats"]["pe_ntm"]["p_med"]),
    ]
    if "ev_ebit_ntm_xsbc" in cp["stats"]:
        metric_rows.insert(2, ("EV / NTM EBIT ex-SBC (like-for-like)", s["ebit_ntm"] + s["sbc_ntm"], "L", "ev", cp["stats"]["ev_ebit_ntm_xsbc"]["p_med"]))
    impl_rows = {}
    for lab, metric, col, kind, chk in metric_rows:
        ws.cell(row=r, column=1, value=lab)
        ws.cell(row=r, column=2, value=float(metric)).font = F_IN
        ws.cell(row=r, column=2).number_format = EUR1 if metric > 100 else "0.00"
        for j, srow in zip((3, 4, 5), (stats_r, stats_r + 1, stats_r + 2)):
            ws.cell(row=r, column=j, value=f"={col}{srow}").number_format = '0.0"x"'
        for j, mc in zip((6, 7, 8), ("C", "D", "E")):
            if kind == "ev":
                f = f"=({mc}{r}*$B{r}+{nonop})/{ref['shares']}"
            else:
                f = f"={mc}{r}*$B{r}"
            ws.cell(row=r, column=j, value=f).number_format = PX
        ws.cell(row=r, column=9, value=float(chk)).number_format = PX
        impl_rows[lab] = r
        r += 1
    # Regression (live SLOPE/INTERCEPT over the core block)
    r += 1
    reg = cp["regression"]
    ws.cell(row=r, column=1, value="Growth-adjusted: EV/NTM sales = a + b x Rule of 40 (core peers)").font = F_B; r += 1
    reg_rows = [
        ("Slope b", f"=SLOPE(J{c0}:J{c1},I{c0}:I{c1})", "0.0000", reg["slope"]),
        ("Intercept a", f"=INTERCEPT(J{c0}:J{c1},I{c0}:I{c1})", "0.000", reg["intercept"]),
        ("R²", f"=RSQ(J{c0}:J{c1},I{c0}:I{c1})", "0.00", reg["r2"]),
        ("Std error of regression", f"=STEYX(J{c0}:J{c1},I{c0}:I{c1})", "0.000", reg["se"]),
        ("SAP NTM growth %", float(reg["sap_growth"]), "0.0", None),
        ("SAP LTM EBIT margin %", float(s["ebit_margin_ltm"]), "0.0", None),
        ("SAP Rule of 40", None, "0.0", reg["sap_rule40"]),
        ("Fitted EV/NTM sales", None, '0.00"x"', reg["fit"]),
        ("Implied price (fit)", None, PX, reg["p_fit"]),
        ("Implied price (fit - 1 SE)", None, PX, reg["p_low"]),
        ("Implied price (fit + 1 SE)", None, PX, reg["p_high"]),
    ]
    base = r
    for i, (lab, f, fmt, chk) in enumerate(reg_rows):
        rr = base + i
        ws.cell(row=rr, column=1, value=lab)
        if f is not None:
            ws.cell(row=rr, column=2, value=f).number_format = fmt
        if chk is not None:
            ws.cell(row=rr, column=3, value=float(chk)).number_format = fmt
    ws.cell(row=base + 6, column=2, value=f"=B{base+4}+B{base+5}").number_format = "0.0"
    ws.cell(row=base + 7, column=2, value=f"=B{base+1}+B{base}*B{base+6}").number_format = '0.00"x"'
    rev_ntm_cell = f"B{impl_rows['EV / NTM sales']}"
    ws.cell(row=base + 8, column=2, value=f"=(B{base+7}*{rev_ntm_cell}+{nonop})/{ref['shares']}").number_format = PX
    ws.cell(row=base + 9, column=2, value=f"=((B{base+7}-B{base+3})*{rev_ntm_cell}+{nonop})/{ref['shares']}").number_format = PX
    ws.cell(row=base + 10, column=2, value=f"=((B{base+7}+B{base+3})*{rev_ntm_cell}+{nonop})/{ref['shares']}").number_format = PX
    r = base + 12
    notes = [
        "Peer selection: S&P Capital IQ 'Top 30 by peer ranking' screen, then (1) business-model screen (enterprise application/platform software with recurring revenue; excludes IT services, hardware, security/infrastructure, EDA, hyper-growth outliers), (2) market cap > €5bn, (3) positive LTM EBIT.",
        "SAP NTM metrics = 1/4 FY2026 + 3/4 FY2027 consensus (anchor 30 Sep 2026). EV-based prices use the same EV-to-equity bridge and diluted shares as the DCF.",
        "EV/EBITDA is shown but not used: D&A is small and partly acquired-intangible amortisation; adds nothing beyond EV/EBIT for asset-light software. P/B not used: book equity is mostly goodwill, internally built intangibles are off balance sheet.",
        ("Like-for-like EV/EBIT: CIQ forward EBIT is broker 'street' EBIT - for US peers non-GAAP excluding SBC (e.g. Salesforce: CIQ NTM EBIT implies ~34% margin vs 21% GAAP). SAP's non-IFRS operating profit includes SBC, so the ex-SBC row applies peer multiples to SAP EBIT before SBC (≈4% of revenue)."
         if res["comps"].get("ebit_basis") == "street" else
         "EV/EBIT (web edition): peer NTM EBIT = GAAP operating margin (SBC included) x NTM revenue, which is like-for-like with SAP's SBC-inclusive non-IFRS EBIT; no ex-SBC adjustment needed."),
        "Peer NTM growth is derived as (EV / (EV/NTM sales)) / LTM revenue - 1; LTM periods end Mar-Sep 2026, so growth for some peers spans > 12 months.",
    ]
    for n in notes:
        ws.cell(row=r, column=1, value="• " + n).alignment = Alignment(wrap_text=False)
        r += 1
    _widths(ws, {"A": 40, "B": 14, "C": 44, **{get_column_letter(j): 12 for j in range(4, 16)}})
    ws.freeze_panes = "C5"
    return {"stats_row": stats_r, "impl": impl_rows, "reg_base": base}


# ======================================================================= value sheets
def _grid(ws, r, title, grid, row_lab, col_lab, row_fmt, col_fmt, base_r=None, base_c=None):
    ws.cell(row=r, column=1, value=title).font = F_B
    ws.cell(row=r + 1, column=1, value=f"{row_lab} ↓ / {col_lab} →").font = Font(italic=True, color="5B6B7F")
    for j, c in enumerate(grid.columns, start=2):
        cc = ws.cell(row=r + 1, column=j, value=float(c)); cc.number_format = col_fmt; cc.font = F_H; cc.fill = FILL_H
    for i, idx in enumerate(grid.index, start=r + 2):
        cc = ws.cell(row=i, column=1, value=float(idx)); cc.number_format = row_fmt; cc.font = F_H; cc.fill = FILL_H
        for j, c in enumerate(grid.columns, start=2):
            v = ws.cell(row=i, column=j, value=float(grid.loc[idx, c])); v.number_format = '"€"#,##0'; v.border = BOX
            if base_r is not None and abs(float(idx) - base_r) < 1e-9 and abs(float(c) - base_c) < 1e-9:
                v.font = F_B
    rng = f"B{r+2}:{get_column_letter(1+len(grid.columns))}{r+1+len(grid.index)}"
    ws.conditional_formatting.add(rng, ColorScaleRule(start_type="min", start_color="F8D7CF", mid_type="percentile", mid_value=50, mid_color="FFFFFF", end_type="max", end_color="BFE6DF"))
    return r + 3 + len(grid.index)


def _sensitivity(wb, res, out_dir):
    ws = wb.create_sheet("Sensitivity")
    a, d = res["assumptions"], res["dcf"]
    _title(ws, "Sensitivity analysis (DCF value per share, €, 13 Nov 2026)",
           "Values computed by the Python engine (same formulas as the DCF sheet); bold = base case")
    r = 4
    r = _grid(ws, r, "1) WACC vs terminal growth", res["grid_wacc_g"], "WACC", "g", PCT2, PCT2, round(d["wacc"], 4), round(a["g"], 4))
    r = _grid(ws, r + 1, "2) WACC vs terminal non-IFRS margin", res["grid_wacc_margin"], "WACC", "terminal margin", PCT2, PCT, round(d["wacc"], 4), round(a["m_terminal"], 4))
    r = _grid(ws, r + 1, "3) Cloud growth shift (2027-30, pp) vs margin shift (2027+ incl. terminal, pp)", res["grid_cloud_margin"], "Δ cloud growth", "Δ margin", PCT, PCT, 0.0, 0.0)
    r = _grid(ws, r + 1, "4) Beta vs equity risk premium", res["grid_erp_beta"], "beta", "ERP", "0.00", PCT, a["beta"], a["erp"])
    ws.cell(row=r + 1, column=1, value="5) Tornado: one input at a time").font = F_B
    _hdr(ws, r + 2, ["Driver", "Low €", "High €", "Range €"])
    for i, (_, t) in enumerate(res["tornado"].iterrows(), start=r + 3):
        ws.cell(row=i, column=1, value=t["driver"])
        for j, k in zip((2, 3, 4), ("low", "high", "range")):
            ws.cell(row=i, column=j, value=float(t[k])).number_format = '"€"#,##0.0'
    img = XLImage(str(out_dir / f"tornado_{res['edition']}.png")); img.width, img.height = 760, 388
    ws.add_image(img, f"J{r+1}")
    _widths(ws, {"A": 34, **{get_column_letter(j): 11 for j in range(2, 10)}})


def _scenarios(wb, res):
    ws = wb.create_sheet("Scenarios")
    _title(ws, "Scenario analysis (values from the Python engine)")
    _hdr(ws, 3, ["Scenario", "Weight", "Value/share €", "WACC", "Terminal g", "Terminal margin", "Revenue 2030E €m", "Narrative"])
    for i, n in enumerate(("Bear", "Base", "Bull"), start=4):
        s = res["scenarios"][n]
        vals = [n, s["weight"], s["price"], s["wacc"], s["g"], s["m_terminal"], s["rev_2030"], s["story"]]
        fmts = [None, PCT, PX, PCT2, PCT2, PCT, EUR, None]
        for j, (v, fm) in enumerate(zip(vals, fmts), start=1):
            c = ws.cell(row=i, column=j, value=float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else v)
            if fm:
                c.number_format = fm
    ws["A8"] = "Probability-weighted value"; ws["A8"].font = F_B
    ws["C8"] = "=SUMPRODUCT(B4:B6,C4:C6)"; ws["C8"].number_format = PX; ws["C8"].font = F_B; ws["C8"].fill = FILL_K
    _widths(ws, {"A": 28, "B": 9, "C": 14, "D": 9, "E": 11, "F": 15, "G": 16, "H": 130})


def _historicals(wb, res):
    ws = wb.create_sheet("Historicals")
    _title(ws, "Historical financials (€m) and consensus")
    hist = res["hist"]
    yrs = sorted(hist)
    keys = [("revenue", "Revenue"), ("cloud", "Cloud"), ("licenses", "Software licences"), ("support", "Software support"),
            ("services", "Services"), ("cost_of_cloud", "Cost of cloud"), ("gross_profit", "Gross profit"), ("rnd", "R&D"),
            ("ebit_ciq", "EBIT (CIQ standard ≈ non-IFRS)"), ("op_ifrs", "IFRS operating profit"), ("restructuring", "Restructuring"),
            ("sbc", "Share-based compensation"), ("sbc_cash_settled", "  of which cash-settled"), ("da", "D&A (as reported)"),
            ("capex", "Capex"), ("cfo", "Operating cash flow"), ("cash_taxes", "Cash taxes paid"), ("buybacks", "Buybacks"),
            ("dividends", "Dividends"), ("contract_liab", "Contract liabilities"), ("receivables", "Trade receivables"),
            ("goodwill", "Goodwill"), ("other_intangibles", "Other intangibles"), ("ppe", "PP&E"), ("total_assets", "Total assets"),
            ("equity", "Equity (parent)"), ("total_debt", "Total debt"), ("cash_sti", "Cash & ST investments"),
            ("lt_investments", "LT investments"), ("backlog", "Order backlog (CIQ)"), ("tax_rate", "Effective tax rate")]
    _hdr(ws, 3, ["Item"] + [f"FY{y}" for y in yrs])
    for i, (k, lab) in enumerate(keys, start=4):
        ws.cell(row=i, column=1, value=lab)
        for j, y in enumerate(yrs, start=2):
            v = hist[y].get(k)
            c = ws.cell(row=i, column=j, value=None if v is None else float(v))
            c.number_format = PCT if k == "tax_rate" else EUR
    r = 6 + len(keys)
    cons = res["consensus"]
    cy = [y for y in sorted(cons) if cons[y].get("revenue")]
    ws.cell(row=r - 1, column=1, value="Consensus (mean)").font = F_B
    _hdr(ws, r, ["Consensus"] + [f"FY{y}{'A' if y == 2025 else 'E'}" for y in cy])
    ck = [("revenue", "Revenue"), ("ebitda", "EBITDA"), ("ebit", "EBIT (non-IFRS)"), ("fcf", "Free cash flow"),
          ("capex", "Capex"), ("da", "D&A"), ("tax_rate", "Tax rate %"), ("eps_norm", "EPS normalised €"),
          ("eps_gaap", "EPS GAAP €"), ("dps", "DPS €"), ("shares", "Diluted shares m")]
    for i, (k, lab) in enumerate(ck, start=r + 1):
        ws.cell(row=i, column=1, value=lab)
        for j, y in enumerate(cy, start=2):
            v = cons[y].get(k)
            ws.cell(row=i, column=j, value=None if v is None else float(v)).number_format = "0.00" if k in ("eps_norm", "eps_gaap", "dps", "tax_rate") else EUR
    r = r + len(ck) + 2
    ws.cell(row=r, column=1, value="Model vs consensus").font = F_B
    _hdr(ws, r + 1, ["", "2026E", "2027E", "2028E", "2029E", "2030E"])
    f = res["dcf"]["forecast"]
    for i, (lab, fk, ck_) in enumerate([("Revenue: model", "revenue", None), ("Revenue: consensus", None, "revenue"),
                                        ("Non-IFRS op. profit: model", "op_nonifrs", None), ("EBIT: consensus", None, "ebit")], start=r + 2):
        ws.cell(row=i, column=1, value=lab)
        for j, y in enumerate(range(2026, 2031), start=2):
            v = f.loc[y, fk] if fk else cons.get(y, {}).get(ck_)
            ws.cell(row=i, column=j, value=None if v is None else float(v)).number_format = EUR
    _widths(ws, {"A": 34, **{get_column_letter(j): 12 for j in range(2, 12)}})


def _business(wb, res, out_dir):
    ws = wb.create_sheet("Business model")
    ds = res["ds"]
    ir = ds["ir"]["q3_2026_statement"]
    fy = ds["ir"]["fy2025_statement"]
    _title(ws, "Business model & cloud transformation KPIs", "Sources: SAP quarterly statements Q4-2025 and Q3-2026 (SEC 6-K), CIQ")
    rows = [
        ("Cloud ERP Suite revenue", fy["fy2024"]["cloud_erp_suite"], fy["fy2025"]["cloud_erp_suite"], ir["ytd9m_prior"]["cloud_erp_suite"], ir["ytd9m"]["cloud_erp_suite"]),
        ("Total cloud revenue", fy["fy2024"]["cloud"], fy["fy2025"]["cloud"], ir["ytd9m_prior"]["cloud"], ir["ytd9m"]["cloud"]),
        ("Software licences", fy["fy2024"]["licenses"], fy["fy2025"]["licenses"], ir["ytd9m_prior"]["licenses"], ir["ytd9m"]["licenses"]),
        ("Software support", fy["fy2024"]["support"], fy["fy2025"]["support"], ir["ytd9m_prior"]["support"], ir["ytd9m"]["support"]),
        ("Services", fy["fy2024"]["services"], fy["fy2025"]["services"], ir["ytd9m_prior"]["services"], ir["ytd9m"]["services"]),
        ("Total revenue", fy["fy2024"]["revenue"], fy["fy2025"]["revenue"], ir["ytd9m_prior"]["revenue"], ir["ytd9m"]["revenue"]),
        ("Cost of cloud", fy["fy2024"]["cost_of_cloud"], fy["fy2025"]["cost_of_cloud"], None, ir["ytd9m"]["cost_of_cloud"]),
        ("Non-IFRS operating profit", fy["fy2024"]["op_nonifrs"], fy["fy2025"]["op_nonifrs"], ir["ytd9m_prior"]["op_nonifrs"], ir["ytd9m"]["op_nonifrs"]),
        ("Share-based compensation", fy["fy2024"]["sbc"], fy["fy2025"]["sbc"], None, ir["ytd9m"]["sbc"]),
        ("Current cloud backlog (period end)", fy["fy2024"]["current_cloud_backlog"], fy["fy2025"]["current_cloud_backlog"], ir["current_cloud_backlog_prior_year"], ir["current_cloud_backlog"]),
    ]
    _hdr(ws, 4, ["€m", "FY2024", "FY2025", "9M-2025", "9M-2026", "Δ FY25", "Δ 9M-26"])
    for i, row in enumerate(rows, start=5):
        ws.cell(row=i, column=1, value=row[0])
        for j, v in enumerate(row[1:], start=2):
            ws.cell(row=i, column=j, value=v).number_format = EUR
        ws.cell(row=i, column=6, value=f"=C{i}/B{i}-1").number_format = PCT
        ws.cell(row=i, column=7, value=f'=IF(D{i}="","",E{i}/D{i}-1)').number_format = PCT
    r = 5 + len(rows) + 1
    kpis = [
        ("Cloud share of revenue", "=C6/C10", "=E6/E10"),
        ("Cloud ERP Suite share of cloud", "=C5/C6", "=E5/E6"),
        ("Cloud gross margin (IFRS)", "=1-C11/C6", "=1-E11/E6"),
        ("Non-IFRS operating margin (incl. SBC)", "=C12/C10", "=E12/E10"),
        ("SBC % revenue", "=C13/C10", "=E13/E10"),
        ("Current cloud backlog / next-year cloud revenue proxy", "=C14/C6", "=E14/(E6/9*12)"),
    ]
    _hdr(ws, r, ["KPI", "FY2025", "9M-2026"])
    for i, (lab, f1, f2) in enumerate(kpis, start=r + 1):
        ws.cell(row=i, column=1, value=lab)
        ws.cell(row=i, column=2, value=f1).number_format = PCT
        ws.cell(row=i, column=3, value=f2).number_format = PCT
    r = r + len(kpis) + 2
    text = [
        "How SAP makes money: ERP-centred enterprise software (finance, supply chain, HR, procurement, CX) sold increasingly as cloud subscriptions (RISE/GROW with SAP, Cloud ERP Suite) instead of perpetual licences + 22% annual maintenance.",
        "Transformation economics: a licence customer moving to cloud lifts lifetime revenue (subscription ≈ 2-3x maintenance) but compresses revenue for 2-3 years (J-curve, 2020-23); 2024-26 is the harvesting phase - margins and FCF expand as the cloud base scales.",
        "Visibility: current cloud backlog €24.1bn (+28%) ≈ 90% of the next 12 months' cloud revenue is contracted; contract liabilities (cash collected in advance) €8.4bn at Q3-26 → negative working capital.",
        "Intangibles: R&D ~18% of revenue is expensed; customer relationships/installed base (switching costs, ECC end of maintenance 2027/2030), data and the partner ecosystem are not on the balance sheet. See Intangibles sheet.",
        "Capital-light: cloud runs largely on hyperscalers (cost of cloud) → capex ~2% of revenue; PP&E only 6% of total assets.",
    ]
    for t in text:
        ws.cell(row=r, column=1, value="• " + t); r += 1
    img = XLImage(str(out_dir / f"revenue_mix_{res['edition']}.png")); img.width, img.height = 720, 336
    ws.add_image(img, f"A{r+1}")
    _widths(ws, {"A": 52, **{get_column_letter(j): 12 for j in range(2, 8)}})


def _intangibles(wb, res):
    ws = wb.create_sheet("Intangibles")
    _title(ws, "Role of intangibles: R&D capitalisation (5-year life) and market vs. book value")
    rd = res["rd"]
    keys = [("revenue", "Revenue", EUR), ("rnd", "R&D expense", EUR), ("rnd_pct_rev", "R&D % revenue", PCT),
            ("rnd_asset", "Capitalised R&D asset (unrecognised)", EUR), ("rnd_amort", "R&D amortisation", EUR),
            ("ebit_reported", "EBIT reported (IFRS)", EUR), ("ebit_adjusted", "EBIT adjusted (R&D capitalised)", EUR),
            ("ebit_margin_reported", "EBIT margin reported", PCT), ("ebit_margin_adjusted", "EBIT margin adjusted", PCT),
            ("invested_capital", "Invested capital (book, incl. goodwill)", EUR), ("invested_capital_adj", "Invested capital + R&D asset", EUR),
            ("ic_ex_goodwill", "Invested capital ex goodwill & acquired intangibles", EUR),
            ("roic_reported", "ROIC reported (after 29% tax)", PCT), ("roic_adjusted", "ROIC with R&D capitalised", PCT),
            ("roic_ex_goodwill", "ROIC ex goodwill (organic capital)", PCT),
            ("intangibles_share_assets", "Goodwill + intangibles / total assets", PCT), ("ppe_share_assets", "PP&E / total assets", PCT)]
    _hdr(ws, 3, ["€m"] + [f"FY{y}" for y in rd.index])
    for i, (k, lab, fm) in enumerate(keys, start=4):
        ws.cell(row=i, column=1, value=lab)
        for j, y in enumerate(rd.index, start=2):
            v = rd.loc[y, k]
            ws.cell(row=i, column=j, value=None if v is None or v != v else float(v)).number_format = fm
    r = 5 + len(keys)
    m = res["mvb"]
    _hdr(ws, r, ["Market vs book (FY2025 balance sheet, current market cap)", "€m / x"])
    items = [("Market capitalisation", m["market_cap"], EUR), ("Book equity", m["book_equity"], EUR),
             ("  of which goodwill + acquired intangibles", m["goodwill_intangibles"], EUR), ("Tangible book equity", m["tangible_book"], EUR),
             ("Unrecognised capitalised R&D (5y)", m["rd_asset_unrecognised"], EUR),
             ("Market value not explained by book equity or R&D asset", m["unexplained_by_balance_sheet"], EUR),
             ("Price / book", m["p_b"], MULT), ("Price / tangible book", m["p_tb"], MULT),
             ("Share of market value off balance sheet", m["share_of_mcap_off_balance_sheet"], PCT)]
    for i, (lab, v, fm) in enumerate(items, start=r + 1):
        ws.cell(row=i, column=1, value=lab); ws.cell(row=i, column=2, value=float(v)).number_format = fm
    r = r + len(items) + 2
    text = [
        "Answer to the case question: yes - SAP's value is overwhelmingly intangible. ~80% of market value is not on the balance sheet; PP&E is ~6% of assets; goodwill + acquired intangibles are ~45% of assets but only ~13% of market value.",
        "Accounting distortion: IFRS expenses internally generated R&D (~€6.6bn p.a.). Capitalising it raises the EBIT margin by ~2-4pp in investment years and adds a ~€19bn asset; reported ROIC overstates and ROIC ex goodwill (>100%) shows the organic business needs almost no tangible capital.",
        "Valuation consequence: book-based metrics (P/B, EV/invested capital) are meaningless; EBITDA understates the reinvestment burden hidden in opex (R&D, S&M for customer acquisition). FCFF already treats R&D as a cash cost, so the DCF is unaffected by capitalisation; the economic 'capex' of SAP is mostly in R&D and SBC, which we keep as costs.",
        "Moat: switching costs of the ERP installed base (~>400k customers), ECC end-of-maintenance deadlines, data gravity and partner ecosystem support the long fade period (to 2035) and a 35% terminal margin; AI-native disruption is the main risk to this intangible base (Bear case).",
    ]
    for t in text:
        ws.cell(row=r, column=1, value="• " + t); r += 1
    _widths(ws, {"A": 60, **{get_column_letter(j): 13 for j in range(2, 8)}})


def _recon(wb, res):
    ws = wb.create_sheet("FCFF bridge")
    _title(ws, "FY2025: FCFF on model definition vs. SAP-reported free cash flow (€m)")
    _hdr(ws, 3, ["Item", "€m"])
    for i, (_, row) in enumerate(res["recon"].iterrows(), start=4):
        ws.cell(row=i, column=1, value=row["item"])
        ws.cell(row=i, column=2, value=float(row["eur_m"])).number_format = EUR
        if row["item"].startswith("="):
            ws.cell(row=i, column=1).font = F_B; ws.cell(row=i, column=2).font = F_B
    ws.cell(row=17, column=1, value="Why the model FCFF is below SAP's FCF: SAP adds back non-cash SBC (we treat SBC as a cost - it dilutes, buybacks offset it with cash) and pays cash taxes below the book rate (timing). Interest is excluded from FCFF by definition.")
    _widths(ws, {"A": 60, "B": 14})


def _sources(wb, res):
    ws = wb.create_sheet("Sources & notes")
    ds = res["ds"]
    _title(ws, "Data sources, conflicts and methodology notes")
    _hdr(ws, 3, ["Data", "Source"])
    r = 4
    for k, v in ds["sources"]:
        ws.cell(row=r, column=1, value=k); ws.cell(row=r, column=2, value=v); r += 1
    r += 1
    ws.cell(row=r, column=1, value="Data conflicts and how they were resolved").font = F_B; r += 1
    conflicts = [
        ("WACC", "Bloomberg 12.6% vs CIQ 7.49% vs model", "Bloomberg's 10.32% market premium rejected; bottom-up CAPM with ERP 5.0%"),
        ("Market cap", "CIQ €247bn (1,154m shares) vs Bloomberg WACC screen €213bn vs 1,229m issued", "Shares outstanding net of treasury (1,154.2m) + 7m dilution"),
        ("Debt", "CIQ Q3 10,588 vs Bloomberg 7,559 vs bond principal 8,050", "Reported financial liabilities 10,588 (bonds + CP + leases) - consistent with IFRS 16 FCFF"),
        ("EBIT basis", "CIQ EBIT 10,418 ≈ non-IFRS 10,419; IFRS 9,830 (SAP) vs 9,617 (CIQ as reported)", "Model builds from non-IFRS (incl. SBC) less acquisition-related charges"),
        ("Tax rate", "27.9% (Bloomberg), 28.7% (CIQ ETR), 29% (guidance), 29.9% (CIQ DCF)", "29% 2026, 29.5% 2027-30, 30% terminal"),
        ("Buyback avg. price", "Bloomberg €212.43 vs SAP Q3 statement €164.79", "Not a valuation input; noted only"),
        ("Valuation date", "13 Nov 2026 is after the data cut-off (5-9 Nov)", "Value anchored at 30 Sep balance sheet, rolled forward 44 days at Ke; price can be overridden at run time"),
    ]
    _hdr(ws, r, ["Item", "Conflict", "Resolution"]); r += 1
    for c in conflicts:
        for j, v in enumerate(c, start=1):
            ws.cell(row=r, column=j, value=v)
        r += 1
    r += 1
    ws.cell(row=r, column=1, value="Colour code: blue = hard-coded input, green = link, black = formula, grey italics = Python engine check").font = Font(italic=True)
    _widths(ws, {"A": 34, "B": 80, "C": 90})


def _summary(wb, ref, wref, dref, cref, res, out_dir):
    ws = wb["Summary"]
    d, sc, cp, mk = res["dcf"], res["scenarios"], res["comps"], res["market"]
    _title(ws, f"SAP SE - company valuation as of 13 Nov 2026 ({res['edition'].upper()} edition)",
           "FCFF DCF + relative valuation | € per share | all figures link to the model sheets")
    _hdr(ws, 4, ["Key outputs", "Value"])
    impl = cref["impl"]; rb = cref["reg_base"]
    rows = [
        ("DCF value per share (base)", f"={dref['price']}", PX),
        ("Reference share price", f"={ref['price']}", PX),
        ("Upside / (downside) vs price", f"={dref['upside']}", PCT),
        ("Probability-weighted scenario value", "=Scenarios!C8", PX),
        ("WACC", f"={wref['wacc']}", PCT2),
        ("Terminal growth", f"={ref['g']}", PCT2),
        ("Enterprise value (€m)", f"={dref['ev']}", EUR),
        ("Equity value (€m)", f"={dref['equity']}", EUR),
        ("Terminal value share of EV", f"={dref['tv_share']}", PCT),
        ("Comps: EV/NTM sales (median)", f"=Comps!G{impl['EV / NTM sales']}", PX),
        ("Comps: EV/NTM EBIT (median)", f"=Comps!G{impl['EV / NTM EBIT']}", PX),
        *([("Comps: EV/NTM EBIT ex-SBC (median)", f"=Comps!G{impl['EV / NTM EBIT ex-SBC (like-for-like)']}", PX)] if 'EV / NTM EBIT ex-SBC (like-for-like)' in impl else []),
        ("Comps: P/NTM EPS (median)", f"=Comps!G{impl['P / NTM EPS']}", PX),
        ("Comps: growth-adjusted EV/sales", f"=Comps!B{rb+8}", PX),
        ("Analyst target mean (n=%d)" % mk["tgt_n"], mk["tgt_mean"], PX),
        ("52-week range", f"€{mk['lo52']:.2f} - €{mk['hi52']:.2f}", None),
    ]
    for i, (lab, f, fm) in enumerate(rows, start=5):
        ws.cell(row=i, column=1, value=lab)
        c = ws.cell(row=i, column=2, value=f)
        if fm:
            c.number_format = fm
    ws["A5"].font = F_B; ws["B5"].font = Font(bold=True, size=12, color=TEAL); ws["B5"].fill = FILL_K
    r = 23
    ws.cell(row=r, column=1, value="Key assumptions").font = F_B
    a = res["assumptions"]
    ka = [
        f"Revenue: cloud +{a['q4g_cloud']*100:.0f}% Q4-26, then {', '.join(f'{x*100:.0f}%' for x in a['g_cloud'])} (2027-30); support {a['g_support'][0]*100:.0f}% to {a['g_support'][-1]*100:.0f}% p.a.; licences -25% p.a. → 2030E revenue €{d['forecast'].loc[2030,'revenue']/1000:.1f}bn" + (f" (consensus €{res['consensus'][2030]['revenue']/1000:.1f}bn)" if res['consensus'].get(2030, {}).get('revenue') else ""),
        f"Non-IFRS operating margin (SBC included) {a['m_nonifrs'][0]*100:.1f}% (2026) → {a['m_nonifrs'][-1]*100:.1f}% (2030) → {a['m_terminal']*100:.0f}% terminal; tax 29-30%",
        f"Capex {a['capex_pct'][0]*100:.1f}-{a['capex_terminal']*100:.1f}% and D&A {a['da_pct'][0]*100:.1f}-{a['da_terminal']*100:.1f}% of revenue; NWC releases {a['nwc_pct']*100:.0f}% of incremental revenue (prepaid subscriptions)",
        f"WACC {d['wacc']*100:.2f}% = Ke {d['ke']*100:.2f}% (Rf {a['rf']*100:.2f}%, β {a['beta']:.2f}, ERP {a['erp']*100:.1f}%) x {d['we']*100:.0f}% + Kd after tax {d['kd_at']*100:.2f}% x {d['wd']*100:.0f}%; g {a['g']*100:.2f}%",
        f"Bridge (30 Sep 2026): + cash & financial assets €{(a['cash']+a['st_inv'])/1000:.1f}bn + investments €{a['lt_inv']/1000:.1f}bn - debt incl. leases €{a['debt']/1000:.1f}bn - pensions/NCI €{(a['pension']+a['nci'])/1000:.1f}bn; {a['shares']:.1f}m diluted shares",
        f"Peers ({cp['stats']['ev_sales_ntm']['n']} core): ORCL, MSFT, CRM, NOW, ADBE, WDAY, ADSK, DSY, PTC, SGE, OTEX, NEM; NTM EV/Sales, EV/EBIT, P/E; EV/EBITDA and P/B ruled out",
    ]
    for i, t in enumerate(ka, start=r + 1):
        ws.cell(row=i, column=1, value="• " + t)
    img = XLImage(str(out_dir / f"football_{res['edition']}.png")); img.width, img.height = 880, 448
    ws.add_image(img, "D4")
    _widths(ws, {"A": 46, "B": 18, "C": 3})


def write(res: dict, path: Path, out_dir: Path) -> Path:
    wb = Workbook()
    wb.active.title = "Summary"
    a = res["assumptions"]
    ref = _assumptions(wb, a, res)
    wref = _wacc(wb, ref, res)
    R = _forecast(wb, ref, res)
    dref = _dcf(wb, ref, wref, R, res)
    _scenarios(wb, res)
    _sensitivity(wb, res, out_dir)
    cref = _comps(wb, ref, res)
    _business(wb, res, out_dir)
    _intangibles(wb, res)
    _historicals(wb, res)
    _recon(wb, res)
    _sources(wb, res)
    _summary(wb, ref, wref, dref, cref, res, out_dir)
    wb.save(path)
    return path
