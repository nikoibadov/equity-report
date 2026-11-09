"""Deterministic FCFF / DCF engine. The Excel workbook re-implements exactly these formulas,
so any change here must be mirrored in excel_writer.py (the test suite checks both agree)."""
from __future__ import annotations

import numpy as np
import pandas as pd

from .assumptions import EXPLICIT, FADE, YEARS, SCENARIOS, SCENARIO_WEIGHTS, Assumptions


def wacc(a: Assumptions) -> dict:
    ke = a["rf"] + a["beta"] * a["erp"]
    kd_at = a["kd"] * (1 - a["tax_wacc"])
    v = a["e_mv"] + a["d_mv"]
    we, wd = a["e_mv"] / v, a["d_mv"] / v
    return {"ke": ke, "kd_at": kd_at, "we": we, "wd": wd, "wacc": we * ke + wd * kd_at}


def _fade(start, end, k, n=5):
    return start + (end - start) * k / n


def forecast(a: Assumptions) -> pd.DataFrame:
    rows = {}
    # 2026 anchor: 9M actual + Q4-2025 x (1 + Q4 growth)
    seg26 = {}
    for s in ("cloud", "licenses", "support", "services"):
        q4_prior = a[f"base_{s}"] - a[f"ytdp_{s}"]
        seg26[s] = a[f"ytd_{s}"] + q4_prior * (1 + a[f"q4g_{s}"])
    seg = {2026: seg26}
    for i, y in enumerate(EXPLICIT):
        seg[y] = {s: seg[y - 1][s] * (1 + a[f"g_{s}"][i]) for s in ("cloud", "licenses", "support", "services")}
    rev = {y: sum(seg[y].values()) for y in [2026] + EXPLICIT}
    rev_base = a["base_cloud"] + a["base_licenses"] + a["base_support"] + a["base_services"]
    g2030 = rev[2030] / rev[2029] - 1
    for k, y in enumerate(FADE, start=1):
        rev[y] = rev[y - 1] * (1 + _fade(g2030, a["g"], k))

    def path(key, term_key):
        vals = list(a[key])  # 2026..2030
        out = dict(zip([2026] + EXPLICIT, vals))
        for k, y in enumerate(FADE, start=1):
            out[y] = _fade(vals[-1], a[term_key], k)
        return out

    m = path("m_nonifrs", "m_terminal")
    acq = path("acq_pct", "acq_terminal")
    tax = path("tax", "tax_terminal")
    da = path("da_pct", "da_terminal")
    cx = path("capex_pct", "capex_terminal")

    prev = rev_base
    for y in YEARS:
        r = rev[y]
        op_nonifrs = r * m[y]
        acq_c = r * acq[y]
        oneoff = a["oneoff_2026"] if y == 2026 else 0.0
        ebit = op_nonifrs - acq_c - oneoff
        nopat = ebit * (1 - tax[y])
        sbc_add = r * a["sbc_pct"] * a["sbc_noncash"] * a["sbc_addback"]
        d_a = r * da[y]
        capex = r * cx[y]
        dnwc = -a["nwc_pct"] * (r - prev)               # negative = cash inflow
        fcff = nopat + sbc_add + d_a - capex - dnwc
        s = seg.get(y, {})
        rows[y] = {
            "cloud": s.get("cloud"), "licenses": s.get("licenses"), "support": s.get("support"), "services": s.get("services"),
            "revenue": r, "growth": r / prev - 1, "cloud_share": (s["cloud"] / r) if s else None,
            "margin_nonifrs": m[y], "op_nonifrs": op_nonifrs, "acq_related": acq_c, "oneoff": oneoff,
            "ebit": ebit, "ebit_margin": ebit / r, "tax_rate": tax[y], "nopat": nopat, "sbc_addback": sbc_add,
            "da": d_a, "capex": capex, "dnwc": dnwc, "fcff": fcff,
        }
        prev = r
    return pd.DataFrame(rows).T


def dcf(a: Assumptions, wacc_override: float | None = None) -> dict:
    w = wacc(a)
    r = w["wacc"] if wacc_override is None else wacc_override
    f = forecast(a)
    stub = a["stub"]
    t, cf = {}, {}
    for y in YEARS:
        if y == 2026:
            cf[y] = f.loc[y, "fcff"] * stub
            t[y] = stub / 2
        else:
            cf[y] = f.loc[y, "fcff"]
            t[y] = stub + (y - 2027) + 0.5
    df = {y: 1 / (1 + r) ** t[y] for y in YEARS}
    pv = {y: cf[y] * df[y] for y in YEARS}
    # Terminal: normalised year T+1 at terminal margin/tax/D&A/capex
    last = f.loc[YEARS[-1]]
    rev_t1 = last["revenue"] * (1 + a["g"])
    ebit_t1 = rev_t1 * (a["m_terminal"] - a["acq_terminal"])
    fcff_t1 = (ebit_t1 * (1 - a["tax_terminal"])
               + rev_t1 * a["sbc_pct"] * a["sbc_noncash"] * a["sbc_addback"]
               + rev_t1 * (a["da_terminal"] - a["capex_terminal"])
               + a["nwc_pct"] * (rev_t1 - last["revenue"]))
    tv = fcff_t1 / (r - a["g"])
    t_tv = stub + (YEARS[-1] - 2027) + 1
    pv_tv = tv / (1 + r) ** t_tv
    ev = sum(pv.values()) + pv_tv
    non_op = a["cash"] + a["st_inv"] + a["lt_inv"]
    claims = a["debt"] + a["pension"] + a["nci"]
    equity = ev + non_op - claims
    ps_anchor = equity / a["shares"]
    roll = (1 + w["ke"]) ** (a["roll_days"] / 365)
    price = ps_anchor * roll
    ntm_rev = 0.25 * f.loc[2026, "revenue"] + 0.75 * f.loc[2027, "revenue"]
    ntm_ebit = 0.25 * f.loc[2026, "op_nonifrs"] + 0.75 * f.loc[2027, "op_nonifrs"]
    return {
        "wacc": r, **{k: v for k, v in w.items() if k != "wacc"}, "forecast": f, "cf": cf, "t": t, "df": df, "pv": pv,
        "sum_pv": sum(pv.values()), "fcff_t1": fcff_t1, "ebit_t1": ebit_t1, "tv": tv, "pv_tv": pv_tv, "t_tv": t_tv,
        "ev": ev, "non_op": non_op, "claims": claims, "equity": equity, "ps_anchor": ps_anchor, "roll": roll,
        "price": price, "tv_share": pv_tv / ev, "exit_ev_ebit": tv / ebit_t1,
        "implied_ev_sales_ntm": ev / ntm_rev, "implied_ev_ebit_ntm": ev / ntm_ebit,
        "upside": price / a["price"] - 1,
    }


# ---------------------------------------------------------------- scenarios
def apply_scenario(a: Assumptions, name: str) -> tuple[Assumptions, float]:
    s = SCENARIOS[name]
    b = a.copy()
    b.set("g_cloud", [x + s["d_cloud"] for x in a["g_cloud"]])
    b.set("m_nonifrs", [a["m_nonifrs"][0]] + [x + s["d_margin"] for x in a["m_nonifrs"][1:]])
    if s["m_terminal"] is not None:
        b.set("m_terminal", s["m_terminal"])
    if s["g"] is not None:
        b.set("g", s["g"])
    return b, wacc(a)["wacc"] + s["d_wacc"]


def scenarios(a: Assumptions) -> dict:
    out = {}
    for n in SCENARIOS:
        b, w = apply_scenario(a, n)
        r = dcf(b, w)
        out[n] = {"price": r["price"], "wacc": w, "g": b["g"], "m_terminal": b["m_terminal"],
                  "rev_2030": r["forecast"].loc[2030, "revenue"], "weight": SCENARIO_WEIGHTS[n], "story": SCENARIOS[n]["story"]}
    out["Weighted"] = {"price": sum(v["price"] * v["weight"] for v in out.values())}
    return out


# ---------------------------------------------------------------- sensitivities
def grid_wacc_g(a, waccs, gs):
    res = pd.DataFrame(index=[round(w, 4) for w in waccs], columns=[round(g, 4) for g in gs], dtype=float)
    for w in waccs:
        for g in gs:
            b = a.copy(); b.set("g", g)
            res.loc[round(w, 4), round(g, 4)] = dcf(b, w)["price"]
    return res


def grid_wacc_margin(a, waccs, margins):
    res = pd.DataFrame(index=[round(w, 4) for w in waccs], columns=[round(m, 4) for m in margins], dtype=float)
    for w in waccs:
        for m in margins:
            b = a.copy(); b.set("m_terminal", m)
            res.loc[round(w, 4), round(m, 4)] = dcf(b, w)["price"]
    return res


def grid_cloud_margin(a, d_cloud, d_margin):
    res = pd.DataFrame(index=[round(x, 4) for x in d_cloud], columns=[round(x, 4) for x in d_margin], dtype=float)
    for dc in d_cloud:
        for dm in d_margin:
            b = a.copy()
            b.set("g_cloud", [x + dc for x in a["g_cloud"]])
            b.set("m_nonifrs", [a["m_nonifrs"][0]] + [x + dm for x in a["m_nonifrs"][1:]])
            b.set("m_terminal", a["m_terminal"] + dm)
            res.loc[round(dc, 4), round(dm, 4)] = dcf(b)["price"]
    return res


def tornado(a) -> pd.DataFrame:
    base = dcf(a)["price"]
    tests = [
        ("Risk-free rate ±50bp", "rf", lambda v: v - 0.005, lambda v: v + 0.005),
        ("Beta ±0.10", "beta", lambda v: v - 0.10, lambda v: v + 0.10),
        ("ERP ±50bp", "erp", lambda v: v - 0.005, lambda v: v + 0.005),
        ("Terminal growth ±25bp", "g", lambda v: v + 0.0025, lambda v: v - 0.0025),
        ("Terminal margin ±2pp", "m_terminal", lambda v: v + 0.02, lambda v: v - 0.02),
        ("Terminal tax ±2pp", "tax_terminal", lambda v: v - 0.02, lambda v: v + 0.02),
        ("Terminal capex ±0.5pp", "capex_terminal", lambda v: v - 0.005, lambda v: v + 0.005),
        ("NWC release 10% / 0%", "nwc_pct", lambda v: 0.10, lambda v: 0.0),
        ("Cloud growth ±2pp p.a.", "g_cloud", lambda v: [x + 0.02 for x in v], lambda v: [x - 0.02 for x in v]),
        ("LT investments 100% / 50%", "lt_inv", lambda v: v, lambda v: v * 0.5),
        ("SBC non-cash add-back on / off", "sbc_addback", lambda v: 1, lambda v: 0),
    ]
    rows = []
    for label, key, up, dn in tests:
        hi = a.copy(); hi.set(key, up(a[key]))
        lo = a.copy(); lo.set(key, dn(a[key]))
        p_hi, p_lo = dcf(hi)["price"], dcf(lo)["price"]
        rows.append({"driver": label, "low": min(p_hi, p_lo), "high": max(p_hi, p_lo), "base": base,
                     "range": abs(p_hi - p_lo)})
    return pd.DataFrame(rows).sort_values("range", ascending=False).reset_index(drop=True)


def fcff_2025_reconciliation(ds: dict, a: Assumptions) -> pd.DataFrame:
    """Rebuild FY2025 FCFF on the model definition and bridge to SAP's reported FCF."""
    h = ds["hist"][2025]
    fy = ds["ir"]["fy2025_statement"]["fy2025"]
    op_ifrs = fy["op_ifrs"]
    tax = 0.29
    nopat = op_ifrs * (1 - tax)
    wc = -675.0 if ds["edition"] == "data" else -675.0   # CF statement: AR -388, payables -308, contract liab +1,336, other -1,315
    fcff = nopat + h["da"] - h["capex"] + wc
    sbc_noncash = h["sbc"] - h["sbc_cash_settled"]
    tax_gap = op_ifrs * tax - h["cash_taxes"]
    other = fy["fcf"] - (fcff + sbc_noncash + tax_gap)
    rows = [
        ("IFRS operating profit FY25", op_ifrs),
        ("- Taxes at 29% on EBIT", -op_ifrs * tax),
        ("= NOPAT", nopat),
        ("+ D&A (as reported)", h["da"]),
        ("- Capex (PP&E + intangibles)", -h["capex"]),
        ("+/- Working-capital change (CF statement)", wc),
        ("= FCFF on model definition", fcff),
        ("+ Non-cash (equity-settled) SBC", sbc_noncash),
        ("+ Cash taxes below 29% book rate", tax_gap),
        ("+/- Interest, lease payments, other", other),
        ("= SAP reported free cash flow", fy["fcf"]),
    ]
    return pd.DataFrame(rows, columns=["item", "eur_m"])
