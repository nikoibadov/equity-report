"""Relative valuation: peer multiples -> implied SAP share price ranges."""
from __future__ import annotations

import numpy as np
import pandas as pd

# Metrics used for the valuation range, and why EV/EBITDA is shown but not used.
METRICS = {
    "ev_sales_ntm": ("EV / NTM sales", "rev_ntm", "ev"),
    "ev_ebit_ntm": ("EV / NTM EBIT", "ebit_ntm", "ev"),
    "pe_ntm": ("P / NTM EPS", "eps_ntm", "price"),
}
EXCLUDED = {
    "ev_ebitda_ntm": "Shown for reference only: D&A at SAP is small and partly amortisation of acquired intangibles, and US peers' EBITDA excludes SBC - EBITDA adds no information beyond EBIT for asset-light software.",
    "pb": "Not used: book equity is dominated by goodwill (€29bn) while internally generated intangibles (R&D, customer base) are off balance sheet.",
}


def sap_ntm(ds: dict) -> dict:
    """SAP next-twelve-months metrics from consensus (anchor 30 Sep 2026: 1/4 FY26 + 3/4 FY27)."""
    c = ds["consensus"]
    w = (0.25, 0.75)
    def blend(k):
        a, b = c[2026].get(k), c[2027].get(k)
        return None if a is None or b is None else w[0] * a + w[1] * b
    return {"rev_ntm": blend("revenue"), "ebit_ntm": blend("ebit"), "ebitda_ntm": blend("ebitda"),
            "eps_ntm": blend("eps_norm"), "rev_ltm": ds["hist_ltm"]["revenue"], "ebit_margin_ltm": ds["hist_ltm"]["ebit_margin"]}


def implied(ds: dict, a, peers: pd.DataFrame) -> dict:
    core = peers[peers["core"]].copy()
    s = sap_ntm(ds)
    non_op = a["cash"] + a["st_inv"] + a["lt_inv"] + a["debt_benefit"] + a["tlcf_asset"]
    claims = a["debt"] + a["pension"] + a["nci"]
    shares = a["shares"]
    out = {"sap": s, "stats": {}, "rows": []}
    for key, (label, metric, kind) in METRICS.items():
        x = core[key].dropna()
        x = x[x > 0]
        q1, med, q3 = np.percentile(x, 25), np.median(x), np.percentile(x, 75)
        def to_price(m):
            if kind == "ev":
                return (m * s[metric] + non_op - claims) / shares
            return m * s[metric]
        out["stats"][key] = {"label": label, "n": len(x), "q1": q1, "median": med, "q3": q3, "mean": x.mean(),
                             "p_q1": to_price(q1), "p_med": to_price(med), "p_q3": to_price(q3),
                             "sap_metric": s[metric], "sap_current": None}
    # Like-for-like EV/EBIT: US peers' street EBIT is typically non-GAAP excluding SBC, SAP's
    # non-IFRS EBIT includes SBC -> also apply peer multiples to SAP EBIT before SBC.
    sbc_ntm = a["sbc_pct"] * s["rev_ntm"]
    out["sap"]["sbc_ntm"] = sbc_ntm
    out["ebit_basis"] = ds.get("peer_ebit_basis", "street")
    if out["ebit_basis"] != "street":
        out["stats"]["ev_ebit_ntm"]["label"] = "EV / NTM EBIT (GAAP margins, SBC-burdened)"
    st = dict(out["stats"]["ev_ebit_ntm"])
    st.update(label="EV / NTM EBIT ex-SBC (like-for-like)", sap_metric=s["ebit_ntm"] + sbc_ntm,
              p_q1=(st["q1"] * (s["ebit_ntm"] + sbc_ntm) + non_op - claims) / shares,
              p_med=(st["median"] * (s["ebit_ntm"] + sbc_ntm) + non_op - claims) / shares,
              p_q3=(st["q3"] * (s["ebit_ntm"] + sbc_ntm) + non_op - claims) / shares)
    if out["ebit_basis"] == "street":
        out["stats"]["ev_ebit_ntm_xsbc"] = st
    # Growth-adjusted: EV/NTM sales regressed on Rule-of-40 (NTM growth + LTM EBIT margin)
    reg = core[["ev_sales_ntm", "rule40", "ticker"]].dropna()
    X = np.column_stack([np.ones(len(reg)), reg["rule40"].values])
    y = reg["ev_sales_ntm"].values
    beta, *_ = np.linalg.lstsq(X, y, rcond=None)
    resid = y - X @ beta
    se = np.sqrt(resid @ resid / (len(y) - 2))
    r2 = 1 - (resid @ resid) / ((y - y.mean()) @ (y - y.mean()))
    sap_g = (s["rev_ntm"] / s["rev_ltm"] - 1) * 100
    sap_r40 = sap_g + s["ebit_margin_ltm"]
    fit = beta[0] + beta[1] * sap_r40
    def p(m):
        return (m * s["rev_ntm"] + non_op - claims) / shares
    out["regression"] = {"intercept": beta[0], "slope": beta[1], "r2": r2, "se": se, "n": len(y),
                         "sap_growth": sap_g, "sap_rule40": sap_r40, "fit": fit,
                         "p_low": p(fit - se), "p_fit": p(fit), "p_high": p(fit + se),
                         "points": reg.assign(fitted=X @ beta)}
    return out
