"""One call from a loaded dataset to the full result bundle used by Excel, markdown and Streamlit."""
from __future__ import annotations

import numpy as np

from . import assumptions as A
from . import comps, intangibles, model


def run(ds: dict, a: A.Assumptions | None = None) -> dict:
    a = a or A.build(ds)
    d = model.dcf(a)
    w = d["wacc"]
    f = d["forecast"]
    ds.setdefault("consensus_gaps", [])
    for y in (2026, 2027):
        c = ds["consensus"].setdefault(y, {})
        if c.get("ebit") is None:
            c["ebit"] = float(f.loc[y, "op_nonifrs"])
            ds["consensus_gaps"].append(f"EBIT {y}: no free consensus, model non-IFRS operating profit used")
    waccs = [round(w + x, 4) for x in np.arange(-0.01, 0.0101, 0.0025)]
    gs = [round(a["g"] + x, 4) for x in np.arange(-0.0075, 0.00751, 0.0025)]
    margins = [round(a["m_terminal"] + x, 4) for x in np.arange(-0.04, 0.0401, 0.01)]
    g_wg = model.grid_wacc_g(a, waccs, gs)
    band_vals = [model.dcf(_with(a, "g", a["g"] + dg), w + dw)["price"] for dw in (-0.005, 0, 0.005) for dg in (-0.0025, 0, 0.0025)]
    res = {
        "edition": ds["edition"], "meta": ds["meta"], "market": ds["market"], "hist": ds["hist"],
        "assumptions": a, "dcf": d, "scenarios": model.scenarios(a),
        "grid_wacc_g": g_wg,
        "grid_wacc_margin": model.grid_wacc_margin(a, waccs, margins),
        "grid_cloud_margin": model.grid_cloud_margin(a, [-0.04, -0.02, 0, 0.02, 0.04], [-0.03, -0.015, 0, 0.015, 0.03]),
        "grid_erp_beta": _grid_erp_beta(a),
        "dcf_band": (min(band_vals), max(band_vals)),
        "tornado": model.tornado(a),
        "comps": comps.implied(ds, a, ds["peers"]),
        "peers": ds["peers"],
        "rd": intangibles.rd_capitalisation(ds, a),
        "mvb": intangibles.market_vs_book(ds, a),
        "recon": model.fcff_2025_reconciliation(ds, a),
        "consensus": ds["consensus"],
        "ds": ds,
    }
    return res


def _with(a, k, v):
    b = a.copy(); b.set(k, v); return b


def _grid_erp_beta(a):
    import pandas as pd
    erps = [0.042, 0.045, 0.05, 0.055, 0.06]
    betas = [0.85, 0.92, 0.98, 1.05, 1.10]
    g = pd.DataFrame(index=betas, columns=erps, dtype=float)
    for b_ in betas:
        for e in erps:
            b = a.copy(); b.set("beta", b_); b.set("erp", e)
            g.loc[b_, e] = model.dcf(b)["price"]
    return g
