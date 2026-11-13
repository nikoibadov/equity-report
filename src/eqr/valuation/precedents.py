"""Precedent transactions: large software deals (public sources) applied to SAP's last-twelve-months revenue,
and the football-field weighting DCF / comparable companies / precedent transactions (team plan: 50 / 25 / 25).

Precedent multiples contain a control premium and are paid for whole companies: they are an upper cross-check on a
minority trading value, not a replacement for it."""
from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
import yaml

PATH = Path(__file__).with_name("loaders") / "precedents.yaml"
WEIGHTS = {"DCF": 0.50, "Comparable companies": 0.25, "Precedent transactions": 0.25}


def load() -> dict:
    return yaml.safe_load(PATH.read_text())


def deals_frame(cfg: dict) -> pd.DataFrame:
    df = pd.DataFrame(cfg["deals"])
    df["ev_rev"] = df["ev_usd_bn"] / df["revenue_usd_bn"]
    cap = cfg["core_growth_max_pct"]
    df["core"] = df["growth_pct"].map(lambda g: bool(pd.notna(g) and g <= cap))
    return df.sort_values("announced", ascending=False).reset_index(drop=True)


def _stats(x: pd.Series) -> dict:
    x = x.dropna()
    return {"n": int(len(x)), "q1": float(np.percentile(x, 25)), "median": float(np.median(x)),
            "q3": float(np.percentile(x, 75)), "mean": float(x.mean()), "min": float(x.min()), "max": float(x.max())}


def comps_value(comps_res: dict) -> tuple[float, list[str]]:
    """Comparable-company value = average of the median-implied prices of EV/NTM sales, EV/NTM EBIT (like-for-like) and P/E."""
    st = comps_res["stats"]
    ebit = "ev_ebit_ntm_xsbc" if "ev_ebit_ntm_xsbc" in st else "ev_ebit_ntm"
    keys = ["ev_sales_ntm", ebit, "pe_ntm"]
    return float(np.mean([st[k]["p_med"] for k in keys])), keys


def run(ds: dict, a, comps_res: dict, dcf_price: float) -> dict:
    cfg = load()
    df = deals_frame(cfg)
    non_op = a["cash"] + a["st_inv"] + a["lt_inv"] + a["debt_benefit"] + a["tlcf_asset"]
    claims = a["debt"] + a["pension"] + a["nci"]
    rev = ds["hist_ltm"]["revenue"]
    stats = {"all": _stats(df["ev_rev"]), "core": _stats(df.loc[df["core"], "ev_rev"])}

    def price(m):
        return (m * rev + non_op - claims) / a["shares"]
    ev_mkt = a["price"] * a["shares"] - non_op + claims
    prices = {"all": {k: price(stats["all"][k]) for k in ("q1", "median", "q3")},
              "core": {k: price(stats["core"][k]) for k in ("q1", "median", "q3")}}
    comps_v, comps_keys = comps_value(comps_res)
    comp = {"DCF": dcf_price, "Comparable companies": comps_v, "Precedent transactions": prices["core"]["median"]}
    weighted = sum(WEIGHTS[k] * v for k, v in comp.items())
    prem = df["premium_pct"].dropna()
    return {"cfg": cfg, "deals": df, "stats": stats, "prices": prices, "rev_ltm": rev, "ev_market": ev_mkt,
            "sap_multiple": ev_mkt / rev, "non_op": non_op, "claims": claims, "weights": dict(WEIGHTS), "components": comp,
            "comps_keys": comps_keys, "weighted": weighted,
            "premium": {"n": int(len(prem)), "median": float(prem.median()) if len(prem) else None,
                        "min": float(prem.min()) if len(prem) else None, "max": float(prem.max()) if len(prem) else None}}
