"""Lecture-aligned extensions (MEAC 'Valuation Methodologies'): sum-of-the-parts, bottom-up beta,
alternative terminal values, WACC-weight iteration and multiple caveats. All are cross-checks around
the base DCF; none of them changes the base case."""
from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd

from . import model

BETA_SNAPSHOT = Path(__file__).parent / "loaders" / "peer_betas.json"

# ---------------------------------------------------------------- 1. sum of the parts
SOTP_GROUPS = {
    "cloud": ("Cloud", ["CRM", "NOW", "WDAY", "ADBE", "ADSK", "INTU"], "Cloud-native application software (subscription, double-digit growth)"),
    "support": ("Software support + licences", ["OTEX", "SGE", "DSY", "PTC"], "Mature software with a large maintenance base (high margin, low growth)"),
    "services": ("Services", ["ACN", "CAP", "SOP"], "IT services (labour-based, low margin)"),
}


def sotp(res_dcf: dict, a, peers: pd.DataFrame) -> dict:
    f = res_dcf["forecast"]
    ntm = {s: 0.25 * f.loc[2026, s] + 0.75 * f.loc[2027, s] for s in ("cloud", "licenses", "support", "services")}
    rev = {"cloud": ntm["cloud"], "support": ntm["support"] + ntm["licenses"], "services": ntm["services"]}
    px = peers.set_index("ticker")
    parts, tot = [], {"q1": 0.0, "median": 0.0, "q3": 0.0}
    for k, (label, tickers, why) in SOTP_GROUPS.items():
        x = px.reindex(tickers)["ev_sales_ntm"].dropna()
        x = x[x > 0]
        q1, med, q3 = (float(np.percentile(x, q)) for q in (25, 50, 75))
        parts.append({"part": label, "rev_ntm": rev[k], "peers": ", ".join(x.index), "why": why, "n": len(x),
                      "q1": q1, "median": med, "q3": q3, "ev_q1": q1 * rev[k], "ev_med": med * rev[k], "ev_q3": q3 * rev[k]})
        for q, m in (("q1", q1), ("median", med), ("q3", q3)):
            tot[q] += m * rev[k]
    non_op = a["cash"] + a["st_inv"] + a["lt_inv"] + a["debt_benefit"] + a["tlcf_asset"]
    claims = a["debt"] + a["pension"] + a["nci"]
    price = {q: (v + non_op - claims) / a["shares"] for q, v in tot.items()}
    total_rev = sum(rev.values())
    # Mix effect: the same three multiples applied to the FY2025, NTM and 2030E revenue mix
    med = {p["part"]: p["median"] for p in parts}
    labels = [SOTP_GROUPS[k][0] for k in ("cloud", "support", "services")]
    def blended(c, sup, svc):
        t = c + sup + svc
        return (med[labels[0]] * c + med[labels[1]] * sup + med[labels[2]] * svc) / t, c / t
    mix = []
    for name, c, sup, svc in (
            ("FY2025A", a["base_cloud"], a["base_support"] + a["base_licenses"], a["base_services"]),
            ("NTM", rev["cloud"], rev["support"], rev["services"]),
            ("FY2030E", f.loc[2030, "cloud"], f.loc[2030, "support"] + f.loc[2030, "licenses"], f.loc[2030, "services"])):
        m, share = blended(c, sup, svc)
        mix.append({"period": name, "cloud_share": share, "blended_ev_sales": m})
    core = peers[peers["core"]]["ev_sales_ntm"].dropna()
    # What the market and the DCF pay for the cloud line if support and services are worth their peer medians
    rest = parts[1]["ev_med"] + parts[2]["ev_med"]
    ev_mkt = a["price"] * a["shares"] - non_op + claims
    implied_cloud = {"market": (ev_mkt - rest) / rev["cloud"], "dcf": (res_dcf["ev"] - rest) / rev["cloud"]}
    g_cloud_ntm = rev["cloud"] / (0.25 * a["base_cloud"] + 0.75 * f.loc[2026, "cloud"]) - 1
    cl = px.reindex(SOTP_GROUPS["cloud"][1])["g_ntm"].dropna()
    return {"parts": parts, "ev": tot, "price": price, "implied_cloud": implied_cloud, "ev_market": ev_mkt,
            "sap_cloud_growth": g_cloud_ntm, "peer_cloud_growth": float(np.median(cl)) / 100 if len(cl) else None, "blended": tot["median"] / total_rev,
            "peer_blended_median": float(np.median(core)), "cloud_share_of_ev": parts[0]["ev_med"] / tot["median"],
            "cloud_share_of_rev": rev["cloud"] / total_rev, "mix": mix}


# ---------------------------------------------------------------- 2. bottom-up beta
def bottom_up_beta(a, peers: pd.DataFrame, t_peer: float = 0.25) -> dict | None:
    if not BETA_SNAPSHOT.exists():
        return None
    snap = json.loads(BETA_SNAPSHOT.read_text())
    rows = []
    for tk in peers[peers["core"]]["ticker"]:
        r = snap["rows"].get(tk)
        if not r or r.get("debt_to_equity") is None:
            continue
        de = r["debt_to_equity"]
        rows.append({"ticker": tk, "index": r["index"], "beta_raw": r["raw"], "beta_adj": r["adjusted"], "d_e": de,
                     "beta_u": r["adjusted"] / (1 + (1 - t_peer) * de)})
    df = pd.DataFrame(rows)
    de_sap = a["d_mv"] / a["e_mv"]
    bu_med, bu_mean = float(df["beta_u"].median()), float(df["beta_u"].mean())
    relever = lambda bu: bu * (1 + (1 - a["tax_wacc"]) * de_sap)  # noqa: E731
    b = a.copy(); b.set("beta", relever(bu_med))
    d = model.dcf(b)
    sap = snap["rows"].get("SAP", {})
    return {"table": df, "t_peer": t_peer, "beta_u_median": bu_med, "beta_u_mean": bu_mean, "d_e_sap": de_sap,
            "beta_relevered": relever(bu_med), "beta_relevered_mean": relever(bu_mean), "beta_base": a["beta"],
            "ke": d["ke"], "wacc": d["wacc"], "price": d["price"], "fetched": snap["fetched"], "method": snap["method"],
            "sap_vs_stoxx_adj": sap.get("adjusted"), "sap_vs_stoxx_raw": sap.get("raw")}


# ---------------------------------------------------------------- 3. terminal-value methods
def terminal_values(a, base: dict, peers: pd.DataFrame, ebit_basis: str) -> dict:
    f = base["forecast"]
    last = f.loc[f.index[-1]]
    w = base["wacc"]
    # (1) zero-growth perpetuity: PV_T = FCFF_T / WACC
    zero = {dw: model.dcf(a, w + dw, tv_override=last["fcff"] / (w + dw)) for dw in (-0.005, 0.0, 0.005)}
    # (3) exit multiple: PV_T = EBITDA x EV/EBITDA; NTM multiple -> EBITDA of T+1
    core = peers[peers["core"]]["ev_ebitda_ntm"].dropna()
    core = core[core > 0]
    q1, med, q3 = (float(np.percentile(core, q)) for q in (25, 50, 75))
    sbc = last["revenue"] * a["sbc_pct"] if ebit_basis == "street" else 0.0     # street EBITDA of US peers excludes SBC
    ebitda_t1 = (last["op_nonifrs"] + last["da"] + sbc) * (1 + a["g"])
    ex = {k: model.dcf(a, tv_override=m * ebitda_t1) for k, m in (("q1", q1), ("median", med), ("q3", q3))}
    implied_g = lambda tv: w - base["fcff_t1"] / tv  # noqa: E731  (growth that a Gordon TV of the same size would need)
    gordon_mult = base["tv"] / ebitda_t1
    return {
        "gordon": {"price": base["price"], "tv": base["tv"], "tv_share": base["tv_share"], "ev_ebitda": gordon_mult},
        "zero": {"price": zero[0.0]["price"], "low": zero[0.005]["price"], "high": zero[-0.005]["price"],
                 "tv": zero[0.0]["tv"], "tv_share": zero[0.0]["tv_share"]},
        "exit": {"price": ex["median"]["price"], "low": ex["q1"]["price"], "high": ex["q3"]["price"], "tv": ex["median"]["tv"],
                 "tv_share": ex["median"]["tv_share"], "mult_q1": q1, "mult_med": med, "mult_q3": q3, "n": len(core),
                 "ebitda_t1": ebitda_t1, "sbc_added": sbc, "implied_g": implied_g(ex["median"]["tv"])},
    }


# ---------------------------------------------------------------- 4. WACC weight circularity
def wacc_iteration(a, tol: float = 1e-7, max_iter: int = 25) -> dict:
    """Guess E (market cap), run the DCF, replace E by the DCF equity value, repeat until the weights stop moving."""
    b = a.copy()
    steps = []
    for i in range(max_iter):
        d = model.dcf(b)
        e_dcf = d["price"] * a["shares"]
        steps.append({"iter": i, "e_used": b["e_mv"], "we": d["we"], "wacc": d["wacc"], "price": d["price"], "e_dcf": e_dcf})
        if abs(e_dcf / b["e_mv"] - 1) < tol:
            break
        b.set("e_mv", e_dcf)
    return {"steps": pd.DataFrame(steps), "wacc": steps[-1]["wacc"], "price": steps[-1]["price"],
            "wacc_start": steps[0]["wacc"], "price_start": steps[0]["price"], "we": steps[-1]["we"]}


# ---------------------------------------------------------------- 6. bridge refinements
def bridge_refinements(a, base: dict, book_equity: float) -> dict:
    pb = a["e_mv"] / book_equity          # SAP's own price-to-book as a proxy for the market value of minorities
    per = lambda x: x / a["shares"] * base["roll"]  # noqa: E731
    return {
        "debt_benefit": a["debt_benefit"], "debt_benefit_ps": per(a["debt_benefit"]),
        "tlcf": a["tlcf_asset"], "tlcf_ps": per(a["tlcf_asset"]),
        "nci_book": a["nci"], "pb": pb, "nci_market": a["nci"] * pb, "nci_ps_at_market": -per(a["nci"] * pb - a["nci"]),
    }


# ---------------------------------------------------------------- 7. P/E and leverage
def pe_leverage(peers: pd.DataFrame, eps_ntm: float, threshold: float = 0.25) -> dict:
    core = peers[peers["core"]].copy()
    core["net_debt_mcap"] = (core["tev"] - core["mcap"]) / core["mcap"]
    x = core.dropna(subset=["pe_ntm"])
    x = x[x["pe_ntm"] > 0]
    flagged = x[x["net_debt_mcap"] > threshold]
    kept = x[x["net_debt_mcap"] <= threshold]["pe_ntm"]
    q = lambda s, p: float(np.percentile(s, p))  # noqa: E731
    return {"threshold": threshold, "flagged": flagged[["ticker", "net_debt_mcap", "pe_ntm", "tev", "mcap"]].to_dict("records"),
            "median_all": float(np.median(x["pe_ntm"])), "n_all": len(x), "n": len(kept),
            "q1": q(kept, 25), "median": q(kept, 50), "q3": q(kept, 75),
            "p_q1": q(kept, 25) * eps_ntm, "p_med": q(kept, 50) * eps_ntm, "p_q3": q(kept, 75) * eps_ntm,
            "table": core[["ticker", "net_debt_mcap", "pe_ntm"]].sort_values("net_debt_mcap", ascending=False)}


MULTIPLE_TABLE = [
    ("EV / EBIT", "Used (lead)", "After D&A, so it reflects the cost of the asset base; capital-structure neutral",
     "Sensitive to accounting: SAP's non-IFRS EBIT includes share-based pay, US peers' street EBIT excludes it → like-for-like ex-SBC variant added"),
    ("EV / Sales", "Used (lead)", "Works for every peer, not affected by SBC, R&D capitalisation or amortisation policy; capital-structure neutral",
     "Ignores profitability and assumes the same growth → corrected with the Rule-of-40 regression and the sum-of-the-parts"),
    ("P / E", "Used with caveat", "Easy to understand, widely quoted",
     "Only comparable at similar leverage: SAP is net cash, Oracle is heavily indebted → levered peers flagged and removed; also affected by tax and accounting policy"),
    ("EV / EBITDA", "Reference only", "Quick, close to cash flow for capital-intensive firms",
     "Adds nothing for asset-light software (D&A ~3% of revenue); ignores capex and tax; peers' EBITDA excludes SBC. Used only as the exit multiple in the terminal-value cross-check"),
    ("P / B", "Excluded", "Useful for banks and asset-heavy firms",
     "Book equity is mostly goodwill; internally generated intangibles (R&D, customer base) are off balance sheet"),
]

NOT_MODELLED = [
    ("Precedent transactions / premiums paid (p47-49)", "Not meaningful for a ~€250bn company: no buyer could pay a control premium for SAP, and the Capital IQ deal screen returned no usable software multiples of this size. Trading multiples are the relevant market evidence."),
    ("APV instead of WACC (p15-16)", "WACC assumes a stable capital structure. SAP's debt is ~4% of capital and stable, so WACC is appropriate; APV matters when leverage changes over time, as in a leveraged buyout."),
    ("Firm size premium (p19)", "Zero. SAP is Europe's most valuable listed company with a liquid free float; a size or illiquidity premium applies to small and private firms."),
    ("Independence (p3)", "'Bankers' objective is to get the deal done, not to find the right price.' This valuation is not a pitch: assumptions are documented with sources, and bear, base and bull cases are shown."),
]


def run(res: dict) -> dict:
    a, base, peers, ds = res["assumptions"], res["dcf"], res["peers"], res["ds"]
    return {
        "sotp": sotp(base, a, peers),
        "beta": bottom_up_beta(a, peers),
        "tv": terminal_values(a, base, peers, ds.get("peer_ebit_basis", "street")),
        "wacc_iter": wacc_iteration(a),
        "bridge": bridge_refinements(a, base, ds["ir"]["q3_2026_statement"]["balance_sheet_30sep2026"]["equity_parent"]),
        "pe": pe_leverage(peers, res["comps"]["sap"]["eps_ntm"]),
        "multiples": MULTIPLE_TABLE, "not_modelled": NOT_MODELLED,
    }
