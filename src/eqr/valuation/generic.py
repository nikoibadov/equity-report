"""Generic FCFF DCF + comps for any listed company (yfinance data), same conventions as the SAP model:
CAPM WACC, mid-year discounting, 5 explicit + 5 fade years, Gordon terminal value, EV-to-equity bridge.

The SAP-specific segment model (model.py) stays the reference for SAP; this module powers the
'any ticker' mode of the Streamlit app."""
from __future__ import annotations

import datetime as dt
import json
import os
import re
from copy import deepcopy
from pathlib import Path

import numpy as np
import pandas as pd
import requests

N_EXPLICIT, N_FADE = 5, 5
TODAY = dt.date.today().isoformat()
CACHE = Path(os.environ.get("VALUATION_CACHE", Path(__file__).resolve().parents[1] / "data" / "web"))
UA = {"User-Agent": "company-valuation research-bot (contact: admin@example.org)"}

INDEX_BY_SUFFIX = {"DE": "^GDAXI", "F": "^GDAXI", "PA": "^FCHI", "L": "^FTSE", "AS": "^AEX", "MC": "^IBEX", "MI": "FTSEMIB.MI",
                   "SW": "^SSMI", "ST": "^OMX", "T": "^N225", "HK": "^HSI", "TO": "^GSPTSE", "AX": "^AXJO", "KS": "^KS11",
                   "CO": "^OMXC25", "HE": "^OMXH25", "OL": "OBX.OL", "BR": "^BFX", "LS": "PSI20.LS", "VI": "^ATX"}
# Damodaran-style synthetic rating spreads (large non-financial firms, Jan 2026 table) by interest coverage
COVERAGE_SPREADS = [(8.5, 0.0040, "AAA"), (6.5, 0.0055, "AA"), (5.5, 0.0070, "A+"), (4.25, 0.0078, "A"), (3.0, 0.0089, "A-"),
                    (2.5, 0.0111, "BBB"), (2.25, 0.0138, "BB+"), (2.0, 0.0184, "BB"), (1.75, 0.0275, "B+"), (1.5, 0.0321, "B"),
                    (1.25, 0.0509, "B-"), (0.8, 0.0885, "CCC"), (0.65, 0.1261, "CC"), (-1e9, 0.1600, "D")]


def _cache(name: str, fn):
    CACHE.mkdir(parents=True, exist_ok=True)
    p = CACHE / f"{re.sub(r'[^A-Za-z0-9_.-]', '_', name)}.json"
    if p.exists():
        return json.loads(p.read_text())
    d = fn()
    p.write_text(json.dumps(d, default=str))
    return d


def _fx(frm: str, to: str) -> float:
    """Units of `to` per 1 `frm`."""
    import yfinance as yf
    if not frm or not to or frm == to:
        return 1.0
    def f():
        return {"r": float(yf.Ticker(f"{frm}{to}=X").history(period="5d")["Close"].iloc[-1])}
    return _cache(f"fx_{frm}{to}_{TODAY}", f)["r"]


def risk_free(currency: str) -> tuple[float, str]:
    try:
        if currency == "EUR":
            url = ("https://api.statistiken.bundesbank.de/rest/data/BBSIS/"
                   "D.I.ZAR.ZI.EUR.S1311.B.A604.R10XX.R.A.A._Z._Z.A?format=csv&lastNObservations=15&lang=en")
            def f():
                txt = requests.get(url, timeout=30).text
                obs = [l.split(",") for l in txt.splitlines() if re.match(r"^\d{4}-\d{2}-\d{2},\d", l)]
                return {"d": obs[-1][0], "y": float(obs[-1][1])}
            r = _cache(f"bund10y_{TODAY}", f)
            return r["y"] / 100, f"Bundesbank 10y Bund {r['d']}"
        if currency == "USD":
            import yfinance as yf
            r = _cache(f"tnx_{TODAY}", lambda: {"y": float(yf.Ticker("^TNX").history(period="5d")["Close"].iloc[-1])})
            return r["y"] / 100, "US 10y Treasury (^TNX)"
    except Exception:  # noqa: BLE001
        pass
    return 0.035, f"Default 3.5% (no automatic 10y source for {currency}; edit)"


# ---------------------------------------------------------------- data
def snapshot(ticker: str) -> dict:
    import yfinance as yf

    def f():
        t = yf.Ticker(ticker)
        i = t.info
        if not i or i.get("quoteType") not in ("EQUITY", None):
            raise ValueError(f"{ticker} is not an equity")
        inc, cf, bs = t.income_stmt, t.cashflow, t.balance_sheet
        def row(df, *names):
            for n in names:
                if n in df.index:
                    return {str(c.date()): (None if pd.isna(v) else float(v)) for c, v in df.loc[n].items()}
            return {}
        try:
            re_ = {k: float(v) for k, v in t.revenue_estimate["avg"].items()}
            ee = {k: float(v) for k, v in t.earnings_estimate["avg"].items()}
        except Exception:  # noqa: BLE001
            re_, ee = {}, {}
        suffix = ticker.split(".")[-1] if "." in ticker else ""
        idx = INDEX_BY_SUFFIX.get(suffix, "^GSPC")
        px = yf.download([ticker, idx], period="5y", interval="1wk", auto_adjust=True, progress=False)["Close"].dropna()
        r = px.pct_change().dropna()
        x, y = r[idx].values, r[ticker].values
        b = float(np.cov(y, x)[0, 1] / np.var(x, ddof=1)) if len(x) > 30 else None
        hist = t.history(period="1y")
        return {
            "ticker": ticker, "info": {k: i.get(k) for k in (
                "longName", "shortName", "currency", "financialCurrency", "sector", "industry", "country", "longBusinessSummary",
                "currentPrice", "regularMarketPrice", "sharesOutstanding", "marketCap", "enterpriseValue", "totalRevenue", "ebitda",
                "operatingMargins", "grossMargins", "revenueGrowth", "totalCash", "totalDebt", "trailingPE", "forwardPE",
                "enterpriseToRevenue", "enterpriseToEbitda", "fiftyTwoWeekHigh", "fiftyTwoWeekLow", "targetMeanPrice",
                "targetMedianPrice", "targetHighPrice", "targetLowPrice", "numberOfAnalystOpinions", "recommendationMean",
                "lastFiscalYearEnd", "nextFiscalYearEnd")},
            "revenue": row(inc, "Total Revenue"), "ebit": row(inc, "Operating Income", "EBIT"),
            "tax_rate": row(inc, "Tax Rate For Calcs"), "interest": row(inc, "Interest Expense"),
            "da": row(cf, "Depreciation And Amortization", "Depreciation Amortization Depletion"),
            "capex": row(cf, "Capital Expenditure"), "dwc": row(cf, "Change In Working Capital"),
            "sbc": row(cf, "Stock Based Compensation"), "nci": row(bs, "Minority Interest"),
            "rev_est": re_, "eps_est": ee, "beta_raw": b, "beta_index": idx, "beta_n": int(len(x)),
            "price_date": str(hist.index[-1].date()) if len(hist) else TODAY,
            "fetched": TODAY,
        }
    return _cache(f"snap_{ticker}_{TODAY}", f)


def _latest(d: dict, n=3):
    vals = [v for k, v in sorted(d.items(), reverse=True) if v is not None][:n]
    return vals


def default_inputs(s: dict) -> dict:
    i = s["info"]
    fin_ccy, q_ccy = i.get("financialCurrency") or i.get("currency"), i.get("currency")
    price = float(i.get("currentPrice") or i.get("regularMarketPrice"))
    if q_ccy == "GBp":
        price, q_ccy = price / 100, "GBP"
    fx = _fx(fin_ccy, q_ccy)                      # quote-currency units per financial-currency unit
    rev_ttm = float(i.get("totalRevenue") or _latest(s["revenue"], 1)[0])
    re_ = s["rev_est"]
    lfy = pd.Timestamp(i["lastFiscalYearEnd"], unit="s") if i.get("lastFiscalYearEnd") else None
    w0 = float(np.clip(((lfy + pd.DateOffset(years=1)) - pd.Timestamp(TODAY)).days / 365, 0, 1)) if lfy is not None else 0.5
    if re_.get("0y") and re_.get("+1y"):
        g1 = (w0 * re_["0y"] + (1 - w0) * re_["+1y"]) / rev_ttm - 1        # NTM vs TTM (same base as rev0)
        g2 = re_["+1y"] / re_["0y"] - 1
    else:
        g1 = i.get("revenueGrowth") or 0.05
        g2 = g1 * 0.8
    g1, g2 = float(np.clip(g1, -0.2, 0.5)), float(np.clip(g2, -0.2, 0.5))
    revs, ebits = _latest(s["revenue"]), _latest(s["ebit"])
    m_ttm = i.get("operatingMargins") or (ebits[0] / revs[0] if revs and ebits else 0.15)
    das, cxs = _latest(s["da"]), _latest(s["capex"])
    da_pct = float(np.mean([d / r for d, r in zip(das, revs)])) if das else 0.04
    cx_pct = float(np.mean([-c / r for c, r in zip(cxs, revs)])) if cxs else 0.04
    taxes = [t for t in _latest(s["tax_rate"]) if t is not None and 0 < t < 0.5]
    tax = float(np.clip(np.mean(taxes), 0.15, 0.30)) if taxes else 0.25
    rf, rf_src = risk_free(fin_ccy)
    beta_raw = s.get("beta_raw") or 1.0
    beta = 0.67 * beta_raw + 0.33
    interest = abs((_latest(s["interest"], 1) or [0])[0] or 0)
    cover = (ebits[0] / interest) if interest and ebits else 20
    spread, rating = next((sp, rt) for th, sp, rt in COVERAGE_SPREADS if cover > th)
    shares = float(i.get("sharesOutstanding"))
    mcap_q = price * shares
    nci = (_latest(s["nci"], 1) or [0])[0] or 0
    path = [g1, g2] + list(np.linspace(g2, (g2 + 0.025) / 2, 4)[1:])
    return {
        "name": i.get("longName") or i.get("shortName") or s["ticker"], "ticker": s["ticker"],
        "fin_ccy": fin_ccy, "q_ccy": q_ccy, "fx": fx, "price": price, "shares": shares,
        "rev0": rev_ttm, "growth": [float(x) for x in path][:N_EXPLICIT],
        "margin0": float(m_ttm), "margin_target": float(m_ttm), "tax": tax, "da_pct": da_pct, "capex_pct": cx_pct,
        "nwc_pct": 0.0, "g": 0.025, "rf": rf, "beta": beta, "erp": 0.05, "kd": rf + spread, "tax_wacc": tax,
        "cash": float(i.get("totalCash") or 0), "debt": float(i.get("totalDebt") or 0), "nci": float(nci),
        "mcap_fin": mcap_q / fx,
        "sources": {
            "rf": rf_src, "beta": f"OLS 5y weekly vs {s['beta_index']} (n={s['beta_n']}), raw {beta_raw:.2f}, Blume-adjusted",
            "erp": "House view 5.0% (IDW/FAUB lower end; Damodaran implied ~4.4-4.6%)",
            "kd": f"Rf + synthetic-rating spread {spread*100:.2f}% (interest cover {cover:.1f}x → {rating})",
            "growth": "Years 1-2 consensus revenue (yfinance), years 3-5 linear towards midpoint of year-2 growth and 2.5%",
            "margin": "TTM operating margin (GAAP, SBC included); target = TTM unless edited",
            "reinvestment": "D&A and capex: average % of revenue over the last 3 fiscal years",
        },
    }


# ---------------------------------------------------------------- model
def wacc(p: dict) -> dict:
    ke = p["rf"] + p["beta"] * p["erp"]
    e, d = p["mcap_fin"], p["debt"]
    we = e / (e + d) if e + d else 1
    return {"ke": ke, "kd_at": p["kd"] * (1 - p["tax_wacc"]), "we": we, "wd": 1 - we,
            "wacc": we * ke + (1 - we) * p["kd"] * (1 - p["tax_wacc"])}


def forecast(p: dict) -> pd.DataFrame:
    rows, rev, prev = {}, p["rev0"], p["rev0"]
    g_last = p["growth"][-1]
    for t in range(1, N_EXPLICIT + N_FADE + 1):
        if t <= N_EXPLICIT:
            g = p["growth"][t - 1]
            m = p["margin0"] + (p["margin_target"] - p["margin0"]) * t / N_EXPLICIT
        else:
            k = t - N_EXPLICIT
            g = g_last + (p["g"] - g_last) * k / N_FADE
            m = p["margin_target"]
        rev = prev * (1 + g)
        ebit = rev * m
        nopat = ebit * (1 - p["tax"])
        da, cx = rev * p["da_pct"], rev * p["capex_pct"]
        dnwc = p["nwc_pct"] * (rev - prev)
        rows[t] = {"revenue": rev, "growth": g, "margin": m, "ebit": ebit, "nopat": nopat, "da": da, "capex": cx,
                   "dnwc": dnwc, "fcff": nopat + da - cx - dnwc}
        prev = rev
    return pd.DataFrame(rows).T


def dcf(p: dict, wacc_override: float | None = None) -> dict:
    w = wacc(p)
    r = w["wacc"] if wacc_override is None else wacc_override
    f = forecast(p)
    t = np.arange(len(f)) + 0.5
    pv = f["fcff"].values / (1 + r) ** t
    last = f.iloc[-1]
    rev_t1 = last["revenue"] * (1 + p["g"])
    fcff_t1 = rev_t1 * p["margin_target"] * (1 - p["tax"]) + rev_t1 * (p["da_pct"] - p["capex_pct"]) - p["nwc_pct"] * (rev_t1 - last["revenue"])
    tv = fcff_t1 / (r - p["g"])
    pv_tv = tv / (1 + r) ** len(f)
    ev = pv.sum() + pv_tv
    equity = ev + p["cash"] - p["debt"] - p["nci"]
    price = equity * p["fx"] / p["shares"]
    return {"wacc": r, **{k: v for k, v in w.items() if k != "wacc"}, "forecast": f, "pv": pv, "tv": tv, "pv_tv": pv_tv,
            "ev": ev, "equity": equity, "price": price, "upside": price / p["price"] - 1, "tv_share": pv_tv / ev,
            "implied_ev_sales_ntm": ev / f.iloc[0]["revenue"], "implied_ev_ebit_ntm": ev / f.iloc[0]["ebit"]}


def scenarios(p: dict) -> dict:
    base_w = wacc(p)["wacc"]
    def mk(dg, dm, dw, dgt):
        q = deepcopy(p)
        q["growth"] = [x + dg for x in p["growth"]]
        q["margin_target"] = p["margin_target"] + dm
        q["g"] = p["g"] + dgt
        return dcf(q, base_w + dw)["price"]
    out = {"Bear": {"price": mk(-0.03, -0.02, 0.005, -0.005), "weight": 0.25},
           "Base": {"price": dcf(p)["price"], "weight": 0.5},
           "Bull": {"price": mk(0.02, 0.015, -0.0025, 0.0025), "weight": 0.25}}
    out["Weighted"] = {"price": sum(v["price"] * v["weight"] for v in out.values())}
    return out


def grid_wacc_g(p: dict, d_w=np.arange(-0.01, 0.0101, 0.0025), d_g=np.arange(-0.0075, 0.00751, 0.0025)) -> pd.DataFrame:
    w0 = wacc(p)["wacc"]
    g = pd.DataFrame(index=[round(w0 + x, 4) for x in d_w], columns=[round(p["g"] + x, 4) for x in d_g], dtype=float)
    for w in g.index:
        for gg in g.columns:
            q = dict(p); q["g"] = gg
            g.loc[w, gg] = dcf(q, w)["price"] if w > gg + 0.005 else np.nan
    return g


# ---------------------------------------------------------------- peers & comps
def peer_row(sym: str) -> dict | None:
    """NTM multiples for one listed company (fiscal-year-weighted consensus, GAAP operating margin)."""
    import yfinance as yf

    def f():
        t = yf.Ticker(sym)
        i = t.info
        re_, ee = t.revenue_estimate, t.earnings_estimate
        lfy = pd.Timestamp(i["lastFiscalYearEnd"], unit="s") if i.get("lastFiscalYearEnd") else None
        fye = (lfy + pd.DateOffset(years=1)) if lfy is not None else None
        w0 = float(np.clip((fye - pd.Timestamp(TODAY)).days / 365, 0, 1)) if fye is not None else 0.5
        rev_ntm = w0 * float(re_.loc["0y", "avg"]) + (1 - w0) * float(re_.loc["+1y", "avg"])
        eps_ntm = w0 * float(ee.loc["0y", "avg"]) + (1 - w0) * float(ee.loc["+1y", "avg"])
        price = float(i.get("currentPrice") or i.get("regularMarketPrice"))
        q_ccy = i.get("currency")
        if q_ccy == "GBp":
            price, q_ccy = price / 100, "GBP"
        fin = i.get("financialCurrency") or q_ccy
        rev_ttm = float(i["totalRevenue"])
        if fin != q_ccy:
            # ADRs / cross listings: Yahoo's EV/revenue mixes quote and reporting currency -> rebuild EV in reporting currency
            ev = (price * float(i["sharesOutstanding"]) * _fx(q_ccy, fin)
                  + float(i.get("totalDebt") or 0) - float(i.get("totalCash") or 0))
        else:
            ev = float(i["enterpriseToRevenue"]) * rev_ttm
        om = i.get("operatingMargins")
        eps_conv = _fx(fin, q_ccy)
        mcap_eur = price * float(i.get("sharesOutstanding") or 0) * _fx(q_ccy, "EUR")
        return {"ticker": sym, "name": i.get("longName") or i.get("shortName"), "country": i.get("country"),
                "industry": i.get("industry"), "mcap_eur": mcap_eur, "ev_sales_ntm": ev / rev_ntm,
                "ev_ebit_ntm": ev / (rev_ntm * om) if om and om > 0 else None,
                "pe_ntm": price / (eps_ntm * eps_conv) if eps_ntm and eps_ntm > 0 else None,
                "g_ntm": (rev_ntm / rev_ttm - 1) * 100, "ebit_margin": (om or 0) * 100,
                "rule40": (rev_ntm / rev_ttm - 1) * 100 + (om or 0) * 100}
    try:
        return _cache(f"peer_{sym}_{TODAY}", f)
    except Exception:  # noqa: BLE001
        return None


def comps(p: dict, s: dict, peers: pd.DataFrame) -> dict:
    re_ = s["rev_est"]
    lfy = pd.Timestamp(s["info"]["lastFiscalYearEnd"], unit="s") if s["info"].get("lastFiscalYearEnd") else None
    w0 = float(np.clip(((lfy + pd.DateOffset(years=1)) - pd.Timestamp(TODAY)).days / 365, 0, 1)) if lfy is not None else 0.5
    rev_ntm = (w0 * re_["0y"] + (1 - w0) * re_["+1y"]) if re_.get("0y") and re_.get("+1y") else p["rev0"] * (1 + p["growth"][0])
    ee = s["eps_est"]
    eps_ntm = (w0 * ee["0y"] + (1 - w0) * ee["+1y"]) * p["fx"] if ee.get("0y") and ee.get("+1y") else None
    ebit_ntm = rev_ntm * p["margin0"]
    bridge = p["cash"] - p["debt"] - p["nci"]
    to_px = lambda ev: (ev + bridge) * p["fx"] / p["shares"]  # noqa: E731
    out = {"stats": {}, "sap": {"rev_ntm": rev_ntm, "ebit_ntm": ebit_ntm, "eps_ntm": eps_ntm}}
    for key, label, metric in (("ev_sales_ntm", "EV / NTM sales", rev_ntm), ("ev_ebit_ntm", "EV / NTM EBIT (GAAP)", ebit_ntm),
                               ("pe_ntm", "P / NTM EPS", eps_ntm)):
        x = peers[key].dropna() if key in peers else pd.Series(dtype=float)
        x = x[x > 0]
        if len(x) < 2 or metric is None:
            continue
        q1, md, q3 = np.percentile(x, 25), np.median(x), np.percentile(x, 75)
        conv = (lambda m: m * metric) if key == "pe_ntm" else (lambda m: to_px(m * metric))
        out["stats"][key] = {"label": label, "n": len(x), "q1": q1, "median": md, "q3": q3,
                             "p_q1": conv(q1), "p_med": conv(md), "p_q3": conv(q3)}
    reg = peers[["ev_sales_ntm", "rule40"]].dropna() if len(peers) else pd.DataFrame()
    if len(reg) >= 4:
        X = np.column_stack([np.ones(len(reg)), reg["rule40"]])
        beta, *_ = np.linalg.lstsq(X, reg["ev_sales_ntm"].values, rcond=None)
        resid = reg["ev_sales_ntm"].values - X @ beta
        se = float(np.sqrt(resid @ resid / max(len(reg) - 2, 1)))
        r40 = (rev_ntm / p["rev0"] - 1) * 100 + p["margin0"] * 100
        fit = beta[0] + beta[1] * r40
        out["regression"] = {"slope": beta[1], "intercept": beta[0], "se": se, "fit": fit, "n": len(reg),
                             "p_fit": to_px(fit * rev_ntm), "p_low": to_px((fit - se) * rev_ntm), "p_high": to_px((fit + se) * rev_ntm)}
    return out


def run(ticker: str, overrides: dict | None = None, peer_tickers: list[str] | None = None) -> dict:
    s = snapshot(ticker)
    p = default_inputs(s)
    if overrides:
        p.update(overrides)
    d = dcf(p)
    sc = scenarios(p)
    w0 = d["wacc"]
    band = [dcf({**p, "g": p["g"] + dg}, w0 + dw)["price"] for dw in (-0.005, 0, 0.005) for dg in (-0.0025, 0, 0.0025)]
    peers = pd.DataFrame([r for r in (peer_row(t) for t in (peer_tickers or [])) if r])
    cp = comps(p, s, peers)
    i = s["info"]
    to_q = (1 / 100) if i.get("currency") == "GBp" else 1
    market = {"price": p["price"], "price_date": s["price_date"],
              "hi52": (i.get("fiftyTwoWeekHigh") or p["price"]) * to_q, "lo52": (i.get("fiftyTwoWeekLow") or p["price"]) * to_q,
              "tgt_mean": i.get("targetMeanPrice"), "tgt_low": i.get("targetLowPrice"), "tgt_high": i.get("targetHighPrice"),
              "tgt_n": i.get("numberOfAnalystOpinions") or 0}
    return {"inputs": p, "snapshot": s, "dcf": d, "scenarios": sc, "dcf_band": (min(band), max(band)),
            "grid_wacc_g": grid_wacc_g(p), "comps": cp, "peers": peers, "market": market}


def football_items(res: dict) -> list[dict]:
    mk, d, sc, cp = res["market"], res["dcf"], res["scenarios"], res["comps"]
    items = [{"label": "52-week range", "low": mk["lo52"], "high": mk["hi52"], "mid": None, "kind": "market"}]
    if mk.get("tgt_low") and mk.get("tgt_high"):
        items.append({"label": f"Analyst targets (n={mk['tgt_n']})", "low": mk["tgt_low"], "high": mk["tgt_high"], "mid": mk["tgt_mean"], "kind": "market"})
    items += [{"label": "DCF: WACC ±0.5pp, g ±0.25pp", "low": res["dcf_band"][0], "high": res["dcf_band"][1], "mid": d["price"], "kind": "dcf"},
              {"label": "DCF: bear / base / bull", "low": sc["Bear"]["price"], "high": sc["Bull"]["price"], "mid": sc["Weighted"]["price"], "kind": "dcf"}]
    for k, st in cp["stats"].items():
        items.append({"label": f"{st['label']} (peer IQR)", "low": st["p_q1"], "high": st["p_q3"], "mid": st["p_med"], "kind": "comps"})
    if "regression" in cp:
        r = cp["regression"]
        items.append({"label": "Growth-adj. EV/Sales (Rule of 40, ±1 SE)", "low": r["p_low"], "high": r["p_high"], "mid": r["p_fit"], "kind": "comps"})
    return items
