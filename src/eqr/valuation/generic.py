"""Generic FCFF DCF + comps for any listed company (Yahoo Finance data), built on the same conventions as the SAP model:

- base year = last reported fiscal year; forecast = 5 explicit + 5 fade years, fading linearly to the terminal values
- value anchored at the latest quarterly balance sheet on or before the valuation date, current fiscal year valued
  only for the part still to come (stub), mid-period discounting, rolled forward to the valuation date at Ke
- CAPM WACC with market-value weights; tax: effective rate in the explicit years, fading to the statutory (marginal)
  rate, which is also used for the debt tax shield
- Gordon terminal value on a normalised year T+1, EV-to-equity bridge with every non-operating asset and claim found
- live Bear/Base/Bull scenario shifts, sensitivities, tornado, Monte Carlo, DDM cross-check, peer comps

The SAP-specific segment model (model.py) stays the reference for SAP; this module powers the 'any ticker' mode.
`generic_book.py` writes the formula-driven workbook that re-implements exactly these formulas."""
from __future__ import annotations

import datetime as dt
import os
import re
from pathlib import Path

import numpy as np
import pandas as pd
import requests

N_EXPLICIT, N_FADE = 5, 5
N_YEARS = N_EXPLICIT + N_FADE
TODAY = dt.date.today().isoformat()
CACHE = Path(os.environ.get("VALUATION_CACHE", Path(__file__).resolve().parents[1] / "data" / "web"))
UA = {"User-Agent": "company-valuation research-bot (contact: admin@example.org)"}

INDEX_BY_SUFFIX = {"DE": "^GDAXI", "F": "^GDAXI", "PA": "^FCHI", "L": "^FTSE", "AS": "^AEX", "MC": "^IBEX", "MI": "FTSEMIB.MI",
                   "SW": "^SSMI", "ST": "^OMX", "T": "^N225", "HK": "^HSI", "TO": "^GSPTSE", "AX": "^AXJO", "KS": "^KS11",
                   "CO": "^OMXC25", "HE": "^OMXH25", "OL": "OBX.OL", "BR": "^BFX", "LS": "PSI20.LS", "VI": "^ATX"}
EXCHANGE_BY_SUFFIX = {"": "US (NYSE / Nasdaq)", "T": "Tokyo", "KS": "Korea (KRX)", "KQ": "KOSDAQ", "TO": "Toronto", "V": "TSX Venture",
                      "DE": "Xetra", "F": "Frankfurt", "PA": "Euronext Paris", "L": "London", "AS": "Euronext Amsterdam",
                      "SW": "SIX Swiss", "HK": "Hong Kong", "AX": "ASX", "MI": "Borsa Italiana", "MC": "Madrid", "ST": "Stockholm",
                      "CO": "Copenhagen", "HE": "Helsinki", "OL": "Oslo", "BR": "Euronext Brussels", "LS": "Euronext Lisbon",
                      "VI": "Vienna", "NS": "NSE India", "BO": "BSE India", "TW": "Taiwan", "TWO": "Taipei Exchange",
                      "SS": "Shanghai", "SZ": "Shenzhen", "SI": "Singapore", "JO": "Johannesburg", "SA": "B3 Brazil",
                      "MX": "Mexico", "IR": "Euronext Dublin", "WA": "Warsaw", "TA": "Tel Aviv", "NZ": "New Zealand"}
# Damodaran-style synthetic rating spreads (large non-financial firms, Jan 2026 table) by interest coverage
COVERAGE_SPREADS = [(8.5, 0.0040, "AAA"), (6.5, 0.0055, "AA"), (5.5, 0.0070, "A+"), (4.25, 0.0078, "A"), (3.0, 0.0089, "A-"),
                    (2.5, 0.0111, "BBB"), (2.25, 0.0138, "BB+"), (2.0, 0.0184, "BB"), (1.75, 0.0275, "B+"), (1.5, 0.0321, "B"),
                    (1.25, 0.0509, "B-"), (0.8, 0.0885, "CCC"), (0.65, 0.1261, "CC"), (-1e9, 0.1600, "D")]
# Combined statutory (marginal) corporate tax rates, 2026 (KPMG / Tax Foundation / OECD tables, rounded)
STATUTORY_TAX = {
    "United States": (0.25, "US federal 21% + average state tax ~4% (Tax Foundation combined rate)"),
    "Germany": (0.2993, "Germany 29.93%: 15% corporate tax + 5.5% solidarity surcharge + 14.105% trade tax (403% multiplier)"),
    "France": (0.258, "France 25.8% incl. social contribution"), "United Kingdom": (0.25, "UK main rate 25%"),
    "Netherlands": (0.258, "Netherlands 25.8%"), "Switzerland": (0.196, "Switzerland ~19.6% (federal + Zurich cantonal)"),
    "Japan": (0.306, "Japan ~30.6% effective statutory (national + local)"), "South Korea": (0.264, "Korea 24% + 10% local surtax"),
    "China": (0.25, "China 25%"), "Taiwan": (0.20, "Taiwan 20%"), "India": (0.252, "India 25.2% (new regime incl. surcharge)"),
    "Canada": (0.265, "Canada ~26.5% (federal + Ontario)"), "Australia": (0.30, "Australia 30%"), "Ireland": (0.15, "Ireland 15% (Pillar Two minimum for large groups)"),
    "Spain": (0.25, "Spain 25%"), "Italy": (0.279, "Italy 24% IRES + 3.9% IRAP"), "Sweden": (0.206, "Sweden 20.6%"),
    "Denmark": (0.22, "Denmark 22%"), "Finland": (0.20, "Finland 20%"), "Norway": (0.22, "Norway 22%"),
    "Belgium": (0.25, "Belgium 25%"), "Austria": (0.23, "Austria 23%"), "Hong Kong": (0.165, "Hong Kong 16.5%"),
    "Singapore": (0.17, "Singapore 17%"), "Brazil": (0.34, "Brazil 34% (IRPJ + CSLL)"), "Mexico": (0.30, "Mexico 30%"),
    "Israel": (0.23, "Israel 23%"), "Luxembourg": (0.249, "Luxembourg ~24.9%"), "Portugal": (0.315, "Portugal ~31.5% incl. surcharges"),
    "Poland": (0.19, "Poland 19%"), "South Africa": (0.27, "South Africa 27%"), "New Zealand": (0.28, "New Zealand 28%"),
}


def _cache(name: str, fn):
    """Daily cache; falls back to the most recent earlier download when today's fetch fails."""
    from .loaders import public_refs
    return public_refs.cached(name, fn, CACHE)


def _fx(frm: str, to: str) -> float:
    """Units of `to` per 1 `frm`."""
    import yfinance as yf
    if not frm or not to or frm == to:
        return 1.0
    from .loaders import public_refs
    a, b = public_refs.eur_rate(frm), public_refs.eur_rate(to)      # ECB reference rates first
    if a and b:
        return b / a
    def f():
        return {"r": float(yf.Ticker(f"{frm}{to}=X").history(period="1mo")["Close"].dropna().iloc[-1])}
    return _cache(f"fx2_{frm}{to}_{TODAY}", f)["r"]


def check_date(vdate: str | None) -> str:
    """Valuation date for any ticker: not in the future, at most five years back (the beta window needs history)."""
    d = dt.date.fromisoformat(str(vdate)[:10]) if vdate else dt.date.fromisoformat(TODAY)
    if d > dt.date.fromisoformat(TODAY):
        raise ValueError(f"Valuation date {d} is in the future")
    if d < dt.date.fromisoformat(TODAY) - dt.timedelta(days=5 * 365):
        raise ValueError(f"Valuation date {d} is more than five years back")
    return d.isoformat()


def _vkey(vdate: str | None) -> str:
    return "" if not vdate or vdate == TODAY else f"_{vdate}"


def exchange_of(sym: str) -> tuple[str, str]:
    """(local code, exchange) from a Yahoo symbol: '6758.T' -> ('6758', 'Tokyo')."""
    code, _, suf = str(sym).partition(".")
    return code, EXCHANGE_BY_SUFFIX.get(suf, suf or EXCHANGE_BY_SUFFIX[""])


def risk_free(currency: str, vdate: str | None = None) -> tuple[float, str]:
    vd = dt.date.fromisoformat(vdate or TODAY)
    try:
        if currency == "EUR":
            url = ("https://api.statistiken.bundesbank.de/rest/data/BBSIS/"
                   f"D.I.ZAR.ZI.EUR.S1311.B.A604.R10XX.R.A.A._Z._Z.A?format=csv&startPeriod={(vd - dt.timedelta(days=21)).isoformat()}"
                   f"&endPeriod={vd.isoformat()}&lang=en")
            def f():
                txt = requests.get(url, timeout=30).text
                obs = [l.split(",") for l in txt.splitlines() if re.match(r"^\d{4}-\d{2}-\d{2},\d", l)]
                return {"d": obs[-1][0], "y": float(obs[-1][1])}
            r = _cache(f"bund10y2_{TODAY}{_vkey(vdate)}", f)
            return r["y"] / 100, f"Bundesbank 10y Bund {r['d']}"
        if currency == "USD":
            import yfinance as yf
            r = _cache(f"tnx3_{TODAY}{_vkey(vdate)}", lambda: {"y": float(yf.Ticker("^TNX").history(
                start=(vd - dt.timedelta(days=21)).isoformat(), end=(vd + dt.timedelta(days=1)).isoformat())["Close"].dropna().iloc[-1])})
            return r["y"] / 100, "US 10y Treasury (^TNX)"
    except Exception:  # noqa: BLE001
        pass
    return 0.035, f"Default 3.5% (no automatic 10y source for {currency}; edit)"


def default_erp(currency: str | None) -> tuple[float, str]:
    """Equity risk premium for the market the cash flows are in. US dollar: Damodaran's implied premium for the
    S&P 500 (what US investors price in today). Euro and others: the 5.0% house view of the SAP model."""
    if currency == "USD":
        try:
            from .loaders import web
            d = web.damodaran()
            return float(d["erp"]), (f"Damodaran implied US equity risk premium of {d['erp_date']} (S&P 500, market-implied). Used for US-dollar "
                                     "companies; the 5.0% house view is calibrated to the German/euro market")
        except Exception:  # noqa: BLE001
            return 0.045, "4.5% placeholder for the US market (Damodaran's implied premium could not be fetched; typical range 4.2-4.8%); edit"
    if currency == "EUR":
        return 0.05, "House view 5.0% for the euro market (lower end of the German IDW/FAUB range of 5.0-6.5%), as in the SAP model"
    return 0.05, f"5.0% mature-market house view (no market-specific source for {currency}); edit"


# ---------------------------------------------------------------- data
_ANNUAL = {
    "revenue": ("inc", ["Total Revenue"]), "gross_profit": ("inc", ["Gross Profit"]), "rnd": ("inc", ["Research And Development"]),
    "ebit": ("inc", ["Operating Income", "EBIT"]), "pretax": ("inc", ["Pretax Income"]), "tax_provision": ("inc", ["Tax Provision"]),
    "net_income": ("inc", ["Net Income Common Stockholders", "Net Income"]), "eps_dil": ("inc", ["Diluted EPS"]),
    "shares_dil": ("inc", ["Diluted Average Shares"]), "tax_rate": ("inc", ["Tax Rate For Calcs"]),
    "interest": ("inc", ["Interest Expense", "Interest Expense Non Operating"]),
    "da": ("cf", ["Depreciation And Amortization", "Depreciation Amortization Depletion"]), "capex": ("cf", ["Capital Expenditure"]),
    "dwc": ("cf", ["Change In Working Capital"]), "sbc": ("cf", ["Stock Based Compensation"]), "ocf": ("cf", ["Operating Cash Flow"]),
    "fcf": ("cf", ["Free Cash Flow"]), "dividends": ("cf", ["Cash Dividends Paid", "Common Stock Dividend Paid"]),
    "buybacks": ("cf", ["Repurchase Of Capital Stock", "Common Stock Payments"]), "cash_taxes": ("cf", ["Income Tax Paid Supplemental Data"]),
    "cash_sti": ("bs", ["Cash Cash Equivalents And Short Term Investments", "Cash And Cash Equivalents"]),
    "lt_inv": ("bs", ["Investmentin Financial Assets", "Available For Sale Securities", "Investments And Advances", "Long Term Equity Investment"]),
    "debt": ("bs", ["Total Debt"]), "equity": ("bs", ["Stockholders Equity", "Common Stock Equity"]), "nci": ("bs", ["Minority Interest"]),
    "pref": ("bs", ["Preferred Stock Equity", "Preferred Stock"]),
    "receivables": ("bs", ["Accounts Receivable", "Receivables"]), "inventory": ("bs", ["Inventory"]),
    "payables": ("bs", ["Accounts Payable", "Payables"]),
    "deferred_rev": ("bs", ["Current Deferred Revenue"]), "ppe": ("bs", ["Net PPE"]),
}
_QUARTER = ("cash_sti", "lt_inv", "debt", "nci", "pref")


def snapshot(ticker: str, vdate: str | None = None) -> dict:
    import yfinance as yf
    vd = dt.date.fromisoformat(vdate or TODAY)

    def f():
        t = yf.Ticker(ticker)
        i = t.info
        if not i or i.get("quoteType") not in ("EQUITY", None):
            raise ValueError(f"{ticker} is not an equity")
        frames = {"inc": t.income_stmt, "cf": t.cashflow, "bs": t.balance_sheet}
        qbs = t.quarterly_balance_sheet
        try:                                      # trailing twelve months from the last four reported quarters
            qi = t.quarterly_income_stmt
            qq = qi.loc[["Total Revenue", "Operating Income"]].dropna(axis=1).iloc[:, :4]
            ttm = ({"rev": float(qq.loc["Total Revenue"].sum()), "ebit": float(qq.loc["Operating Income"].sum()),
                    "end": str(qq.columns[0].date())} if qq.shape[1] == 4 and vd.isoformat() >= TODAY else {})
        except Exception:  # noqa: BLE001
            ttm = {}

        def row(df, names):
            for n in names:
                if df is not None and n in df.index:
                    out = {str(c.date()): (None if pd.isna(v) else float(v)) for c, v in df.loc[n].items()}
                    if any(v is not None for v in out.values()):
                        return out
            return {}
        try:
            re_ = {k: float(v) for k, v in t.revenue_estimate["avg"].items()}
            ee = {k: float(v) for k, v in t.earnings_estimate["avg"].items()}
        except Exception:  # noqa: BLE001
            re_, ee = {}, {}
        suffix = ticker.split(".")[-1] if "." in ticker else ""
        idx = INDEX_BY_SUFFIX.get(suffix, "^GSPC")
        px = yf.download([ticker, idx], start=(vd - dt.timedelta(days=5 * 365 + 7)).isoformat(), end=(vd + dt.timedelta(days=1)).isoformat(),
                         interval="1wk", auto_adjust=True, progress=False)["Close"].dropna()
        r = px.pct_change().dropna()
        x, y = r[idx].values, r[ticker].values
        b = float(np.cov(y, x)[0, 1] / np.var(x, ddof=1)) if len(x) > 30 else None
        r2 = float(np.corrcoef(x, y)[0, 1] ** 2) if len(x) > 30 else None
        hist = t.history(start=(vd - dt.timedelta(days=365)).isoformat(), end=(vd + dt.timedelta(days=1)).isoformat())["Close"].dropna()
        out = {
            "ticker": ticker, "info": {k: i.get(k) for k in (
                "longName", "shortName", "currency", "financialCurrency", "sector", "industry", "country", "longBusinessSummary",
                "currentPrice", "regularMarketPrice", "sharesOutstanding", "impliedSharesOutstanding", "marketCap", "enterpriseValue",
                "totalRevenue", "ebitda", "operatingMargins", "grossMargins", "revenueGrowth", "totalCash", "totalDebt", "trailingPE",
                "forwardPE", "enterpriseToRevenue", "enterpriseToEbitda", "fiftyTwoWeekHigh", "fiftyTwoWeekLow", "targetMeanPrice",
                "targetMedianPrice", "targetHighPrice", "targetLowPrice", "numberOfAnalystOpinions", "recommendationMean",
                "lastFiscalYearEnd", "nextFiscalYearEnd", "mostRecentQuarter", "dividendRate", "trailingAnnualDividendRate",
                "payoutRatio", "exchange", "fullExchangeName")},
            "q": {k: row(qbs, _ANNUAL[k][1]) for k in _QUARTER},
            "ttm": ttm, "rev_est": re_, "eps_est": ee, "beta_raw": b, "beta_r2": r2, "beta_index": idx, "beta_n": int(len(x)),
            "price_date": str(hist.index[-1].date()) if len(hist) else TODAY,
            "vdate": vd.isoformat(),
            # as of an earlier valuation date: close, 52-week range from the history up to that date (else live quote values)
            "price_v": float(hist.iloc[-1]) if len(hist) and vd.isoformat() < TODAY else None,
            "hi_v": float(hist.max()) if len(hist) and vd.isoformat() < TODAY else None,
            "lo_v": float(hist.min()) if len(hist) and vd.isoformat() < TODAY else None,
            "fetched": TODAY,
        }
        for k, (frame, names) in _ANNUAL.items():
            out[k] = row(frames[frame], names)
        return out
    return _cache(f"snap5_{ticker}_{TODAY}{_vkey(vdate)}", f)


def _latest(d: dict, n=3, upto: str | None = None):
    return [v for k, v in sorted(d.items(), reverse=True) if v is not None and (upto is None or k <= upto)][:n]


def _at(d: dict, date: str | None):
    return d.get(date) if date else None


def statutory_tax(country: str | None) -> tuple[float, str]:
    if country in STATUTORY_TAX:
        return STATUTORY_TAX[country]
    return 0.25, f"No statutory-rate entry for {country or 'unknown country'}: 25% placeholder (OECD average ~24%); edit"


def default_inputs(s: dict) -> dict:
    i = s["info"]
    fin_ccy, q_ccy = i.get("financialCurrency") or i.get("currency"), i.get("currency")
    pence = q_ccy == "GBp"
    price = float(s.get("price_v") or i.get("currentPrice") or i.get("regularMarketPrice"))     # price_v: close on the valuation date
    live_px = float(i.get("currentPrice") or i.get("regularMarketPrice") or price)
    if pence:
        price, live_px, q_ccy = price / 100, live_px / 100, "GBP"
    vd = s.get("vdate") or TODAY
    fx = _fx(fin_ccy, q_ccy)                      # quote-currency units per financial-currency unit
    q = s.get("q", {})

    # ---- balance-sheet anchor: latest quarter end on or before the valuation date
    qdates = sorted({d for k in ("cash_sti", "debt") for d, v in q.get(k, {}).items() if v is not None and d <= vd})
    bs_date = qdates[-1] if qdates else None

    def bs(k, info_key=None):
        v = _at(q.get(k, {}), bs_date)
        if v is None:
            v = (_latest(s.get(k, {}), 1, vd) or [None])[0]
        if v is None and info_key:
            v = i.get(info_key)
        return float(v or 0)

    # ---- base fiscal year: latest reported fiscal year on or before the anchor
    anchor = bs_date or vd
    fy_dates = sorted(d for d, v in s.get("revenue", {}).items() if v is not None and d <= anchor)
    if fy_dates:
        fy0_date = fy_dates[-1]
        rev0 = float(s["revenue"][fy0_date])
        fy1_end = (pd.Timestamp(fy0_date) + pd.DateOffset(years=1)).date()
        if bs_date is None:
            bs_date = vd
        stub = float(np.clip((fy1_end - dt.date.fromisoformat(bs_date)).days / 365, 0.01, 1.0))
    else:                                         # no annual statements: TTM base, no stub
        fy0_date, rev0, stub = None, float(i.get("totalRevenue") or 0), 1.0
        bs_date = bs_date or vd
    fy0 = int(fy0_date[:4]) if fy0_date else dt.date.fromisoformat(vd).year - 1
    roll_days = float((dt.date.fromisoformat(vd) - dt.date.fromisoformat(bs_date)).days)

    # ---- revenue path: known later actuals (only for earlier valuation dates), then consensus
    later = [s["revenue"][d] for d in sorted(s.get("revenue", {})) if fy0_date and d > fy0_date and s["revenue"][d] is not None]
    re_ = s.get("rev_est", {})
    known = later + [re_[k] for k in ("0y", "+1y") if re_.get(k)]
    if len(known) >= 2 and rev0:
        g1, g2 = known[0] / rev0 - 1, known[1] / known[0] - 1
        g_src = "FY1-FY2: analyst consensus revenue (Yahoo Finance)" + (" incl. later reported actuals (look-ahead)" if later else "")
    else:
        g1 = i.get("revenueGrowth") or 0.05
        g2 = g1 * 0.8
        g_src = "FY1: Yahoo trailing revenue growth (no consensus); FY2 = 80% of FY1"
    g1, g2 = float(np.clip(g1, -0.2, 0.5)), float(np.clip(g2, -0.2, 0.5))
    g_term = 0.025
    g5 = min((g2 + g_term) / 2, 0.15)             # FY5: halfway from FY2 growth to terminal growth, at most 15%
    path = [g1, g2] + [float(x) for x in np.linspace(min(g2, 0.25), g5, 4)[1:]]   # beyond consensus, growth above 25% is not extrapolated

    # ---- profitability and reinvestment: last three reported fiscal years
    up = fy0_date or vd
    revs, ebits = _latest(s.get("revenue", {}), 3, up), _latest(s.get("ebit", {}), 3, up)
    fy_m = [e / r for e, r in zip(ebits, revs) if r]
    ttm = s.get("ttm") or {}
    if ttm.get("rev") and ttm.get("end", "") >= (fy0_date or ""):
        m_ttm, m_src = ttm["ebit"] / ttm["rev"], f"last four quarters to {ttm['end']}"
    elif fy_m:
        m_ttm, m_src = fy_m[0], f"fiscal year {fy0}"
    else:
        m_ttm, m_src = (i.get("operatingMargins") or 0.15), "Yahoo's latest-quarter figure"
    # target: today's margin, unless it sits more than 5 points outside the range of the reported fiscal years
    # (a cyclical peak or trough); then halfway back to the average of those years
    cyc = bool(fy_m) and (m_ttm > max(fy_m) + 0.05 or m_ttm < min(fy_m) - 0.05)
    m_target = float(np.mean([m_ttm, np.mean(fy_m)])) if cyc else float(m_ttm)
    das, cxs = _latest(s.get("da", {}), 3, up), _latest(s.get("capex", {}), 3, up)
    da_pct = float(np.mean([d / r for d, r in zip(das, revs)])) if das else 0.04
    cx_pct = float(np.mean([-c / r for c, r in zip(cxs, revs)])) if cxs else 0.04
    ppe = (_latest(s.get("ppe", {}), 1, up) or [None])[0]
    ppe_rev = float(ppe / revs[0]) if ppe and revs else None
    cx_term = da_pct + g_term * ppe_rev if ppe_rev is not None else max(cx_pct, da_pct)
    sbcs = _latest(s.get("sbc", {}), 3, up)
    sbc_pct = float(np.mean([x / r for x, r in zip(sbcs, revs)])) if sbcs else 0.0
    taxes = [t for t in _latest(s.get("tax_rate", {}), 3, up) if t is not None and 0 < t < 0.5]
    tax = float(np.clip(np.mean(taxes), 0.10, 0.35)) if taxes else 0.25
    tax_stat, tax_src = statutory_tax(i.get("country"))
    # NWC: trade working capital (receivables + inventory - payables) as % of revenue, last three fiscal years;
    # the investment each year is that ratio times the revenue increase
    twc = []
    for d_ in sorted(d for d, v in s.get("revenue", {}).items() if v and d <= up)[-3:]:
        ar, inv, ap, dr = (s.get(k, {}).get(d_) for k in ("receivables", "inventory", "payables", "deferred_rev"))
        if ar is not None or ap is not None:
            twc.append(((ar or 0) + (inv or 0) - (ap or 0) - (dr or 0)) / s["revenue"][d_])
    nwc_pct = float(np.clip(np.mean(twc), -0.10, 0.25)) if twc else 0.0

    # ---- cost of capital
    rf, rf_src = risk_free(fin_ccy, s.get("vdate"))
    erp, erp_src = default_erp(fin_ccy)
    beta_raw = s.get("beta_raw") or 1.0
    beta = 0.67 * beta_raw + 0.33
    interest = abs((_latest(s.get("interest", {}), 1, up) or [0])[0] or 0)
    cover = (ebits[0] / interest) if interest and ebits else 20
    spread, rating = next((sp, rt) for th, sp, rt in COVERAGE_SPREADS if cover > th)

    # ---- shares: all share classes (market cap / price), else listed shares
    mcap_live = float(i.get("marketCap") or 0)
    shares = mcap_live / live_px if mcap_live and live_px else float(i.get("sharesOutstanding") or i.get("impliedSharesOutstanding"))
    mcap_q = price * shares

    # ---- dividends (DDM cross-check)
    dps0 = float(i.get("dividendRate") or i.get("trailingAnnualDividendRate") or 0) / (100 if pence else 1)
    ee = s.get("eps_est", {})
    g_div = float(np.clip(ee["+1y"] / ee["0y"] - 1, 0.0, 0.15)) if ee.get("0y") and ee.get("+1y") and ee["0y"] > 0 else 0.05

    cash, lt_inv, debt = bs("cash_sti", "totalCash"), bs("lt_inv"), bs("debt", "totalDebt")
    nci, pref = bs("nci"), bs("pref")
    p = {
        "name": i.get("longName") or i.get("shortName") or s["ticker"], "ticker": s["ticker"],
        "fin_ccy": fin_ccy, "q_ccy": q_ccy, "fx": fx, "price": price, "shares": shares,
        "fy0": fy0, "fy0_date": fy0_date, "bs_date": bs_date, "valuation_date": vd, "stub": stub, "roll_days": roll_days,
        "rev0": rev0, "growth": path[:N_EXPLICIT],
        "margin0": float(m_ttm), "margin_target": m_target,
        "tax": tax, "tax_terminal": tax_stat, "tax_wacc": tax_stat,
        "da_pct": da_pct, "da_terminal": da_pct, "capex_pct": cx_pct, "capex_terminal": float(cx_term),
        "nwc_pct": nwc_pct, "sbc_pct": sbc_pct, "sbc_addback": 0,
        "g": g_term, "rf": rf, "beta": beta, "erp": erp, "kd": rf + spread,
        "cash": cash, "lt_inv": lt_inv, "other_nonop": 0.0, "debt": debt, "pension": 0.0, "nci": nci, "pref": pref,
        "mcap_fin": mcap_q / fx, "dps0": dps0, "g_div": g_div,
        "d_growth": 0.0, "d_margin": 0.0,
        "sources": {
            "rf": rf_src, "beta": f"OLS 5y weekly vs {s.get('beta_index', '?')} (n={s.get('beta_n', 0)}), raw {beta_raw:.2f}, Blume-adjusted 0.67 x raw + 0.33",
            "erp": erp_src,
            "kd": f"Rf + synthetic-rating spread {spread*100:.2f}% (interest cover {cover:.1f}x -> {rating}, Damodaran table)",
            "growth": f"{g_src}; FY3-FY5 linear from FY2 growth (at most 25%) towards the midpoint of FY2 and terminal growth (at most 15% in FY5); "
                      "FY6-FY10 fade linearly to g",
            "margin": f"FY1 = trailing-twelve-month operating margin ({m_src}; GAAP/IFRS, SBC included). "
                      + (f"It is more than 5 points outside the range of the last {len(fy_m)} fiscal years ({min(fy_m) * 100:.1f}% - {max(fy_m) * 100:.1f}%), "
                         f"so the target from FY5 is set halfway back to their average ({np.mean(fy_m) * 100:.1f}%): a cyclical peak or trough is not extrapolated"
                         if cyc else "Target from FY5 = the same margin (within the range of recent fiscal years); edit for a margin story"),
            "tax": f"Explicit years: average effective tax rate of the last {len(taxes)} fiscal years (Yahoo 'Tax Rate For Calcs'), "
                   "fading in FY6-FY10 to the statutory rate",
            "tax_terminal": f"Statutory (marginal) rate: {tax_src}. Lecture: the marginal home-country rate is used in perpetuity and "
                            "for the debt tax shield, because deductions save tax at the marginal rate",
            "reinvestment": f"D&A and capex: average % of revenue over the last {len(das)} fiscal years. Terminal capex = D&A + g x net PP&E / revenue"
                            + (f" ({ppe_rev * 100:.0f}%)" if ppe_rev is not None else "")
                            + ": in a steady state the asset base grows with revenue, so capex covers depreciation plus that growth "
                              "(today's capex can sit far above or below it during a build-out or a pause)",
            "nwc": f"Trade working capital (receivables + inventory - payables - deferred revenue) / revenue, average of the last {len(twc)} fiscal years, capped at -10% / +25%; "
                   "invested (or released, if negative) in proportion to revenue growth",
            "sbc": "Stock-based compensation is treated as a real expense (it is in operating profit); the add-back switch is off",
            "bridge": f"Latest quarterly balance sheet on or before the valuation date ({bs_date}): cash and short-term investments, "
                      "long-term financial investments, total debt incl. leases, minority interests, preferred stock (Yahoo Finance). "
                      "Pensions and other items are not in Yahoo's data: add them from the filings",
            "timing": f"Value anchored at {bs_date}; the current fiscal year FY{fy0 + 1} is valued for the {stub:.2f} years still to come, "
                      f"mid-period discounting, then rolled forward {roll_days:.0f} days to the valuation date at the cost of equity",
            "shares": "All share classes: market capitalisation / share price (Yahoo Finance)",
            "dividend": "Indicated annual dividend (Yahoo Finance); dividend growth FY1-FY5 = consensus EPS growth FY2 vs FY1, capped at 0-15%",
        },
    }
    return p


# ---------------------------------------------------------------- model
_DEFAULTS = {"stub": 1.0, "roll_days": 0.0, "fy0": 0, "da_terminal": None, "capex_terminal": None, "tax_terminal": None,
             "sbc_pct": 0.0, "sbc_addback": 0, "lt_inv": 0.0, "other_nonop": 0.0, "pension": 0.0, "nci": 0.0, "pref": 0.0,
             "d_growth": 0.0, "d_margin": 0.0, "dps0": 0.0, "g_div": 0.05}


def norm(p: dict) -> dict:
    """Fill keys older callers do not set (paired terminal values default to the explicit ones)."""
    q = dict(p)
    for k, v in _DEFAULTS.items():
        if q.get(k) is None:
            q[k] = v
    for k, base in (("da_terminal", "da_pct"), ("capex_terminal", "capex_pct"), ("tax_terminal", "tax")):
        if p.get(k) is None:
            q[k] = q[base]
    q.setdefault("tax_wacc", q["tax"])
    return q


def wacc(p: dict) -> dict:
    ke = p["rf"] + p["beta"] * p["erp"]
    e, d = p["mcap_fin"], p["debt"]
    we = e / (e + d) if e + d else 1
    return {"ke": ke, "kd_at": p["kd"] * (1 - p["tax_wacc"]), "we": we, "wd": 1 - we,
            "wacc": we * ke + (1 - we) * p["kd"] * (1 - p["tax_wacc"])}


def _fade(v5, term, k):
    return v5 + (term - v5) * k / N_FADE


def _rows(p: dict) -> list[dict]:
    """Forecast FY1..FY10. Mirrors the Forecast sheet of the workbook cell by cell."""
    g5 = p["growth"][N_EXPLICIT - 1] + p["d_growth"]
    mt = p["margin_target"] + p["d_margin"]
    rows, prev = [], p["rev0"]
    for t in range(1, N_YEARS + 1):
        if t <= N_EXPLICIT:
            g = p["growth"][t - 1] + (p["d_growth"] if t >= 2 else 0.0)
            m = p["margin0"] + (mt - p["margin0"]) * (t - 1) / (N_EXPLICIT - 1)
            tx, da_p, cx_p = p["tax"], p["da_pct"], p["capex_pct"]
        else:
            k = t - N_EXPLICIT
            g = _fade(g5, p["g"], k)
            m = mt
            tx, da_p, cx_p = _fade(p["tax"], p["tax_terminal"], k), _fade(p["da_pct"], p["da_terminal"], k), _fade(p["capex_pct"], p["capex_terminal"], k)
        rev = prev * (1 + g)
        ebit = rev * m
        nopat = ebit * (1 - tx)
        sbc = rev * p["sbc_pct"] * p["sbc_addback"]
        da, cx = rev * da_p, rev * cx_p
        dnwc = p["nwc_pct"] * (rev - prev)
        rows.append({"revenue": rev, "growth": g, "margin": m, "ebit": ebit, "tax_rate": tx, "nopat": nopat, "sbc": sbc,
                     "da_pct": da_p, "da": da, "capex_pct": cx_p, "capex": cx, "dnwc": dnwc, "fcff": nopat + sbc + da - cx - dnwc})
        prev = rev
    return rows


def forecast(p: dict) -> pd.DataFrame:
    p = norm(p)
    return pd.DataFrame(_rows(p), index=[p["fy0"] + t for t in range(1, N_YEARS + 1)])


def dcf(p: dict, wacc_override: float | None = None, full: bool = True) -> dict:
    p = norm(p)
    w = wacc(p)
    r = w["wacc"] if wacc_override is None else wacc_override
    rows = _rows(p)
    stub = p["stub"]
    t = np.array([stub / 2] + [stub + k + 0.5 for k in range(N_YEARS - 1)])
    fcff = np.array([x["fcff"] for x in rows])
    cf = fcff * np.array([stub] + [1.0] * (N_YEARS - 1))
    pv = cf / (1 + r) ** t
    last = rows[-1]
    mt = p["margin_target"] + p["d_margin"]
    rev_t1 = last["revenue"] * (1 + p["g"])
    ebit_t1 = rev_t1 * mt
    fcff_t1 = (ebit_t1 * (1 - p["tax_terminal"]) + rev_t1 * p["sbc_pct"] * p["sbc_addback"]
               + rev_t1 * (p["da_terminal"] - p["capex_terminal"]) - p["nwc_pct"] * (rev_t1 - last["revenue"]))
    tv = fcff_t1 / (r - p["g"])
    t_tv = stub + N_YEARS - 1
    pv_tv = tv / (1 + r) ** t_tv
    ev = pv.sum() + pv_tv
    non_op = p["cash"] + p["lt_inv"] + p["other_nonop"]
    claims = p["debt"] + p["pension"] + p["nci"] + p["pref"]
    equity = ev + non_op - claims
    ps_anchor = equity * p["fx"] / p["shares"]
    roll = (1 + w["ke"]) ** (p["roll_days"] / 365)
    price = ps_anchor * roll
    out = {"wacc": r, **{k: v for k, v in w.items() if k != "wacc"}, "price": price, "ev": ev, "equity": equity,
           "upside": price / p["price"] - 1, "tv_share": pv_tv / ev}
    if not full:
        return out
    ntm_rev = stub * rows[0]["revenue"] + (1 - stub) * rows[1]["revenue"]
    ntm_ebit = stub * rows[0]["ebit"] + (1 - stub) * rows[1]["ebit"]
    out.update({"forecast": pd.DataFrame(rows, index=[p["fy0"] + k for k in range(1, N_YEARS + 1)]), "cf": cf, "t": t, "pv": pv,
                "sum_pv": float(pv.sum()), "fcff_t1": fcff_t1, "ebit_t1": ebit_t1, "tv": tv, "pv_tv": pv_tv, "t_tv": t_tv,
                "non_op": non_op, "claims": claims, "ps_anchor": ps_anchor, "roll": roll, "exit_ev_ebit": tv / ebit_t1,
                "implied_ev_sales_ntm": ev / ntm_rev, "implied_ev_ebit_ntm": ev / ntm_ebit if ntm_ebit > 0 else None})
    return out


# ---------------------------------------------------------------- scenarios (same shifts as the workbook's selector)
SCENARIOS = {
    "Bear": {"d_growth": -0.03, "d_margin": -0.02, "d_g": -0.005, "d_wacc": 0.005, "weight": 0.25,
             "story": "Growth 3pp lower from FY2, target margin 2pp lower, terminal growth -0.5pp, WACC +0.5pp"},
    "Base": {"d_growth": 0.0, "d_margin": 0.0, "d_g": 0.0, "d_wacc": 0.0, "weight": 0.50,
             "story": "Consensus revenue for FY1-FY2, margin and reinvestment at their recent averages"},
    "Bull": {"d_growth": 0.02, "d_margin": 0.015, "d_g": 0.0025, "d_wacc": -0.0025, "weight": 0.25,
             "story": "Growth 2pp higher from FY2, target margin 1.5pp higher, terminal growth +0.25pp, WACC -0.25pp"},
}


def scenario_inputs(p: dict, name: str) -> tuple[dict, float]:
    s = SCENARIOS[name]
    q = {**norm(p), "d_growth": s["d_growth"], "d_margin": s["d_margin"], "g": p["g"] + s["d_g"]}
    return q, wacc(norm(p))["wacc"] + s["d_wacc"]


def scenarios(p: dict) -> dict:
    out = {}
    for n, s in SCENARIOS.items():
        q, w = scenario_inputs(p, n)
        d = dcf(q, w)
        out[n] = {"price": d["price"], "weight": s["weight"], "wacc": w, "g": q["g"], "m_target": q["margin_target"] + q["d_margin"],
                  "rev5": float(d["forecast"].iloc[N_EXPLICIT - 1]["revenue"]), "story": s["story"]}
    out["Weighted"] = {"price": sum(v["price"] * v["weight"] for v in out.values())}
    return out


# ---------------------------------------------------------------- sensitivities
def _grid(rows, cols, fn) -> pd.DataFrame:
    g = pd.DataFrame(index=[round(x, 4) for x in rows], columns=[round(x, 4) for x in cols], dtype=float)
    for a in g.index:
        for b in g.columns:
            g.loc[a, b] = fn(a, b)
    return g


def grid_wacc_g(p: dict, d_w=np.arange(-0.01, 0.0101, 0.0025), d_g=np.arange(-0.0075, 0.00751, 0.0025)) -> pd.DataFrame:
    p = norm(p)
    w0 = wacc(p)["wacc"]
    return _grid([w0 + x for x in d_w], [p["g"] + x for x in d_g],
                 lambda w, g: dcf({**p, "g": g}, w, full=False)["price"] if w > g + 0.005 else np.nan)


def grid_wacc_margin(p: dict) -> pd.DataFrame:
    p = norm(p)
    w0 = wacc(p)["wacc"]
    return _grid([w0 + x for x in np.arange(-0.01, 0.0101, 0.0025)], [p["margin_target"] + x for x in np.arange(-0.04, 0.0401, 0.01)],
                 lambda w, m: dcf({**p, "margin_target": m}, w, full=False)["price"] if w > p["g"] + 0.005 else np.nan)


def grid_growth_margin(p: dict) -> pd.DataFrame:
    p = norm(p)
    return _grid([-0.04, -0.02, 0.0, 0.02, 0.04], [-0.03, -0.015, 0.0, 0.015, 0.03],
                 lambda dg, dm: dcf({**p, "d_growth": dg, "d_margin": dm}, full=False)["price"])


def grid_beta_erp(p: dict) -> pd.DataFrame:
    p = norm(p)
    return _grid([p["beta"] + x for x in (-0.2, -0.1, 0, 0.1, 0.2)], [p["erp"] + x for x in (-0.01, -0.005, 0, 0.005, 0.01)],
                 lambda b, e: dcf({**p, "beta": b, "erp": e}, full=False)["price"])


def tornado(p: dict) -> pd.DataFrame:
    p = norm(p)
    base = dcf(p, full=False)["price"]
    tests = [
        ("Risk-free rate ±50bp", "rf", -0.005, 0.005), ("Beta ±0.10", "beta", -0.10, 0.10), ("ERP ±50bp", "erp", -0.005, 0.005),
        ("Terminal growth ±25bp", "g", 0.0025, -0.0025), ("Target EBIT margin ±2pp", "margin_target", 0.02, -0.02),
        ("Terminal tax rate ±2pp", "tax_terminal", -0.02, 0.02), ("Terminal capex ±0.5pp of revenue", "capex_terminal", -0.005, 0.005),
        ("Revenue growth FY2-FY5 ±2pp", "d_growth", 0.02, -0.02), ("NWC intensity ±5pp", "nwc_pct", -0.05, 0.05),
    ]
    rows = []
    for label, key, up, dn in tests:
        a, b = dcf({**p, key: p[key] + up}, full=False)["price"], dcf({**p, key: p[key] + dn}, full=False)["price"]
        rows.append({"driver": label, "low": min(a, b), "high": max(a, b), "base": base, "range": abs(a - b)})
    if p["sbc_pct"] > 0:
        a = dcf({**p, "sbc_addback": 1 - p["sbc_addback"]}, full=False)["price"]
        rows.append({"driver": "SBC add-back on / off", "low": min(a, base), "high": max(a, base), "base": base, "range": abs(a - base)})
    return pd.DataFrame(rows).sort_values("range", ascending=False).reset_index(drop=True)


MC_DRIVERS = [("d_wacc", "WACC shift", 0.005), ("d_g", "Terminal growth shift", 0.0025),
              ("d_mt", "Target margin shift", 0.015), ("d_growth", "Revenue growth shift FY2-FY5", 0.015)]


def monte_carlo(p: dict, n: int = 2000, seed: int = 7) -> dict:
    p = norm(p)
    rng = np.random.default_rng(seed)
    draws = {k: np.clip(rng.normal(0, sd, n), -2.5 * sd, 2.5 * sd) for k, _, sd in MC_DRIVERS}
    w0 = wacc(p)["wacc"]
    px = np.empty(n)
    for j in range(n):
        w = w0 + draws["d_wacc"][j]
        g = min(p["g"] + draws["d_g"][j], w - 0.01)
        draws["d_g"][j] = g - p["g"]
        px[j] = dcf({**p, "g": g, "margin_target": p["margin_target"] + draws["d_mt"][j], "d_growth": draws["d_growth"][j]}, w, full=False)["price"]
    df = pd.DataFrame({**draws, "price": px})
    pct = {q: float(np.percentile(px, q)) for q in (5, 25, 50, 75, 95)}
    return {"n": n, "seed": seed, "draws": df, "pct": pct, "mean": float(px.mean()), "sd": float(px.std(ddof=1)),
            "p_above_price": float((px > p["price"]).mean()), "base": dcf(p, full=False)["price"], "drivers": MC_DRIVERS}


# ---------------------------------------------------------------- reverse DCF: what the share price implies
def _solve(fn, lo: float, hi: float, target: float, n: int = 70) -> float | None:
    """Bisection for fn(x) = target on [lo, hi]; None if the target is not bracketed."""
    flo, fhi = fn(lo) - target, fn(hi) - target
    if not (np.isfinite(flo) and np.isfinite(fhi)) or flo * fhi > 0:
        return None
    for _ in range(n):
        mid = (lo + hi) / 2
        fm = fn(mid) - target
        if flo * fm <= 0:
            hi = mid
        else:
            lo, flo = mid, fm
    return (lo + hi) / 2


def reverse_dcf(p: dict, hist: pd.DataFrame | None = None) -> dict:
    """Solve the DCF backwards: the value each driver would need, one at a time and all else at base, for the DCF
    value to equal the share price. None = no value in a sensible range does it."""
    p = norm(p)
    price, w0 = p["price"], wacc(p)["wacc"]
    px = lambda q, w=None: dcf(q, w, full=False)["price"]  # noqa: E731
    m = _solve(lambda x: px({**p, "margin_target": x}), -0.20, 0.95, price)
    dg = _solve(lambda x: px({**p, "d_growth": x}), -0.30, 0.60, price)
    w = _solve(lambda x: px(p, x), p["g"] + 0.005, 0.30, price)
    g = _solve(lambda x: px({**p, "g": x}), -0.03, w0 - 0.005, price)
    base_g = float(np.mean(p["growth"][1:]))
    ctx = {}
    if hist is not None and len(hist) and "revenue" in hist:
        h = hist[hist["revenue"].notna()]
        if "ebit" in h and h["ebit"].notna().any():
            mm = (h["ebit"] / h["revenue"]).dropna()
            ctx["margin_hist"] = (float(mm.min()), float(mm.max()))
        if len(h) >= 2:
            n_ = len(h) - 1
            ctx["rev_cagr"] = float((h["revenue"].iloc[-1] / h["revenue"].iloc[0]) ** (1 / n_) - 1)
            ctx["rev_years"] = n_
    return {"price": price, "base_price": px(p),
            "margin_target": {"base": p["margin_target"], "implied": m},
            "growth": {"base": base_g, "implied": None if dg is None else base_g + dg, "shift": dg},
            "wacc": {"base": w0, "implied": w},
            "g": {"base": p["g"], "implied": g}, "context": ctx}


class ModelNotApplicable(ValueError):
    """The one-segment FCFF model does not fit this company (bank, insurer, REIT, loss-maker)."""


def applicability(s: dict) -> list[str]:
    """Reasons why the FCFF model should not be used for this company (empty list = fine)."""
    i = s.get("info", {})
    ind, sec = (i.get("industry") or ""), (i.get("sector") or "")
    out = []
    if ind.startswith("Banks") or "Insurance" in ind:
        out.append(f"{ind}: for banks and insurers debt is an operating item, so free cash flow to the firm and WACC are not meaningful. "
                   "Use a dividend discount or excess-return model")
    if ind.startswith("REIT") or (sec == "Real Estate" and "REIT" in ind):
        out.append(f"{ind}: REITs are valued on funds from operations and net asset value, not on an FCFF DCF")
    if ind in ("Auto Manufacturers", "Farm & Heavy Construction Machinery") and (i.get("totalDebt") or 0) > 0.5 * (i.get("marketCap") or float("inf")):
        out.append(f"{ind} with debt above half of market value: most likely a captive financing arm, whose customer loans, leased assets and "
                   "funding debt distort capex, working capital and the WACC. Value the industrial and financing businesses separately")
    ebit = (_latest(s.get("ebit", {}), 1) or [None])[0]
    om = i.get("operatingMargins")
    if (ebit is not None and ebit <= 0) or (om is not None and om <= 0):
        out.append("Operating profit is negative: a margin-on-revenue forecast from the trailing margin gives a negative value. "
                   "A loss-making company needs an explicit path to profitability")
    if not (i.get("totalRevenue") or _latest(s.get("revenue", {}), 1)):
        out.append("No revenue reported")
    return out


def ddm(p: dict) -> dict | None:
    """Two-stage dividend discount model at the cost of equity (cross-check; ignores buybacks)."""
    p = norm(p)
    if not p["dps0"] or p["dps0"] <= 0:
        return None
    ke = wacc(p)["ke"]
    dps, rows, prev = [], [], p["dps0"]
    for t in range(1, N_YEARS + 1):
        g = p["g_div"] if t <= N_EXPLICIT else _fade(p["g_div"], p["g"], t - N_EXPLICIT)
        d = prev * (1 + g)
        tt = 0.5 + (t - 1)
        rows.append({"growth": g, "dps": d, "t": tt, "pv": d / (1 + ke) ** tt})
        prev = d
    dps_t1 = prev * (1 + p["g"])
    tv = dps_t1 / (ke - p["g"])
    pv_tv = tv / (1 + ke) ** (rows[-1]["t"] + 0.5)
    pv_div = sum(r["pv"] for r in rows)
    return {"table": pd.DataFrame(rows), "ke": ke, "dps_t1": dps_t1, "tv": tv, "pv_tv": pv_tv, "pv_div": pv_div, "price": pv_div + pv_tv}


def history(s: dict) -> pd.DataFrame:
    """Annual statement lines (raw currency units) for the fiscal years Yahoo reports, oldest first."""
    keys = [k for k in _ANNUAL if s.get(k)]
    dates = sorted({d for k in keys for d, v in s[k].items() if v is not None})
    df = pd.DataFrame({k: [s[k].get(d) for d in dates] for k in keys}, index=dates)
    return df[df["revenue"].notna()] if "revenue" in df else df


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
        # marketCap covers all share classes (sharesOutstanding only the listed one); Yahoo quotes it in the main unit (GBP, not pence)
        mcap_q = float(i.get("marketCap") or 0) or price * float(i.get("sharesOutstanding") or 0)
        if fin != q_ccy:
            # ADRs / cross listings: Yahoo's EV/revenue mixes quote and reporting currency -> rebuild EV in reporting currency
            ev = (mcap_q * _fx(q_ccy, fin)
                  + float(i.get("totalDebt") or 0) - float(i.get("totalCash") or 0))
        else:
            ev = float(i["enterpriseToRevenue"]) * rev_ttm
        om = i.get("operatingMargins")
        eps_conv = _fx(fin, q_ccy)
        mcap_eur = mcap_q * _fx(q_ccy, "EUR")
        return {"ticker": sym, "name": i.get("longName") or i.get("shortName"), "country": i.get("country"),
                "industry": i.get("industry"), "mcap_eur": mcap_eur, "ev_sales_ntm": ev / rev_ntm,
                "ev_ebit_ntm": ev / (rev_ntm * om) if om and om > 0 else None,
                "pe_ntm": price / (eps_ntm * eps_conv) if eps_ntm and eps_ntm > 0 else None,
                "g_ntm": (rev_ntm / rev_ttm - 1) * 100, "ebit_margin": (om or 0) * 100,
                "rule40": (rev_ntm / rev_ttm - 1) * 100 + (om or 0) * 100}
    try:
        return _cache(f"peer2_{sym}_{TODAY}", f)
    except Exception:  # noqa: BLE001
        return None


COMP_KEYS = (("ev_sales_ntm", "EV / NTM sales"), ("ev_ebit_ntm", "EV / NTM EBIT (GAAP)"), ("pe_ntm", "P / NTM EPS"))


def comps(p: dict, s: dict, peers: pd.DataFrame) -> dict:
    p = norm(p)
    re_ = s.get("rev_est", {})
    info = s.get("info", {})
    lfy = pd.Timestamp(info["lastFiscalYearEnd"], unit="s") if info.get("lastFiscalYearEnd") else None
    w0 = float(np.clip(((lfy + pd.DateOffset(years=1)) - pd.Timestamp(s.get("vdate") or TODAY)).days / 365, 0, 1)) if lfy is not None else 0.5
    rev_ntm = (w0 * re_["0y"] + (1 - w0) * re_["+1y"]) if re_.get("0y") and re_.get("+1y") else p["rev0"] * (1 + p["growth"][0])
    ee = s.get("eps_est", {})
    eps_ntm = (w0 * ee["0y"] + (1 - w0) * ee["+1y"]) * p["fx"] if ee.get("0y") and ee.get("+1y") else None
    ebit_ntm = rev_ntm * p["margin0"]
    bridge = p["cash"] + p["lt_inv"] + p["other_nonop"] - p["debt"] - p["pension"] - p["nci"] - p["pref"]
    to_px = lambda ev: (ev + bridge) * p["fx"] / p["shares"]  # noqa: E731
    tgt = {"rev_ntm": rev_ntm, "ebit_ntm": ebit_ntm, "eps_ntm": eps_ntm, "bridge": bridge}
    out = {"stats": {}, "sap": tgt, "target": tgt}
    for key, label in COMP_KEYS:
        metric = {"ev_sales_ntm": rev_ntm, "ev_ebit_ntm": ebit_ntm, "pe_ntm": eps_ntm}[key]
        x = peers[key].dropna() if key in peers else pd.Series(dtype=float)
        x = x[x > 0]
        if len(x) < 2 or metric is None or metric <= 0:
            continue
        q1, md, q3 = np.percentile(x, 25), np.median(x), np.percentile(x, 75)
        conv = (lambda m: m * metric) if key == "pe_ntm" else (lambda m: to_px(m * metric))
        out["stats"][key] = {"label": label, "n": len(x), "q1": q1, "median": md, "q3": q3, "mean": float(x.mean()), "metric": metric,
                             "p_q1": conv(q1), "p_med": conv(md), "p_q3": conv(q3)}
    reg = peers[["ev_sales_ntm", "rule40"]].dropna() if len(peers) and "rule40" in peers else pd.DataFrame()
    if len(reg) >= 4:
        X = np.column_stack([np.ones(len(reg)), reg["rule40"]])
        beta, *_ = np.linalg.lstsq(X, reg["ev_sales_ntm"].values, rcond=None)
        resid = reg["ev_sales_ntm"].values - X @ beta
        se = float(np.sqrt(resid @ resid / max(len(reg) - 2, 1)))
        r40 = (rev_ntm / p["rev0"] - 1) * 100 + p["margin0"] * 100
        fit = beta[0] + beta[1] * r40
        out["regression"] = {"slope": beta[1], "intercept": beta[0], "se": se, "fit": fit, "n": len(reg), "rule40": r40,
                             "p_fit": to_px(fit * rev_ntm), "p_low": to_px((fit - se) * rev_ntm), "p_high": to_px((fit + se) * rev_ntm)}
    return out


def comps_value(cp: dict) -> float | None:
    meds = [st["p_med"] for st in cp.get("stats", {}).values()]
    return sum(meds) / len(meds) if meds else None


def run(ticker: str, overrides: dict | None = None, peer_tickers: list[str] | None = None, vdate: str | None = None,
        extras: bool = True, force: bool = False) -> dict:
    vdate = check_date(vdate)
    s = snapshot(ticker, vdate)
    problems = applicability(s)
    if problems and not force:
        raise ModelNotApplicable(f"{ticker}: " + "; ".join(problems))
    p = default_inputs(s)
    if overrides:
        p.update(overrides)
    p = norm(p)
    d = dcf(p)
    if (d["fcff_t1"] <= 0 or d["ev"] <= 0) and not force:
        raise ModelNotApplicable(
            f"{ticker}: the forecast free cash flow is negative in the long run (capex {p['capex_terminal'] * 100:.1f}% of revenue and working capital "
            f"{p['nwc_pct'] * 100:.0f}% of growth exceed after-tax profit). Typical for a company with a financing arm, whose customer loans "
            "and leased assets sit in receivables and capex; value the industrial and financing businesses separately")
    if abs(p["growth"][0]) > 0.30:
        problems = problems + [f"FY1 revenue growth is {p['growth'][0] * 100:+.0f}% against the last fiscal year: check that the consensus figure "
                               "uses the same revenue definition as the reported one"]
    sc = scenarios(p)
    w0 = d["wacc"]
    band = [dcf({**p, "g": p["g"] + dg}, w0 + dw, full=False)["price"] for dw in (-0.005, 0, 0.005) for dg in (-0.0025, 0, 0.0025)]
    peers = pd.DataFrame([r for r in (peer_row(t) for t in (peer_tickers or [])) if r])
    cp = comps(p, s, peers)
    i = s["info"]
    to_q = (1 / 100) if i.get("currency") == "GBp" else 1
    market = {"price": p["price"], "price_date": s["price_date"],
              "hi52": (s.get("hi_v") or (i.get("fiftyTwoWeekHigh") or p["price"] / to_q) * to_q),
              "lo52": (s.get("lo_v") or (i.get("fiftyTwoWeekLow") or p["price"] / to_q) * to_q),
              "tgt_mean": (i.get("targetMeanPrice") or 0) * to_q or None, "tgt_low": (i.get("targetLowPrice") or 0) * to_q or None,
              "tgt_high": (i.get("targetHighPrice") or 0) * to_q or None, "tgt_n": i.get("numberOfAnalystOpinions") or 0}
    meta = {"valuation_date": vdate, "lookahead": vdate < TODAY}
    res = {"inputs": p, "snapshot": s, "dcf": d, "scenarios": sc, "dcf_band": (min(band), max(band)),
           "grid_wacc_g": grid_wacc_g(p), "comps": cp, "peers": peers, "market": market, "meta": meta}
    if extras:
        res.update({"grid_wacc_margin": grid_wacc_margin(p), "grid_growth_margin": grid_growth_margin(p), "grid_beta_erp": grid_beta_erp(p),
                    "tornado": tornado(p), "mc": monte_carlo(p), "ddm": ddm(p), "hist": history(s)})
        res["reverse"] = reverse_dcf(p, res["hist"])
    if problems:
        meta["warnings"] = problems
    return res


def football_items(res: dict) -> list[dict]:
    mk, d, sc, cp = res["market"], res["dcf"], res["scenarios"], res["comps"]
    items = [{"label": "52-week range", "low": mk["lo52"], "high": mk["hi52"], "mid": None, "kind": "market"}]
    if mk.get("tgt_low") and mk.get("tgt_high"):
        items.append({"label": f"Analyst targets (n={mk['tgt_n']})", "low": mk["tgt_low"], "high": mk["tgt_high"], "mid": mk["tgt_mean"], "kind": "market"})
    items += [{"label": "DCF: WACC ±0.5pp, g ±0.25pp", "low": res["dcf_band"][0], "high": res["dcf_band"][1], "mid": d["price"], "kind": "dcf"},
              {"label": "DCF: bear / base / bull", "low": sc["Bear"]["price"], "high": sc["Bull"]["price"], "mid": sc["Weighted"]["price"], "kind": "dcf"}]
    if res.get("mc"):
        mc = res["mc"]
        items.append({"label": "DCF: Monte Carlo P5-P95", "low": mc["pct"][5], "high": mc["pct"][95], "mid": mc["pct"][50], "kind": "dcf"})
    for k, st in cp["stats"].items():
        items.append({"label": f"{st['label']} (peer IQR)", "low": st["p_q1"], "high": st["p_q3"], "mid": st["p_med"], "kind": "comps"})
    if "regression" in cp:
        r = cp["regression"]
        items.append({"label": "Growth-adj. EV/Sales (Rule of 40, ±1 SE)", "low": r["p_low"], "high": r["p_high"], "mid": r["p_fit"], "kind": "comps"})
    return items
