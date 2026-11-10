"""Web edition loader: public sources only.

- SAP IR quarterly statements (SEC 6-K, transcribed in sap_ir.yaml)
- SEC EDGAR XBRL company facts (ifrs-full) for multi-year history
- yfinance: price, 52w range, analyst targets/estimates, statements, beta regression, peers
- Deutsche Bundesbank API: 10y Bund yield
- Damodaran (NYU Stern): implied ERP and rating default spreads (cross-checks)
Every network fetch is cached in data/web/ with the fetch date, so reruns are reproducible."""
from __future__ import annotations

import datetime as dt
import io
import json
import os
import re
from pathlib import Path

import numpy as np
import pandas as pd
import requests

from .common import PEER_CLASS, load_sap_ir, peer_frame

CACHE = Path(os.environ.get("VALUATION_CACHE", Path(__file__).resolve().parents[2] / "data" / "web"))
UA = {"User-Agent": "SAP-valuation-coursework research-bot (contact: admin@example.org)"}
TODAY = dt.date.today().isoformat()

# yfinance symbols for the Capital IQ peer list (+ Intuit, a public-only addition)
YF_PEERS = {
    "ORCL": "ORCL", "MSFT": "MSFT", "CRM": "CRM", "NOW": "NOW", "ADBE": "ADBE", "WDAY": "WDAY", "ADSK": "ADSK",
    "DSY": "DSY.PA", "PTC": "PTC", "SGE": "SGE.L", "OTEX": "OTEX", "NEM": "NEM.DE", "INTU": "INTU",
    "ACN": "ACN", "CAP": "CAP.PA", "SOP": "SOP.PA", "IBM": "IBM", "DELL": "DELL", "PANW": "PANW", "CRWD": "CRWD",
    "NET": "NET", "DDOG": "DDOG", "ESTC": "ESTC", "PLTR": "PLTR", "TEAM": "TEAM", "HUBS": "HUBS", "DOCU": "DOCU",
    "PCTY": "PCTY", "CDNS": "CDNS", "SNPS": "SNPS", "TMV": "TMV.DE", "SAP": "SAP.DE",
}


def _cache_json(name: str, fn):
    CACHE.mkdir(parents=True, exist_ok=True)
    p = CACHE / f"{name}.json"
    if p.exists():
        return json.loads(p.read_text())
    data = fn()
    p.write_text(json.dumps(data, default=str, indent=1))
    return data


# ------------------------------------------------------------------ public sources
def bund_10y() -> dict:
    def f():
        url = ("https://api.statistiken.bundesbank.de/rest/data/BBSIS/"
               "D.I.ZAR.ZI.EUR.S1311.B.A604.R10XX.R.A.A._Z._Z.A?format=csv&lastNObservations=15&lang=en")
        txt = requests.get(url, timeout=30).text
        obs = [l.split(",") for l in txt.splitlines() if re.match(r"^\d{4}-\d{2}-\d{2},\d", l)]
        d, v = obs[-1][0], float(obs[-1][1])
        return {"date": d, "yield": v, "url": url, "fetched": TODAY}
    return _cache_json(f"bundesbank_10y_{TODAY}", f)


def damodaran() -> dict:
    def f():
        h = requests.get("https://pages.stern.nyu.edu/~adamodar/New_Home_Page/home.htm", headers=UA, timeout=30).text
        t = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", h))
        m = re.search(r"Implied ERP on ([A-Za-z]+ \d+, \d{4})\s*=\s*([\d.]+)\s*%.*?default spread\s*\(\s*([\d.]+)\s*%", t, re.S)
        r = requests.get("https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ratings.html", headers=UA, timeout=30).text
        tab = [x for x in pd.read_html(io.StringIO(r)) if "A1/A+" in x.to_string()][0]
        spread = None
        for _, row in tab.iterrows():
            if str(row.iloc[2]).strip() == "A1/A+":
                spread = float(str(row.iloc[3]).strip("%")) / 100
                break
        return {"erp_date": m.group(1), "erp": float(m.group(2)) / 100, "us_default_spread": float(m.group(3)) / 100,
                "a_plus_spread": spread, "fetched": TODAY,
                "urls": ["https://pages.stern.nyu.edu/~adamodar/New_Home_Page/home.htm",
                         "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ratings.html"]}
    return _cache_json(f"damodaran_{TODAY}", f)


def sec_facts() -> dict:
    p = CACHE / "sec_companyfacts_sap.json"
    if not p.exists():
        CACHE.mkdir(parents=True, exist_ok=True)
        r = requests.get("https://data.sec.gov/api/xbrl/companyfacts/CIK0001000184.json", headers=UA, timeout=60)
        p.write_bytes(r.content)
    facts = json.loads(p.read_text())["facts"]["ifrs-full"]

    def series(tag, instant=False):
        if tag not in facts:
            return {}
        out = {}
        for x in list(facts[tag]["units"].values())[0]:
            fr = x.get("frame") or ""
            m = re.match(r"CY(\d{4})(Q4I)?$", fr)
            if m and (bool(m.group(2)) == instant):
                out[int(m.group(1))] = x["val"] / 1e6
        return out
    return {
        "revenue": series("Revenue"), "rnd": series("ResearchAndDevelopmentExpense"),
        "op_ifrs": series("ProfitLossFromOperatingActivities"), "cfo": series("CashFlowsFromUsedInOperatingActivities"),
        "cash_taxes": series("IncomeTaxesPaidRefundClassifiedAsOperatingActivities"),
        "goodwill": series("Goodwill", True), "other_intangibles": series("IntangibleAssetsOtherThanGoodwill", True),
        "total_assets": series("Assets", True), "equity": series("EquityAttributableToOwnersOfParent", True),
        "cash": series("CashAndCashEquivalents", True), "contract_liab": series("CurrentContractLiabilities", True),
        "receivables": series("TradeAndOtherCurrentReceivables", True),
    }


def _yf_sap() -> dict:
    import yfinance as yf

    def f():
        t = yf.Ticker("SAP.DE")
        i = t.info
        hist = t.history(period="1mo")
        closes = hist["Close"].dropna()
        bs, cf, inc = t.balance_sheet, t.cashflow, t.income_stmt
        def row(df, name):
            return {int(c.year): (None if pd.isna(v) else float(v) / 1e6) for c, v in df.loc[name].items()} if name in df.index else {}
        re_ = t.revenue_estimate
        ee = t.earnings_estimate
        return {
            # Yahoo can append an unfinished trading day with an empty close: use the last valid close
            "price": float(closes.iloc[-1]) if len(closes) else float(i.get("currentPrice") or i.get("regularMarketPrice")),
            "price_date": str(closes.index[-1].date()) if len(closes) else TODAY,
            "hi52": i.get("fiftyTwoWeekHigh"), "lo52": i.get("fiftyTwoWeekLow"), "shares_out": i.get("sharesOutstanding") / 1e6,
            "tgt_mean": i.get("targetMeanPrice"), "tgt_median": i.get("targetMedianPrice"), "tgt_high": i.get("targetHighPrice"),
            "tgt_low": i.get("targetLowPrice"), "tgt_n": i.get("numberOfAnalystOpinions"), "rec": i.get("recommendationMean"),
            "rev_est": {k: float(v) / 1e6 for k, v in re_["avg"].items()}, "eps_est": {k: float(v) for k, v in ee["avg"].items()},
            "da": row(cf, "Depreciation And Amortization"), "capex": {k: -v for k, v in row(cf, "Capital Expenditure").items() if v is not None},
            "sbc": row(cf, "Stock Based Compensation"), "total_debt": row(bs, "Total Debt"), "nci": row(bs, "Minority Interest"),
            "lt_investments": row(bs, "Investments And Advances"), "ppe": row(bs, "Net PPE"),
            "cash_sti": row(bs, "Cash Cash Equivalents And Short Term Investments"), "payables": row(bs, "Accounts Payable"),
            "gross_profit": row(inc, "Gross Profit"), "net_income": row(inc, "Net Income"),
            "buybacks": {k: -v for k, v in row(cf, "Repurchase Of Capital Stock").items() if v is not None},
            "dividends": {k: -v for k, v in row(cf, "Cash Dividends Paid").items() if v is not None},
            "fetched": TODAY,
        }
    d = _cache_json(f"yf_sap_v2_{TODAY}", f)
    if not (isinstance(d.get("price"), (int, float)) and np.isfinite(d["price"]) and d["price"] > 0):
        raise ValueError("No valid SAP.DE price from Yahoo Finance right now; try again in a minute")
    return d


def beta_regression() -> dict:
    """5y weekly SAP.DE vs DAX, raw and Blume-adjusted."""
    import yfinance as yf

    def f():
        px = yf.download(["SAP.DE", "^GDAXI"], period="5y", interval="1wk", auto_adjust=True, progress=False)["Close"].dropna()
        r = px.pct_change().dropna()
        x, y = r["^GDAXI"].values, r["SAP.DE"].values
        b = np.cov(y, x)[0, 1] / np.var(x, ddof=1)
        resid = y - (y.mean() - b * x.mean()) - b * x
        se = np.sqrt(resid.var(ddof=2) / ((x - x.mean()) ** 2).sum())
        r2 = np.corrcoef(x, y)[0, 1] ** 2
        return {"raw": float(b), "adjusted": float(0.67 * b + 0.33), "se": float(se), "r2": float(r2), "n": int(len(x)),
                "start": str(r.index[0].date()), "end": str(r.index[-1].date()), "fetched": TODAY}
    return _cache_json(f"beta_{TODAY}", f)


def peers() -> pd.DataFrame:
    import yfinance as yf

    def f():
        fx = {}
        for c in ("USD", "GBP", "CAD"):
            fx[c] = float(yf.Ticker(f"EUR{c}=X").history(period="1mo")["Close"].dropna().iloc[-1])
        fx["EUR"] = 1.0
        out = []
        today = pd.Timestamp(TODAY)
        for tk, sym in YF_PEERS.items():
            try:
                t = yf.Ticker(sym)
                i = t.info
                re_, ee = t.revenue_estimate, t.earnings_estimate
                fin_ccy = i.get("financialCurrency") or i.get("currency")
                q_ccy = i.get("currency")
                nfy = pd.Timestamp(i.get("nextFiscalYearEnd", 0), unit="s") if i.get("nextFiscalYearEnd") else None
                lfy = pd.Timestamp(i.get("lastFiscalYearEnd", 0), unit="s") if i.get("lastFiscalYearEnd") else None
                # weight of the current fiscal year (0y) in the next twelve months
                fye_cur = (lfy + pd.DateOffset(years=1)) if lfy is not None else nfy
                w0 = float(np.clip((fye_cur - today).days / 365.0, 0, 1)) if fye_cur is not None else 0.5
                rev0, rev1 = float(re_.loc["0y", "avg"]), float(re_.loc["+1y", "avg"])
                eps0, eps1 = float(ee.loc["0y", "avg"]), float(ee.loc["+1y", "avg"])
                rev_ntm = w0 * rev0 + (1 - w0) * rev1
                eps_ntm = w0 * eps0 + (1 - w0) * eps1
                price = float(i.get("currentPrice") or i.get("regularMarketPrice"))
                if q_ccy == "GBp":
                    price /= 100
                ev_ttm_sales = i.get("enterpriseToRevenue")
                rev_ttm = float(i.get("totalRevenue"))
                ev = ev_ttm_sales * rev_ttm                        # in financial currency
                op_m = i.get("operatingMargins")
                mcap_fin = price * i.get("sharesOutstanding", 0) * (fx.get(fin_ccy, 1) / fx.get("GBP" if q_ccy == "GBp" else q_ccy, 1))
                to_eur = 1 / fx.get(fin_ccy, 1)
                out.append({
                    "name": i.get("longName") or i.get("shortName"), "exchange": sym, "ticker": tk,
                    "mcap": mcap_fin * to_eur / 1e6, "tev": ev * to_eur / 1e6, "rev_ltm": rev_ttm * to_eur / 1e6,
                    "ebitda_ltm": (i.get("ebitda") or 0) * to_eur / 1e6, "ebit_ltm": rev_ttm * (op_m or 0) * to_eur / 1e6,
                    "gross_margin": (i.get("grossMargins") or 0) * 100, "ebit_margin": (op_m or 0) * 100,
                    "rev_growth_ltm": (i.get("revenueGrowth") or 0) * 100,
                    "ev_ebitda_ltm": i.get("enterpriseToEbitda"), "ev_sales_ltm": ev_ttm_sales,
                    "ev_ebit_ltm": (ev_ttm_sales / op_m) if op_m and op_m > 0 else None, "pe_ltm": i.get("trailingPE"),
                    "pe_ntm": (price / eps_ntm) if eps_ntm and eps_ntm > 0 else None,
                    "ev_sales_ntm": ev / rev_ntm if rev_ntm else None,
                    # NTM EBIT proxy: GAAP operating margin (SBC included, like SAP non-IFRS) x NTM revenue
                    "ev_ebit_ntm": (ev / (rev_ntm * op_m)) if op_m and op_m > 0 else None,
                    "ev_ebitda_ntm": (ev / (rev_ntm * (i.get("ebitda") or 0) / rev_ttm)) if i.get("ebitda") else None,
                    "w0": w0, "fin_ccy": fin_ccy,
                })
            except Exception as e:  # noqa: BLE001
                out.append({"name": sym, "exchange": sym, "ticker": tk, "error": str(e)})
        return {"fx": fx, "rows": out, "fetched": TODAY}
    data = _cache_json(f"yf_peers_v2_{TODAY}", f)
    rows = [r for r in data["rows"] if "error" not in r]
    return peer_frame(rows)


# ------------------------------------------------------------------ assemble
def load() -> dict:
    ir = load_sap_ir()
    sec = sec_facts()
    y = _yf_sap()
    b10 = bund_10y()
    dam = damodaran()
    beta = beta_regression()
    fy = ir["fy2025_statement"]
    years = range(2021, 2026)
    hist = {}
    for t in years:
        segs = fy.get(f"fy{t}", {})
        g = lambda d, k=t: (d.get(str(k)) if str(k) in d else d.get(k))  # noqa: E731  (json keys are str)
        hist[t] = {
            "revenue": sec["revenue"].get(t), "cloud": segs.get("cloud"), "licenses": segs.get("licenses"),
            "support": segs.get("support"), "services": segs.get("services"), "cost_of_cloud": segs.get("cost_of_cloud"),
            "op_ifrs": sec["op_ifrs"].get(t), "ebit_ciq": segs.get("op_nonifrs"), "rnd": sec["rnd"].get(t),
            "sbc": g(y["sbc"]), "sbc_cash_settled": 817.0 if t == 2025 else None, "da": g(y["da"]), "capex": g(y["capex"]),
            "cfo": sec["cfo"].get(t), "cash_taxes": sec["cash_taxes"].get(t), "gross_profit": g(y["gross_profit"]),
            "net_income": g(y["net_income"]), "buybacks": g(y["buybacks"]), "dividends": g(y["dividends"]),
            "contract_liab": sec["contract_liab"].get(t), "receivables": sec["receivables"].get(t),
            "goodwill": sec["goodwill"].get(t), "other_intangibles": sec["other_intangibles"].get(t),
            "ppe": g(y["ppe"]), "total_assets": sec["total_assets"].get(t), "equity": sec["equity"].get(t),
            "total_debt": g(y["total_debt"]), "cash_sti": g(y["cash_sti"]), "lt_investments": g(y["lt_investments"]),
            "nci": g(y["nci"]), "backlog": segs.get("current_cloud_backlog"), "tax_rate": None, "restructuring": None,
            "payables": g(y["payables"]),
        }
    # FY2025 cash-settled SBC: SAP cash-flow statement (SEC 6-K) 'Share-based payments' outflow 817
    q3 = ir["q3_2026_statement"]
    rev_est, eps_est = y["rev_est"], y["eps_est"]
    nonifrs_2026 = 11850.0   # outlook €11.8-12.2bn cc less ~2pp FX (SAP Q3-26 statement)
    cons = {
        2025: {"revenue": 36800.0, "ebit": fy["fy2025"]["op_nonifrs"], "fcf": fy["fy2025"]["fcf"], "eps_norm": None},
        2026: {"revenue": rev_est.get("0y"), "ebit": nonifrs_2026, "eps_norm": eps_est.get("0y"), "ebitda": None},
        2027: {"revenue": rev_est.get("+1y"), "ebit": None, "eps_norm": eps_est.get("+1y"), "ebitda": None},
    }
    bs = q3["balance_sheet_30sep2026"]
    lt_ratio = (y["lt_investments"].get("2025") or y["lt_investments"].get(2025)) / 7269.0
    bridge = {
        "date": "2026-09-30", "cash": bs["cash"], "st_investments": bs["other_fin_assets_current"],
        "lt_investments": round(bs["other_fin_assets_noncurrent"] * lt_ratio, 0),
        "debt": bs["fin_liabilities_current"] + bs["fin_liabilities_noncurrent"], "pension": 177.0, "nci": bs["nci"],
        "notes": {
            "cash": "SAP Q3-26 statement",
            "lt_investments": f"Non-current other financial assets 8,300 x {lt_ratio:.3f} (FY25 yfinance 'Investments and advances' / reported OFA 7,269)",
            "debt": "SAP Q3-26 financial liabilities 2,053 + 8,535 (bonds, CP, leases)",
            "pension": "FY25 annual report: DBO 2,358 - plan assets 2,181",
        },
    }
    shares_d = y["shares_out"] + (q3["shares"]["wavg_diluted_q3_m"] - q3["shares"]["wavg_basic_q3_m"])
    market = {"price": y["price"], "price_date": y["price_date"], "price_source": "yfinance SAP.DE close",
              "shares_out": y["shares_out"], "hi52": y["hi52"], "lo52": y["lo52"], "tgt_mean": y["tgt_mean"],
              "tgt_median": y["tgt_median"], "tgt_high": y["tgt_high"], "tgt_low": y["tgt_low"], "tgt_n": y["tgt_n"],
              "rec": f"{y['rec']:.2f}"}
    kd = b10["yield"] / 100 + (dam["a_plus_spread"] or 0.007)
    wacc_in = {
        "rf": (b10["yield"] / 100, f"Deutsche Bundesbank 10y Bund yield {b10['date']} (BBSIS API)"),
        "beta": (beta["adjusted"], f"OLS 5y weekly SAP.DE vs DAX ({beta['start']} to {beta['end']}, n={beta['n']}): raw {beta['raw']:.2f}, Blume-adjusted, SE {beta['se']:.3f}"),
        "erp": (0.050, f"House view 5.0% = low end of IDW/FAUB 5.0-6.5%; Damodaran implied ERP {dam['erp']*100:.2f}% on {dam['erp_date']} (+{dam['us_default_spread']*100:.2f}% default-spread adj. = {(dam['erp']+dam['us_default_spread'])*100:.2f}%)"),
        "kd": (kd, f"Rf {b10['yield']:.2f}% + Damodaran A1/A+ default spread {dam['a_plus_spread']*100:.2f}% (S&P A+ / Moody's A1)"),
        "tax": (0.29, "SAP 2026 outlook non-IFRS tax rate ~29%"),
    }
    rnd_long = dict(sec["rnd"])
    pe = peers()
    ltm_rev = 39387.0
    return {
        "edition": "web",
        "peer_ebit_basis": "gaap",   # NTM EBIT = GAAP operating margin (SBC included) x NTM revenue
        "meta": {"name": "SAP SE", "ticker": "SAP.DE", "currency": "EUR", "data_date": TODAY},
        "market": market, "hist": hist, "rnd_long": rnd_long, "consensus": cons, "ir": ir,
        "hist_ltm": {"revenue": ltm_rev, "ebit": 11470.0, "ebit_margin": 29.12,
                     "source": "FY25 36,800 - 9M25 27,116 + 9M26 29,703 (SAP statements); EBIT LTM non-IFRS 10,419-7,590+8,639"},
        "bridge": bridge, "shares_diluted": shares_d, "wacc_inputs": wacc_in, "peers": pe,
        "web_meta": {"bund": b10, "damodaran": dam, "beta": beta},
        "reference": {"damodaran_erp": dam["erp"] + dam["us_default_spread"]},
        "sources": [
            ("Segment revenue FY24-25, 9M-26, backlog, balance sheet 30 Sep 2026", "SAP Quarterly Statements Q4-2025 and Q3-2026 (SEC Form 6-K)"),
            ("Multi-year history (revenue, R&D, IFRS operating profit, goodwill, equity, cash taxes)", "SEC EDGAR XBRL company facts, CIK 1000184"),
            ("D&A, capex, SBC, debt, investments", "yfinance (Yahoo Finance) annual statements"),
            ("Price, 52-week range, analyst targets, revenue/EPS estimates", f"yfinance, fetched {TODAY}"),
            ("Risk-free rate", f"Deutsche Bundesbank 10y Bund {b10['date']}: {b10['yield']:.2f}%"),
            ("Beta", f"Own OLS regression, 5y weekly vs DAX (yfinance prices): adjusted {beta['adjusted']:.2f}"),
            ("ERP / credit spread", f"House ERP 5.0%; Damodaran implied {dam['erp']*100:.2f}% ({dam['erp_date']}); A1/A+ spread {dam['a_plus_spread']*100:.2f}%"),
            ("Peers", "Same 30-name universe as the Capital IQ screen + Intuit; yfinance multiples; NTM = fiscal-year-weighted consensus"),
        ],
    }
