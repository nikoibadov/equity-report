"""Build data/global_universe.parquet: a global listed-company universe for peer selection in the
Company Valuation mode (beyond the S&P 500).

Constituents come from Wikipedia index pages, plus a curated list of large software/tech names that sit
outside these indices. Fundamentals come from yfinance; amounts are converted to EUR.

    .venv/bin/python scripts/build_global_universe.py            # full build (~1,300 names, several minutes)
    .venv/bin/python scripts/build_global_universe.py --limit 50  # quick test
"""
from __future__ import annotations

import argparse
import io
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import pandas as pd
import requests

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "global_universe.parquet"
UA = {"User-Agent": "Mozilla/5.0 (company-valuation coursework; index constituents)"}

# page -> (yfinance suffix to append, or None if tickers already carry one; normaliser)
PAGES = {
    "DAX": (None, None), "MDAX": (".DE", None), "CAC_40": (None, None), "Euro_Stoxx_50": (None, None),
    "AEX_index": (None, None), "IBEX_35": (None, None), "FTSE_MIB": (None, None), "OMX_Stockholm_30": (None, None),
    "OMX_Helsinki_25": (None, None), "FTSE_100_Index": (".L", "dash"), "Swiss_Market_Index": (".SW", None),
    "OMX_Copenhagen_25": (".CO", "space"), "S%26P/TSX_60": (".TO", "dash"), "S%26P/ASX_50": (".AX", None),
    "NIFTY_50": (".NS", None), "Hang_Seng_Index": (".HK", "hk"),
}
EXTRA = [
    # Japan / Korea / Taiwan large caps (no machine-readable Nikkei list on Wikipedia)
    "6758.T", "6861.T", "9984.T", "6098.T", "4307.T", "9613.T", "6702.T", "6701.T", "4684.T", "4716.T", "7203.T", "8035.T",
    "6501.T", "6954.T", "9432.T", "9433.T", "4063.T", "6367.T", "7974.T", "8306.T", "4502.T", "6981.T", "3659.T", "4385.T",
    "005930.KS", "000660.KS", "035420.KS", "035720.KS", "2330.TW", "2317.TW", "2454.TW",
    # Software / internet outside the index pages above
    "SNOW", "TEAM", "HUBS", "DOCU", "PCTY", "ESTC", "MDB", "NET", "ZS", "OKTA", "U", "PATH", "GTLB", "APP", "SPOT", "ARM",
    "SHOP.TO", "CSU.TO", "OTEX", "TEMN.SW", "AOF.DE", "NEM.DE", "TMV.DE", "SOP.PA", "CAP.PA", "DSY.PA", "SGE.L", "WISE.L",
    "MONY.L", "SAP.DE", "IFX.DE", "ADYEN.AS", "WKL.AS", "PRX.AS", "NXT.L", "AUTO.L", "REL.L", "EXPN.L", "LSEG.L",
    "CHKP", "CYBR", "MNDY", "WIX", "NICE", "GLOB", "EPAM", "INFY", "WIT", "TCS.NS", "HCLTECH.NS", "BABA", "PDD", "BIDU",
    "0700.HK", "9988.HK", "3690.HK", "1810.HK", "XRO.AX", "WTC.AX", "TNE.AX", "SDR.L", "NOVO-B.CO", "ASML.AS", "TTE.PA",
]


def _norm(t: str, suffix: str | None, mode: str | None) -> str | None:
    t = str(t).replace("\xa0", " ").strip()
    if not t or t.lower() == "nan":
        return None
    if mode == "hk":
        m = re.search(r"(\d+)", t)
        return f"{int(m.group(1)):04d}.HK" if m else None
    if mode == "space":
        t = t.replace(" ", "-")
    if mode == "dash":
        t = t.rstrip(".").replace(".", "-")
    t = t.split()[0]
    return t + suffix if suffix and not t.endswith(suffix) else t


def constituents() -> dict[str, set[str]]:
    out: dict[str, set[str]] = {}
    sp = pd.read_csv(ROOT / "data" / "sp500_constituents.csv")
    for s in sp["Symbol"]:
        out.setdefault(str(s).replace(".", "-"), set()).add("S&P 500")
    for page, (suffix, mode) in PAGES.items():
        try:
            h = requests.get(f"https://en.wikipedia.org/wiki/{page}", headers=UA, timeout=30).text
            tables = pd.read_html(io.StringIO(h))
            def tick_col(t):
                for c in t.columns:
                    cl = str(c).lower()
                    if "ticker" in cl or "symbol" in cl or cl in ("code", "epic"):
                        return c
            cand = [t for t in tables if tick_col(t) is not None]
            t = max(cand, key=len)
            name = page.replace("%26", "&").replace("_", " ")
            for raw in t[tick_col(t)]:
                sym = _norm(raw, suffix, mode)
                if sym:
                    out.setdefault(sym, set()).add(name)
        except Exception as e:  # noqa: BLE001
            print(f"  ! {page}: {e}")
    for s in EXTRA:
        out.setdefault(s, set()).add("Curated")
    return out


_FX: dict[str, float] = {}


def eur_per(ccy: str) -> float:
    import yfinance as yf
    if ccy in (None, "EUR"):
        return 1.0
    if ccy not in _FX:
        _FX[ccy] = 1 / float(yf.Ticker(f"EUR{ccy}=X").history(period="5d")["Close"].iloc[-1])
    return _FX[ccy]


def one(sym: str) -> dict | None:
    import yfinance as yf
    for attempt in range(5):
        try:
            i = yf.Ticker(sym).info
            if not i or i.get("quoteType") != "EQUITY":
                return None
            q = i.get("currency")
            fin = i.get("financialCurrency") or q
            price = i.get("currentPrice") or i.get("regularMarketPrice")
            if not price or not i.get("sharesOutstanding"):
                return None
            q_ccy = "GBP" if q == "GBp" else ("ZAR" if q == "ZAc" else q)
            px = price / 100 if q in ("GBp", "ZAc") else price
            mcap = px * i["sharesOutstanding"] * eur_per(q_ccy)
            rev = i.get("totalRevenue")
            return {
                "symbol": sym, "name": i.get("longName") or i.get("shortName"), "country": i.get("country"),
                "exchange": i.get("exchange"), "currency": q_ccy, "fin_currency": fin,
                "sector": i.get("sector"), "industry": i.get("industry"),
                "summary": (i.get("longBusinessSummary") or "")[:2000],
                "mcap_eur": mcap, "ev_eur": (i.get("enterpriseToRevenue") or 0) * (rev or 0) * eur_per(fin) or None,
                "revenue_eur": (rev or 0) * eur_per(fin) or None,
                "op_margin": i.get("operatingMargins"), "gross_margin": i.get("grossMargins"), "rev_growth": i.get("revenueGrowth"),
                "ev_rev": i.get("enterpriseToRevenue"), "ev_ebitda": i.get("enterpriseToEbitda"),
                "trailing_pe": i.get("trailingPE"), "forward_pe": i.get("forwardPE"), "beta": i.get("beta"),
            }
        except Exception as e:  # noqa: BLE001
            # Yahoo throttles bursts: back off hard on rate limits, briefly otherwise
            time.sleep(30 * (attempt + 1) if "Rate" in type(e).__name__ or "Too Many" in str(e) else 1.5)
    return None


def main(limit: int | None = None, workers: int = 8, resume: bool = False):
    cons = constituents()
    syms = sorted(cons)[:limit] if limit else sorted(cons)
    old = pd.read_parquet(OUT) if resume and OUT.exists() else pd.DataFrame()
    if len(old):
        syms = [s for s in syms if s not in set(old["symbol"])]
        print(f"resume: {len(old)} already built")
    print(f"{len(syms)} symbols from {len(PAGES) + 2} sources")
    rows = []
    with ThreadPoolExecutor(workers) as ex:
        futs = {ex.submit(one, s): s for s in syms}
        for n, f in enumerate(as_completed(futs), 1):
            r = f.result()
            if r:
                r["indices"] = ", ".join(sorted(cons[futs[f]]))
                rows.append(r)
            if n % 100 == 0:
                print(f"  {n}/{len(syms)} done, {len(rows)} usable", flush=True)
    df = pd.concat([old, pd.DataFrame(rows)], ignore_index=True).drop_duplicates("symbol", keep="last")
    df["built"] = time.strftime("%Y-%m-%d")
    OUT.parent.mkdir(exist_ok=True)
    df.to_parquet(OUT, index=False)
    print(f"wrote {OUT} with {len(df)} companies from {df['country'].nunique()} countries")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int)
    ap.add_argument("--workers", type=int, default=4)
    ap.add_argument("--resume", action="store_true", help="only fetch symbols missing from the existing parquet")
    a = ap.parse_args()
    main(a.limit, a.workers, a.resume)
