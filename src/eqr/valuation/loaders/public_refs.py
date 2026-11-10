"""Free official reference data used as cross-checks in both editions (no API key needed).

- ECB Data Portal: euro-area AAA government yield curve (10y spot), Survey of Professional Forecasters
  (long-term inflation and real GDP growth), euro foreign-exchange reference rates.
- Damodaran (NYU Stern): industry averages for 'Software (System & Application)': unlevered beta and EV/EBITDA.

Order of use: today's download -> the most recent earlier download -> the snapshot shipped next to this file."""
from __future__ import annotations

import datetime as dt
import io
import json
import os
import re
from pathlib import Path

import pandas as pd
import requests

CACHE = Path(os.environ.get("VALUATION_CACHE", Path(__file__).resolve().parents[2] / "data" / "web"))
SNAPSHOT = Path(__file__).with_name("public_refs.json")
UA = {"User-Agent": "valuation-coursework research-bot"}
ECB = "https://data-api.ecb.europa.eu/service/data/"
DAMODARAN = "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/"
INDUSTRY = "Software (System & Application)"


def cached(name: str, fn, cache: Path | None = None, valid=None):
    """Daily cache with a 'last good download' fallback: `name` must end in _<ISO date>."""
    cache = cache or CACHE
    cache.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r"[^A-Za-z0-9_.-]", "_", name)
    p = cache / f"{safe}.json"
    if p.exists():
        return json.loads(p.read_text())
    try:
        d = fn()
        if valid is not None and not valid(d):
            raise ValueError(f"{name}: download failed validation")
        p.write_text(json.dumps(d, default=str, indent=1))
        return d
    except Exception:
        prefix = re.sub(r"\d{4}-\d{2}-\d{2}$", "", safe)
        older = sorted(q for q in cache.glob(f"{prefix}*.json") if re.fullmatch(re.escape(prefix) + r"\d{4}-\d{2}-\d{2}", q.stem))
        if not older:
            raise
        d = json.loads(older[-1].read_text())
        if isinstance(d, dict):
            d["stale_from"] = older[-1].stem[len(prefix):]
        return d


def _ecb(series: str, n: int = 1) -> pd.DataFrame:
    r = requests.get(f"{ECB}{series}?lastNObservations={n}&format=csvdata", headers=UA, timeout=45)
    r.raise_for_status()
    return pd.read_csv(io.StringIO(r.text))


def ecb_fx() -> dict:
    """Units of foreign currency per 1 EUR (ECB reference rates)."""
    d = _ecb("EXR/D.USD+GBP+CAD+CHF+JPY+SEK+AUD+HKD+INR+KRW+CNY.EUR.SP00.A")
    return {"date": str(d["TIME_PERIOD"].max()), "rates": {r["CURRENCY"]: float(r["OBS_VALUE"]) for _, r in d.iterrows()}}


def _fetch() -> dict:
    out = {"fetched": dt.date.today().isoformat()}
    y = _ecb("YC/B.U2.EUR.4F.G_N_A.SV_C_YM.SR_10Y")
    out["ecb_aaa_10y"] = {"date": str(y["TIME_PERIOD"].iloc[-1]), "yield": float(y["OBS_VALUE"].iloc[-1]) / 100,
                          "source": "ECB euro-area AAA government yield curve, 10y spot rate"}
    inf, gdp = _ecb("SPF/Q.U2.HICP.POINT.LT.Q.AVG"), _ecb("SPF/Q.U2.RGDP.POINT.LT.Q.AVG")
    out["ecb_spf"] = {"round": str(inf["TIME_PERIOD"].iloc[-1]), "inflation_lt": float(inf["OBS_VALUE"].iloc[-1]) / 100,
                      "real_gdp_lt": float(gdp["OBS_VALUE"].iloc[-1]) / 100,
                      "source": "ECB Survey of Professional Forecasters, long-term expectations (euro area)"}
    norm = lambda s: re.sub(r"\s+", " ", str(s)).strip()  # noqa: E731
    pct = lambda s: float(str(s).strip("%")) / 100  # noqa: E731
    b = pd.read_html(io.StringIO(requests.get(DAMODARAN + "Betas.html", headers=UA, timeout=45).text))[0]
    row = b[b.iloc[:, 0].map(norm) == INDUSTRY].iloc[0]
    m = pd.read_html(io.StringIO(requests.get(DAMODARAN + "vebitda.html", headers=UA, timeout=45).text))[0]
    mrow = m[m.iloc[:, 0].map(norm) == INDUSTRY].iloc[0]
    out["damodaran_industry"] = {
        "industry": INDUSTRY, "region": "US", "n_firms": int(row.iloc[1]), "beta": float(row.iloc[2]), "d_e": pct(row.iloc[3]),
        "beta_unlevered": float(row.iloc[5]), "beta_unlevered_cash_adj": float(row.iloc[7]),
        "ev_ebitda": float(mrow.iloc[2]), "ev_ebit": float(mrow.iloc[3]),
        "source": "Damodaran, NYU Stern: industry betas and EV multiples (US, latest annual update)"}
    try:
        out["ecb_fx"] = ecb_fx()
    except Exception:  # noqa: BLE001  (the FX endpoint is slow at times; FX then comes from Yahoo Finance)
        out["ecb_fx"] = None
    return out


_MEMO: dict = {}


def load(refresh_snapshot: bool = False) -> dict | None:
    """Reference data, or None when nothing is reachable and no snapshot exists. One attempt per process."""
    if "d" not in _MEMO or refresh_snapshot:
        _MEMO["d"] = _load(refresh_snapshot)
    return _MEMO["d"]


def _load(refresh_snapshot: bool) -> dict | None:
    try:
        d = cached(f"public_refs_{dt.date.today().isoformat()}", _fetch)
        if refresh_snapshot and not d.get("stale_from"):
            SNAPSHOT.write_text(json.dumps(d, indent=1))
        return d
    except Exception:  # noqa: BLE001
        if SNAPSHOT.exists():
            d = json.loads(SNAPSHOT.read_text()); d["stale_from"] = d.get("fetched")
            return d
        return None


def eur_rate(ccy: str) -> float | None:
    """Units of `ccy` per 1 EUR from the ECB reference rates, if available."""
    if ccy == "EUR":
        return 1.0
    d = load()
    fx = (d or {}).get("ecb_fx") or {}
    return (fx.get("rates") or {}).get(ccy)
