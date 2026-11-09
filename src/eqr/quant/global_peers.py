"""Peer suggestions from the global universe (data/global_universe.parquet), not only the S&P 500.

Signals (reusing eqr.quant.similarity): business-description TF-IDF, industry/sector match on Yahoo
classifications (consistent for all countries), size closeness and financial-profile closeness."""
from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

from . import similarity

ROOT = Path(__file__).resolve().parents[3]
UNIVERSE = ROOT / "data" / "global_universe.parquet"
WEIGHTS = {"business": 0.45, "classification": 0.20, "size": 0.15, "fundamentals": 0.20}


def load_universe() -> pd.DataFrame:
    if not UNIVERSE.exists():
        return pd.DataFrame(columns=["symbol", "name", "country", "sector", "industry", "summary", "mcap_eur"])
    return pd.read_parquet(UNIVERSE)


def _same_company(a: str, b: str) -> bool:
    norm = lambda s: "".join(ch for ch in str(s).lower() if ch.isalnum())[:10]  # noqa: E731
    return norm(a) == norm(b)


def suggest(target: dict, universe: pd.DataFrame | None = None, n: int = 20, min_mcap_eur: float = 1e9) -> pd.DataFrame:
    """target: dict with symbol, name, summary, sector, industry, mcap_eur, op_margin, gross_margin, rev_growth."""
    u = (universe if universe is not None else load_universe()).copy()
    if u.empty:
        return u
    u = u[(u["symbol"] != target["symbol"]) & (u["mcap_eur"].fillna(0) >= min_mcap_eur)]
    u = u[~u["name"].map(lambda x: _same_company(x, target.get("name", "")))]
    u = pd.concat([u, pd.DataFrame([target])], ignore_index=True).drop_duplicates("symbol", keep="last").set_index("symbol")
    t = target["symbol"]
    sig = {
        "business": similarity.business(u["summary"].fillna("").to_dict(), t),
        "size": similarity.size(u["mcap_eur"].drop(index=t), float(target["mcap_eur"])),
    }
    ind = (u["industry"] == target.get("industry")).astype(float)
    sec = (u["sector"] == target.get("sector")).astype(float) * 0.5
    sig["classification"] = np.maximum(ind, sec).drop(index=t)
    cols = [c for c in ("op_margin", "gross_margin", "rev_growth") if c in u]
    ranks = u[cols].astype(float).rank(pct=True)
    diff = (ranks.drop(index=t) - ranks.loc[t]).abs()
    sig["fundamentals"] = (1 - diff.mean(axis=1)).where(diff.notna().sum(axis=1) >= 2)
    b = similarity.blend(sig, WEIGHTS)
    out = b.join(u[["name", "country", "sector", "industry", "mcap_eur"]])
    out = out[~out["name"].map(lambda x: "".join(c for c in str(x).lower() if c.isalnum())[:12]).duplicated()]  # one listing per company
    out = out.head(n)
    return out.reset_index().rename(columns={"index": "symbol"})[
        ["symbol", "name", "country", "industry", "mcap_eur", "similarity", "why"]]
