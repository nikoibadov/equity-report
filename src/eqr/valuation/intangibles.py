"""Intangibles lens: how much of SAP's value sits off balance sheet, and how expensing R&D
distorts margins, invested capital and ROIC (Damodaran-style R&D capitalisation)."""
from __future__ import annotations

import pandas as pd

LIFE = 5   # amortisable life of R&D (years), enterprise software


NEED = ("op_ifrs", "equity", "total_debt", "nci", "cash_sti", "lt_investments", "goodwill", "other_intangibles", "ppe", "total_assets", "revenue")


def rd_capitalisation(ds: dict, a, years=range(2021, 2026)) -> pd.DataFrame:
    rnd = {int(k): v for k, v in ds["rnd_long"].items()}
    h = ds["hist"]
    rows = {}
    years = [t for t in years if t in h and all(h[t].get(k) is not None for k in NEED)
             and all((t - j) in rnd for j in range(LIFE + 1))]
    for t in years:
        asset = sum(rnd.get(t - j, 0) * (LIFE - j) / LIFE for j in range(LIFE))
        amort = sum(rnd.get(t - j, 0) / LIFE for j in range(1, LIFE + 1))
        ebit = h[t]["op_ifrs"]
        tax = 0.29
        ic = (h[t]["equity"] + h[t]["total_debt"] + h[t]["nci"] - h[t]["cash_sti"] - h[t]["lt_investments"])
        ic_xgw = ic - h[t]["goodwill"] - h[t]["other_intangibles"]
        adj_ebit = ebit + rnd[t] - amort
        rows[t] = {
            "revenue": h[t]["revenue"], "rnd": rnd[t], "rnd_pct_rev": rnd[t] / h[t]["revenue"],
            "rnd_asset": asset, "rnd_amort": amort,
            "ebit_reported": ebit, "ebit_margin_reported": ebit / h[t]["revenue"],
            "ebit_adjusted": adj_ebit, "ebit_margin_adjusted": adj_ebit / h[t]["revenue"],
            "invested_capital": ic, "invested_capital_adj": ic + asset, "ic_ex_goodwill": ic_xgw,
            "roic_reported": ebit * (1 - tax) / ic, "roic_adjusted": adj_ebit * (1 - tax) / (ic + asset),
            "roic_ex_goodwill": ebit * (1 - tax) / ic_xgw if ic_xgw > 0 else None,
            "goodwill": h[t]["goodwill"], "other_intangibles": h[t]["other_intangibles"], "ppe": h[t]["ppe"],
            "total_assets": h[t]["total_assets"],
            "intangibles_share_assets": (h[t]["goodwill"] + h[t]["other_intangibles"]) / h[t]["total_assets"],
            "ppe_share_assets": h[t]["ppe"] / h[t]["total_assets"],
        }
    return pd.DataFrame(rows).T


def market_vs_book(ds: dict, a) -> dict:
    h = ds["hist"][2025]
    mcap = a["e_mv"]
    book = h["equity"]
    tangible = book - h["goodwill"] - h["other_intangibles"]
    rd = rd_capitalisation(ds, a).loc[2025, "rnd_asset"]
    return {
        "market_cap": mcap, "book_equity": book, "tangible_book": tangible,
        "goodwill_intangibles": h["goodwill"] + h["other_intangibles"],
        "rd_asset_unrecognised": rd,
        "unexplained_by_balance_sheet": mcap - book - rd,
        "p_b": mcap / book, "p_tb": mcap / tangible,
        "share_of_mcap_off_balance_sheet": (mcap - book) / mcap,
    }
