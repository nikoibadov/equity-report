"""Static charts (matplotlib) for the workbook, the summary markdown and the Streamlit app."""
from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402

INK, TEAL, MINT, AMBER, SLATE, CORAL, GRID = "#0E1B2C", "#0F8B8D", "#8FD3C8", "#F4A259", "#5B6B7F", "#E4572E", "#E3E8EF"


def football_items(res: dict) -> list[dict]:
    """Assemble football-field bars from a valuation result bundle."""
    mk, d, sc, cp, sens = res["market"], res["dcf"], res["scenarios"], res["comps"], res["grid_wacc_g"]
    st = cp["stats"]
    items = [
        {"label": "52-week range", "low": mk["lo52"], "high": mk["hi52"], "mid": None, "kind": "market"},
        {"label": f"Analyst targets (n={mk['tgt_n']})", "low": mk["tgt_low"], "high": mk["tgt_high"], "mid": mk["tgt_mean"], "kind": "market"},
        {"label": "DCF: WACC ±0.5pp, g ±0.25pp", "low": res["dcf_band"][0], "high": res["dcf_band"][1], "mid": d["price"], "kind": "dcf"},
        {"label": "DCF: bear / base / bull", "low": sc["Bear"]["price"], "high": sc["Bull"]["price"], "mid": sc["Weighted"]["price"], "kind": "dcf"},
        {"label": "EV / NTM sales (peer IQR)", "low": st["ev_sales_ntm"]["p_q1"], "high": st["ev_sales_ntm"]["p_q3"], "mid": st["ev_sales_ntm"]["p_med"], "kind": "comps"},
        {"label": ("EV / NTM EBIT (peer IQR)" if "ev_ebit_ntm_xsbc" in st else "EV / NTM EBIT, GAAP (peer IQR)"), "low": st["ev_ebit_ntm"]["p_q1"], "high": st["ev_ebit_ntm"]["p_q3"], "mid": st["ev_ebit_ntm"]["p_med"], "kind": "comps"},
        *([{"label": "EV / NTM EBIT ex-SBC (like-for-like)", "low": st["ev_ebit_ntm_xsbc"]["p_q1"], "high": st["ev_ebit_ntm_xsbc"]["p_q3"], "mid": st["ev_ebit_ntm_xsbc"]["p_med"], "kind": "comps"}] if "ev_ebit_ntm_xsbc" in st else []),
        {"label": "P / NTM EPS (peer IQR)", "low": st["pe_ntm"]["p_q1"], "high": st["pe_ntm"]["p_q3"], "mid": st["pe_ntm"]["p_med"], "kind": "comps"},
        {"label": "Growth-adj. EV/Sales (Rule of 40, ±1 SE)", "low": cp["regression"]["p_low"], "high": cp["regression"]["p_high"], "mid": cp["regression"]["p_fit"], "kind": "comps"},
    ]
    return items


def football(res: dict, path, title: str | None = None, items: list[dict] | None = None, currency: str = "€"):
    items = items or football_items(res)
    price = res["market"]["price"]
    colors = {"market": SLATE, "dcf": TEAL, "comps": AMBER}
    fig, ax = plt.subplots(figsize=(11, 5.6), dpi=150)
    for i, it in enumerate(reversed(items)):
        ax.barh(i, it["high"] - it["low"], left=it["low"], height=0.55, color=colors[it["kind"]], alpha=0.9)
        pad = (max(x["high"] for x in items) - min(x["low"] for x in items)) * 0.012
        ax.text(it["low"] - pad, i, f"{currency}{it['low']:.0f}", va="center", ha="right", fontsize=8, color=INK)
        ax.text(it["high"] + pad, i, f"{currency}{it['high']:.0f}", va="center", ha="left", fontsize=8, color=INK)
        if it["mid"] is not None:
            ax.plot([it["mid"]] * 2, [i - 0.3, i + 0.3], color=INK, lw=2)
    ax.set_yticks(range(len(items)))
    ax.set_yticklabels([it["label"] for it in reversed(items)], fontsize=9)
    ax.axvline(price, color=CORAL, ls="--", lw=1.4)
    ax.text(price, -0.75, f" Price {currency}{price:.2f}", color=CORAL, fontsize=9, va="center")
    ax.axvline(res["dcf"]["price"], color=TEAL, ls=":", lw=1.2)
    ax.set_xlabel(f"{currency} per share")
    lo = min(it["low"] for it in items); hi = max(it["high"] for it in items)
    ax.set_xlim(lo - (hi - lo) * 0.12, hi + (hi - lo) * 0.12)
    ax.grid(axis="x", color=GRID); ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    from matplotlib.patches import Patch
    ax.legend(handles=[Patch(color=SLATE, label="Market reference"), Patch(color=TEAL, label="DCF (FCFF)"),
                       Patch(color=AMBER, label="Relative valuation")], loc="lower right", fontsize=8, frameon=False)
    ax.set_title(title or "SAP SE: valuation football field (€/share, 13 Nov 2026)", fontsize=11, color=INK, loc="left")
    fig.tight_layout()
    fig.savefig(path)
    plt.close(fig)
    return path


def revenue_mix(res: dict, path: Path) -> Path:
    hist, f = res["hist"], res["dcf"]["forecast"]
    years = [y for y in (2021, 2022, 2023, 2024, 2025) if hist[y].get("cloud") is not None] + [2026, 2027, 2028, 2029, 2030]
    data = {k: [] for k in ("cloud", "licenses", "support", "services")}
    for y in years:
        src = hist[y] if y <= 2025 else f.loc[y]
        for k in data:
            data[k].append((src[k] or 0) / 1000)
    fig, ax = plt.subplots(figsize=(9, 4.2), dpi=150)
    bottom = [0] * len(years)
    for k, c in zip(data, (TEAL, SLATE, MINT, AMBER)):
        ax.bar([str(y) + ("E" if y > 2025 else "") for y in years], data[k], bottom=bottom, color=c, label=k.title())
        bottom = [b + v for b, v in zip(bottom, data[k])]
    ax2 = ax.twinx()
    share = [data["cloud"][i] / bottom[i] * 100 for i in range(len(years))]
    ax2.plot(range(len(years)), share, color=INK, marker="o", lw=1.5, label="Cloud share (%)")
    ax2.set_ylim(0, 100); ax2.set_ylabel("Cloud share of revenue (%)")
    ax.set_ylabel("Revenue (€bn)")
    ax.legend(loc="upper left", fontsize=8, frameon=False)
    ax.set_title("Cloud transformation: revenue mix 2021A-2030E", fontsize=11, color=INK, loc="left")
    for s in ("top",):
        ax.spines[s].set_visible(False)
    fig.tight_layout(); fig.savefig(path); plt.close(fig)
    return path


def tornado(res: dict, path: Path) -> Path:
    t = res["tornado"].iloc[::-1]
    base = res["dcf"]["price"]
    fig, ax = plt.subplots(figsize=(9, 4.6), dpi=150)
    for i, (_, r) in enumerate(t.iterrows()):
        ax.barh(i, r["high"] - base, left=base, color=TEAL)
        ax.barh(i, r["low"] - base, left=base, color=CORAL)
        ax.text(r["high"] + 2, i, f"{r['high']:.0f}", va="center", fontsize=7)
        ax.text(r["low"] - 2, i, f"{r['low']:.0f}", va="center", ha="right", fontsize=7)
    ax.axvline(base, color=INK, lw=1)
    ax.set_yticks(range(len(t))); ax.set_yticklabels(t["driver"], fontsize=8)
    ax.set_xlabel("DCF value per share (€)")
    ax.set_title(f"DCF sensitivity to single inputs (base €{base:.0f})", fontsize=11, color=INK, loc="left")
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    fig.tight_layout(); fig.savefig(path); plt.close(fig)
    return path
