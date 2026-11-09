"""Company Valuation mode: FCFF DCF + relative valuation with a global peer universe.

SAP.DE: the full segment model from the AIFA Assignment 2 package (Data or Web edition).
Any other ticker: the generic DCF (same conventions) on public data."""
from __future__ import annotations

import io
import os
import tempfile
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

SAP_MODELS = {"sap_data": "SAP segment model · Data edition (Capital IQ + Bloomberg)",
              "sap_web": "SAP segment model · Web edition (public sources)",
              "generic": "Generic DCF (any listed company)"}


def _chart(fn, res, **kw) -> bytes:
    buf = io.BytesIO()
    fn(res, buf, **kw)
    return buf.getvalue()


# ------------------------------------------------------------------ cached computations
@st.cache_resource(show_spinner=False)
def _load_ds(edition: str):
    if edition == "data":
        from eqr.valuation.loaders import capiq
        return capiq.load()
    from eqr.valuation.loaders import web
    return web.load()


@st.cache_data(show_spinner=False)
def _run_sap(edition: str, overrides: tuple, extra_peers: tuple, drop_peers: tuple):
    import copy
    from eqr.valuation import assumptions, pipeline
    from eqr.valuation.generic import peer_row
    ds = copy.deepcopy(_load_ds(edition))
    a = assumptions.build(ds)
    ov = dict(overrides)
    for k in ("rf", "beta", "erp", "kd", "g", "m_terminal", "nwc_pct", "sbc_addback", "lt_inv", "price"):
        if k in ov:
            a.set(k, ov[k])
    if "price" in ov:
        a.set("e_mv", ov["price"] * ds["market"]["shares_out"])
        ds["market"]["price"] = ov["price"]
    if ov.get("d_cloud"):
        a.set("g_cloud", [x + ov["d_cloud"] for x in a["g_cloud"]])
    if ov.get("d_margin"):
        a.set("m_nonifrs", [a["m_nonifrs"][0]] + [x + ov["d_margin"] for x in a["m_nonifrs"][1:]])
    peers = ds["peers"]
    peers.loc[peers["ticker"].isin(drop_peers), "core"] = False
    rows = [r for r in (peer_row(t) for t in extra_peers) if r]
    if rows:
        add = pd.DataFrame(rows).rename(columns={"mcap_eur": "mcap"})
        add["core"], add["group"] = True, "Added from global universe (yfinance, GAAP basis)"
        add["exchange"] = add["ticker"]
        ds["peers"] = pd.concat([peers, add], ignore_index=True)
    return pipeline.run(ds, a)


@st.cache_data(show_spinner=False)
def _run_generic(ticker: str, overrides: tuple, peers: tuple):
    from eqr.valuation import generic
    return generic.run(ticker, dict(overrides), list(peers))


@st.cache_resource(show_spinner=False)
def _universe():
    from eqr.quant import global_peers
    return global_peers.load_universe()


def _suggest(ticker: str) -> pd.DataFrame:
    try:
        return _suggest_cached(ticker)
    except Exception as e:  # noqa: BLE001  (e.g. Yahoo rate limit) - peers are optional
        st.caption(f"Peer suggestions unavailable right now ({type(e).__name__}).")
        return pd.DataFrame(columns=["symbol", "name", "country", "industry", "mcap_eur", "similarity", "why"])


@st.cache_data(show_spinner=False)
def _suggest_cached(ticker: str):
    from eqr.quant import global_peers
    from eqr.valuation import generic
    u = _universe()
    if u.empty:
        return pd.DataFrame()
    row = u[u["symbol"] == ticker]
    if len(row):
        target = row.iloc[0].to_dict()
    elif ticker == "SAP.DE":
        target = {"symbol": "SAP.DE", "name": "SAP SE", "sector": "Technology", "industry": "Software - Application",
                  "summary": ("SAP provides enterprise application software and cloud services: ERP (S/4HANA, Cloud ERP Suite), "
                              "finance, supply chain management, procurement and spend management, human capital management, "
                              "customer experience, business technology platform, database and analytics, business AI, "
                              "subscription software sold to enterprises worldwide with support and implementation services."),
                  "mcap_eur": 2.47e11, "op_margin": 0.29, "gross_margin": 0.735, "rev_growth": 0.13}
    else:
        s = generic.snapshot(ticker)
        i = s["info"]
        p = generic.default_inputs(s)
        target = {"symbol": ticker, "name": i.get("longName") or ticker, "summary": i.get("longBusinessSummary") or "",
                  "sector": i.get("sector"), "industry": i.get("industry"),
                  "mcap_eur": p["price"] * p["shares"] * generic._fx(p["q_ccy"], "EUR"),
                  "op_margin": i.get("operatingMargins"), "gross_margin": i.get("grossMargins"), "rev_growth": i.get("revenueGrowth")}
    return global_peers.suggest(target, u, n=25)


# ------------------------------------------------------------------ view
def render(root: Path):
    os.environ.setdefault("VALUATION_CACHE", str(root / "data" / "valuation_cache"))
    from eqr.valuation.loaders import capiq

    st.title("Company valuation")
    st.caption("FCFF DCF + relative valuation · global peer universe · AIFA Group 5 (academic project; not investment advice)")

    with st.sidebar:
        ticker = st.text_input("Ticker", "SAP.DE", key="val_ticker",
                               help="Yahoo Finance symbol, e.g. SAP.DE, MSFT, DSY.PA, 6758.T").strip().upper()
        opts = []
        if ticker in ("SAP.DE", "SAP"):
            if capiq.available():
                opts.append("sap_data")
            opts.append("sap_web")
        opts.append("generic")
        model = st.radio("Model", opts, format_func=SAP_MODELS.get, key=f"val_model_{ticker}")
        st.caption("The Data edition reads the licensed Capital IQ/Bloomberg exports from the local "
                   "'Data Assigment 2' folder and is only offered when they are present.")
        n_univ = len(_universe())
        st.caption(f"Global peer universe: {n_univ:,} listed companies" if n_univ else
                   "Global peer universe not built yet: run scripts/build_global_universe.py")
        go = st.button("Run valuation", type="primary", icon=":material/calculate:", width="stretch")

    state_key = f"val_ran::{ticker}::{model}"
    if go:
        st.session_state[state_key] = True
    if not st.session_state.get(state_key) and not (ticker == "SAP.DE" and model in ("sap_data", "sap_web")):
        st.info("Pick a ticker and model in the sidebar, then click **Run valuation**.", icon=":material/arrow_back:")
        return
    try:
        if model in ("sap_data", "sap_web"):
            _sap_view(model.split("_")[1], root)
        else:
            _generic_view(ticker)
    except Exception as e:  # noqa: BLE001
        st.error(f"{type(e).__name__}: {e}", title=f"Could not value {ticker}")


# ------------------------------------------------------------------ SAP segment model
def _sap_view(edition: str, root: Path):
    from eqr.valuation import charts
    with st.spinner("Loading data and running the model..."):
        ds = _load_ds(edition)
        base = _run_sap(edition, (), (), ())
    a0 = base["assumptions"]

    with st.expander("Adjust key assumptions", icon=":material/tune:", expanded=False):
        with st.form(f"sap_form_{edition}"):
            c1, c2, c3, c4 = st.columns(4)
            rf = c1.number_input("Risk-free rate %", value=round(a0["rf"] * 100, 2), step=0.05, format="%.2f")
            beta = c1.number_input("Beta", value=round(a0["beta"], 2), step=0.02, format="%.2f")
            erp = c2.number_input("Equity risk premium %", value=round(a0["erp"] * 100, 2), step=0.1, format="%.2f")
            kd = c2.number_input("Pre-tax cost of debt %", value=round(a0["kd"] * 100, 2), step=0.05, format="%.2f")
            g = c3.number_input("Terminal growth %", value=round(a0["g"] * 100, 2), step=0.05, format="%.2f")
            mt = c3.number_input("Terminal non-IFRS margin %", value=round(a0["m_terminal"] * 100, 1), step=0.5, format="%.1f")
            dc = c4.slider("Cloud growth shift 2027-30 (pp)", -6.0, 6.0, 0.0, 0.5)
            dm = c4.slider("Margin shift 2027+ (pp)", -4.0, 4.0, 0.0, 0.25)
            c5, c6, c7 = st.columns(3)
            nwc = c5.number_input("NWC release % of Δrevenue", value=round(a0["nwc_pct"] * 100, 1), step=1.0, format="%.1f")
            sbc = c6.toggle("Add back non-cash SBC", value=bool(a0["sbc_addback"]))
            price = c7.number_input("Reference price €", value=round(float(a0["price"]), 2), step=0.5, format="%.2f",
                                    help="Set the 13 Nov 2026 close here once known")
            core = base["peers"][base["peers"]["core"]]["ticker"].tolist()
            keep = st.multiselect("Core peers", core, default=core)
            sugg = _suggest("SAP.DE")
            add = st.multiselect("Add peers from the global universe (yfinance NTM multiples)",
                                 sugg["symbol"].tolist() if len(sugg) else [], default=[],
                                 format_func=lambda s: f"{s} · {sugg.set_index('symbol').loc[s, 'name']}" if len(sugg) and s in set(sugg['symbol']) else s,
                                 accept_new_options=True)
            st.form_submit_button("Recalculate", type="primary", icon=":material/refresh:")
    ov = {"rf": rf / 100, "beta": beta, "erp": erp / 100, "kd": kd / 100, "g": g / 100, "m_terminal": mt / 100,
          "nwc_pct": nwc / 100, "sbc_addback": int(sbc), "d_cloud": dc / 100, "d_margin": dm / 100}
    if abs(price - round(float(a0["price"]), 2)) > 1e-9:
        ov["price"] = price
    # Compare with the same rounded defaults the widgets show, so an untouched form means "base case".
    defaults = {"rf": round(a0["rf"] * 100, 2) / 100, "beta": round(a0["beta"], 2), "erp": round(a0["erp"] * 100, 2) / 100,
                "kd": round(a0["kd"] * 100, 2) / 100, "g": round(a0["g"] * 100, 2) / 100,
                "m_terminal": round(a0["m_terminal"] * 100, 1) / 100, "nwc_pct": round(a0["nwc_pct"] * 100, 1) / 100,
                "sbc_addback": int(a0["sbc_addback"]), "d_cloud": 0.0, "d_margin": 0.0}
    changed = {k: v for k, v in ov.items() if k == "price" or abs(v - defaults[k]) > 1e-9}
    drop = tuple(sorted(set(core) - set(keep)))
    with st.spinner("Recalculating..."):
        res = _run_sap(edition, tuple(sorted(changed.items())), tuple(add), drop) if (changed or add or drop) else base
    d, sc, mk, a = res["dcf"], res["scenarios"], res["market"], res["assumptions"]

    with st.container(horizontal=True):
        st.metric("DCF value / share", f"€{d['price']:.2f}", f"{d['upside']*100:+.1f}% vs €{a['price']:.2f}", border=True)
        st.metric("Scenario-weighted", f"€{sc['Weighted']['price']:.2f}", border=True)
        st.metric("WACC", f"{d['wacc']*100:.2f}%", f"g {a['g']*100:.2f}%", delta_color="off", border=True)
        st.metric("Analyst target (mean)", f"€{mk['tgt_mean']:.2f}", f"n={mk['tgt_n']}", delta_color="off", border=True)
        st.metric("EV", f"€{d['ev']/1000:.1f}bn", f"TV {d['tv_share']*100:.0f}% of EV", delta_color="off", border=True)
    if changed or add or drop:
        st.caption(f"Base case: €{base['dcf']['price']:.2f}. Your changes: " + ", ".join(
            [f"{k}={v:.4g}" for k, v in changed.items()] + ([f"+{len(add)} peers"] if add else []) + ([f"-{len(drop)} peers"] if drop else [])))

    t1, t2, t3, t4, t5, t6 = st.tabs(["Football field", "DCF & FCFF", "Sensitivity", "Peers", "Business & intangibles", "Download"])
    with t1:
        st.image(_chart(charts.football, res, title=f"SAP SE football field · {edition.upper()} edition (€/share)"))
        st.dataframe(pd.DataFrame(charts.football_items(res))[["label", "low", "mid", "high"]].rename(
            columns={"label": "Method", "low": "Low €", "mid": "Mid €", "high": "High €"}).round(1), hide_index=True)
        scen = pd.DataFrame({n: {"Weight": sc[n]["weight"], "Value €": sc[n]["price"], "WACC": sc[n]["wacc"], "g": sc[n]["g"],
                                 "Terminal margin": sc[n]["m_terminal"], "Story": sc[n]["story"]} for n in ("Bear", "Base", "Bull")}).T
        st.dataframe(scen, column_config={"Value €": st.column_config.NumberColumn(format="€%.0f"),
                                          "Weight": st.column_config.NumberColumn(format="percent"),
                                          "WACC": st.column_config.NumberColumn(format="percent"),
                                          "g": st.column_config.NumberColumn(format="percent"),
                                          "Terminal margin": st.column_config.NumberColumn(format="percent")})
    with t2:
        f = d["forecast"].copy()
        show = f[["cloud", "licenses", "support", "services", "revenue", "growth", "op_nonifrs", "ebit", "nopat", "da", "capex", "dnwc", "fcff"]].T
        show.columns = [f"{c}E" for c in show.columns]
        st.subheader("FCFF build (€m)")
        show.index = ["Cloud", "Software licences", "Software support", "Services", "Revenue", "Growth", "Non-IFRS op. profit",
                      "EBIT (IFRS basis)", "NOPAT", "+ D&A", "- Capex", "- ΔNWC", "FCFF"]
        show = show.apply(pd.to_numeric, errors="coerce")
        st.dataframe(show.style.format(lambda v: f"{v*100:.1f}%" if abs(v) < 1.5 else f"{v:,.0f}", na_rep="–"))
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("WACC")
            st.dataframe(pd.DataFrame({"Value": {
                "Risk-free": f"{a['rf']*100:.2f}%", "Beta": f"{a['beta']:.2f}", "ERP": f"{a['erp']*100:.2f}%",
                "Cost of equity": f"{d['ke']*100:.2f}%", "Kd after tax": f"{d['kd_at']*100:.2f}%",
                "E / V": f"{d['we']*100:.1f}%", "WACC": f"{d['wacc']*100:.2f}%"}}))
        with c2:
            st.subheader("EV to equity (30 Sep 2026)")
            st.dataframe(pd.DataFrame({"€m": {
                "PV of FCFF 2026-35": d["sum_pv"], "PV of terminal value": d["pv_tv"], "Enterprise value": d["ev"],
                "+ cash & financial assets": d["non_op"], "- debt, pensions, NCI": -d["claims"], "Equity value": d["equity"],
                "Value/share 30 Sep (€)": d["ps_anchor"], "Value/share 13 Nov (€)": d["price"]}}).style.format("{:,.2f}"))
        with st.expander("All assumptions with sources and rationale"):
            st.dataframe(pd.DataFrame([{"Key": k, "Value": str(v.value), "Unit": v.unit, "Source": v.source, "Rationale": v.rationale}
                                       for k, v in a.items.items()]), hide_index=True)
    with t3:
        grid = res["grid_wacc_g"].copy()
        grid.index = [f"{x*100:.2f}%" for x in grid.index]
        grid.columns = [f"{x*100:.2f}%" for x in grid.columns]
        st.subheader("Value per share: WACC (rows) vs terminal growth (columns)")
        st.dataframe(grid.style.format("€{:.0f}").background_gradient(cmap="RdYlGn", axis=None))
        eb = res["grid_erp_beta"].copy()
        eb.index = [f"β {x:.2f}" for x in eb.index]
        eb.columns = [f"ERP {x*100:.1f}%" for x in eb.columns]
        st.subheader("Beta vs ERP")
        st.dataframe(eb.style.format("€{:.0f}").background_gradient(cmap="RdYlGn", axis=None))
        st.image(_chart(charts.tornado, res))
    with t4:
        cp = res["comps"]
        stats = pd.DataFrame({v["label"]: {"Peers": v["n"], "Q1": v["q1"], "Median": v["median"], "Q3": v["q3"],
                                            "Price @Q1": v["p_q1"], "Price @median": v["p_med"], "Price @Q3": v["p_q3"]}
                              for v in cp["stats"].values()}).T
        st.subheader("Implied value from peer multiples (NTM)")
        st.dataframe(stats.style.format({"Peers": "{:.0f}", "Q1": "{:.1f}x", "Median": "{:.1f}x", "Q3": "{:.1f}x",
                                         "Price @Q1": "€{:.0f}", "Price @median": "€{:.0f}", "Price @Q3": "€{:.0f}"}))
        r = cp["regression"]
        st.caption(f"Growth-adjusted: EV/NTM sales = {r['intercept']:.2f} + {r['slope']:.3f} × Rule of 40 (R² {r['r2']:.2f}, n={r['n']}); "
                   f"SAP Rule of 40 = {r['sap_rule40']:.1f} → {r['fit']:.2f}x → €{r['p_fit']:.0f} (±1 SE €{r['p_low']:.0f}-€{r['p_high']:.0f}). "
                   f"EBIT basis: {'street (US peers ex-SBC)' if cp.get('ebit_basis') == 'street' else 'GAAP (SBC included)'}.")
        pe = res["peers"]
        cols = ["ticker", "name", "group", "core", "mcap", "ev_sales_ntm", "ev_ebit_ntm", "pe_ntm", "g_ntm", "ebit_margin", "rule40"]
        st.dataframe(pe[[c for c in cols if c in pe]].sort_values(["core", "mcap"], ascending=False), hide_index=True,
                     column_config={"mcap": st.column_config.NumberColumn("Mkt cap €m", format="%,.0f"),
                                    "ev_sales_ntm": st.column_config.NumberColumn("EV/NTM sales", format="%.1fx"),
                                    "ev_ebit_ntm": st.column_config.NumberColumn("EV/NTM EBIT", format="%.1fx"),
                                    "pe_ntm": st.column_config.NumberColumn("P/NTM EPS", format="%.1fx"),
                                    "g_ntm": st.column_config.NumberColumn("NTM growth %", format="%.1f"),
                                    "ebit_margin": st.column_config.NumberColumn("EBIT margin %", format="%.1f"),
                                    "rule40": st.column_config.NumberColumn("Rule of 40", format="%.1f")})
        _peer_finder("SAP.DE")
    with t5:
        st.image(_chart(charts.revenue_mix, res))
        rd = res["rd"].T
        st.subheader("Intangibles: R&D capitalised over 5 years")
        rd = rd.apply(pd.to_numeric, errors="coerce")
        st.dataframe(rd.style.format(lambda v: f"{v*100:.1f}%" if abs(v) < 2 else f"{v:,.0f}", na_rep="–"))
        m = res["mvb"]
        st.markdown(f"- **{m['share_of_mcap_off_balance_sheet']*100:.0f}%** of SAP's market value is not on the balance sheet "
                    f"(P/B {m['p_b']:.1f}x, P/tangible book {m['p_tb']:.1f}x).\n"
                    f"- Capitalising R&D adds an unrecognised asset of **€{m['rd_asset_unrecognised']/1000:.1f}bn**; "
                    f"goodwill + acquired intangibles are €{m['goodwill_intangibles']/1000:.1f}bn.\n"
                    "- Consequence: book-based multiples (P/B, EV/IC) are uninformative; FCFF treats R&D as a cash cost, so the DCF is not affected.")
    with t6:
        _sap_downloads(res, edition)


def _sap_downloads(res, edition):
    from eqr.valuation import charts, excel_writer
    if st.button("Build Excel valuation package", icon=":material/table_view:"):
        with st.spinner("Writing workbook..."):
            with tempfile.TemporaryDirectory() as tmp:
                out = Path(tmp)
                charts.football(res, out / f"football_{edition}.png")
                charts.revenue_mix(res, out / f"revenue_mix_{edition}.png")
                charts.tornado(res, out / f"tornado_{edition}.png")
                x = excel_writer.write(res, out / f"SAP_Valuation_{edition}.xlsx", out)
                st.session_state[f"xlsx_{edition}"] = x.read_bytes()
    if st.session_state.get(f"xlsx_{edition}"):
        st.download_button("Download SAP_Valuation.xlsx", st.session_state[f"xlsx_{edition}"],
                           file_name=f"SAP_Valuation_{edition}.xlsx", icon=":material/download:",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("Formula-driven workbook (Assumptions → WACC → Forecast → DCF), sensitivity grids, comps, intangibles, sources.")


# ------------------------------------------------------------------ peer finder (global universe)
def _peer_finder(ticker: str) -> list[str]:
    st.subheader("Global peer finder")
    sugg = _suggest(ticker)
    if sugg.empty:
        if _universe().empty:
            st.info("Global universe not built. Run `scripts/build_global_universe.py`.")
        else:
            st.info("No suggestions right now (data source busy); try again shortly.")
        return []
    st.caption("Similarity over the global universe (S&P 500, STOXX/European indices, FTSE 100, TSX 60, ASX 50, Hang Seng, "
               "NIFTY 50, Japan/Korea/Taiwan large caps, curated software): business description, industry, size, financial profile.")
    st.dataframe(sugg, hide_index=True, column_config={
        "mcap_eur": st.column_config.NumberColumn("Mkt cap €", format="compact"),
        "similarity": st.column_config.ProgressColumn("Similarity", min_value=0.0, max_value=1.0, format="%.2f")})
    return sugg["symbol"].tolist()


# ------------------------------------------------------------------ generic model
def _generic_view(ticker: str):
    from eqr.valuation import charts, generic
    with st.spinner(f"Loading {ticker}..."):
        s = generic.snapshot(ticker)
        p0 = generic.default_inputs(s)
        sugg = _suggest(ticker)
    ccy = {"EUR": "€", "USD": "$", "GBP": "£", "JPY": "¥"}.get(p0["q_ccy"], p0["q_ccy"] + " ")
    default_peers = sugg["symbol"].head(10).tolist() if len(sugg) else []
    with st.expander("Assumptions and peers", icon=":material/tune:", expanded=False):
        with st.form(f"gen_form_{ticker}"):
            st.caption("Revenue growth, years 1-5 (%)")
            gc = st.columns(5)
            growth = [gc[i].number_input(f"Y{i+1}", value=round(p0["growth"][i] * 100, 1), step=0.5, format="%.1f", key=f"g{i}_{ticker}") for i in range(5)]
            c1, c2, c3, c4 = st.columns(4)
            mt = c1.number_input("Target EBIT margin % (Y5+)", value=round(p0["margin_target"] * 100, 1), step=0.5, format="%.1f")
            tax = c1.number_input("Tax rate %", value=round(p0["tax"] * 100, 1), step=0.5, format="%.1f")
            da = c2.number_input("D&A % revenue", value=round(p0["da_pct"] * 100, 1), step=0.1, format="%.1f")
            cx = c2.number_input("Capex % revenue", value=round(p0["capex_pct"] * 100, 1), step=0.1, format="%.1f")
            nwc = c3.number_input("NWC investment % of Δrevenue", value=round(p0["nwc_pct"] * 100, 1), step=1.0, format="%.1f")
            g = c3.number_input("Terminal growth %", value=round(p0["g"] * 100, 2), step=0.05, format="%.2f")
            rf = c4.number_input("Risk-free %", value=round(p0["rf"] * 100, 2), step=0.05, format="%.2f", help=p0["sources"]["rf"])
            beta = c4.number_input("Beta", value=round(p0["beta"], 2), step=0.02, format="%.2f", help=p0["sources"]["beta"])
            c5, c6 = st.columns(2)
            erp = c5.number_input("ERP %", value=round(p0["erp"] * 100, 2), step=0.1, format="%.2f", help=p0["sources"]["erp"])
            kd = c6.number_input("Pre-tax cost of debt %", value=round(p0["kd"] * 100, 2), step=0.05, format="%.2f", help=p0["sources"]["kd"])
            opts = list(dict.fromkeys((sugg["symbol"].tolist() if len(sugg) else [])))
            name_of = sugg.set_index("symbol")["name"].to_dict() if len(sugg) else {}
            peers = st.multiselect("Peers (suggested from the global universe; type any Yahoo ticker to add)", opts,
                                   default=default_peers, format_func=lambda x: f"{x} · {name_of.get(x, '')}".rstrip(" ·"),
                                   accept_new_options=True)
            st.form_submit_button("Recalculate", type="primary", icon=":material/refresh:")
    ov = {"growth": [x / 100 for x in growth], "margin_target": mt / 100, "tax": tax / 100, "tax_wacc": tax / 100,
          "da_pct": da / 100, "capex_pct": cx / 100, "nwc_pct": nwc / 100, "g": g / 100, "rf": rf / 100, "beta": beta,
          "erp": erp / 100, "kd": kd / 100}
    ov_t = tuple(sorted((k, tuple(v) if isinstance(v, list) else v) for k, v in ov.items()))
    with st.spinner("Valuing and fetching peer multiples..."):
        res = _run_generic(ticker, tuple((k, list(v) if isinstance(v, tuple) else v) for k, v in ov_t), tuple(peers))
    d, sc, mk, p = res["dcf"], res["scenarios"], res["market"], res["inputs"]
    with st.container(horizontal=True):
        st.metric("DCF value / share", f"{ccy}{d['price']:.2f}", f"{d['upside']*100:+.1f}% vs {ccy}{p['price']:.2f}", border=True)
        st.metric("Scenario-weighted", f"{ccy}{sc['Weighted']['price']:.2f}", border=True)
        st.metric("WACC", f"{d['wacc']*100:.2f}%", f"g {p['g']*100:.2f}%", delta_color="off", border=True)
        if mk.get("tgt_mean"):
            st.metric("Analyst target (mean)", f"{ccy}{mk['tgt_mean']:.2f}", f"n={mk['tgt_n']}", delta_color="off", border=True)
        st.metric("TV share of EV", f"{d['tv_share']*100:.0f}%", border=True)
    st.caption(f"{p['name']} · financials in {p['fin_ccy']}, price in {p['q_ccy']} · data {s['price_date']} (yfinance)")
    t1, t2, t3, t4, t5 = st.tabs(["Football field", "DCF & FCFF", "Sensitivity", "Peers", "Download"])
    with t1:
        items = generic.football_items(res)
        st.image(_chart(charts.football, res, items=items, currency=ccy, title=f"{p['name']}: valuation football field ({ccy}/share)"))
        st.dataframe(pd.DataFrame(items)[["label", "low", "mid", "high"]].round(2), hide_index=True)
    with t2:
        f = d["forecast"].T
        f.columns = [f"Y{c}" for c in f.columns]
        st.subheader(f"FCFF build ({p['fin_ccy']} m)")
        st.dataframe((f / np.where(f.index.isin(["growth", "margin"]), 1, 1e6)[:, None]).style.format(
            lambda v: f"{v*100:.1f}%" if abs(v) < 1.5 else f"{v:,.0f}"))
        st.dataframe(pd.DataFrame({"Source": p["sources"]}))
    with t3:
        grid = res["grid_wacc_g"].copy()
        grid.index = [f"{x*100:.2f}%" for x in grid.index]
        grid.columns = [f"{x*100:.2f}%" for x in grid.columns]
        st.dataframe(grid.style.format(lambda v: "" if pd.isna(v) else f"{ccy}{v:.0f}").background_gradient(cmap="RdYlGn", axis=None))
    with t4:
        cp = res["comps"]
        if cp["stats"]:
            st.dataframe(pd.DataFrame({v["label"]: {"Peers": v["n"], "Q1": v["q1"], "Median": v["median"], "Q3": v["q3"],
                                                     f"Price @median ({ccy})": v["p_med"]} for v in cp["stats"].values()}).T.round(2))
        if len(res["peers"]):
            st.dataframe(res["peers"], hide_index=True, column_config={
                "mcap_eur": st.column_config.NumberColumn("Mkt cap €", format="compact"),
                "ev_sales_ntm": st.column_config.NumberColumn("EV/NTM sales", format="%.1fx"),
                "ev_ebit_ntm": st.column_config.NumberColumn("EV/NTM EBIT", format="%.1fx"),
                "pe_ntm": st.column_config.NumberColumn("P/NTM EPS", format="%.1fx"),
                "g_ntm": st.column_config.NumberColumn("NTM growth %", format="%.1f"),
                "ebit_margin": st.column_config.NumberColumn("EBIT margin %", format="%.1f"),
                "rule40": st.column_config.NumberColumn("Rule of 40", format="%.1f")})
        _peer_finder(ticker)
    with t5:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as xw:
            pd.DataFrame({"value": {k: str(v) for k, v in p.items() if k != "sources"}}).to_excel(xw, sheet_name="Inputs")
            d["forecast"].to_excel(xw, sheet_name="Forecast")
            res["grid_wacc_g"].to_excel(xw, sheet_name="WACC x g")
            res["peers"].to_excel(xw, sheet_name="Peers", index=False)
            pd.DataFrame(generic.football_items(res)).to_excel(xw, sheet_name="Football field", index=False)
        st.download_button("Download valuation (xlsx)", buf.getvalue(), file_name=f"{ticker}_valuation.xlsx",
                           icon=":material/download:", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
