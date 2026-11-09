import os
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

try:  # Streamlit Cloud: keys live in app Secrets, never in the repo
    for k in ("ANTHROPIC_API_KEY", "SEC_USER_AGENT", "FINNHUB_API_KEY", "FMP_API_KEY"):
        if k in st.secrets:
            os.environ[k] = st.secrets[k]
except Exception:
    pass

# Streamlit only reloads files next to this script. After a code update (git pull on Streamlit Cloud) the old
# eqr modules would stay in memory while templates are read fresh from disk, so drop them when src/ changed.
_stamp = max(p.stat().st_mtime for p in (ROOT / "src" / "eqr").rglob("*.py"))
if "eqr" in sys.modules and getattr(sys.modules["eqr"], "_code_stamp", None) != _stamp:
    for _m in [m for m in sys.modules if m == "eqr" or m.startswith("eqr.")]:
        del sys.modules[_m]

import eqr  # noqa: E402
from eqr import sources  # noqa: E402  (loads .env)

eqr._code_stamp = _stamp

st.set_page_config(page_title="AIFA Group 5 · Equity Scorecard", page_icon=":material/monitoring:", layout="wide")

# Mode switch (top centre): the equity scorecard and the company valuation app share one deployment.
_, _mid, _ = st.columns([1, 2, 1])
with _mid:
    mode = st.segmented_control("Mode", ["Equity Report", "Company Valuation"], default="Equity Report", key="mode",
                                label_visibility="collapsed", width="stretch")
if mode == "Company Valuation":
    import importlib

    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import valuation_view  # noqa: E402

    _vstamp = Path(valuation_view.__file__).stat().st_mtime
    if getattr(valuation_view, "_stamp", None) not in (None, _vstamp):
        valuation_view = importlib.reload(valuation_view)
    valuation_view._stamp = _vstamp
    valuation_view.render(ROOT)
    st.stop()

st.title("Equity scorecard")
st.caption("AIFA Group 5 · Goethe University Frankfurt · detailed stock report modelled on LSEG Stock Reports Plus (academic project)")

status = sources.status()
with st.sidebar:
    ticker = st.text_input("Ticker", "HON", help="Yahoo Finance symbol, e.g. HON, AAPL, SAP.DE").strip().upper()
    live = st.toggle("Use latest live data", value=False,
                     help="Off: report date 15 Oct 2026 with data to the 14 Oct close, comparable with the LSEG benchmark.")
    offline = st.toggle("Offline mode (cached data only)", value=False)

    st.subheader("Data sources")
    usable = [k for k in sources.ANALYST_SOURCES if status[k][0]]
    analyst = st.selectbox("Analyst consensus", usable, format_func=lambda k: sources.ANALYST_SOURCES[k],
                           help="Other sources appear here once their key or login is set up (see README).")
    insider = st.segmented_control("Insider trades", ["sec", "yahoo"], default="sec" if status["sec"][0] else "yahoo",
                                   format_func=lambda k: sources.INSIDER_SOURCES[k])
    with st.expander("Source status", icon=":material/database:"):
        for k, (ok, why) in status.items():
            label = sources.ANALYST_SOURCES.get(k, sources.INSIDER_SOURCES.get(k, k))
            st.markdown(f"{':material/check_circle:' if ok else ':material/radio_button_unchecked:'} **{label}** – {why}")

    go = st.button("Generate report", type="primary", icon=":material/description:", width="stretch")
    ready = sorted(p.stem.replace("_report", "") for p in (ROOT / "output").glob("*_report.pdf"))
    preload = st.selectbox("Open prebuilt PDF", ready, index=None, placeholder="Choose a ticker") if ready else None

os.environ["EQR_AS_OF"] = "live" if live else os.environ.get("EQR_AS_OF_DEFAULT", "2026-10-15")
os.environ["EQR_OFFLINE"] = "1" if offline else "0"
os.environ["EQR_ANALYST_SOURCE"] = analyst or "yahoo"
os.environ["EQR_INSIDER_SOURCE"] = insider or "yahoo"


def show_pdf(data: bytes, name: str):
    st.download_button("Download PDF", data, file_name=name, mime="application/pdf", icon=":material/download:")
    try:
        st.pdf(data, height=900, alt=f"{name} report")
    except Exception:  # streamlit-pdf not installed
        import base64
        st.iframe(f"data:application/pdf;base64,{base64.b64encode(data).decode()}", height=900)


def show(pdf: Path, ctx: dict | None):
    data = pdf.read_bytes()
    if not ctx:
        show_pdf(data, pdf.name)
        return
    with st.container(horizontal=True):
        st.metric("Average score", f"{ctx['avg_num']:.1f}" if ctx["avg_num"] is not None else "NR", border=True)
        st.metric("Outlook", ctx["outlook"], border=True)
        st.metric("Optimized score", ctx["opt_score"], border=True)
        st.metric("Consensus", ctx["consensus"].split(" (")[0], border=True)
        st.metric("Report date", ctx["date"], border=True)
    summary, check, report = st.tabs(["Summary", "Source cross-check", "PDF report"])
    with summary:
        left, right = st.columns([1, 2])
        with left:
            comp = {v["title"]: float(v["score"]) for k, v in ctx["comp"].items() if k != "average" and v["score"] != "NR"}
            st.bar_chart(pd.Series(comp, name="Score"), horizontal=True, color="#0F8B8D")
        with right:
            st.subheader("Highlights")
            for b in ctx["highlights"]:
                st.markdown(f"- {b}")
            notes = ctx.get("page_notes", {}).get("overview", []) + ctx.get("page_notes", {}).get("relative_valuation", [])
            if notes:
                st.warning("\n\n".join(dict.fromkeys(notes)), title="Why some numbers look unusual", icon=":material/info:")
    with check:
        if len(ctx.get("cmp_head", [])) > 1:
            st.dataframe(pd.DataFrame(ctx["cmp_rows"], columns=["Item"] + ctx["cmp_head"]), hide_index=True, alt="Same figures from each data source")
            for c in ctx["cmp_cites"]:
                st.caption(c)
        else:
            st.info("Only one source returned fundamentals for this ticker (SEC EDGAR covers US filers only).")
        st.caption(f"Analyst data: {ctx.get('an_cite', '')}. {ctx.get('an_note', '')}")
        st.caption(f"Insider data: {ctx.get('ins_cite', '')}")
    with report:
        show_pdf(data, pdf.name)


if go and ticker:
    from eqr.pipeline import run_full
    try:
        with st.status(f"Building report for {ticker}", expanded=True) as box:
            pdf, ctx = run_full(ticker, progress=lambda m: st.write(m))
            box.update(label="Report ready", state="complete", expanded=False)
        show(pdf, ctx)
    except Exception as e:
        st.error(f"{type(e).__name__}: {e}", title=f"Could not build {ticker}")
        fallback = ROOT / "output" / f"{ticker}_report.pdf"
        if fallback.exists():
            st.warning("Showing the prebuilt report instead.")
            show(fallback, None)
elif preload:
    show(ROOT / "output" / f"{preload}_report.pdf", None)
else:
    st.info("Enter a ticker and click **Generate report**, or open a prebuilt PDF from the sidebar.", icon=":material/arrow_back:")
