"""Data edition loader: S&P Capital IQ exports + Bloomberg screenshots (transcribed)
+ SAP's own IR releases for the few gaps. Everything is normalised to EUR millions."""
from __future__ import annotations

import os
import re
from pathlib import Path

import openpyxl
import yaml

from . import public_refs
from .common import load_sap_ir, peer_frame

def _find_root() -> Path:
    """Folder with the licensed CIQ/Bloomberg exports: $SAP_DATA_DIR or the nearest 'Data Assigment 2' up the tree."""
    if os.environ.get("SAP_DATA_DIR"):
        return Path(os.environ["SAP_DATA_DIR"])
    for p in Path(__file__).resolve().parents:
        if (p / "Data Assigment 2").exists():
            return p / "Data Assigment 2"
    return Path(__file__).resolve().parents[3] / "Data Assigment 2"


ROOT = _find_root()
CIQ = ROOT / "S&P Capital IQ" / "SAP"
REPORT = CIQ / "SAPSEXTRASAP_Report_11-05-2026 (1).xlsx"
CONSENSUS = CIQ / "SPGlobal_SAPSE_Consensus_05-Nov-2026.xlsx"
QTR = CIQ / "SPGlobal_SAPSE_FinancialHighlights_05-Nov-2026.xlsx"
PEERS = CIQ / "Peer 30.xlsx"
BBG = Path(__file__).with_name("bloomberg.yaml")


def available() -> bool:
    """Data edition needs the licensed CIQ exports and the (unpublished) Bloomberg transcription."""
    return REPORT.exists() and BBG.exists()


def _rows(path: Path, sheet: str | None = None):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    return [list(r) for r in ws.iter_rows(values_only=True)]


def _num(v):
    """CIQ cells: numbers, 'NA', or strings like '1,234.5 E' / '(3.2) A'."""
    if v is None or isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    if s in ("NA", "NM", "", "-"):
        return None
    m = re.match(r"^\(?([-\d,\.]+)\)?\s*[AE]?$", s)
    if not m:
        return None
    x = float(m.group(1).replace(",", ""))
    return -x if s.startswith("(") else x


def _table(rows, header_row: int, label_col=0):
    """Map label -> {period_label: value} for a CIQ statement sheet (1-based header row)."""
    hdr = rows[header_row - 1]
    out = {}
    for r in rows[header_row:]:
        lab = r[label_col]
        if lab is None:
            continue
        lab = str(lab).strip()
        out.setdefault(lab, {h: r[i] for i, h in enumerate(hdr) if i != label_col and h})
    return out


def _fy(table, label, scale=1e-3):
    """Return {year:int -> value in €m} from a '2025 FY' style table (values in €000)."""
    vals = table.get(label, {})
    res = {}
    for h, v in vals.items():
        m = re.match(r"(\d{4})", str(h))
        if m and _num(v) is not None:
            res[int(m.group(1))] = _num(v) * scale
    return res


def load() -> dict:
    bbg = yaml.safe_load(BBG.read_text())
    ir = load_sap_ir()

    rows_is = _rows(REPORT, "Income Statement")
    rows_cf = _rows(REPORT, "Cash Flow")
    rows_bs = _rows(REPORT, "Balance Sheet")
    rows_isr = _rows(REPORT, "Income Statement (As Reported)")
    rows_cfr = _rows(REPORT, "Cash Flow (As Reported)")
    IS, CF, BS = _table(rows_is, 13), _table(rows_cf, 13), _table(rows_bs, 13)

    # As-reported sheets: header row 13 holds period-end datetimes; values in €m.
    def _asrep(rows, label):
        hdr = rows[12]
        for r in rows[13:]:
            if r[0] and str(r[0]).strip() == label:
                return {hdr[i].year: _num(r[i]) for i in range(1, len(r)) if hdr[i] is not None and _num(r[i]) is not None}
        return {}

    years = range(2021, 2026)
    hist = {}
    rev = _fy(IS, "Total Revenue")
    for y in years:
        hist[y] = {
            "revenue": rev.get(y),
            "gross_profit": _fy(IS, "Gross Profit").get(y),
            "rnd": _fy(IS, "R & D Exp.").get(y),
            "sga": _fy(IS, "Selling General & Admin Exp.").get(y),
            "ebit_ciq": _fy(IS, "EBIT").get(y),                       # ~ SAP non-IFRS basis
            "ebitda_ciq": _fy(IS, "EBITDA").get(y),
            "restructuring": -(_fy(IS, "Restructuring Charges").get(y) or 0),
            "sbc": _fy(IS, "Stock-Based Comp., Total").get(y),
            "net_income": _fy(IS, "Net Income").get(y),
            "tax_rate": (_num(IS["Effective Tax Rate (%)"].get(f"{y} FY")) or 0) / 100,
            "cfo": _fy(CF, "Cash from Ops.").get(y),
            "capex": -(_fy(CF, "Capital Expenditure").get(y) or 0),
            "cash_taxes": _fy(CF, "Cash Taxes Paid").get(y),
            "buybacks": -(_fy(CF, "Repurchase of Common Stock").get(y) or 0),
            "dividends": -(_fy(CF, "Common Dividends Paid").get(y) or 0),
            "chg_unearned": _fy(CF, "Change in Unearned Rev.").get(y),
            "cash_sti": _fy(BS, "Cash & Short-term Investments").get(y),
            "lt_investments": _fy(BS, "Long-term Investments").get(y),
            "goodwill": _fy(BS, "Goodwill").get(y),
            "other_intangibles": _fy(BS, "Other Intangibles").get(y),
            "ppe": _fy(BS, "Net Property, Plant & Equipment").get(y),
            "total_assets": _fy(BS, "Total Assets").get(y),
            "receivables": _fy(BS, "Accounts Receivable").get(y),
            "payables": _fy(BS, "Accounts Payable").get(y),
            "contract_liab": (_fy(BS, "Unearned Revenue, Current").get(y) or 0) + (_fy(BS, "Unearned Revenue, Non-Current").get(y) or 0),
            "total_debt": _fy(BS, "Total Debt").get(y),
            "equity": _fy(BS, "Total Common Equity").get(y),
            "nci": _fy(BS, "Total Minority Interest").get(y),
            "backlog": _fy(BS, "Order Backlog").get(y),
            "employees": (_fy(BS, "Full Time Employees (actual)", 1).get(y)),
        }
    # As-reported lines (€m)
    seg = {
        "cloud": _asrep(rows_isr, "Cloud"),
        "licenses": _asrep(rows_isr, "Software"),
        "support": _asrep(rows_isr, "Support Revenue"),
        "services": _asrep(rows_isr, "Services"),
        "cost_of_cloud": {k: -v for k, v in _asrep(rows_isr, "Cost of Cloud").items()},
        "op_ifrs": _asrep(rows_isr, "Operating Income (Loss)"),
        "rnd": {k: -v for k, v in _asrep(rows_isr, "Research and Development").items()},
        "cost_of_services": {k: -v for k, v in _asrep(rows_isr, "Cost of Services").items()},
        "cost_of_lic_support": {k: -v for k, v in _asrep(rows_isr, "Cost of Software Licenses and Support").items()},
    }
    da_rep = _asrep(rows_cfr, "Depreciation and Amortization")
    sbc_cash = _asrep(rows_cfr, "Share Based Payments")
    lease_pay = _asrep(rows_cfr, "Payment of Lease Liabilities")
    for y in years:
        h = hist[y]
        for k in ("cloud", "licenses", "support", "services", "cost_of_cloud", "op_ifrs", "cost_of_services", "cost_of_lic_support"):
            h[k] = seg[k].get(y)
        h["da"] = da_rep.get(y)
        h["sbc_cash_settled"] = -(sbc_cash.get(y) or 0)
        h["lease_payments"] = -(lease_pay.get(y) or 0)
    # R&D history for the intangibles side analysis (2016-2025, CIQ standard, €m)
    rnd_long = _fy(IS, "R & D Exp.")   # CIQ standard R&D (2016-2025)

    # ---------- consensus (CIQ, €m) ----------
    crow = _rows(CONSENSUS)
    hdr = crow[17]
    fy_cols = {int(re.search(r"FY (\d{4})", str(h)).group(1)): i for i, h in enumerate(hdr) if h and str(h).startswith("FY ")}
    def cmetric(label, section_start=222):
        for r in crow[section_start - 1:]:
            if r[0] and str(r[0]).strip() == label:
                return {y: _num(r[i]) for y, i in fy_cols.items()}
        return {}
    def cmetric_ps(label):
        for r in crow[18:91]:
            if r[0] and str(r[0]).strip() == label:
                return {y: _num(r[i]) for y, i in fy_cols.items()}
        return {}
    cons = {}
    m = {k: cmetric(k) for k in ("Revenue", "EBITDA", "EBIT", "Free Cash Flow", "Capital Expenditures",
                                  "Depreciation & Amortization", "Effective Tax Rate (%)", "Net Income Normalized",
                                  "Research & Development", "Cash From Operations", "Gross Margin (%)")}
    ps = {k: cmetric_ps(k) for k in ("EPS Normalized", "EPS (GAAP)", "Dividend per share", "Avg. Diluted Shares Outstanding")}
    for y in range(2025, 2032):
        cons[y] = {
            "revenue": m["Revenue"].get(y), "ebitda": m["EBITDA"].get(y), "ebit": m["EBIT"].get(y),
            "fcf": m["Free Cash Flow"].get(y), "capex": m["Capital Expenditures"].get(y),
            "da": m["Depreciation & Amortization"].get(y), "tax_rate": m["Effective Tax Rate (%)"].get(y),
            "net_income": m["Net Income Normalized"].get(y), "rnd": m["Research & Development"].get(y),
            "eps_norm": ps["EPS Normalized"].get(y), "eps_gaap": ps["EPS (GAAP)"].get(y),
            "dps": ps["Dividend per share"].get(y), "shares": ps["Avg. Diluted Shares Outstanding"].get(y),
        }

    # ---------- market / sentiment ----------
    sent = _rows(REPORT, "Sentiment Analysis")
    def srow(lbl):
        for r in sent:
            if r[0] == lbl:
                return r
    hi, lo = [float(x) for x in str(sent[28][4]).split("/")]
    ks = _rows(REPORT, "Key Stats")
    kv = {str(r[0]).strip(): r[1] for r in ks if r[0]}
    market = {
        "price": float(kv["Closing Price (€)"]), "price_date": "2026-11-04", "price_source": "S&P Capital IQ Key Stats",
        "shares_out": kv["Common Shares Outstanding (actual)"] / 1e6,
        "hi52": hi, "lo52": lo,
        "tgt_mean": srow("Mean")[1], "tgt_median": srow("Median")[1], "tgt_high": srow("High")[1],
        "tgt_low": srow("Low")[1], "tgt_n": 27, "rec": "1.50 (Buy)",
        "bbg_price": bbg["price"]["last"],
    }

    # ---------- bridge (30 Sep 2026) ----------
    bs = ir["q3_2026_statement"]["balance_sheet_30sep2026"]
    lt_ratio = hist[2025]["lt_investments"] / 7269.0     # CIQ LT investments / reported non-current OFA (FY25)
    bridge = {
        "date": "2026-09-30",
        "cash": bs["cash"],
        "st_investments": bs["other_fin_assets_current"],
        "lt_investments": round(bs["other_fin_assets_noncurrent"] * lt_ratio, 0),
        "debt": None,
        "pension": 177.0,
        "nci": bs["nci"],
        "notes": {
            "cash": "Q3-26 statement; equals CIQ cash & ST inv. 10,997 incl. current OFA",
            "lt_investments": f"Non-current other financial assets 8,300 x {lt_ratio:.3f} (FY25 CIQ LT investments / reported OFA) - equity stakes (Sapphire Ventures) and debt investments; excluded from EBIT, so added as non-operating",
            "debt": "CIQ total debt Q3-26 = reported financial liabilities 2,053+8,535; includes bonds 8.05bn, CP 0.49bn, lease liabilities ~1.7bn",
            "pension": "Unfunded DBO FY25 (PBO 2,358 - plan assets 2,181), CIQ Pension Details",
            "nci": "Q3-26 statement",
        },
    }
    debt_val = None
    for r in ks:
        if r[0] and "Total Debt" in str(r[0]):
            debt_val = r[1]
    bridge["debt"] = float(debt_val) / 1e3

    shares_diluted = market["shares_out"] + (ir["q3_2026_statement"]["shares"]["wavg_diluted_q3_m"] - ir["q3_2026_statement"]["shares"]["wavg_basic_q3_m"])

    # ---------- WACC inputs ----------
    wacc_in = {
        "rf": (bbg["bund_curve"]["y10"] / 100, "Bloomberg GC: 10y German Bund BVAL bid yield, 5 Nov 2026 (risk-free rate 2.png)"),
        "beta": (bbg["beta"]["y5_adjusted"], "Bloomberg BETA vs DAX, 5y weekly, adjusted (Blume); SE 0.087 vs 0.173 for 2y"),
        "erp": (0.050, "House view 5.0% = low end of IDW/FAUB 5.0-6.5% range for German valuations; above CIQ DCF template 4.23% and Damodaran implied 4.63% (1 Nov 2026, incl. 0.22% default-spread adj.); Bloomberg's 10.32% market premium rejected as implausible"),
        "kd": (bbg["bond"]["sap_3_5_2033_yield"] / 100, "Bloomberg YAS: SAP 3.5% 06/2033 bond yield 3.92% (longest SAP bond, A+/A1)"),
        "tax": (0.30, "Marginal (statutory) rate, Germany ~30% (15% corporate tax + 5.5% solidarity surcharge + ~14% trade tax; KPMG tax-rate table). Lecture: marginal home-country rate preferred over the effective rate (SAP ETR ~29%)"),
    }

    peers = _load_peers()

    return {
        "edition": "data",
        "peer_ebit_basis": "street",  # CIQ forward EBIT = broker 'street' EBIT (US peers: non-GAAP, ex-SBC)
        "meta": {"name": "SAP SE", "ticker": "SAP GY / XTRA:SAP", "currency": "EUR", "data_date": "2026-11-05"},
        "hist_ltm": {"revenue": 39387.0, "ebit": 11470.0, "ebit_margin": 29.12, "source": "CIQ Key Stats 2026 FQ3 LTM"},
        "market": market, "hist": hist, "rnd_long": rnd_long, "consensus": cons, "ir": ir,
        "bridge": bridge, "shares_diluted": shares_diluted, "wacc_inputs": wacc_in, "peers": peers,
        "bbg": bbg,
        "refs": public_refs.load(),      # free official cross-checks (ECB, Damodaran industry data); None if offline
        "reference": {"ciq_dcf_price": 268.72, "ciq_wacc": 0.0749, "bbg_wacc": bbg["wacc_screen"]["wacc"] / 100,
                      "ciq_quick_val": 222.93},
        "sources": [
            ("Historical financials FY21-25", "S&P Capital IQ SAP report export, 5 Nov 2026 (Income Statement, Cash Flow, Balance Sheet, As Reported sheets)"),
            ("Q3-26 / 9M-26 segment revenue, backlog, balance sheet", "SAP Quarterly Statement Q3 2026 (SEC 6-K) - gap fill"),
            ("Consensus FY26-31", "S&P Capital IQ Consensus export, 5 Nov 2026 (mean estimates)"),
            ("Price, targets, 52w range", "S&P Capital IQ Key Stats / Sentiment Analysis, 4-5 Nov 2026"),
            ("Risk-free, beta, Kd, credit", "Bloomberg screens (GC, BETA, YAS, CRPR, DDIS), 5 Nov 2026"),
            ("ERP", "House view 5.0% (IDW/FAUB range 5.0-6.5%); cross-checks CIQ 4.23%, Damodaran 4.63%, Bloomberg 10.32% (rejected)"),
            ("Cross-checks (public)", "ECB Data Portal (AAA yield curve, Survey of Professional Forecasters); Damodaran industry averages for software; peer betas from Yahoo Finance prices"),
            ("Peers (30)", "S&P Capital IQ screen 'Top 30 by Peer Ranking', LTM & forward multiples in EUR"),
        ],
    }


def _load_peers():
    s1 = _rows(PEERS, "Sheet1")
    s2 = _rows(PEERS, "Sheet2")
    s3 = _rows(PEERS, "Sheet3")
    recs = []
    for r1, r2, r3 in zip(s1[7:38], s2[7:38], s3[7:38]):
        if not r1[0]:
            continue
        name = r1[0]
        m = re.search(r"\(([^:]+):([^)]+)\)", name)
        recs.append({
            "name": re.sub(r"\s*\(.*\)$", "", name), "exchange": m.group(1) if m else "", "ticker": m.group(2) if m else name,
            "mcap": _num(r1[4]), "tev": _num(r1[7]),                                   # €m
            "rev_ltm": (_num(r1[10]) or 0) / 1e3, "ebitda_ltm": (_num(r1[11]) or 0) / 1e3, "ebit_ltm": (_num(r1[12]) or 0) / 1e3,
            "gross_margin": _num(r2[2]), "ebit_margin": _num(r2[4]), "rev_growth_ltm": _num(r2[6]),
            "ev_ebitda_ltm": _num(r3[2]), "ev_sales_ltm": _num(r3[3]), "ev_ebit_ltm": _num(r3[4]), "pe_ltm": _num(r3[5]),
            "pe_ntm": _num(r3[7]), "ev_sales_ntm": _num(r3[9]), "ev_ebitda_ntm": _num(r3[10]), "ev_ebit_ntm": _num(r3[11]),
        })
    return peer_frame(recs)
