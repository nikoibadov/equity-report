"""Assumption set for the SAP DCF. Every input carries a value, a source and a rationale.

The same structure is used by both editions; only the market inputs (WACC, price, bridge)
and the base-year data come from the edition-specific loader."""
from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass, field

EXPLICIT = [2027, 2028, 2029, 2030]          # segment-driven years after the 2026 anchor
FADE = [2031, 2032, 2033, 2034, 2035]        # linear fade to terminal values
YEARS = [2026] + EXPLICIT + FADE


@dataclass
class A:
    value: object
    unit: str = ""
    source: str = ""
    rationale: str = ""


@dataclass
class Assumptions:
    items: dict = field(default_factory=dict)

    def __getitem__(self, k):
        return self.items[k].value

    def set(self, k, v):
        self.items[k].value = v

    def copy(self):
        return deepcopy(self)


def build(ds: dict) -> Assumptions:
    """Assemble the base-case assumptions from a loaded dataset (either edition)."""
    ir = ds["ir"]["q3_2026_statement"]
    fy25 = ds["ir"]["fy2025_statement"]["fy2025"]
    h25 = ds["hist"][2025]
    y9 = ir["ytd9m"]
    y9p = ir["ytd9m_prior"]
    w = ds["wacc_inputs"]
    b = ds["bridge"]
    mk = ds["market"]
    src_hist = "S&P Capital IQ (as reported)" if ds["edition"] == "data" else "SAP Q4-2025 statement (SEC 6-K)"

    it = {
        # ---------- base year (FY2025 actual) ----------
        "base_cloud": A(h25["cloud"], "€m", src_hist, "FY2025 cloud revenue"),
        "base_licenses": A(h25["licenses"], "€m", src_hist, "FY2025 software licences"),
        "base_support": A(h25["support"], "€m", src_hist, "FY2025 software support"),
        "base_services": A(h25["services"], "€m", src_hist, "FY2025 services"),
        # ---------- 2026 anchor: 9M actual + Q4 assumption ----------
        "ytd_cloud": A(y9["cloud"], "€m", "SAP Q3-26 statement", "9M-2026 actual"),
        "ytd_licenses": A(y9["licenses"], "€m", "SAP Q3-26 statement", "9M-2026 actual"),
        "ytd_support": A(y9["support"], "€m", "SAP Q3-26 statement", "9M-2026 actual"),
        "ytd_services": A(y9["services"], "€m", "SAP Q3-26 statement", "9M-2026 actual"),
        "ytdp_cloud": A(y9p["cloud"], "€m", "SAP Q3-26 statement", "9M-2025 comparative (Q4-25 = FY25 - 9M25)"),
        "ytdp_licenses": A(y9p["licenses"], "€m", "SAP Q3-26 statement", "9M-2025 comparative"),
        "ytdp_support": A(y9p["support"], "€m", "SAP Q3-26 statement", "9M-2025 comparative"),
        "ytdp_services": A(y9p["services"], "€m", "SAP Q3-26 statement", "9M-2025 comparative"),
        "q4g_cloud": A(0.24, "%", "Own estimate", "Q3 cloud +26% reported (+25% cc); current cloud backlog +28% gives visibility; guidance 'upper end' of €25.8-26.2bn cc less ~1.5pp FX"),
        "q4g_licenses": A(-0.20, "%", "Own estimate", "9M licences -27%; on-prem licence sales being replaced by RISE/GROW cloud contracts"),
        "q4g_support": A(-0.06, "%", "Own estimate", "Q3 support -5%, 9M -8%; on-prem customers converting to cloud (ECC end of maintenance 2027)"),
        "q4g_services": A(-0.04, "%", "Own estimate", "Q3 and 9M services -4%; partner-led implementations"),
        # ---------- explicit segment growth 2027-2030 ----------
        "g_cloud": A([0.22, 0.20, 0.17, 0.15], "%", "Own estimate vs consensus", "Fade from ~23-25% as Cloud ERP Suite base grows; supported by €24.1bn current cloud backlog (+28%) and ECC-to-S/4HANA migration wave up to 2027/2030"),
        "g_licenses": A([-0.25, -0.25, -0.25, -0.25], "%", "Own estimate", "Structural decline of perpetual licences"),
        "g_support": A([-0.08, -0.09, -0.10, -0.10], "%", "SAP 2027 ambition", "SAP: 'software support decline rate to accelerate' as installed base moves to cloud"),
        "g_services": A([0.00, 0.01, 0.02, 0.02], "%", "Own estimate", "Low-margin, partner-led; stabilises with migration demand"),
        # ---------- profitability ----------
        "m_nonifrs": A([0.293, 0.308, 0.320, 0.330, 0.340], "% rev", "SAP outlook / consensus", "2026: non-IFRS op. profit outlook €11.8-12.2bn cc less ~2pp FX ≈ €11.85bn (consensus 11.9bn); 2027: opex to grow at 80-90% of revenue growth; path tracks consensus EBIT margin 30.5%/31.7%/33.7%/34.1%. SAP non-IFRS INCLUDES share-based compensation"),
        "m_terminal": A(0.350, "% rev", "Own estimate / peers", "Mature subscription software margin; ORCL 34%, ADBE 36%, MSFT 47% LTM EBIT margin; scale benefits on cloud gross margin (~74%)"),
        "acq_pct": A([0.0094, 0.008, 0.007, 0.006, 0.005], "% rev", "SAP outlook", "Acquisition-related charges (mainly amortisation of acquired intangibles) excluded from non-IFRS: 2026 outlook €340-420m"),
        "acq_terminal": A(0.003, "% rev", "Own estimate", "Ongoing bolt-on M&A (WalkMe, LeanIX, Reltio type)"),
        "oneoff_2026": A(49.0, "€m", "SAP outlook", "2026 restructuring (0-20) + Teradata litigation (29)"),
        "tax": A([0.29, 0.295, 0.295, 0.295, 0.295], "%", "SAP outlook / consensus", "2026 non-IFRS ETR ~29%; consensus 29.4-30.4% for 2027-30"),
        "tax_terminal": A(0.30, "%", "German statutory", "Blended statutory rate (Germany ~30%)"),
        "da_pct": A([0.033, 0.032, 0.030, 0.030, 0.029], "% rev", "Consensus", "Consensus D&A €1.34bn/1.45bn/1.46bn/1.67bn/1.80bn"),
        "da_terminal": A(0.022, "% rev", "Own estimate", "Converges towards maintenance capex; hyperscaler-hosted cloud keeps the asset base light"),
        "capex_pct": A([0.020, 0.022, 0.022, 0.021, 0.021], "% rev", "Consensus", "Consensus capex €0.79bn/0.99bn/1.12bn/1.16bn/1.29bn; SAP runs most cloud on hyperscalers (cost of cloud), not own data centres"),
        "capex_terminal": A(0.024, "% rev", "Own estimate", "Slightly above D&A to fund growth at g"),
        "nwc_pct": A(0.05, "% of Δrev", "Own estimate (historicals)", "Subscription billing in advance: contract liabilities ≈18% of revenue vs receivables ≈16% and payables ≈7%; net operating WC ≈ -9% of revenue. 5% of incremental revenue released as cash (conservative vs -9%, other assets e.g. capitalised contract costs grow too)"),
        "sbc_pct": A(0.040, "% rev", "SAP Q3-26 / FY25", "SBC €1.70bn (4.6%) FY25, €1.23bn 9M-26 (4.1%)"),
        "sbc_noncash": A(0.52, "share", src_hist, "FY25: expense 1,695 vs cash-settled 817 → ~52% equity-settled (non-cash)"),
        "sbc_addback": A(0, "0/1", "Methodology", "0 = SBC treated as a real economic cost (base case; buybacks offset dilution). 1 = add back the non-cash part (sensitivity only)"),
        # ---------- terminal & timing ----------
        "g": A(0.0275, "%", "Own estimate", "Long-run nominal growth: ECB 2% inflation + ~0.75% real; below 10y Bund (3.3%); SAP still structurally growing in ERP cloud but mature by 2035"),
        "stub": A(0.25, "years", "Methodology", "Valuation anchored on last balance sheet (30 Sep 2026); Q4-2026 = 25% of FY2026 FCFF"),
        "roll_days": A(44, "days", "Methodology", "Equity value rolled forward from 30 Sep to 13 Nov 2026 at cost of equity"),
        # ---------- WACC ----------
        "rf": A(w["rf"][0], "%", w["rf"][1], "Risk-free: 10-year German government yield (EUR cash flows, AAA)"),
        "beta": A(w["beta"][0], "x", w["beta"][1], "Levered equity beta"),
        "erp": A(w["erp"][0], "%", w["erp"][1], "Mature-market equity risk premium; no country premium (Germany AAA, global revenue)"),
        "kd": A(w["kd"][0], "%", w["kd"][1], "Pre-tax cost of debt = market yield on SAP's longest bond"),
        "tax_wacc": A(w["tax"][0], "%", w["tax"][1], "Tax shield on interest"),
        "e_mv": A(mk["price"] * mk["shares_out"], "€m", mk["price_source"], "Market value of equity = price x shares outstanding"),
        "d_mv": A(b["debt"], "€m", "Q3-26 balance sheet", "Book value of financial debt incl. leases (bonds trade near par: mostly 92-100%)"),
        # ---------- bridge ----------
        "cash": A(b["cash"], "€m", "SAP Q3-26 statement", b.get("notes", {}).get("cash", "")),
        "st_inv": A(b["st_investments"], "€m", "SAP Q3-26 statement", "Current other financial assets (time deposits, debt securities)"),
        "lt_inv": A(b["lt_investments"], "€m", "SAP Q3-26 / CIQ", b.get("notes", {}).get("lt_investments", "")),
        "debt": A(b["debt"], "€m", "SAP Q3-26 / CIQ", b.get("notes", {}).get("debt", "")),
        "pension": A(b["pension"], "€m", "CIQ Pension Details", b.get("notes", {}).get("pension", "")),
        "nci": A(b["nci"], "€m", "SAP Q3-26 statement", "Non-controlling interests at book value"),
        "shares": A(ds["shares_diluted"], "m", "CIQ / SAP Q3-26", "Shares outstanding (1,154.2m, net of 74.8m treasury) + 7m dilution from share-based awards (Q3 diluted 1,157m - basic 1,150m)"),
        "price": A(mk["price"], "€", mk["price_source"], f"Reference market price ({mk['price_date']})"),
    }
    return Assumptions(it)


SCENARIOS = {
    "Bear": {"d_cloud": -0.03, "d_margin": -0.02, "m_terminal": 0.32, "g": 0.0225, "d_wacc": 0.005,
             "story": "Cloud growth fades 3pp faster (AI-native competition, slower ECC migration), margins 2pp lower, 32% terminal margin, g 2.25%, WACC +50bp"},
    "Base": {"d_cloud": 0.0, "d_margin": 0.0, "m_terminal": None, "g": None, "d_wacc": 0.0,
             "story": "Assumptions as set out on the Assumptions sheet"},
    "Bull": {"d_cloud": 0.02, "d_margin": 0.015, "m_terminal": 0.37, "g": 0.03, "d_wacc": -0.0025,
             "story": "Business AI (Joule) and Business Data Cloud upsell keep cloud growth 2pp higher, margins +1.5pp, 37% terminal margin, g 3.0%, WACC -25bp"},
}
SCENARIO_WEIGHTS = {"Bear": 0.25, "Base": 0.50, "Bull": 0.25}
