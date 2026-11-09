"""Helpers shared by both editions: SAP IR facts and the peer screen."""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import yaml

IR = Path(__file__).with_name("sap_ir.yaml")


def load_sap_ir() -> dict:
    return yaml.safe_load(IR.read_text())


# Business-model classification for the peer screen. Applied identically in both editions.
# group: why a name is (not) comparable to SAP's ERP-centred, subscription-led model.
PEER_CLASS = {
    # Core peers: enterprise application / platform software with recurring revenue
    "ORCL": ("Core: ERP & database / cloud apps", True),
    "MSFT": ("Core: business apps (Dynamics) & cloud platform", True),
    "CRM": ("Core: enterprise cloud applications (CRM)", True),
    "WDAY": ("Core: cloud HCM & finance ERP", True),
    "NOW": ("Core: enterprise workflow platform", True),
    "SGE": ("Core: SME accounting/ERP", True),
    "DSY": ("Core: enterprise PLM / industrial software", True),
    "ADSK": ("Core: design & engineering software", True),
    "PTC": ("Core: PLM / industrial software", True),
    "ADBE": ("Core: enterprise content & experience software", True),
    "OTEX": ("Core: enterprise information management", True),
    "NEM": ("Core: AEC software (European mid-cap)", True),
    "INTU": ("Core: financial / SME accounting software", True),
    # Excluded
    "ACN": ("Excluded: IT services (labour-based)", False),
    "CAP": ("Excluded: IT services", False),
    "SOP": ("Excluded: IT services", False),
    "IBM": ("Excluded: IT services/hardware mix", False),
    "DELL": ("Excluded: hardware", False),
    "PANW": ("Excluded: cybersecurity, hyper-growth multiple", False),
    "CRWD": ("Excluded: cybersecurity, loss-making", False),
    "NET": ("Excluded: infrastructure, loss-making", False),
    "DDOG": ("Excluded: observability, EBIT near zero", False),
    "ESTC": ("Excluded: search infra, loss-making", False),
    "PLTR": ("Excluded: AI hyper-growth outlier (EV/S >30x)", False),
    "TEAM": ("Excluded: dev tools, negative earnings", False),
    "HUBS": ("Secondary: SMB CRM, EBIT margin <10%", False),
    "DOCU": ("Secondary: point solution (e-signature)", False),
    "PCTY": ("Secondary: US SMB payroll", False),
    "CDNS": ("Excluded: EDA (semiconductor design tools)", False),
    "SNPS": ("Excluded: EDA (semiconductor design tools)", False),
    "TMV": ("Excluded: small-cap remote-access tool", False),
    "SAP": ("Target", False),
}


def peer_frame(recs: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(recs)
    # Derived NTM revenue and growth (both editions): NTM revenue = TEV / (TEV / NTM sales)
    df["rev_ntm"] = df["tev"] / df["ev_sales_ntm"]
    df["g_ntm"] = (df["rev_ntm"] / df["rev_ltm"] - 1) * 100
    df["rule40"] = df["g_ntm"] + df["ebit_margin"]
    df["group"] = df["ticker"].map(lambda t: PEER_CLASS.get(t, ("Unclassified", False))[0])
    df["core"] = df["ticker"].map(lambda t: PEER_CLASS.get(t, ("", False))[1])
    # Hard filters on top of the business-model screen
    df.loc[df["mcap"] < 5000, "core"] = False
    df.loc[(df["ebit_ltm"] <= 0) | df["ebit_ltm"].isna(), "core"] = False
    return df
