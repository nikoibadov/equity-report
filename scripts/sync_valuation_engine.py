"""Copy the valuation engine from the SAP_Valuation package (single source of truth) into src/eqr/valuation.

    .venv/bin/python scripts/sync_valuation_engine.py [--src ../SAP_Valuation/valuation]
"""
import argparse
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DST = ROOT / "src" / "eqr" / "valuation"

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=str(ROOT.parent / "SAP_Valuation" / "valuation"))
    src = Path(ap.parse_args().src)
    n = 0
    for f in src.rglob("*"):
        if f.suffix in (".py", ".yaml") and "__pycache__" not in f.parts:
            out = DST / f.relative_to(src)
            out.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(f, out)
            n += 1
    print(f"synced {n} files {src} -> {DST}")
