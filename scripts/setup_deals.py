"""Create local, Git-ignored run directories for an authorized deal/month."""
from __future__ import annotations

import argparse
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEAL_PATTERN = re.compile(r"AALC\d{4}")


def prepare_run(deal: str, month: str) -> Path:
    if not DEAL_PATTERN.fullmatch(deal):
        raise ValueError(f"Invalid deal ID: {deal}")
    if not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", month):
        raise ValueError(f"Invalid month: {month}; expected YYYY-MM")
    if not (ROOT / "deals" / deal / "deal.yaml").is_file():
        raise ValueError(f"Unknown deal: {deal}")
    base = ROOT / "runs" / deal / month
    for section in ("inputs", "benchmark", "output", "reports"):
        (base / section).mkdir(parents=True, exist_ok=True)
    return base


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--deal", required=True)
    parser.add_argument("--month", required=True)
    args = parser.parse_args()
    print(prepare_run(args.deal, args.month))


if __name__ == "__main__":
    main()
