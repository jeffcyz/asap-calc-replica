"""Create Git-trackable run folder scaffolds without committing run data."""

from __future__ import annotations

import argparse
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEAL_PATTERN = re.compile(r"AALC\d{4}")
MONTH_PATTERN = re.compile(r"\d{4}-(0[1-9]|1[0-2])")
RUN_SUBDIRECTORIES = (
    "inputs",
    "inputs/original",
    "inputs/csv",
    "benchmark",
    "output",
    "reports",
)


def discover_deals() -> list[str]:
    """Return configured AALC deal IDs, excluding _template."""
    root = ROOT / "deals"
    if not root.is_dir():
        return []
    return sorted(
        item.name for item in root.iterdir()
        if item.is_dir()
        and DEAL_PATTERN.fullmatch(item.name)
        and (item / "deal.yaml").is_file()
    )


def prepare_run(deal: str, month: str) -> Path:
    """Create .gitkeep markers, without modifying real data files."""
    if not DEAL_PATTERN.fullmatch(deal):
        raise ValueError(f"Invalid deal ID: {deal}")
    if not MONTH_PATTERN.fullmatch(month):
        raise ValueError(f"Invalid month: {month}; expected YYYY-MM")
    if not (ROOT / "deals" / deal / "deal.yaml").is_file():
        raise ValueError(f"Unknown deal: {deal}")
    base = ROOT / "runs" / deal / month
    for section in RUN_SUBDIRECTORIES:
        directory = base / section
        directory.mkdir(parents=True, exist_ok=True)
        (directory / ".gitkeep").touch(exist_ok=True)
    return base


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    scope = parser.add_mutually_exclusive_group(required=True)
    scope.add_argument("--deal", help="A configured deal ID")
    scope.add_argument("--all", action="store_true", help="All configured AALC deals")
    parser.add_argument("--month", required=True, help="Payment month YYYY-MM")
    args = parser.parse_args()
    deals = discover_deals() if args.all else [args.deal]
    if not deals:
        parser.error("No configured AALC deals found")
    for deal in deals:
        print(prepare_run(deal, args.month))


if __name__ == "__main__":
    main()
