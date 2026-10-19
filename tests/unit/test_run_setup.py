"""Synthetic tests of run directory initialization, without ASAP data."""

from pathlib import Path

import pytest

from scripts import setup_deals


@pytest.fixture
def demo_root(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    monkeypatch.setattr(setup_deals, "ROOT", tmp_path)
    for deal in ("AALC2201", "AALC2404"):
        folder = tmp_path / "deals" / deal
        folder.mkdir(parents=True)
        (folder / "deal.yaml").write_text("deal_id: DEMO\n", encoding="utf-8")
    return tmp_path


def test_find_configured_deals(demo_root: Path) -> None:
    assert setup_deals.discover_deals() == ["AALC2201", "AALC2404"]


def test_create_all_expected_directories(demo_root: Path) -> None:
    base = setup_deals.prepare_run("AALC2201", "2026-10")
    for section in setup_deals.RUN_SUBDIRECTORIES:
        assert (base / section / ".gitkeep").is_file()


def test_source_data_is_untouched(demo_root: Path) -> None:
    base = setup_deals.prepare_run("AALC2201", "2026-10")
    original = base / "inputs" / "original" / "source.xlsx"
    original.write_bytes(b"private workbook data")
    setup_deals.prepare_run("AALC2201", "2026-10")
    assert original.read_bytes() == b"private workbook data"


@pytest.mark.parametrize("deal,month", [
    ("_template", "2026-10"),
    ("AALC2201", "2026-13"),
    ("AALC9999", "2026-10"),
])
def test_reject_bad_identifiers(demo_root: Path, deal: str, month: str) -> None:
    with pytest.raises(ValueError):
        setup_deals.prepare_run(deal, month)
