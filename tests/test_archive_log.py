"""Tests for archive (folder + zip) and log (JSON snapshot)."""

import json

import pytest

from dca.archive import archive_previous_output, ensure_fresh_output, zip_dir
from dca.config import DATA_DIR
from dca.excel_io import load_deal
from dca.log import dump_log, load_log

FILLED_EXAMPLE = DATA_DIR / "input_example_FILLED.xlsx"


# ─── log.py ─────────────────────────────────────────────────────

def test_dump_and_load_log_roundtrip(tmp_path):
    deal = load_deal(FILLED_EXAMPLE)
    log_path = tmp_path / "log.json"
    dump_log(deal, log_path)

    assert log_path.exists()
    loaded = load_log(log_path)
    # Key fields preserved
    assert loaded.agreement_num == deal.agreement_num
    assert loaded.agreement_date == deal.agreement_date
    assert loaded.party_a_name_eng == deal.party_a_name_eng
    assert loaded.party_b_name_eng == deal.party_b_name_eng


def test_log_json_is_human_readable(tmp_path):
    deal = load_deal(FILLED_EXAMPLE)
    log_path = tmp_path / "log.json"
    dump_log(deal, log_path)

    text = log_path.read_text(encoding="utf-8")
    assert "agreement_num" in text
    assert "_metadata" in text
    # Cyrillic preserved (not \u escapes)
    assert "Изделие X" in text


def test_load_log_missing_file_raises(tmp_path):
    with pytest.raises(FileNotFoundError):
        load_log(tmp_path / "nonexistent.json")


def test_load_log_accepts_raw_dict(tmp_path):
    """If JSON is just the Deal dict (no _metadata wrapper), should still work."""
    deal = load_deal(FILLED_EXAMPLE)
    raw_data = deal.model_dump(mode="json")
    log_path = tmp_path / "log.json"
    log_path.write_text(json.dumps(raw_data), encoding="utf-8")

    loaded = load_log(log_path)
    assert loaded.agreement_num == deal.agreement_num


def test_load_log_rejects_non_json(tmp_path):
    log_path = tmp_path / "log.json"
    log_path.write_text("not valid {json", encoding="utf-8")
    with pytest.raises(ValueError, match="not valid JSON"):
        load_log(log_path)


def test_load_log_rejects_top_level_list(tmp_path):
    log_path = tmp_path / "log.json"
    log_path.write_text("[1, 2, 3]", encoding="utf-8")
    with pytest.raises(ValueError, match="JSON object at top level"):
        load_log(log_path)


def test_load_log_rejects_wrapped_with_no_deal(tmp_path):
    """{_metadata: ...} without 'deal' is malformed — must error explicitly,
    not pass an empty dict into Deal() and produce a wall of Pydantic errors."""
    log_path = tmp_path / "log.json"
    log_path.write_text('{"_metadata": {"foo": "bar"}}', encoding="utf-8")
    with pytest.raises(ValueError, match="no valid 'deal' object"):
        load_log(log_path)


def test_load_log_rejects_random_dict(tmp_path):
    """Dict with no agreement_num and no wrapping keys → clear error, not Pydantic noise."""
    log_path = tmp_path / "log.json"
    log_path.write_text('{"foo": "bar", "baz": 42}', encoding="utf-8")
    with pytest.raises(ValueError, match="missing 'agreement_num'"):
        load_log(log_path)


# ─── archive.py — archive_previous_output ───────────────────────

def test_archive_no_output_returns_none(tmp_path):
    deal_dir = tmp_path / "deal1"
    deal_dir.mkdir()
    result = archive_previous_output(deal_dir)
    assert result is None


def test_archive_empty_output_dir_returns_none(tmp_path):
    deal_dir = tmp_path / "deal1"
    (deal_dir / "output").mkdir(parents=True)
    result = archive_previous_output(deal_dir)
    assert result is None
    assert not (deal_dir / "output").exists()  # cleaned up


def test_archive_output_with_only_dotfiles_treated_as_empty(tmp_path):
    """macOS Finder creates `.DS_Store` when user opens a folder. We must NOT
    treat such a folder as non-empty (would create spurious `_archive/v1/` etc).
    """
    deal_dir = tmp_path / "deal1"
    output = deal_dir / "output"
    output.mkdir(parents=True)
    (output / ".DS_Store").write_bytes(b"fake macOS metadata")
    (output / ".hidden").write_bytes(b"any dotfile")

    result = archive_previous_output(deal_dir)
    assert result is None
    assert not output.exists()  # wiped (including dotfiles)
    assert not (deal_dir / "_archive").exists()  # no spurious archive cycle


def test_archive_non_empty_output_moves_to_v1(tmp_path):
    deal_dir = tmp_path / "deal1"
    output = deal_dir / "output"
    output.mkdir(parents=True)
    (output / "Agreement.docx").write_bytes(b"PK\x03\x04dummy")

    result = archive_previous_output(deal_dir)
    assert result == deal_dir / "_archive" / "v1"
    assert (deal_dir / "_archive" / "v1" / "Agreement.docx").exists()
    assert not output.exists()  # moved away


def test_archive_increments_version_number(tmp_path):
    deal_dir = tmp_path / "deal1"
    # Pre-existing v1 archive
    (deal_dir / "_archive" / "v1").mkdir(parents=True)
    # New output to archive
    output = deal_dir / "output"
    output.mkdir(parents=True)
    (output / "a.docx").write_bytes(b"a")

    result = archive_previous_output(deal_dir)
    assert result == deal_dir / "_archive" / "v2"


# ─── ensure_fresh_output ────────────────────────────────────────

def test_ensure_fresh_output_creates_empty_dir(tmp_path):
    deal_dir = tmp_path / "deal1"
    deal_dir.mkdir()
    out = ensure_fresh_output(deal_dir)
    assert out == deal_dir / "output"
    assert out.exists()
    assert not any(out.iterdir())


def test_ensure_fresh_output_archives_old(tmp_path):
    deal_dir = tmp_path / "deal1"
    output = deal_dir / "output"
    output.mkdir(parents=True)
    (output / "old.docx").write_bytes(b"old")

    out = ensure_fresh_output(deal_dir)
    assert out.exists()
    assert not any(out.iterdir())  # fresh empty
    assert (deal_dir / "_archive" / "v1" / "old.docx").exists()  # old archived


# ─── zip_dir ────────────────────────────────────────────────────

def test_zip_dir_includes_all_files(tmp_path):
    src = tmp_path / "src"
    src.mkdir()
    (src / "Agreement.docx").write_bytes(b"c")
    (src / "Annex.docx").write_bytes(b"i")
    (src / "log.json").write_text("{}")

    import zipfile as zf
    dest = tmp_path / "package.zip"
    result = zip_dir(src, dest)
    assert result == dest
    assert dest.exists()

    with zf.ZipFile(dest) as z:
        names = z.namelist()
    assert set(names) == {"Agreement.docx", "Annex.docx", "log.json"}


def test_zip_dir_does_not_include_itself(tmp_path):
    """If dest_path is inside src_dir, exclude it from zip."""
    src = tmp_path / "src"
    src.mkdir()
    (src / "a.txt").write_text("a")

    import zipfile as zf
    dest = src / "package.zip"
    zip_dir(src, dest)
    with zf.ZipFile(dest) as z:
        names = z.namelist()
    assert "package.zip" not in names
    assert names == ["a.txt"]


def test_zip_dir_missing_src_raises(tmp_path):
    with pytest.raises(FileNotFoundError):
        zip_dir(tmp_path / "nonexistent", tmp_path / "out.zip")


# ─── End-to-end pipeline simulation ─────────────────────────────

def test_full_pipeline_archive_then_zip(tmp_path):
    """Simulate: run 1 → archive previous (no-op), generate output, zip; run 2 → archive v1, regenerate."""
    deal = load_deal(FILLED_EXAMPLE)
    deal_dir = tmp_path / f"{deal.agreement_num}_{int(deal.quantity_kg)}kg"

    # Run 1
    output_dir = ensure_fresh_output(deal_dir)
    (output_dir / "Agreement.docx").write_bytes(b"run1")
    dump_log(deal, output_dir / "log.json")
    zip_dir(output_dir, output_dir / f"Package_{deal.agreement_num}.zip")
    assert (deal_dir / "output" / "Agreement.docx").exists()
    assert not (deal_dir / "_archive").exists()  # no prior runs

    # Run 2 — should move first run's output to _archive/v1/
    output_dir = ensure_fresh_output(deal_dir)
    (output_dir / "Agreement.docx").write_bytes(b"run2")
    assert (deal_dir / "_archive" / "v1" / "Agreement.docx").read_bytes() == b"run1"
    assert (deal_dir / "output" / "Agreement.docx").read_bytes() == b"run2"
