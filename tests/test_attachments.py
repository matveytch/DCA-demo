"""Attachment pack (dca.attachments) and soft checks (dca.checks): AC8, AC9, AC31."""

from __future__ import annotations

import unicodedata
from datetime import date

import pytest

from dca.attachments import copy_attachment_files, list_attachment_items
from dca.checks import attachment_warnings, batch_company_warnings, date_warnings, write_check_file
from dca.models import AnnexLine
from tests.helpers import COMPANY, deal, example_deal, incoming_deal


@pytest.fixture
def scans(tmp_path):
    d = tmp_path / "scans"
    (d / "П-001").mkdir(parents=True)
    (d / "П-001" / "соглашение В-17.pdf").write_bytes(b"17")
    (d / "П-001" / "приложение 1 к В-17.pdf").write_bytes(b"17-1")
    (d / "соглашение В-21.pdf").write_bytes(b"21")
    return d


@pytest.fixture
def company(tmp_path):
    c = tmp_path / COMPANY
    c.mkdir()
    (c / "company_profile.png").write_bytes(b"profile")       # any extension
    return c


def test_items_follow_the_cover_list_numbering():
    items = list_attachment_items(example_deal())
    assert [(i["n"], i["kind"], i["pkg"]) for i in items] == [
        (2, "agreement", 1), (3, "annex", 1), (4, "agreement", 2), (5, "company_profile", None)]


def test_copy_renames_by_number_and_label(tmp_path, scans, company):
    manifest = copy_attachment_files(example_deal(), tmp_path / "out", scans, company)
    names = sorted(p.name for p in manifest["target_dir"].iterdir())
    assert names[0] == "2. копия соглашения №В-17-2025 от 15.09.2025.pdf"     # «/» is not allowed in a name
    assert names[-1] == "5. карточка компании.png"
    assert manifest["missing"] == [] and manifest["skipped"] == [] and manifest["ambiguous"] == []


def test_missing_scan_and_missing_profile_are_reported(tmp_path, scans):
    (scans / "соглашение В-21.pdf").unlink()
    empty_company = tmp_path / "empty"
    empty_company.mkdir()
    manifest = copy_attachment_files(example_deal(), tmp_path / "out", scans, empty_company)
    assert manifest["missing"] == ["соглашение В-21.pdf", "empty/company_profile.pdf"]


def test_the_batch_folder_wins_over_a_different_file_elsewhere(tmp_path, scans, company):
    (scans / "соглашение В-17.pdf").write_bytes(b"another document")
    manifest = copy_attachment_files(example_deal(), tmp_path / "out", scans, company)
    copied = manifest["target_dir"] / "2. копия соглашения №В-17-2025 от 15.09.2025.pdf"
    assert copied.read_bytes() == b"17"


def test_different_files_with_one_name_are_never_guessed(tmp_path, scans, company):
    (scans / "x").mkdir()
    (scans / "x" / "соглашение В-21.pdf").write_bytes(b"not the same")
    manifest = copy_attachment_files(example_deal(), tmp_path / "out", scans, company)
    assert [a["source_name"] for a in manifest["ambiguous"]] == ["соглашение В-21.pdf"]


def test_identical_copies_of_one_scan_are_fine(tmp_path, scans, company):
    (scans / "x").mkdir()
    (scans / "x" / "соглашение В-21.pdf").write_bytes(b"21")
    assert copy_attachment_files(example_deal(), tmp_path / "out", scans, company)["ambiguous"] == []


def test_decomposed_letters_in_a_file_name_still_match(tmp_path, company):
    scans = tmp_path / "scans"
    scans.mkdir()
    name = "соглашение В-21 (й).pdf"
    (scans / unicodedata.normalize("NFD", name)).write_bytes(b"21")
    d = deal(batch_agreement_file_1=name, batch_annex_num_1=None, batch_scans_folder_1=None)
    manifest = copy_attachment_files(d, tmp_path / "out", scans, company)
    assert manifest["missing"] == [] and len(manifest["copied"]) == 2


# ─── soft checks ───────────────────────────────────────────────

def test_complete_deal_has_no_notices():
    d = example_deal()
    assert attachment_warnings(d) == [] and batch_company_warnings(d) == [] and date_warnings(d) == []


def test_batch_of_another_company_and_without_owner():
    warnings = batch_company_warnings(deal(batch_company_key_1="omega", batch_company_name_ru_1="Общество «Омега»"))
    assert warnings == ["Партия #1 «П-001» — ПАРТИЯ ДРУГОЙ КОМПАНИИ: Общество «Омега» (omega), "
                        "а сторона А этой сделки — alpha."]
    assert "не указано, какой компании" in batch_company_warnings(deal(batch_company_key_1=None))[0]


def test_incomplete_batch_and_no_batch_at_all():
    w = attachment_warnings(deal(batch_party_name_ru_1=None, batch_annex_date_1=None))
    assert len(w) == 1 and "контрагент (batch_party_name_ru)" in w[0] and "заполняются вместе" in w[0]
    none = attachment_warnings(deal(batch_key_1=None, batch_use_kg_1=None))
    assert none and "нет ни одной партии" in none[0]
    assert attachment_warnings(incoming_deal()) == []


def test_date_contradictions_are_warnings():
    early_end = date_warnings(deal(due_date=date(2026, 5, 1)))
    assert any("раньше даты соглашения" in w for w in early_end)
    late_annex = date_warnings(deal(annexes=[
        AnnexLine(annex_num="1", annex_date=date(2026, 5, 17), quantity_kg=500),
        AnnexLine(annex_num="2", annex_date=date(2026, 8, 1), quantity_kg=500)]))
    assert late_annex == [late_annex[0]] and "Приложение №2 от 01.08.2026 позже" in late_annex[0]
    assert date_warnings(deal(due_date=date(2026, 5, 25))) == []      # equal to the transfer date: fine


def test_check_file_only_when_there_is_something_to_check(tmp_path):
    assert write_check_file(tmp_path, []) is None and not list(tmp_path.iterdir())
    path = write_check_file(tmp_path, ["первое", "второе"])
    assert path.read_text(encoding="utf-8").splitlines()[2:] == ["• первое", "• второе"]
