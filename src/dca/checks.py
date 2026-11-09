"""Soft checks shown after `dca generate` (documents are still produced).

They end up in two places so they are hard to miss: a yellow block at the very
end of the terminal output and `output_<num>/❗ПРОВЕРЬТЕ.txt`.
"""

from __future__ import annotations

from pathlib import Path

from dca.models import Deal

CHECK_FILENAME = "❗ПРОВЕРЬТЕ.txt"


def batch_company_warnings(deal: Deal) -> list[str]:
    """Batches that belong to another of our companies, or whose owner is not set.

    Compared by key (batch_company_key vs the deal's party_a_key) — exact, so
    different spellings of a company name never raise a false alarm.
    """
    if deal.deal_kind != "outgoing":
        return []
    out = []
    for n in deal.active_batches():
        key = getattr(deal, f"batch_key_{n}")
        owner = getattr(deal, f"batch_company_key_{n}")
        if not owner:
            out.append(f"Партия #{n} «{key}»: не указано, какой компании она принадлежит "
                       f"(batch_company_key в «Партиях» пуст).")
        elif owner != deal.party_a_key:
            name = getattr(deal, f"batch_company_name_ru_{n}") or owner
            out.append(f"Партия #{n} «{key}» — ПАРТИЯ ДРУГОЙ КОМПАНИИ: {name} ({owner}), "
                       f"а сторона А этой сделки — {deal.party_a_key}.")
    return out


# Fields of a batch that the cover list prints as is: an empty one leaves
# «№ от ;» in the document. field → what to tell the operator.
BATCH_REQUIRED_FIELDS = {
    "batch_agreement_num": "номер соглашения партии",
    "batch_agreement_date": "дата соглашения партии",
    "batch_party_name_ru": "контрагент (batch_party_name_ru)",
    "batch_party_location_ru": "место контрагента (batch_party_location_ru)",
    "batch_use_kg": "количество из партии (batch_use_kg на «Сделке»)",
}


def attachment_warnings(deal: Deal) -> list[str]:
    """Batches the cover list can't be filled from.

    A missing batch or field silently gives «№ от ;» in the cover list, and the
    numbering of the scans no longer matches it.
    """
    if deal.deal_kind != "outgoing":
        return []
    active = deal.active_batches()
    if not active:
        return ["В сделке нет ни одной партии — опись приложений получится без копий документов "
                "партий. Выберите партию(и) на «Сделке»."]
    out = []
    for n in active:
        key = getattr(deal, f"batch_key_{n}")
        missing = [label for field, label in BATCH_REQUIRED_FIELDS.items()
                   if getattr(deal, f"{field}_{n}") in (None, "")]
        if bool(getattr(deal, f"batch_annex_num_{n}")) != bool(getattr(deal, f"batch_annex_date_{n}")):
            missing.append("номер и дата приложения партии заполняются вместе")
        if missing:
            out.append(f"Партия #{n} «{key}» заполнена не полностью — в документах будут пустые "
                       f"места: {'; '.join(missing)}.")
    return out


def _ru(d) -> str:
    return f"{d:%d.%m.%Y}"


def transfer_date_warnings(deal: Deal) -> list[str]:
    """The transfer date after due_date: the agreement no longer runs then."""
    end = deal.due_date
    handed = deal.transfer.transfer_date if deal.transfer is not None else None
    if end is None or handed is None or handed <= end:
        return []
    return [f"Дата передачи {_ru(handed)} позже срока исполнения {_ru(end)}: "
            f"в этот день соглашение уже не действует. Исправьте transfer_date "
            f"на «Передаче» или due_date на «Сделке»."]


def date_warnings(deal: Deal) -> list[str]:
    """Dates that contradict the agreement's term.

    due_date is both the deadline and the end of the agreement; a date equal to
    it is fine.
    """
    end = deal.due_date
    if end is None:
        return []
    out = []
    if end < deal.agreement_date:
        out.append(f"Срок исполнения {_ru(end)} раньше даты соглашения "
                   f"{_ru(deal.agreement_date)}: соглашение закончится раньше, чем его подписали. "
                   f"Исправьте due_date на «Сделке».")
    for annex in deal.resolved_annexes():
        if annex.annex_date > end:
            out.append(f"Приложение №{annex.annex_num} от {_ru(annex.annex_date)} позже "
                       f"срока исполнения {_ru(end)}: в этот день соглашение уже не действует. "
                       f"Исправьте дату приложения или due_date на «Сделке».")
    return out + transfer_date_warnings(deal)


def write_check_file(output_dir: Path, notices: list[str]) -> Path | None:
    """`❗ПРОВЕРЬТЕ.txt` with every notice; no notices → no file."""
    if not notices:
        return None
    path = output_dir / CHECK_FILENAME
    lines = ["Проверьте перед отправкой документов:", ""] + [f"• {n}" for n in notices]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path
