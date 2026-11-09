"""Pydantic models for a Deal.

Single flat `Deal` class containing all SOURCE fields (what the operator
enters). Computed values for templates are built separately in `context.py`.

Validation:
  - default-filling (annex_num = leading part of agreement_num,
    annex_date = agreement_date, due_date = agreement_date + 2 months)
  - business rules (party_c_* required iff variant='B', batch quantities sum
    to the deal quantity, annex quantities sum to the agreement quantity)
"""

from __future__ import annotations

import math
import unicodedata
from datetime import date
from decimal import Decimal
from typing import ClassVar, Literal

from dateutil.relativedelta import relativedelta
from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

# ─── Type aliases for enums ─────────────────────────────────────
Currency = Literal["USD", "EUR"]
UnitOfMeasure = Literal["kg", "pcs", "t"]
Variant = Literal["A", "B"]

BATCH_SLOTS = (1, 2, 3, 4, 5)


# ─── Annex line — one of N annexes per agreement ────────────────
class AnnexLine(BaseModel):
    """One annex of an agreement (its own number, date and quantity).

    An agreement can have several annexes. All share the same item, price,
    parties and place — only the number, date and quantity differ. Read from
    the «Приложения» sheet.
    """

    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    annex_num: str = Field(..., min_length=1)
    annex_date: date
    quantity_kg: float = Field(..., gt=0, allow_inf_nan=False)

    @field_validator("annex_num")
    @classmethod
    def _nfc(cls, v: str) -> str:
        # «й» typed as one letter or as «и» + combining breve looks the same and is
        # the same file name on disk — keep one spelling.
        return unicodedata.normalize("NFC", v)


# ─── Transfer act ───────────────────────────────────────────────
class TransferInfo(BaseModel):
    """Data of the transfer act from the «Передача» sheet: one act per deal.

    Every field is optional at load time, so a half-filled sheet never blocks
    `dca generate`. Completeness and date logic are checked only when the act
    is rendered — see `problems()`.
    """

    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    transfer_num: str | None = None
    transfer_date: date | None = None
    place_from_ru: str | None = None
    place_to_ru: str | None = None
    handed_by_ru: str | None = None     # representative of party A (full name)
    received_by_ru: str | None = None   # representative of party B (full name)
    note_ru: str | None = None          # optional multi-line note (box 4)

    REQUIRED: ClassVar[tuple[str, ...]] = (
        "transfer_num", "transfer_date", "place_from_ru", "place_to_ru",
        "handed_by_ru", "received_by_ru",
    )

    def missing(self) -> list[str]:
        """Required fields still empty — the sheet is not filled yet (normal early in a deal)."""
        return [f"{f} is required" for f in self.REQUIRED if getattr(self, f) in (None, "")]

    def problems(self, agreement_date: date) -> list[str]:
        """Everything that blocks rendering the transfer act (empty = OK)."""
        out = self.missing()
        if self.transfer_date and self.transfer_date < agreement_date:
            out.append(f"transfer_date {self.transfer_date:%d.%m.%Y} is before agreement_date "
                       f"{agreement_date:%d.%m.%Y}")
        if self.handed_by_ru and self.received_by_ru and self.handed_by_ru == self.received_by_ru:
            out.append("handed_by_ru and received_by_ru must be different people")
        return out


# ─── Amendment ──────────────────────────────────────────────────
def _exact(x: float) -> Decimal:
    """A float as the operator typed it (0.07 → Decimal('0.07'))."""
    return Decimal(repr(float(x)))


class AmendmentInfo(BaseModel):
    """The «Изменение» sheet: what the amendment changes.

    Old values default to the agreement («Сделка»): empty old quantity = the deal's
    quantity, empty old price = the deal's price, empty new price = the old price.
    Prices are USD per 1 kg. All checks run only when `dca amend` renders —
    see `problems()`.
    """

    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    amendment_num: str | None = None
    amendment_date: date | None = None
    old_quantity_kg: float | None = Field(default=None, gt=0, allow_inf_nan=False)
    new_quantity_kg: float | None = Field(default=None, gt=0, allow_inf_nan=False)
    old_price_per_unit: float | None = Field(default=None, gt=0, allow_inf_nan=False)
    new_price_per_unit: float | None = Field(default=None, gt=0, allow_inf_nan=False)

    REQUIRED: ClassVar[tuple[str, ...]] = ("amendment_num", "amendment_date", "new_quantity_kg")

    def old_quantity(self, deal_quantity: float) -> float:
        return self.old_quantity_kg or deal_quantity

    def old_price(self, deal_price: float) -> float:
        return self.old_price_per_unit or deal_price

    def new_price(self, deal_price: float) -> float:
        return self.new_price_per_unit or self.old_price(deal_price)

    def quantity_changed(self, deal_quantity: float) -> bool:
        # Documents print whole kg rounded up: compare what is printed.
        new = self.new_quantity_kg
        return new is not None and math.ceil(self.old_quantity(deal_quantity)) != math.ceil(new)

    def price_changed(self, deal_price: float) -> bool:
        return _exact(self.old_price(deal_price)) != _exact(self.new_price(deal_price))

    def missing(self) -> list[str]:
        """Required fields still empty."""
        return [f"{f} is required" for f in self.REQUIRED if getattr(self, f) in (None, "")]

    def problems(self, agreement_date: date, deal_quantity: float, deal_price: float) -> list[str]:
        """Everything that blocks rendering the amendment (empty = OK)."""
        out = self.missing()
        if self.amendment_date and self.amendment_date < agreement_date:
            out.append(f"amendment_date {self.amendment_date:%d.%m.%Y} is before agreement_date "
                       f"{agreement_date:%d.%m.%Y}")
        if (self.new_quantity_kg is not None and not self.quantity_changed(deal_quantity)
                and not self.price_changed(deal_price)):
            out.append("nothing changes: new_quantity_kg and the price are the same as before")
        return out


# ─── Deal model — single flat container of all source data ─────
class Deal(BaseModel):
    """All source data for one deal, as filled by the operator.

    Validators fill defaults and enforce business rules.
    """

    model_config = ConfigDict(
        str_strip_whitespace=True,   # auto-trim whitespace
        validate_assignment=True,     # validate when fields reassigned
        extra="forbid",               # error on unknown fields
    )

    # ─── Deal kind ───────────────────────────────────────────────
    # "outgoing" = full package: agreement, annex×N, cover list, attachments, ledger
    #              (dca new / generate).
    # "incoming" = agreement + annex×N only (dca new-incoming / generate-incoming).
    #              Several outgoing-only fields below become optional
    #              (see _check_outgoing_required_fields).
    deal_kind: Literal["outgoing", "incoming"] = "outgoing"

    # ─── Deal meta ───────────────────────────────────────────────
    agreement_num: str = Field(..., min_length=1)
    agreement_date: date
    # Single-annex fields (default). When the «Приложения» sheet lists several
    # annexes, `annexes` below is populated and these act only as a fallback for
    # single-annex deals (resolved_annexes).
    annex_num: str | None = None
    annex_date: date | None = None
    annexes: list[AnnexLine] = Field(default_factory=list)
    # Transfer act from the «Передача» sheet; None = not filled.
    transfer: TransferInfo | None = None
    # Amendment from the «Изменение» sheet; None = not filled.
    amendment: AmendmentInfo | None = None
    # Optional at type level; REQUIRED for outgoing deals (see _check_outgoing_required_fields).
    variant: Variant | None = None
    due_date: date | None = None

    # ─── Item ────────────────────────────────────────────────────
    item_key: str | None = None   # chosen «Изделия» record (kept for traceability)
    item_name_ru: str = Field(..., min_length=1)
    item_name_eng: str = Field(..., min_length=1)
    item_description_ru: str = Field(..., min_length=1)
    item_description_eng: str = Field(..., min_length=1)
    unit_of_measure: UnitOfMeasure = "kg"
    price_per_unit: float = Field(..., gt=0, allow_inf_nan=False)
    currency: Currency = "USD"
    quantity_kg: float = Field(..., gt=0, allow_inf_nan=False)

    # ─── Place ───────────────────────────────────────────────────
    place_ru: str = Field(..., min_length=1)
    place_eng: str = Field(..., min_length=1)

    # ─── Party A ─────────────────────────────────────────────────
    # Key of the chosen «Стороны А» record and its company folder (set only for OUR
    # companies): an outgoing deal is filed under <folder>/Outgoing by `dca generate`.
    party_a_key: str | None = None
    party_a_folder_name: str | None = None
    party_a_name_ru: str = Field(..., min_length=1)
    party_a_name_eng: str = Field(..., min_length=1)
    party_a_city_ru: str = Field(..., min_length=1)
    party_a_city_eng: str = Field(..., min_length=1)
    party_a_represented_by_ru: str = Field(..., min_length=1)
    party_a_represented_by_eng: str = Field(..., min_length=1)
    party_a_address_full_ru: str = Field(..., min_length=1)
    party_a_address_full_eng: str = Field(..., min_length=1)
    party_a_details_full_ru: str = Field(..., min_length=1)
    party_a_details_full_eng: str = Field(..., min_length=1)
    party_a_signatory_name_ru: str = Field(..., min_length=1)
    party_a_signatory_name_eng: str = Field(..., min_length=1)

    # ─── Party B ─────────────────────────────────────────────────
    # party_b_folder_name is set for OUR companies (as party B of incoming deals):
    # an incoming deal is filed under <folder>/Incoming.
    party_b_key: str | None = None
    party_b_folder_name: str | None = None
    party_b_name_ru: str = Field(..., min_length=1)
    party_b_name_eng: str = Field(..., min_length=1)
    party_b_represented_by_ru: str = Field(..., min_length=1)
    party_b_represented_by_eng: str = Field(..., min_length=1)
    party_b_address_full_ru: str = Field(..., min_length=1)
    party_b_address_full_eng: str = Field(..., min_length=1)
    party_b_details_full_ru: str = Field(..., min_length=1)
    party_b_details_full_eng: str = Field(..., min_length=1)
    party_b_signatory_name_ru: str = Field(..., min_length=1)
    party_b_signatory_name_eng: str = Field(..., min_length=1)

    # ─── Party C (variant B only; required iff variant='B') ─────
    party_c_name: str | None = None
    party_c_address: str | None = None
    party_c_details: str | None = None

    # ─── Cover list (outgoing deals) ─────────────────────────────
    letter_num: str | None = None  # empty → "1"
    letter_date: date | None = None

    # ─── Source batches (up to 5; each picked from the «Партии» sheet) ─
    # Two operator-entered fields per slot on «Сделка»:
    #   - batch_key_N:    which batch this slot uses
    #   - batch_use_kg_N: how many kg from THIS batch go into THIS deal
    # The other fields are looked up from the «Партии» sheet by key.
    # Batches must be filled WITHOUT GAPS: if #3 is set, #1 and #2 must be too.
    # Sum of batch_use_kg across active batches must equal quantity_kg.
    batch_key_1: str | None = None
    batch_key_2: str | None = None
    batch_key_3: str | None = None
    batch_key_4: str | None = None
    batch_key_5: str | None = None
    batch_use_kg_1: float | None = Field(None, gt=0, allow_inf_nan=False)
    batch_use_kg_2: float | None = Field(None, gt=0, allow_inf_nan=False)
    batch_use_kg_3: float | None = Field(None, gt=0, allow_inf_nan=False)
    batch_use_kg_4: float | None = Field(None, gt=0, allow_inf_nan=False)
    batch_use_kg_5: float | None = Field(None, gt=0, allow_inf_nan=False)

    # Batch #1 — looked up from «Партии» by batch_key_1
    batch_agreement_num_1: str | None = None
    batch_agreement_date_1: date | None = None
    batch_annex_num_1: str | None = None    # optional annex of the batch agreement
    batch_annex_date_1: date | None = None
    batch_party_name_ru_1: str | None = None
    batch_party_location_ru_1: str | None = None
    batch_quantity_kg_1: float | None = Field(None, gt=0, allow_inf_nan=False)  # the batch's whole quantity
    batch_agreement_file_1: str | None = None        # scan file names (attachment pack)
    batch_annex_file_1: str | None = None
    batch_company_key_1: str | None = None           # which of OUR companies owns the batch
    batch_company_name_ru_1: str | None = None
    batch_scans_folder_1: str | None = None          # folder in Data/scans with this batch's scans

    # Batch #2
    batch_agreement_num_2: str | None = None
    batch_agreement_date_2: date | None = None
    batch_annex_num_2: str | None = None
    batch_annex_date_2: date | None = None
    batch_party_name_ru_2: str | None = None
    batch_party_location_ru_2: str | None = None
    batch_quantity_kg_2: float | None = Field(None, gt=0, allow_inf_nan=False)
    batch_agreement_file_2: str | None = None
    batch_annex_file_2: str | None = None
    batch_company_key_2: str | None = None
    batch_company_name_ru_2: str | None = None
    batch_scans_folder_2: str | None = None

    # Batch #3
    batch_agreement_num_3: str | None = None
    batch_agreement_date_3: date | None = None
    batch_annex_num_3: str | None = None
    batch_annex_date_3: date | None = None
    batch_party_name_ru_3: str | None = None
    batch_party_location_ru_3: str | None = None
    batch_quantity_kg_3: float | None = Field(None, gt=0, allow_inf_nan=False)
    batch_agreement_file_3: str | None = None
    batch_annex_file_3: str | None = None
    batch_company_key_3: str | None = None
    batch_company_name_ru_3: str | None = None
    batch_scans_folder_3: str | None = None

    # Batch #4
    batch_agreement_num_4: str | None = None
    batch_agreement_date_4: date | None = None
    batch_annex_num_4: str | None = None
    batch_annex_date_4: date | None = None
    batch_party_name_ru_4: str | None = None
    batch_party_location_ru_4: str | None = None
    batch_quantity_kg_4: float | None = Field(None, gt=0, allow_inf_nan=False)
    batch_agreement_file_4: str | None = None
    batch_annex_file_4: str | None = None
    batch_company_key_4: str | None = None
    batch_company_name_ru_4: str | None = None
    batch_scans_folder_4: str | None = None

    # Batch #5
    batch_agreement_num_5: str | None = None
    batch_agreement_date_5: date | None = None
    batch_annex_num_5: str | None = None
    batch_annex_date_5: date | None = None
    batch_party_name_ru_5: str | None = None
    batch_party_location_ru_5: str | None = None
    batch_quantity_kg_5: float | None = Field(None, gt=0, allow_inf_nan=False)
    batch_agreement_file_5: str | None = None
    batch_annex_file_5: str | None = None
    batch_company_key_5: str | None = None
    batch_company_name_ru_5: str | None = None
    batch_scans_folder_5: str | None = None

    # ─── Validators: defaults ────────────────────────────────────

    @model_validator(mode="after")
    def _fill_defaults(self) -> Deal:
        """Fill chained defaults.

        annex_num  ← prefix of agreement_num up to (but excluding) the first hyphen.
                     Examples: '1-2026-1705' → '1'; 'A7-2026' → 'A7'; 'NOHYPHEN' → 'NOHYPHEN'
        annex_date ← agreement_date
        due_date   ← agreement_date + 2 calendar months
        """
        if not self.annex_num:
            self.annex_num = self.agreement_num.split("-", 1)[0]
        if self.annex_date is None:
            self.annex_date = self.agreement_date
        if self.due_date is None:
            self.due_date = self.agreement_date + relativedelta(months=2)
        return self

    # ─── Validators: business rules ──────────────────────────────

    @model_validator(mode="after")
    def _check_supported_units(self) -> Deal:
        """Only kg and USD are supported for now.

        The templates print the price «за 1 кг» and the amount in words in US dollars /
        cents; anything else would be printed wrongly, so it is refused.
        """
        problems = []
        if self.unit_of_measure != "kg":
            problems.append(f"unit_of_measure '{self.unit_of_measure}' пока не поддерживается — документы "
                            f"рассчитаны на цену за кг; укажите kg")
        if self.currency != "USD":
            problems.append(f"currency '{self.currency}' пока не поддерживается — документы формируются "
                            f"в USD (сумма прописью, центы)")
        if problems:
            raise ValueError("; ".join(problems))
        return self

    @model_validator(mode="after")
    def _check_outgoing_required_fields(self) -> Deal:
        """Fields that are optional at type level but required for outgoing deals."""
        if self.deal_kind != "outgoing":
            return self
        if self.variant is None:
            raise ValueError("deal_kind='outgoing' requires variant (A or B). "
                             "Fill it, or use an incoming deal (dca new-incoming).")
        return self

    @model_validator(mode="after")
    def _check_party_c_iff_b(self) -> Deal:
        """Variant B adds a clause about party C: all three party C fields are printed."""
        missing = [f for f in ("party_c_name", "party_c_address", "party_c_details") if not getattr(self, f)]
        if self.variant == "B" and missing:
            raise ValueError(f"variant='B' requires {', '.join(missing)}. Provide them or change variant.")
        return self

    @model_validator(mode="after")
    def _check_batches_no_gaps(self) -> Deal:
        """Batches must be filled WITHOUT GAPS: 1; 1+2; 1+2+3 … but NOT 1+3."""
        active = self.active_batches()
        if not active:
            return self
        expected = set(range(1, max(active) + 1))
        if set(active) != expected:
            missing = sorted(expected - set(active))
            raise ValueError(
                f"Batches must be filled sequentially: gap detected. "
                f"Active slots: {active}, missing: {missing}. "
                f"Move higher batches down to fill the gaps."
            )
        return self

    @model_validator(mode="after")
    def _check_batches_quantity_sum(self) -> Deal:
        """Sum of batch quantities must equal quantity_kg — compared AS RENDERED.

        Documents print ceil(quantity_kg) and ceil(batch_use_kg_N) independently.
        On raw floats quantity 1000.5 = 500.3 + 500.2 would pass, but the documents
        would show 1001 = 501 + 501 = 1002. So the check uses the ceil'd numbers.
        """
        active = self.active_batches()
        if not active:
            return self
        ceiled: list[int] = []
        for n in active:
            w = getattr(self, f"batch_use_kg_{n}")
            if w is None:
                raise ValueError(
                    f"Batch #{n} is selected (batch_key_{n} set) "
                    f"but batch_use_kg_{n} is empty. Fill the quantity."
                )
            ceiled.append(math.ceil(w))
        total = sum(ceiled)
        deal_total = math.ceil(self.quantity_kg)
        if total != deal_total:
            details = " + ".join(str(w) for w in ceiled)
            raise ValueError(
                f"Sum of batch quantities (as shown in documents) "
                f"{details} = {total} kg does NOT match "
                f"deal quantity_kg = {deal_total} kg (rendered). "
                f"Adjust batch quantities so they sum exactly after rounding-up."
            )
        return self

    @model_validator(mode="after")
    def _check_annexes_quantity_sum(self) -> Deal:
        """Several annexes: their quantities must sum to the agreement quantity_kg
        (AS RENDERED, same rationale as _check_batches_quantity_sum), and annex
        numbers must be unique (they become output file names).

        No-op for single-annex deals (empty `annexes`).
        """
        if not self.annexes:
            return self
        nums = [a.annex_num for a in self.annexes]
        dupes = sorted({n for n in nums if nums.count(n) > 1})
        if dupes:
            raise ValueError(f"Annex numbers must be unique within an agreement; duplicates: {', '.join(dupes)}.")
        ceiled = [math.ceil(a.quantity_kg) for a in self.annexes]
        total = sum(ceiled)
        deal_total = math.ceil(self.quantity_kg)
        if total != deal_total:
            details = " + ".join(str(w) for w in ceiled)
            raise ValueError(
                f"Sum of annex quantities (as shown in documents) "
                f"{details} = {total} kg does NOT match agreement "
                f"quantity_kg = {deal_total} kg (rendered). "
                f"Adjust annex quantities so they sum exactly after rounding-up."
            )
        return self

    # ─── Helpers ─────────────────────────────────────────────────

    def active_batches(self) -> list[int]:
        """Slot numbers whose batch_key_N is set."""
        return [n for n in BATCH_SLOTS if getattr(self, f"batch_key_{n}")]

    def resolved_annexes(self) -> list[AnnexLine]:
        """The effective list of annexes to render (always ≥ 1).

        Multi-annex deals return `annexes` as-is. Single-annex deals (empty list)
        get ONE annex from the deal-level fields (annex_num / annex_date /
        quantity_kg), already defaulted by the validators.
        """
        if self.annexes:
            return self.annexes
        assert self.annex_num is not None  # filled by _fill_defaults
        assert self.annex_date is not None
        return [AnnexLine(annex_num=self.annex_num, annex_date=self.annex_date, quantity_kg=self.quantity_kg)]

