"""Render .docx templates with a Deal via docxtpl.

Public API:
  render_template(template_path, deal)  → bytes
  render_all(deal)                       → dict[str, bytes]
  active_kinds_for(deal)                 → list[str]   # which templates to render
  render_transfer(deal)                  → dict[str, bytes]  # transfer act (on demand)
  render_amendment(deal)                 → bytes             # amendment (on demand)

The output filename for each template (based on agreement_num) is produced
by the helper `output_filename_for(kind, agreement_num)`.
"""

from __future__ import annotations

import io
import unicodedata
from pathlib import Path

import jinja2
from docxtpl import DocxTemplate

from dca.config import AMENDMENT_TEMPLATE, OUTPUT_FILES, TEMPLATE_FILES, TEMPLATES_DIR
from dca.context import build_annex_contexts, build_render_context
from dca.models import Deal

# ─── Public API ─────────────────────────────────────────────────


def active_kinds_for(deal: Deal) -> list[str]:
    """Which template kinds to render for this deal.

    Outgoing deals: agreement, annex×N, cover list.
    Incoming deals: agreement, annex×N (the agreement template switches its
    wording by deal_kind).
    """
    if deal.deal_kind == "incoming":
        return ["agreement", "annex"]
    return list(TEMPLATE_FILES)


def render_with_context(template_path: Path | str, context: dict) -> bytes:  # type: ignore[type-arg]
    """Render one template with an already-built context dict; return bytes (.docx).

    Strict: a template variable the context doesn't define is an error (Jinja would
    silently print nothing — a typo or a field dropped from the context would make
    a party's name vanish from the document).
    """
    template_path = Path(template_path)
    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")

    doc = DocxTemplate(str(template_path))
    missing = sorted(v for v in doc.get_undeclared_template_variables() if v not in context)
    if missing:
        raise ValueError(
            f"Template «{template_path.name}» uses variables the engine does not provide: "
            f"{', '.join(missing)} — a typo in the template, or a field missing from the data."
        )
    # StrictUndefined: a nested name ({{ b.typo }}) is an error too, not an empty
    # string — the check above only sees top-level names.
    try:
        doc.render(context, jinja_env=jinja2.Environment(undefined=jinja2.StrictUndefined))
    except jinja2.UndefinedError as e:
        raise ValueError(f"Template «{template_path.name}»: {e} — a typo in the template, "
                         f"or a field missing from the data.") from e

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


def render_template(template_path: Path | str, deal: Deal) -> bytes:
    """Render one template with the deal's (aggregate) data; return bytes (.docx)."""
    if not Path(template_path).exists():
        raise FileNotFoundError(f"Template not found: {template_path}")
    return render_with_context(template_path, build_render_context(deal))


def render_all(deal: Deal) -> dict[str, bytes]:
    """Render all documents of the deal's package.

    Non-annex documents render ONCE with the agreement-aggregate context. The
    annex renders ONCE PER annex (deal.resolved_annexes()), each with its own
    number/date/quantity/total — one file per annex.
    """
    out: dict[str, bytes] = {}
    for kind in active_kinds_for(deal):
        if kind == "annex":
            out.update(render_annex_set(deal))
        else:
            out[output_filename_for(kind, deal.agreement_num)] = render_template(
                TEMPLATES_DIR / TEMPLATE_FILES[kind], deal
            )
    return out


def render_annex_set(deal: Deal) -> dict[str, bytes]:
    """Render the annex document(s) — one file per resolved annex.

    Single annex → `Annex_<num>.docx`; several → one `Annex_<annexnum>_<num>.docx` each.
    """
    template_path = TEMPLATES_DIR / TEMPLATE_FILES["annex"]
    annexes = deal.resolved_annexes()
    names = [_annex_filename(deal.agreement_num, a.annex_num, single=len(annexes) == 1) for a in annexes]
    # Distinct numbers can collapse into ONE file name (A/B and A:B → A_B; case on macOS):
    # the second document would silently replace the first.
    seen: dict[str, str] = {}
    for annex, name in zip(annexes, names, strict=True):
        key = unicodedata.normalize("NFC", name).casefold()   # as the file system compares
        if key in seen:
            raise ValueError(
                f"Annexes «{seen[key]}» and «{annex.annex_num}» would be saved under "
                f"the same file name {name} — give them numbers that differ in letters or digits."
            )
        seen[key] = annex.annex_num
    out: dict[str, bytes] = {}
    for name, actx in zip(names, build_annex_contexts(deal), strict=True):
        out[name] = render_with_context(template_path, actx)
    return out


def output_filename_for(kind: str, agreement_num: str) -> str:
    """Filename for a rendered document (kind: agreement / annex / cover_list / log / package)."""
    if kind not in OUTPUT_FILES:
        raise ValueError(f"Unknown template kind: {kind!r}. Known: {list(OUTPUT_FILES)}")
    from dca.utils.slug import slug
    return OUTPUT_FILES[kind].format(num=slug(agreement_num))


def _annex_filename(agreement_num: str, annex_num: str, *, single: bool) -> str:
    """Single-annex deal → `Annex_<agreement>.docx`; several → `Annex_<annexnum>_<agreement>.docx`."""
    from dca.utils.slug import slug
    safe = slug(agreement_num)
    if single:
        return OUTPUT_FILES["annex"].format(num=safe)
    return f"Annex_{slug(annex_num)}_{safe}.docx"


def render_amendment(deal: Deal) -> bytes:
    """Render the amendment; rendered on demand by `dca amend`.

    Every field comes from the deal and its «Изменение» sheet — old/new quantity,
    price, totals in words, clause numbers (build_amendment_context). Raises
    ValueError listing every problem when the sheet is empty or incomplete.
    """
    from dca.context import build_amendment_context

    return render_with_context(TEMPLATES_DIR / AMENDMENT_TEMPLATE, build_amendment_context(deal))


def render_transfer(deal: Deal) -> dict[str, bytes]:
    """Render the transfer act → {filename: bytes}.

    Raises ValueError (from build_transfer_context) when the «Передача» sheet is
    incomplete or a text doesn't fit the form — never renders a partial act.
    """
    from dca.config import TRANSFER_OUTPUT, TRANSFER_TEMPLATE
    from dca.context import build_transfer_context
    from dca.utils.slug import slug

    ctx = build_transfer_context(deal)
    return {TRANSFER_OUTPUT.format(num=slug(deal.agreement_num)):
            render_with_context(TEMPLATES_DIR / TRANSFER_TEMPLATE, ctx)}
