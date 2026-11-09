"""Attachment pack: gather scan-copies of the batch documents, rename them to
match the numbering of the cover list, and place them in
`output_<num>/attachments_<num>/`.

The renaming mirrors the Jinja counter inside `templates/cover_list.docx`.
Numbering is determined PURELY from `Deal` fields (no need to parse the
rendered cover list).

Cover-list order (must stay in sync with the template):
  1. соглашение                                        <- no file (Agreement_<num>.docx generated)
  for each active batch N (1..5):
    next. копия соглашения партии N                    <- batch_agreement_file_N
    [next. копия приложения к соглашению партии N]     <- only if batch_annex_num_N is set
  last.   карточка компании                            <- company_profile.<ext> in the company folder
"""

from __future__ import annotations

import re
import shutil
import unicodedata
from pathlib import Path
from typing import Any

from dca.models import Deal
from dca.utils import i18n

# ─── Helpers ────────────────────────────────────────────────────

# Characters that are unsafe in filenames on common filesystems
# (Windows is strictest; macOS/Linux are looser, but we target Windows-safe).
_FILENAME_BAD_CHARS = re.compile(r'[/\\:\*\?"<>\|]')

# Operator-managed scan in each company folder (any extension). Same wording as
# the last item in templates/cover_list.docx — keep them in sync.
COMPANY_PROFILE_FILENAME = "company_profile.pdf"
COMPANY_PROFILE_LABEL = "карточка компании"


def _sanitize(s: str) -> str:
    """Replace filesystem-unsafe chars with `-`. Also collapse whitespace runs."""
    s = _FILENAME_BAD_CHARS.sub("-", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _finder_view(fs_name: str) -> str:
    """macOS quirk: Finder shows `/` to the user, but HFS+/APFS stores it as `:`.

    `os.listdir` / `Path.rglob` return the FS form (with `:`). When the operator
    types the value into Excel they type what they see in Finder (with `/`).

    To compare correctly we convert FS-form → Finder-form by swapping `:` → `/`.
    This is a deterministic 1-to-1 mapping, NOT fuzzy guessing.
    """
    return fs_name.replace(":", "/")


def _canonical_name(name: str) -> str:
    """Index key — Finder-view + NFC normalization.

    APFS/HFS+ store filenames as NFD (decomposed: `й` = `и` + combining breve).
    Python string literals and Excel cells are usually NFC (precomposed `й`).
    The two are visually identical but bitwise different — a naive
    `dict[fs_name]` lookup misses every filename containing `й`, `ё`, etc.

    Normalize both sides to NFC so the lookup matches.
    """
    return unicodedata.normalize("NFC", _finder_view(name))


def list_attachment_items(deal: Deal) -> list[dict[str, Any]]:
    """Build the ordered list of files to copy for this deal.

    Each item: {"n": number in the cover list, "kind", "pkg": batch slot or None,
    "source_name": file name in Data/scans (None if the operator left it empty),
    "label": text of the cover-list line}. Items without `source_name` are still
    returned so the caller can warn.
    """
    def ru(d: Any) -> str:
        return i18n.fmt_ru(d) if d else "-"

    items: list[dict[str, Any]] = []
    n = 1                                   # 1. the agreement itself — generated, no file
    for pkg in deal.active_batches():
        num = getattr(deal, f"batch_agreement_num_{pkg}")
        date = getattr(deal, f"batch_agreement_date_{pkg}")
        annex_num = getattr(deal, f"batch_annex_num_{pkg}")
        annex_date = getattr(deal, f"batch_annex_date_{pkg}")
        n += 1
        items.append({
            "n": n, "kind": "agreement", "pkg": pkg,
            "source_name": getattr(deal, f"batch_agreement_file_{pkg}"),
            "label": f"копия соглашения №{num} от {ru(date)}",
        })
        if annex_num:
            n += 1
            items.append({
                "n": n, "kind": "annex", "pkg": pkg,
                "source_name": getattr(deal, f"batch_annex_file_{pkg}"),
                "label": f"копия приложения №{annex_num} от {ru(annex_date)} к соглашению №{num} от {ru(date)}",
            })
    n += 1                                  # last: the company profile from the company folder
    items.append({
        "n": n, "kind": "company_profile", "pkg": None,
        "source_name": COMPANY_PROFILE_FILENAME, "label": COMPANY_PROFILE_LABEL,
    })
    return items


def _build_filename_index(
    scans_dir: Path,
) -> tuple[dict[str, Path], dict[str, list[Path]]]:
    """Walk `scans_dir` recursively.

    Returns (index, duplicates):
      - index: {canonical_name: path}
        Key is the Finder-visible name (with `/` where the filesystem stores `:`)
        normalized to NFC (so `й` matches whether the file came from APFS-NFD or
        an Excel-typed NFC literal).
        For each name, the path stored is the **lexicographically smallest** full
        path (deterministic — same xlsx + same scans/ tree → same selection).
      - duplicates: {canonical_name: [path, path, ...]}
        All paths sharing a name (only present when count ≥ 2). Caller can use
        this to warn the operator about ambiguity before silently picking one.
    """
    index: dict[str, Path] = {}
    by_name: dict[str, list[Path]] = {}
    if not scans_dir.exists():
        return index, {}
    for p in scans_dir.rglob("*"):
        if p.is_file():
            name = _canonical_name(p.name)
            by_name.setdefault(name, []).append(p)
    duplicates: dict[str, list[Path]] = {}
    for name, paths in by_name.items():
        paths_sorted = sorted(paths)
        index[name] = paths_sorted[0]
        if len(paths_sorted) > 1:
            duplicates[name] = paths_sorted
    return index, duplicates


def _same_content(paths: list[Path]) -> bool:
    """True if all files are byte-for-byte identical (the same scan in several folders)."""
    import hashlib

    digests = {hashlib.sha256(p.read_bytes()).hexdigest() for p in paths}
    return len(digests) == 1


def _batch_folder(scans_dir: Path, folder_name: str | None) -> Path | None:
    """The batch's own scans folder inside scans_dir (None if not set / outside it)."""
    if not folder_name or not str(folder_name).strip():
        return None
    folder = (scans_dir / str(folder_name).strip()).resolve()
    try:
        folder.relative_to(scans_dir.resolve())
    except ValueError:
        return None
    return folder


def _resolve_scan(
    key: str, own_folder: Path | None, file_index: dict[str, Path], duplicates: dict[str, list[Path]]
) -> tuple[Path | None, list[Path]]:
    """Which file to copy for scan name `key` → (path, []) or (None, conflicting paths).

    One document may belong to several batches and lie, under the same name, in
    each of their folders — that is normal:
      1. the batch's own folder (batch_scans_folder) wins, whatever lies elsewhere;
      2. otherwise several files with this name are fine when their content is identical;
      3. different files with one name are still never guessed: (None, their paths).
    (None, []) = not found at all.
    """
    if own_folder is not None and own_folder.is_dir():
        own = sorted(p for p in own_folder.rglob("*") if p.is_file() and _canonical_name(p.name) == key)
        if own:
            return (own[0], []) if len(own) == 1 or _same_content(own) else (None, own)
    candidates = duplicates.get(key) or ([file_index[key]] if key in file_index else [])
    if not candidates:
        return None, []
    if len(candidates) > 1 and not _same_content(candidates):
        return None, candidates
    return candidates[0], []


def _company_profile(company_dir: Path) -> Path | None:
    """The company profile in a company folder: the fixed name, any extension."""
    if not company_dir.is_dir():
        return None
    want = _canonical_name(Path(COMPANY_PROFILE_FILENAME).stem)
    found = sorted(p for p in company_dir.iterdir() if p.is_file() and _canonical_name(p.stem) == want)
    return found[0] if found else None


def copy_attachment_files(
    deal: Deal, output_dir: Path, scans_dir: Path, company_dir: Path | None = None
) -> dict[str, Any]:
    """Copy + rename attachment files into `output_dir/attachments_<num>/`.

    Batch scans are looked up in `scans_dir` (recursively, by file name). The
    company profile comes from `company_dir` — party A's company folder
    — with any extension; without `company_dir` it is looked up in `scans_dir`.

    Returns a manifest:
      {
        "target_dir": Path,
        "copied":     [(source_name, target_name), ...],
        "missing":    [source_name, ...],     # listed in Excel but not found
        "skipped":    [item, ...],            # no source_name in Excel
      }
    """
    from dca.utils.slug import slug

    target_dir = output_dir / f"attachments_{slug(deal.agreement_num)}"
    target_dir.mkdir(parents=True, exist_ok=True)

    file_index, all_duplicates = _build_filename_index(scans_dir)

    copied: list[tuple[str, str]] = []
    missing: list[str] = []
    skipped: list[dict] = []
    # Only flag duplicates for filenames *this deal actually references* —
    # other duplicates in Data/scans are not the operator's problem right now.
    ambiguous: list[dict] = []
    missing_folders: list[str] = []

    for item in list_attachment_items(deal):
        src_name = item.get("source_name")
        if not src_name:
            skipped.append(item)
            continue
        if item["kind"] == "company_profile" and company_dir is not None:
            cert = _company_profile(company_dir)
            if cert is None:
                missing.append(f"{company_dir.name}/{src_name}")
                continue
            target_name = _sanitize(f"{item['n']}. {item['label']}") + cert.suffix
            shutil.copy2(cert, target_dir / target_name)
            copied.append((f"{company_dir.name}/{_canonical_name(cert.name)}", target_name))
            continue
        # The Excel value is the Finder-visible name (with `/`).
        # Index is keyed by canonical name (Finder-view + NFC), so we normalize
        # the lookup key too — otherwise NFC source vs NFD filesystem misses.
        lookup_key = _canonical_name(src_name)
        folder_name = getattr(deal, f"batch_scans_folder_{item['pkg']}", None) if item["pkg"] else None
        own_folder = _batch_folder(scans_dir, folder_name)
        if folder_name and (own_folder is None or not own_folder.is_dir()):
            note = f"партия #{item['pkg']}: {folder_name}"
            if note not in missing_folders:
                missing_folders.append(note)
        src_path, conflicting = _resolve_scan(lookup_key, own_folder, file_index, all_duplicates)
        if conflicting:
            # Different files share this name: we don't guess which one is the right
            # document — nothing is copied, the operator gets every path.
            ambiguous.append({
                "source_name": src_name,
                "picked":      None,
                "all_paths":   conflicting,
            })
            continue
        if src_path is None:
            missing.append(src_name)
            continue

        ext = src_path.suffix  # ".pdf", ".jpg", ...
        target_name = _sanitize(f"{item['n']}. {item['label']}") + ext
        target_path = target_dir / target_name
        shutil.copy2(src_path, target_path)
        # Report the canonical name (Finder-view + NFC) so the log is stable
        # regardless of whether the source filename was stored in NFD or NFC.
        copied.append((_canonical_name(src_path.name), target_name))

    return {
        "target_dir": target_dir,
        "copied":  copied,
        "missing": missing,
        "skipped": skipped,
        "ambiguous": ambiguous,
        "missing_folders": missing_folders,
    }
