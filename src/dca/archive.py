"""Output folder management and ZIP packaging.

Functions:
  archive_previous_output(deal_dir) — if `output/` exists, move it to `_archive/v{N}/`
  ensure_fresh_output(deal_dir)     — call above, then make a fresh empty `output/`
  zip_dir(src_dir, dest_path)       — zip the directory contents
"""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path


def archive_previous_output(
    deal_dir: Path | str, output_dir_name: str = "output", prefix: str = "v"
) -> Path | None:
    """If `<deal_dir>/<output_dir_name>/` exists and has real files, move it to `_archive/v{N}/`.

    Hidden files (dotfiles like macOS `.DS_Store`) are ignored when deciding
    whether the dir is "empty": a folder containing only `.DS_Store` is treated
    as empty (no archive cycle), and the file is wiped together with the dir.
    Without this, opening the deal folder in Finder creates `.DS_Store` and
    causes spurious `_archive/v1/`, `v2/`, ... full of just `.DS_Store`.

    Returns the path of the archived version directory, or None if nothing to archive.
    """
    deal_dir = Path(deal_dir)
    output_dir = deal_dir / output_dir_name
    if not output_dir.exists():
        return None

    real_contents = [p for p in output_dir.iterdir() if not p.name.startswith(".")]
    if not real_contents:
        # Effectively empty (only dotfiles like .DS_Store) — wipe and return
        shutil.rmtree(output_dir)
        return None

    archive_root = deal_dir / "_archive"
    archive_root.mkdir(exist_ok=True)

    # Find next <prefix>N slot
    existing_versions = []
    for sub in archive_root.iterdir():
        if sub.is_dir() and sub.name.startswith(prefix):
            try:
                existing_versions.append(int(sub.name[len(prefix):]))
            except ValueError:
                continue
    next_n = max(existing_versions, default=0) + 1
    dest = archive_root / f"{prefix}{next_n}"

    shutil.move(str(output_dir), str(dest))
    return dest


def ensure_fresh_output(deal_dir: Path | str, output_dir_name: str = "output") -> Path:
    """Archive previous output (if any) and return a fresh empty output dir path."""
    deal_dir = Path(deal_dir)
    archive_previous_output(deal_dir, output_dir_name)
    output_dir = deal_dir / output_dir_name
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


def zip_dir(src_dir: Path | str, dest_path: Path | str) -> Path:
    """Zip everything in `src_dir` (RECURSIVE — preserves subfolder structure)
    into `dest_path`.

    Subfolders (e.g. `attachments_<num>/`) are included as nested paths.
    The zip file itself is excluded if it lives inside `src_dir`.

    Returns dest_path.
    """
    src_dir = Path(src_dir).resolve()
    dest_path = Path(dest_path).resolve()
    if not src_dir.exists():
        raise FileNotFoundError(f"src_dir not found: {src_dir}")

    dest_path.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(dest_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(src_dir.rglob("*")):
            if not path.is_file():
                continue
            if path == dest_path:
                # Don't include the zip itself
                continue
            arcname = path.relative_to(src_dir)
            zf.write(path, arcname=str(arcname))
    return dest_path
