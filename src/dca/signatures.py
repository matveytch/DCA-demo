"""Which version of the documents the signed files belong to (red team 2026-10-28, R03).

`signed_<num>/.signed_version.json` binds the files in `signed_<num>/` to the
documents that were published when those files appeared (or changed). After every
build the new documents are compared with that bound version: while any of them
differs — new data OR a changed template — the «signed for the PREVIOUS version»
warning repeats on every run, until the signed files are replaced.

Documents are compared by their content (the XML parts of the .docx), not by the
.docx bytes: the zip stores write times, so the same document differs byte-wise
on every render.
"""

from __future__ import annotations

import hashlib
import io
import json
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any

MANIFEST = ".signed_version.json"


def docx_fingerprint(data: bytes) -> str:
    h = hashlib.sha256()
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        for name in sorted(z.namelist()):
            if name.startswith("docProps/"):
                continue                      # creation / modification stamps
            h.update(name.encode() + b"\0" + z.read(name) + b"\0")
    return h.hexdigest()


def documents(folder: Path) -> dict[str, str]:
    """Relative path → fingerprint for every .docx in an output folder (incl. the transfer act)."""
    if not folder.is_dir():
        return {}
    out = {}
    for p in sorted(folder.rglob("*.docx")):
        if p.name.startswith(("~$", ".")):
            continue
        try:
            out[p.relative_to(folder).as_posix()] = docx_fingerprint(p.read_bytes())
        except (OSError, zipfile.BadZipFile):
            continue
    return out


def signed_files(signed_dir: Path) -> list[list[Any]]:
    return [[p.name, p.stat().st_size, hashlib.sha256(p.read_bytes()).hexdigest()]
            for p in sorted(signed_dir.iterdir()) if p.is_file() and not p.name.startswith(".")]


def bind(signed_dir: Path, published: Path) -> dict[str, Any] | None:
    """Before a build: (re)bind the signed files to the published output if they are
    new or changed. None when nothing is signed."""
    files = signed_files(signed_dir)
    if not files:
        return None
    path = signed_dir / MANIFEST
    try:
        manifest = json.loads(path.read_text(encoding="utf-8")) if path.exists() else None
    except (OSError, ValueError):
        manifest = None
    if manifest is None or manifest.get("signed_files") != files:
        manifest = {"signed_files": files, "documents": documents(published),
                    "bound_at": datetime.now().isoformat(timespec="seconds")}
        path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest


def stale_documents(manifest: dict[str, Any], built: Path, signed_dir: Path) -> list[str]:
    """After a build: the signed-for documents that are now different (or gone)."""
    bound: dict[str, str] = manifest.get("documents") or {}
    new = documents(built)
    if not bound:                      # signed before any output existed: bind to this build
        manifest["documents"] = new
        (signed_dir / MANIFEST).write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        return []
    return [name for name, fp in bound.items() if new.get(name) != fp]
