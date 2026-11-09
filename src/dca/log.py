"""JSON snapshot of a Deal — dump and load for re-generation.

`dump_log(deal, path)` — serialize Deal to JSON.
`load_log(path)` → Deal — read JSON snapshot, validate, return Deal.

Use case: operator runs `dca generate`, log.json is written alongside docs.
Later, to re-generate from same data: `dca from-log path/to/log.json`.
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

from dca.models import Deal

# Raise when Deal gains a field: an older DCA (extra="forbid") can't read such a log,
# so it refuses it with «update DCA» instead of a validation error.
LOG_SCHEMA_VERSION = 1


def dump_log(
    deal: Deal, path: Path | str, input_xlsx: Path | None = None, deal_id: str | None = None
) -> None:
    """Write Deal as pretty-printed JSON. Dates serialized as ISO strings.

    `_metadata` records what produced the package (audit 2026-10-28 A19): schema
    and DCA version, git commit of the code, sha256 of every template and of the
    input Excel — so a later re-render can be checked against the original.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    data = deal.model_dump(mode="json")
    # model_dump(mode="json") already converts date → "YYYY-MM-DD"; good.
    payload = {
        "_metadata": {
            "generator": "DCA",
            "deal_id": deal_id,
            "schema_version": LOG_SCHEMA_VERSION,
            "dca_version": _dca_version(),
            "code_commit": _git_commit(),
            "saved_at": datetime.now().isoformat(timespec="seconds"),
            "templates_sha256": _templates_sha256(),
            "input_file": input_xlsx.name if input_xlsx else None,
            "input_sha256": _sha256(input_xlsx) if input_xlsx and input_xlsx.exists() else None,
        },
        "deal": data,
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _sha256(path: Path) -> str:
    import hashlib
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _templates_sha256() -> dict[str, str]:
    from dca.config import TEMPLATES_DIR
    return {p.name: _sha256(p) for p in sorted(TEMPLATES_DIR.glob("*.docx")) if not p.name.startswith("~$")}


def _dca_version() -> str:
    import dca
    return str(getattr(dca, "__version__", "?"))


def _git_commit() -> str | None:
    """Commit of the code that produced the package (None outside a git checkout)."""
    import subprocess

    from dca.config import PROJECT_ROOT
    try:
        out = subprocess.run(["git", "-C", str(PROJECT_ROOT), "rev-parse", "HEAD"],
                             capture_output=True, text=True, timeout=5)
    except (OSError, subprocess.SubprocessError):
        return None
    return out.stdout.strip() or None if out.returncode == 0 else None


def load_log(path: Path | str) -> Deal:
    """Read JSON snapshot, validate, return Deal (see `load_log_with_metadata`)."""
    return load_log_with_metadata(path)[0]


def load_log_with_metadata(path: Path | str) -> tuple[Deal, dict[str, Any]]:
    """Read JSON snapshot, validate, return Deal.

    Accepts two formats:
      1. Wrapped (current `dump_log` output): `{"_metadata": {...}, "deal": {...}}`
      2. Raw Deal dict (legacy / hand-written): `{"agreement_num": ..., ...}`

    A file that is JSON but neither shape (e.g. a list, a string, a dict with
    `_metadata` but no `deal` key) raises ValueError with a clear message —
    Pydantic's own error in that case would be a wall of cryptic "Field required"
    lines, which is not useful for an operator who fat-fingered the path.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Log file not found: {path}")

    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise ValueError(f"Log file is not valid JSON: {path} — {e}") from e

    if not isinstance(raw, dict):
        raise ValueError(
            f"Log file must contain a JSON object at top level, got "
            f"{type(raw).__name__}: {path}"
        )

    # Determine format
    metadata: dict[str, Any] = {}
    if "_metadata" in raw or "deal" in raw:
        # Wrapped format expected — "deal" key MUST be present
        deal_data = raw.get("deal")
        if not isinstance(deal_data, dict):
            raise ValueError(
                f"Log file looks wrapped but has no valid 'deal' object: {path}"
            )
        metadata = raw.get("_metadata") if isinstance(raw.get("_metadata"), dict) else {}
        version = metadata.get("schema_version")
        # A snapshot from a newer DCA may mean fields this version reads wrongly —
        # never accept it silently (red team R11).
        if isinstance(version, int) and version > LOG_SCHEMA_VERSION:
            raise ValueError(
                f"log.json schema_version {version} is newer than this DCA supports "
                f"({LOG_SCHEMA_VERSION}) — update DCA (git pull) before re-rendering: {path}"
            )
    else:
        # Raw Deal dict — minimal sanity check
        if "agreement_num" not in raw:
            raise ValueError(
                f"Log file is a JSON object but missing 'agreement_num' — "
                f"doesn't look like a Deal snapshot: {path}"
            )
        deal_data = raw

    return Deal(**deal_data), metadata


def changed_templates(metadata: dict[str, Any]) -> list[str]:
    """Templates whose sha256 differs from the snapshot's (or that are gone now)."""
    saved = metadata.get("templates_sha256")
    if not isinstance(saved, dict):
        return []
    now = _templates_sha256()
    return sorted(name for name, digest in saved.items() if now.get(name) != digest)
