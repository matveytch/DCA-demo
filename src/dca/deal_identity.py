"""Deal identity: `<deal folder>/.dca-deal.json` .

The folder name and the agreement number used to be the deal's only identity: a
renamed folder or a `from-log` run in another folder slipped past the duplicate
check, and a `deal_kind` row in Excel could turn an outgoing deal into an incoming one.

The identity file fixes both: a `deal_id` that never changes (renaming or moving
the folder keeps it; copying the folder copies it, so a copy is recognised), and
the deal's kind, decided by the command that created it — never by the Excel.
"""

from __future__ import annotations

import json
import re
import unicodedata
import uuid
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path

IDENTITY_FILE = ".dca-deal.json"
_FOLDER_RE = re.compile(r"^(?P<num>.+?)_(pending|incoming_pending|\d+kg)$")


@dataclass
class DealIdentity:
    deal_id: str
    agreement_num: str
    kind: str                 # "outgoing" | "incoming"
    created_at: str


def new_identity(agreement_num: str, kind: str) -> DealIdentity:
    return DealIdentity(uuid.uuid4().hex, agreement_num, kind, datetime.now().isoformat(timespec="seconds"))


def read_identity(deal_dir: Path) -> DealIdentity | None:
    path = deal_dir / IDENTITY_FILE
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return DealIdentity(**{k: str(data[k]) for k in ("deal_id", "agreement_num", "kind", "created_at")})
    except (OSError, ValueError, KeyError, TypeError) as e:
        raise ValueError(f"{path} is damaged ({e}) — restore it from a backup; do not delete it.") from e


def write_identity(deal_dir: Path, identity: DealIdentity) -> None:
    path = deal_dir / IDENTITY_FILE
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(asdict(identity), ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def number_key(agreement_num: str) -> str:
    """Comparison key of a agreement number: one per visually-equal spelling."""
    from dca.utils.slug import slug

    return slug(unicodedata.normalize("NFC", str(agreement_num))).casefold()


def known_number(deal_dir: Path) -> str | None:
    """The agreement number a deal folder holds: identity file → log.json → folder name."""
    try:
        identity = read_identity(deal_dir)
    except ValueError:
        identity = None
    if identity is not None:
        return identity.agreement_num
    for log in sorted(deal_dir.glob("output_*/log.json")):
        try:
            payload = json.loads(log.read_text(encoding="utf-8"))
            deal = payload.get("deal", payload)
            if deal.get("agreement_num"):
                return str(deal["agreement_num"])
        except (OSError, ValueError, AttributeError):
            continue
    m = _FOLDER_RE.match(deal_dir.name)
    return m.group("num") if m else None
