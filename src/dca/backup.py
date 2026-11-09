"""Encrypted, versioned backups of the operator's data with restic.

What is backed up: the data root (company folders, «Новые сделки», Data/, the
marker) — not the code folder (the code is on GitHub). Where: a restic repository,
by default in the Google Drive folder (`DCA-backups`), so the copy leaves the
Mac encrypted (AES-256) and Google Drive for desktop uploads it.

- restic (free, open source) deduplicates: daily versions cost little more than
  the data itself; retention is 7 daily + 4 weekly + 12 monthly.
- The password lives in the macOS Keychain (service `dca-backup`) for the
  scheduled run; the operator keeps a copy outside the Mac.
- `Data/_backup_status.json` records every run; `dca doctor` and `generate`
  warn only when the last run failed or the last good backup is too old.
- Restores go to a NEW folder only, never over the working data.

Sync (iCloud «Desktop & Documents») is not a backup: it copies deletions and
damage too. These versions are.
"""

from __future__ import annotations

import fnmatch
import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from dca.paths import DataRoot, marker_in

KEYCHAIN_SERVICE = "dca-backup"
KEYCHAIN_ACCOUNT = "dca"
REPO_DIR_NAME = "DCA-backups"
REPO_MARKER_KEY = "backup_repo"
STATUS_FILE = "_backup_status.json"      # in Data/
LOG_FILE = "_backup.log"                  # in Data/, output of the scheduled run
TAG = "dca"
KEEP = ("--keep-daily", "7", "--keep-weekly", "4", "--keep-monthly", "12")
STALE_DAYS = 3
SCHEDULE_LABEL = "com.dca.backup"
SCHEDULE_HOUR = 19
CODE_DIR_NAME = "dca"   # the code checkout inside the data root (code_dir in .dca.toml)
# Never worth keeping: Excel lock files, DCA lock files and half-built outputs, Finder junk.
EXCLUDES = ("~$*", "*.lock", "_сборка_*", ".DS_Store")   # _сборка_*: a crashed build folder
# A scheduled run gets a minimal PATH from launchd: look where Homebrew puts restic.
RESTIC_CANDIDATES = ("/opt/homebrew/bin/restic", "/usr/local/bin/restic")


class BackupError(RuntimeError):
    """A backup step failed; the message says what to do."""


# ─── configuration ─────────────────────────────────────────────


def find_restic() -> str | None:
    found = shutil.which("restic")
    if found:
        return found
    return next((c for c in RESTIC_CANDIDATES if Path(c).is_file()), None)


def google_drive_dirs() -> list[Path]:
    """«My Drive» folders of Google Drive for desktop (any account, any UI language)."""
    base = Path.home().joinpath("Library", "CloudStorage")
    if not base.is_dir():
        return []
    out = []
    for account in sorted(base.glob("GoogleDrive-*")):
        out += [d for d in (account.joinpath("My Drive"), account.joinpath("Мой диск")) if d.is_dir()]
    return out


def configured_repo(root: DataRoot) -> Path | None:
    """The repository path written by `dca backup --setup` into the data-root marker."""
    marker = marker_in(root.path)
    if marker is None:
        return None
    m = re.search(rf'^\s*{REPO_MARKER_KEY}\s*=\s*"(.*)"\s*$', marker.read_text(encoding="utf-8"), re.M)
    return Path(m.group(1)) if m else None


def save_repo(root: DataRoot, repo: Path) -> Path:
    """Write `backup_repo = "<path>"` into the data-root marker (replacing an old value)."""
    marker = marker_in(root.path) or root.path.joinpath(".dca.toml")
    text = marker.read_text(encoding="utf-8") if marker.exists() else ""
    line = f'{REPO_MARKER_KEY} = "{repo}"'
    pattern = re.compile(rf"^\s*{REPO_MARKER_KEY}\s*=.*$", re.M)
    if pattern.search(text):
        text = pattern.sub(line, text)
    else:
        text = text + ("" if not text or text.endswith("\n") else "\n") + line + "\n"
    marker.write_text(text, encoding="utf-8")
    return marker


def store_password() -> None:
    """Let `security` ask for the password itself (twice) and keep it in the Keychain.

    DCA never sees or prints the password. `-U` replaces an earlier one.
    """
    proc = subprocess.run(["security", "add-generic-password", "-U", "-s", KEYCHAIN_SERVICE,
                           "-a", KEYCHAIN_ACCOUNT, "-w"], check=False)
    if proc.returncode != 0:
        raise BackupError("пароль не сохранён в Связке ключей — запустите `dca backup --setup` ещё раз")


def _env(repo: Path) -> dict[str, str]:
    env = dict(os.environ)
    env["RESTIC_REPOSITORY"] = str(repo)
    if not (env.get("RESTIC_PASSWORD") or env.get("RESTIC_PASSWORD_FILE") or env.get("RESTIC_PASSWORD_COMMAND")):
        env["RESTIC_PASSWORD_COMMAND"] = f"security find-generic-password -s {KEYCHAIN_SERVICE} -a {KEYCHAIN_ACCOUNT} -w"
    return env


def _restic(repo: Path, *args: str, timeout: int = 6 * 3600) -> subprocess.CompletedProcess[str]:
    restic = find_restic()
    if restic is None:
        raise BackupError("restic не установлен — выполните в Терминале: brew install restic")
    proc = subprocess.run([restic, *args], env=_env(repo), capture_output=True, text=True,
                          timeout=timeout, check=False)
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout).strip().splitlines()
        raise BackupError(f"restic {args[0]}: {detail[-1] if detail else f'код {proc.returncode}'}")
    return proc


def init_repo(repo: Path) -> None:
    repo.mkdir(parents=True, exist_ok=True)
    if repo.joinpath("config").exists():
        return                                  # already a restic repository
    _restic(repo, "init")


# ─── what to back up ───────────────────────────────────────────


def _excluded(path: Path, root: DataRoot) -> bool:
    rel = path.relative_to(root.path)
    if rel.parts and rel.parts[0] == CODE_DIR_NAME:
        return True
    return any(fnmatch.fnmatch(part, pat) for part in rel.parts for pat in EXCLUDES)


def data_files(root: DataRoot) -> list[Path]:
    """Every file a backup holds, as paths under the data root."""
    return sorted(p for p in root.path.rglob("*") if p.is_file() and not _excluded(p, root))


def open_in_excel(root: DataRoot) -> list[str]:
    """Workbooks Excel holds open now (`~$` lock files): the backup gets their last SAVED state."""
    return sorted(p.relative_to(root.path).as_posix() for p in root.path.rglob("~$*")
                  if p.is_file() and not p.relative_to(root.path).parts[0] == CODE_DIR_NAME)


# ─── status ────────────────────────────────────────────────────


def status_path(root: DataRoot) -> Path:
    return root.data_dir.joinpath(STATUS_FILE)


def read_status(root: DataRoot) -> dict[str, Any]:
    try:
        return dict(json.loads(status_path(root).read_text(encoding="utf-8")))
    except (OSError, ValueError):
        return {}


def _write_status(root: DataRoot, status: dict[str, Any]) -> None:
    root.data_dir.mkdir(parents=True, exist_ok=True)
    status_path(root).write_text(json.dumps(status, ensure_ascii=False, indent=2), encoding="utf-8")


def stale_notice(root: DataRoot, now: datetime | None = None) -> str | None:
    """One line when backups need attention, else None.

    Only for a configured backup, and only when the last run failed or the last
    good backup is older than STALE_DAYS — never on every run (no banner blindness).
    """
    if configured_repo(root) is None:
        return None
    now = now or datetime.now()
    status = read_status(root)
    last_ok = status.get("last_success")
    if status and status.get("ok") is False:
        return f"БЭКАП НЕ УДАЛСЯ {status.get('last_attempt', '')[:16]}: {status.get('error')} — запустите dca backup"
    if not last_ok:
        return "БЭКАП ещё ни разу не выполнен — запустите dca backup"
    age = now - datetime.fromisoformat(last_ok)
    if age > timedelta(days=STALE_DAYS):
        return f"БЭКАП устарел: последний удачный {last_ok[:10]} ({age.days} дн. назад) — запустите dca backup"
    return None


# ─── backup / check / restore ──────────────────────────────────


def run_backup(root: DataRoot, now: datetime | None = None) -> dict[str, Any]:
    """Back up, apply retention, check the repository; record the result in the status file.

    Returns the new status. Raises BackupError (after recording the failure).
    """
    now = now or datetime.now()
    status = read_status(root)
    status.update(last_attempt=now.isoformat(timespec="seconds"), ok=False, error=None)
    try:
        repo = configured_repo(root)
        if repo is None:
            raise BackupError("бэкап не настроен — выполните dca backup --setup")
        excel = open_in_excel(root)
        args = ["backup", str(root.path), "--tag", TAG, "--json",
                "--exclude", str(root.path.joinpath(CODE_DIR_NAME))]
        for pat in EXCLUDES:
            args += ["--exclude", pat]
        proc = _restic(repo, *args)
        summary: dict[str, Any] = next((json.loads(line) for line in reversed(proc.stdout.splitlines())
                        if line.strip().startswith("{") and '"summary"' in line), {})
        _restic(repo, "forget", "--tag", TAG, *KEEP, "--prune")
        _restic(repo, "check")
    except (BackupError, OSError, subprocess.TimeoutExpired) as e:
        status["error"] = str(e)
        _write_status(root, status)
        raise BackupError(str(e)) from e
    # Precise start time: every file not modified after it is in this snapshot (restore_check).
    status.update(ok=True, last_success=status["last_attempt"], snapshot_time=now.isoformat(), repo=str(repo),
                  snapshot=summary.get("snapshot_id"), files=summary.get("total_files_processed"),
                  bytes=summary.get("total_bytes_processed"), excel_open=excel)
    _write_status(root, status)
    return status


def snapshots(root: DataRoot) -> list[dict[str, Any]]:
    repo = configured_repo(root)
    if repo is None:
        raise BackupError("бэкап не настроен — выполните dca backup --setup")
    return list(json.loads(_restic(repo, "snapshots", "--tag", TAG, "--json").stdout or "[]"))


def restore(root: DataRoot, snapshot: str, target: Path) -> Path:
    """Restore a snapshot into a NEW folder (never over the working data). → the restored data root."""
    target = target.resolve()
    if target == root.path or root.path in target.parents or target in root.path.parents:
        raise BackupError(f"восстанавливайте в отдельную папку вне {root.path} — рабочие данные не трогаются")
    if target.exists() and any(target.iterdir()):
        raise BackupError(f"папка {target} не пустая — укажите новую")
    repo = configured_repo(root)
    if repo is None:
        raise BackupError("бэкап не настроен — выполните dca backup --setup")
    target.mkdir(parents=True, exist_ok=True)
    _restic(repo, "restore", snapshot, "--target", str(target))
    # restic keeps absolute paths: the data root comes back under target/<its full path>.
    return target.joinpath(*root.path.parts[1:])


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def restore_check(root: DataRoot) -> dict[str, Any]:
    """Restore the latest snapshot into a temporary folder and compare it with the data.

    Files unchanged since the snapshot must be byte-identical; the ledger and the
    master database must open. The temporary folder is removed. → report.
    """
    from openpyxl import load_workbook

    status = read_status(root)
    taken = status.get("snapshot_time") or status.get("last_success")
    since = datetime.fromisoformat(taken) if taken else None
    tmp = Path(tempfile.mkdtemp(prefix="dca-restore-test-"))
    report: dict[str, Any] = {"checked": 0, "missing": [], "different": [], "unreadable": []}
    try:
        restored = restore(root, "latest", tmp.joinpath("restore"))
        for f in data_files(root):
            if since and datetime.fromtimestamp(f.stat().st_mtime) > since:
                continue                         # changed after the backup — newer than the copy
            rel = f.relative_to(root.path)
            copy = restored.joinpath(rel)
            report["checked"] += 1
            if not copy.is_file():
                report["missing"].append(rel.as_posix())
            elif _sha256(copy) != _sha256(f):
                report["different"].append(rel.as_posix())
        for book in (root.ledger_path, root.master_db_path):
            copy = restored.joinpath(book.relative_to(root.path))
            if copy.exists():
                try:
                    load_workbook(copy, read_only=True).close()
                except Exception as e:
                    report["unreadable"].append(f"{book.name}: {e}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    if report["checked"] == 0:
        report["unreadable"].append("ни одного файла для сравнения — сначала запустите dca backup")
    report["ok"] = not (report["missing"] or report["different"] or report["unreadable"])
    status = read_status(root)
    status["last_restore_test"] = {"at": datetime.now().isoformat(timespec="seconds"), "ok": report["ok"],
                                   "checked": report["checked"]}
    _write_status(root, status)
    return report


# ─── schedule (macOS LaunchAgent) ──────────────────────────────


def launch_agents_dir() -> Path:
    return Path(os.environ.get("DCA_LAUNCH_AGENTS_DIR") or Path.home().joinpath("Library", "LaunchAgents"))


def plist_path() -> Path:
    return launch_agents_dir().joinpath(f"{SCHEDULE_LABEL}.plist")


def plist_text(launcher: Path, root: DataRoot) -> str:
    log = root.data_dir.joinpath(LOG_FILE)
    path = "/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin"
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>{SCHEDULE_LABEL}</string>
  <key>ProgramArguments</key>
  <array><string>{launcher}</string><string>backup</string></array>
  <key>StartCalendarInterval</key>
  <dict><key>Hour</key><integer>{SCHEDULE_HOUR}</integer><key>Minute</key><integer>0</integer></dict>
  <key>WorkingDirectory</key><string>{root.path}</string>
  <key>EnvironmentVariables</key><dict><key>PATH</key><string>{path}</string></dict>
  <key>StandardOutPath</key><string>{log}</string>
  <key>StandardErrorPath</key><string>{log}</string>
</dict>
</plist>
"""


def _launchctl(*args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(["launchctl", *args], capture_output=True, text=True, check=False)


def install_schedule(launcher: Path, root: DataRoot) -> Path:
    """Write the LaunchAgent (every day at SCHEDULE_HOUR:00; a missed run happens after wake) and load it."""
    if not launcher.is_file():
        raise BackupError(f"нет скрипта запуска {launcher}")
    plist = plist_path()
    plist.parent.mkdir(parents=True, exist_ok=True)
    plist.write_text(plist_text(launcher, root), encoding="utf-8")
    domain = f"gui/{os.getuid()}"
    _launchctl("bootout", f"{domain}/{SCHEDULE_LABEL}")            # an older copy, if loaded
    proc = _launchctl("bootstrap", domain, str(plist))
    if proc.returncode != 0:
        raise BackupError(f"расписание записано ({plist}), но не загружено: {proc.stderr.strip()}")
    return plist


def remove_schedule() -> bool:
    plist = plist_path()
    _launchctl("bootout", f"gui/{os.getuid()}/{SCHEDULE_LABEL}")
    if plist.exists():
        plist.unlink()
        return True
    return False


def scheduled_launcher() -> Path | None:
    """The launcher the installed schedule runs (None if there is no schedule)."""
    plist = plist_path()
    if not plist.exists():
        return None
    m = re.search(r"<array><string>(.*?)</string>", plist.read_text(encoding="utf-8"))
    return Path(m.group(1)) if m else None
