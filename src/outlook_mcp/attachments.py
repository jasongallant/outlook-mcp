"""Attachment helpers: safe local saves for downloads, path checks for sends.

Attachment names come from whoever sent the mail, so they are treated as
hostile when building local paths. Downloads can hold sensitive records, so
they are private to the user and purged after ``RETENTION_SECONDS``.
"""

from __future__ import annotations

import contextlib
import hashlib
import mimetypes
import os
import re
import shutil
import time
from pathlib import Path
from typing import Any, BinaryIO, Iterable

from exchangelib import FileAttachment, ItemAttachment

from . import config as _cfg

MAX_NAME_BYTES = 200
RETENTION_SECONDS = 7 * 24 * 3600
_SUBDIR_RE = re.compile(r"^att-[0-9a-f]{16}$")
# Under ~/Library, but these hold the user's documents (OneDrive, iCloud Drive).
_LIBRARY_DOCUMENT_DIRS = ("CloudStorage", "Mobile Documents")


def attachment_to_dict(att: Any) -> dict[str, Any]:
    """Metadata for one attachment. Reads nothing from the server."""
    return {
        "attachment_id": att.attachment_id.id if att.attachment_id else None,
        "name": getattr(att, "name", None),
        "content_type": getattr(att, "content_type", None),
        "size": getattr(att, "size", None),
        "is_inline": getattr(att, "is_inline", None),
        "kind": "file" if isinstance(att, FileAttachment) else "item",
    }


def safe_name(name: str | None, fallback: str = "attachment") -> str:
    """Reduce a sender-supplied name to a single harmless path component."""
    base = (name or "").replace("\\", "/").rsplit("/", 1)[-1]
    cleaned = "".join(c for c in base if c.isprintable()).strip(" .")
    if not cleaned:
        return fallback
    if len(cleaned.encode()) <= MAX_NAME_BYTES:
        return cleaned
    stem, ext = os.path.splitext(cleaned)
    ext = ext.encode()[:20].decode(errors="ignore")
    room = MAX_NAME_BYTES - len(ext.encode())
    return stem.encode()[:room].decode(errors="ignore") + ext


def purge_old(base: Path, now: float | None = None) -> None:
    """Remove this server's own download folders past retention. Touches nothing else."""
    cutoff = (now if now is not None else time.time()) - RETENTION_SECONDS
    with contextlib.suppress(OSError), os.scandir(base) as entries:
        for entry in entries:
            if not _SUBDIR_RE.match(entry.name) or not entry.is_dir(follow_symlinks=False):
                continue
            with contextlib.suppress(OSError):
                if entry.stat(follow_symlinks=False).st_mtime < cutoff:
                    shutil.rmtree(entry.path)


def _write_content(att: Any, out: BinaryIO) -> None:
    if isinstance(att, FileAttachment):
        with att.fp as fp:  # streams, so large files stay out of memory
            shutil.copyfileobj(fp, out)
        return
    mime = att.item.mime_content
    if not mime:
        raise ValueError(f"Attached item {att.name!r} has no MIME content to save.")
    out.write(mime.encode() if isinstance(mime, str) else mime)


def save_attachment(att: Any) -> dict[str, Any]:
    """Download one attachment to ``attachment_dir`` and return where it landed."""
    base = _cfg.get_attachment_dir()
    base.mkdir(parents=True, exist_ok=True, mode=0o700)
    purge_old(base)

    key = hashlib.sha256(att.attachment_id.id.encode()).hexdigest()[:16]
    target = base / f"att-{key}"
    target.mkdir(exist_ok=True, mode=0o700)

    name = att.name or "attachment"
    if isinstance(att, ItemAttachment) and not name.lower().endswith(".eml"):
        name += ".eml"
    path = target / safe_name(name)
    if path.resolve().parent != target.resolve():
        raise ValueError(f"Refusing to save outside {target}: {att.name!r}")

    tmp = path.with_name(path.name + ".part")
    fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC | os.O_NOFOLLOW, 0o600)
    try:
        with os.fdopen(fd, "wb") as out:
            _write_content(att, out)
        os.replace(tmp, path)
    except BaseException:
        with contextlib.suppress(OSError):
            tmp.unlink()
        raise
    return {
        "path": str(path),
        "name": att.name,
        "content_type": att.content_type,
        "size": path.stat().st_size,
    }


def check_send_path(raw: str) -> Path:
    """Resolve a local path to attach, refusing common secret locations.

    A malicious email could tell an agent to attach and send a local file. This
    blocks the usual targets, but not every sensitive file. The client's approval
    prompt on the send tools is the real protection.
    """
    resolved = Path(raw).expanduser().resolve()
    home = Path.home().resolve()
    # Files this server downloaded live under a hidden cache folder by default.
    downloaded = resolved.is_relative_to(_cfg.get_attachment_dir().resolve())
    if not downloaded and any(p.startswith(".") for p in resolved.parts):
        raise ValueError(f"Refusing to attach a file under a hidden path: {raw}")
    library = home / "Library"
    if (
        not downloaded
        and resolved.is_relative_to(library)
        and not any(resolved.is_relative_to(library / d) for d in _LIBRARY_DOCUMENT_DIRS)
    ):
        raise ValueError(f"Refusing to attach a file under ~/Library: {raw}")
    if not resolved.is_file():
        raise ValueError(f"Attachment is not a regular file: {raw}")
    return resolved


def file_attachments(paths: Iterable[str] | None) -> list[FileAttachment]:
    """Build attachments from local paths. Raises before anything is created or sent."""
    checked = [check_send_path(p) for p in (paths or [])]
    return [
        FileAttachment(
            name=path.name,
            content=path.read_bytes(),
            content_type=mimetypes.guess_type(path.name)[0] or "application/octet-stream",
        )
        for path in checked
    ]
