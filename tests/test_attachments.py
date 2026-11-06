"""Attachment path safety and local-save behaviour. No network."""

from __future__ import annotations

import io
import os
import stat
import time
from types import SimpleNamespace

import pytest
from exchangelib import FileAttachment, ItemAttachment
from exchangelib.attachments import AttachmentId

from outlook_mcp import attachments as at


@pytest.fixture
def home(tmp_path, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path))
    monkeypatch.setattr(
        at._cfg, "get_attachment_dir", lambda: tmp_path / ".cache" / "outlook-mcp" / "attachments"
    )
    return tmp_path


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("report.pdf", "report.pdf"),
        ("../../etc/passwd", "passwd"),
        ("..\\..\\boot.ini", "boot.ini"),
        ("/abs/path/x.txt", "x.txt"),
        (".hidden", "hidden"),
        ("..", "attachment"),
        ("", "attachment"),
        (None, "attachment"),
        ("a\x00b\nc.txt", "abc.txt"),
        ("evil‮gnp.exe", "evilgnp.exe"),
    ],
)
def test_safe_name(raw, expected):
    assert at.safe_name(raw) == expected


def test_safe_name_caps_length_in_bytes_and_keeps_extension():
    out = at.safe_name("é" * 500 + ".pdf")
    assert len(out.encode()) <= at.MAX_NAME_BYTES
    assert out.endswith(".pdf")


def test_attachment_to_dict_reports_id_and_kind():
    file_att = FileAttachment(
        name="a.pdf", content_type="application/pdf", attachment_id=AttachmentId(id="A1")
    )
    item_att = ItemAttachment(name="fwd", attachment_id=AttachmentId(id="A2"))
    assert at.attachment_to_dict(file_att)["attachment_id"] == "A1"
    assert at.attachment_to_dict(file_att)["kind"] == "file"
    assert at.attachment_to_dict(item_att)["kind"] == "item"


def _file_att(monkeypatch, name, data):
    att = FileAttachment(name=name, content_type="text/plain", attachment_id=AttachmentId(id="A1"))

    class _Stream:
        def __enter__(self):
            return io.BytesIO(data)

        def __exit__(self, *exc):
            return False

    monkeypatch.setattr(FileAttachment, "fp", property(lambda self: _Stream()))
    return att


def test_save_writes_private_file_inside_target(home, monkeypatch):
    out = at.save_attachment(_file_att(monkeypatch, "../../../evil.txt", b"hello"))
    path = home / ".cache" / "outlook-mcp" / "attachments"
    saved = next(path.glob("att-*/evil.txt"))
    assert out["path"] == str(saved)
    assert out["size"] == 5
    assert saved.read_bytes() == b"hello"
    assert stat.S_IMODE(saved.stat().st_mode) == 0o600
    assert stat.S_IMODE(saved.parent.stat().st_mode) == 0o700
    assert not list(path.glob("att-*/*.part"))


def test_save_failure_leaves_no_partial_file(home, monkeypatch):
    att = _file_att(monkeypatch, "x.txt", b"")

    def boom(_att, _out):
        raise OSError("connection lost")

    monkeypatch.setattr(at, "_write_content", boom)
    with pytest.raises(OSError):
        at.save_attachment(att)
    assert not list((home / ".cache").rglob("x.txt*"))


@pytest.mark.parametrize("mime", [b"Subject: hi\r\n", "Subject: hi\r\n"])
def test_save_item_attachment_as_eml(home, monkeypatch, mime):
    att = ItemAttachment(name="Forwarded", content_type=None, attachment_id=AttachmentId(id="A2"))
    monkeypatch.setattr(
        ItemAttachment, "item", property(lambda self: SimpleNamespace(mime_content=mime))
    )
    out = at.save_attachment(att)
    assert out["path"].endswith("/Forwarded.eml")
    assert out["size"] == len(b"Subject: hi\r\n")


def test_purge_removes_only_old_own_folders(tmp_path):
    now = time.time()
    old = now - at.RETENTION_SECONDS - 10
    own_old = tmp_path / "att-0123456789abcdef"
    own_new = tmp_path / "att-fedcba9876543210"
    other_old = tmp_path / "keep-me"
    near_miss = tmp_path / "att-notahash"
    for d in (own_old, own_new, other_old, near_miss):
        d.mkdir()
    for d in (own_old, other_old, near_miss):
        os.utime(d, (old, old))
    at.purge_old(tmp_path, now=now)
    assert not own_old.exists()
    assert own_new.exists()
    assert other_old.exists()
    assert near_miss.exists()


def test_purge_tolerates_missing_dir(tmp_path):
    at.purge_old(tmp_path / "nope")


@pytest.mark.parametrize(
    "rel", [".ssh/config", ".aws/credentials", ".config/outlook-mcp/settings.json", "a/.env"]
)
def test_send_refuses_hidden_paths_even_when_absent(home, rel):
    with pytest.raises(ValueError, match="hidden"):
        at.check_send_path(str(home / rel))


def test_send_refuses_library_even_when_absent(home):
    with pytest.raises(ValueError, match="Library"):
        at.check_send_path(str(home / "Library" / "Keychains" / "login.keychain-db"))


def test_send_refuses_symlink_into_hidden_path(home):
    secret = home / ".ssh" / "id_rsa"
    secret.parent.mkdir()
    secret.write_text("secret")
    link = home / "innocent.txt"
    link.symlink_to(secret)
    with pytest.raises(ValueError, match="hidden"):
        at.check_send_path(str(link))


@pytest.mark.parametrize("sub", ["Library/CloudStorage/OneDrive", "Library/Mobile Documents/x"])
def test_send_allows_cloud_documents_under_library(home, sub):
    doc = home / sub / "doc.pdf"
    doc.parent.mkdir(parents=True)
    doc.write_bytes(b"x")
    assert at.check_send_path(str(doc)) == doc.resolve()


def test_send_allows_previously_downloaded_file(home):
    doc = home / ".cache" / "outlook-mcp" / "attachments" / "att-0123456789abcdef" / "a.pdf"
    doc.parent.mkdir(parents=True)
    doc.write_bytes(b"x")
    assert at.check_send_path(str(doc)) == doc.resolve()


def test_send_refuses_missing_file_and_directory(home):
    with pytest.raises(ValueError, match="regular file"):
        at.check_send_path(str(home / "missing.pdf"))
    (home / "dir").mkdir()
    with pytest.raises(ValueError, match="regular file"):
        at.check_send_path(str(home / "dir"))


def test_file_attachments_builds_all_or_raises(home):
    good = home / "a.pdf"
    good.write_bytes(b"pdf")
    built = at.file_attachments([str(good)])
    assert [(a.name, a.content, a.content_type) for a in built] == [
        ("a.pdf", b"pdf", "application/pdf")
    ]
    assert at.file_attachments(None) == []
    with pytest.raises(ValueError):
        at.file_attachments([str(good), str(home / "missing.pdf")])
