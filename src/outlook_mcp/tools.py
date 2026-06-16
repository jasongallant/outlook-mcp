"""MCP tool definitions for Microsoft 365 (Outlook + Calendar + Contacts).

- ``get_user_profile`` uses Microsoft Graph (Mac Mail's OAuth grant has
  ``User.Read`` on Graph).
- Everything else uses Exchange Web Services via ``exchangelib`` because the
  Mac Mail grant has ``EWS.AccessAsUser.All`` (not Graph mail/calendar perms).

Read tools registered always. Write tools (send_email, reply, mark_read,
delete_message, create_event, update_event, respond_to_event) are gated
behind ``OUTLOOK_MCP_WRITES`` env var so they never fire by accident.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any

from exchangelib import (
    Account,
    CalendarItem,
    EWSDateTime,
    EWSTimeZone,
    HTMLBody,
    Mailbox,
    Message,
    Q,
)
from exchangelib.items import AcceptItem, DeclineItem, TentativelyAcceptItem
from mcp.server.fastmcp import FastMCP

from . import graph_client as g
from .config import get_timezone, writes_enabled
from .ews_client import with_account


def _msg_to_dict(m: Message) -> dict[str, Any]:
    sender = m.sender
    return {
        "id": m.id,
        "subject": m.subject,
        "from_name": sender.name if sender else None,
        "from_address": sender.email_address if sender else None,
        "received": m.datetime_received.isoformat() if m.datetime_received else None,
        "has_attachments": m.has_attachments,
        "is_read": m.is_read,
        "importance": m.importance,
        "preview": (m.text_body or "")[:280] if hasattr(m, "text_body") else None,
        "to": [r.email_address for r in (m.to_recipients or [])],
        "cc": [r.email_address for r in (m.cc_recipients or [])],
        "web_link": m.web_link if hasattr(m, "web_link") else None,
    }


def _msg_full(m: Message) -> dict[str, Any]:
    base = _msg_to_dict(m)
    base["body_text"] = m.text_body or (str(m.body) if m.body else "")
    return base


def _event_to_dict(e: CalendarItem) -> dict[str, Any]:
    organizer = e.organizer
    return {
        "id": e.id,
        "subject": e.subject,
        "start": e.start.isoformat() if e.start else None,
        "end": e.end.isoformat() if e.end else None,
        "location": getattr(e, "location", None),
        "organizer_name": organizer.name if organizer else None,
        "organizer_address": organizer.email_address if organizer else None,
        "is_online_meeting": getattr(e, "is_online_meeting", None),
        "online_join_url": getattr(e, "online_meeting_join_url", None),
        "response_status": getattr(e, "my_response_type", None),
        "show_as": getattr(e, "legacy_free_busy_status", None),
        "is_cancelled": getattr(e, "is_cancelled", None),
        "web_link": getattr(e, "web_link", None),
    }


def register_tools(mcp: FastMCP) -> None:
    # ---------------- READ ----------------

    @mcp.tool(description="Return the authenticated user's Microsoft 365 profile (via Graph).")
    def get_user_profile() -> dict:
        d = g.get("/me")
        return {
            "id": d.get("id"),
            "display_name": d.get("displayName"),
            "given_name": d.get("givenName"),
            "surname": d.get("surname"),
            "mail": d.get("mail"),
            "user_principal_name": d.get("userPrincipalName"),
            "job_title": d.get("jobTitle"),
            "office_location": d.get("officeLocation"),
        }

    @mcp.tool(description="List top-level mail folders (Inbox, Sent Items, Archive, etc.).")
    def list_mail_folders() -> list[dict]:
        def _act(a: Account) -> list[dict]:
            return [
                {
                    "id": f.id,
                    "display_name": f.name,
                    "total": f.total_count,
                    "unread": f.unread_count,
                    "child_folder_count": f.child_folder_count,
                }
                for f in a.msg_folder_root.children
            ]

        return with_account(_act)

    @mcp.tool(
        description=(
            "List recent messages. folder is one of inbox, sent, drafts, archive, junk. "
            "unread_only filters to is_read=false."
        ),
    )
    def list_messages(
        folder: str = "inbox",
        limit: int = 25,
        unread_only: bool = False,
    ) -> list[dict]:
        def _act(a: Account) -> list[dict]:
            f = {
                "inbox": a.inbox,
                "sent": a.sent,
                "drafts": a.drafts,
                "archive": getattr(a, "archive", None) or a.inbox,
                "junk": a.junk,
                "trash": a.trash,
            }.get(folder.lower(), a.inbox)
            qs = f.all().order_by("-datetime_received")
            if unread_only:
                qs = qs.filter(is_read=False)
            return [_msg_to_dict(m) for m in qs[:limit]]

        return with_account(_act)

    @mcp.tool(description="Fetch a single message by EWS id, including the plain-text body.")
    def get_message(message_id: str) -> dict:
        def _act(a: Account) -> dict:
            m = a.inbox.get(id=message_id)
            return _msg_full(m)

        return with_account(_act)

    @mcp.tool(
        description=(
            "Search messages by substring across subject + body. Pass a simple keyword query."
        ),
    )
    def search_messages(query: str, limit: int = 25) -> list[dict]:
        def _act(a: Account) -> list[dict]:
            qs = (
                a.inbox.filter(Q(subject__icontains=query) | Q(body__icontains=query))
                .order_by("-datetime_received")
            )
            return [_msg_to_dict(m) for m in qs[:limit]]

        return with_account(_act)

    @mcp.tool(description="List attachments on a message (metadata only — no content).")
    def list_attachments(message_id: str) -> list[dict]:
        def _act(a: Account) -> list[dict]:
            m = a.inbox.get(id=message_id)
            return [
                {
                    "name": getattr(att, "name", None),
                    "content_type": getattr(att, "content_type", None),
                    "size": getattr(att, "size", None),
                    "is_inline": getattr(att, "is_inline", None),
                }
                for att in (m.attachments or [])
            ]

        return with_account(_act)

    @mcp.tool(
        description=(
            "List calendar events between ISO timestamps "
            "(e.g. '2026-05-01T00:00:00'). Times interpreted in the timezone set in settings.json."
        ),
    )
    def list_events(start: str, end: str, limit: int = 50) -> list[dict]:
        def _act(a: Account) -> list[dict]:
            tz = EWSTimeZone(get_timezone())
            s = EWSDateTime.from_datetime(datetime.fromisoformat(start)).astimezone(tz)
            e = EWSDateTime.from_datetime(datetime.fromisoformat(end)).astimezone(tz)
            qs = a.calendar.view(start=s, end=e).order_by("start")
            return [_event_to_dict(ev) for ev in qs[:limit]]

        return with_account(_act)

    @mcp.tool(description="Fetch a single calendar event by EWS id.")
    def get_event(event_id: str) -> dict:
        def _act(a: Account) -> dict:
            ev = a.calendar.get(id=event_id)
            return _event_to_dict(ev)

        return with_account(_act)

    @mcp.tool(description="List Outlook contacts (People).")
    def list_contacts(limit: int = 50) -> list[dict]:
        def _act(a: Account) -> list[dict]:
            out: list[dict] = []
            for c in a.contacts.all()[:limit]:
                emails = []
                for ea in (c.email_addresses or []):
                    if ea.email:
                        emails.append(ea.email)
                out.append(
                    {
                        "id": c.id,
                        "display_name": c.display_name,
                        "given_name": getattr(c, "given_name", None),
                        "surname": getattr(c, "surname", None),
                        "emails": emails,
                        "company": getattr(c, "company_name", None),
                        "job_title": getattr(c, "job_title", None),
                    }
                )
            return out

        return with_account(_act)

    # ---------------- WRITE (gated) ----------------

    if not writes_enabled():
        return

    @mcp.tool(
        description=(
            "DESTRUCTIVE: Send an email from the user's mailbox. to/cc/bcc are lists of "
            "plain addresses. body_html=true treats body as HTML."
        ),
    )
    def send_email(
        to: list[str],
        subject: str,
        body: str,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        body_html: bool = False,
    ) -> dict:
        def _act(a: Account) -> dict:
            m = Message(
                account=a,
                subject=subject,
                body=HTMLBody(body) if body_html else body,
                to_recipients=[Mailbox(email_address=x) for x in to],
                cc_recipients=[Mailbox(email_address=x) for x in (cc or [])],
                bcc_recipients=[Mailbox(email_address=x) for x in (bcc or [])],
            )
            m.send_and_save()
            return {"sent": True, "to": to, "subject": subject}

        return with_account(_act)

    @mcp.tool(
        description=(
            "DESTRUCTIVE: Reply to a message. reply_all=True replies to everyone on thread. "
            "comment is the reply body text."
        ),
    )
    def reply_to_message(message_id: str, comment: str, reply_all: bool = False) -> dict:
        def _act(a: Account) -> dict:
            m = a.inbox.get(id=message_id)
            reply = m.create_reply_all(subject=f"RE: {m.subject}", body=comment) if reply_all else m.create_reply(subject=f"RE: {m.subject}", body=comment)
            reply.send()
            return {"replied": True, "reply_all": reply_all, "message_id": message_id}

        return with_account(_act)

    @mcp.tool(description="Mark a message read or unread. Reversible.")
    def mark_message_read(message_id: str, is_read: bool = True) -> dict:
        def _act(a: Account) -> dict:
            m = a.inbox.get(id=message_id)
            m.is_read = is_read
            m.save()
            return {"id": message_id, "is_read": is_read}

        return with_account(_act)

    @mcp.tool(
        description="DESTRUCTIVE: Move a message to Deleted Items (recoverable from trash).",
    )
    def delete_message(message_id: str) -> dict:
        def _act(a: Account) -> dict:
            m = a.inbox.get(id=message_id)
            m.move_to_trash()
            return {"deleted": True, "message_id": message_id}

        return with_account(_act)

    @mcp.tool(
        description=(
            "Create a calendar event on the primary calendar. start/end are ISO 8601 "
            "(local, in the timezone set in settings.json). attendees is a list of email addresses."
        ),
    )
    def create_event(
        subject: str,
        start: str,
        end: str,
        location: str | None = None,
        body: str | None = None,
        attendees: list[str] | None = None,
    ) -> dict:
        def _act(a: Account) -> dict:
            tz = EWSTimeZone(get_timezone())
            ev = CalendarItem(
                account=a,
                folder=a.calendar,
                subject=subject,
                start=EWSDateTime.from_datetime(datetime.fromisoformat(start)).astimezone(tz),
                end=EWSDateTime.from_datetime(datetime.fromisoformat(end)).astimezone(tz),
                location=location or "",
                body=body or "",
                required_attendees=[Mailbox(email_address=x) for x in (attendees or [])],
            )
            ev.save(send_meeting_invitations="SendOnlyToAll" if attendees else "SendToNone")
            return _event_to_dict(ev)

        return with_account(_act)

    @mcp.tool(
        description="DESTRUCTIVE: Patch fields on an existing event (subject, start, end, location, body).",
    )
    def update_event(
        event_id: str,
        subject: str | None = None,
        start: str | None = None,
        end: str | None = None,
        location: str | None = None,
        body: str | None = None,
    ) -> dict:
        def _act(a: Account) -> dict:
            ev = a.calendar.get(id=event_id)
            tz = EWSTimeZone(get_timezone())
            if subject is not None:
                ev.subject = subject
            if start is not None:
                ev.start = EWSDateTime.from_datetime(datetime.fromisoformat(start)).astimezone(tz)
            if end is not None:
                ev.end = EWSDateTime.from_datetime(datetime.fromisoformat(end)).astimezone(tz)
            if location is not None:
                ev.location = location
            if body is not None:
                ev.body = body
            ev.save(send_meeting_invitations_or_cancellations="SendOnlyToAll")
            return _event_to_dict(ev)

        return with_account(_act)

    @mcp.tool(
        description=(
            "DESTRUCTIVE: Respond to a meeting invite. response is one of "
            "accept, decline, tentativelyAccept."
        ),
    )
    def respond_to_event(event_id: str, response: str, comment: str = "") -> dict:
        if response not in {"accept", "decline", "tentativelyAccept"}:
            raise ValueError("response must be accept|decline|tentativelyAccept")

        def _act(a: Account) -> dict:
            ev = a.calendar.get(id=event_id)
            cls_map = {
                "accept": AcceptItem,
                "decline": DeclineItem,
                "tentativelyAccept": TentativelyAcceptItem,
            }
            cls_map[response](
                account=a,
                reference_item_id=ev.id,
                body=comment,
            ).send()
            return {"event_id": event_id, "response": response}

        return with_account(_act)
