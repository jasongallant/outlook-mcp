"""MCP tool definitions for Microsoft 365 (Outlook + Calendar + Contacts).

- ``get_user_profile`` uses Microsoft Graph (Mac Mail's OAuth grant has
  ``User.Read`` on Graph).
- Everything else — mail, calendar, contacts, and Microsoft To Do — uses
  Exchange Web Services via ``exchangelib`` because the Mac Mail grant has
  ``EWS.AccessAsUser.All`` but no Graph mail/calendar/Tasks perms (the Graph
  ``/me/todo/*`` endpoints 403; To Do lists are plain IPF.Task mailbox folders
  reachable over EWS).

Read tools registered always. Write tools (send_email, reply, mark_read,
delete_message, create_event, update_event, respond_to_event, create_task,
update_task, complete_task, delete_task, create_todo_list) are gated behind
``OUTLOOK_MCP_WRITES`` env var so they never fire by accident.
"""

from __future__ import annotations

from datetime import date, datetime, time
from decimal import Decimal
from typing import Any

from exchangelib import (
    Account,
    CalendarItem,
    EWSDate,
    EWSDateTime,
    EWSTimeZone,
    HTMLBody,
    Mailbox,
    Message,
    Q,
    Task,
)
from exchangelib.folders import Folder
from exchangelib.items import AcceptItem, DeclineItem, TentativelyAcceptItem

# EWS Task DueDate/StartDate are date-backed by a datetime. exchangelib sends
# that datetime at UTC midnight, which a server in a negative-offset mailbox
# timezone rolls back to the previous calendar day on readback (and in the
# To Do app). Backing the date with noon UTC keeps it on the intended day for
# every real-world timezone (UTC-11 .. UTC+11). Matches how the To Do app
# stores due dates.
for _fname in ("due_date", "start_date"):
    Task.get_field_by_fieldname(_fname)._default_time = time(12, 0)
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


# ---- Microsoft To Do (EWS Task folders) ----------------------------------
#
# To Do lists are plain IPF.Task folders in the mailbox: the default "Tasks"
# folder (``account.tasks``) plus its child folders (one per custom list).
# Reached over EWS like mail/calendar — the Mac Mail Graph grant has no
# Tasks.ReadWrite scope, so the Graph /me/todo/* endpoints 403.

_IMPORTANCE = {"low": "Low", "normal": "Normal", "high": "High"}


def _todo_folders(a: Account) -> list[Folder]:
    """Default Tasks folder first, then each child task folder (a To Do list)."""
    folders = [a.tasks]
    for child in a.tasks.children:
        if str(getattr(child, "folder_class", "") or "").startswith("IPF.Task"):
            folders.append(child)
    return folders


def _resolve_list(a: Account, name: str | None) -> Folder:
    """Return the task folder for a list name (case-insensitive). None → default."""
    if not name or not name.strip():
        return a.tasks
    want = name.strip().lower()
    for f in _todo_folders(a):
        if f.name.lower() == want:
            return f
    available = ", ".join(repr(f.name) for f in _todo_folders(a))
    raise ValueError(f"To Do list {name!r} not found. Available: {available}")


def _parse_due(s: str) -> EWSDate:
    """Parse a due date. Accepts 'YYYY-MM-DD' or a full ISO datetime (date part used)."""
    return EWSDate.from_date(date.fromisoformat(s[:10]))


def _parse_reminder(s: str) -> EWSDateTime:
    """Parse an ISO datetime reminder into the configured timezone."""
    tz = EWSTimeZone(get_timezone())
    return EWSDateTime.from_datetime(datetime.fromisoformat(s)).astimezone(tz)


def _task_to_dict(t: Task, list_name: str | None = None) -> dict[str, Any]:
    due = t.due_date
    reminder = t.reminder_due_by if t.reminder_is_set else None
    return {
        "id": t.id,
        "subject": t.subject,
        "list": list_name if list_name is not None else getattr(getattr(t, "folder", None), "name", None),
        "status": t.status,
        "is_complete": bool(getattr(t, "is_complete", None)),
        "percent_complete": int(t.percent_complete) if t.percent_complete is not None else None,
        "importance": t.importance,
        "due": due.isoformat() if due else None,
        "start": t.start_date.isoformat() if t.start_date else None,
        "reminder": reminder.isoformat() if reminder else None,
        "complete_date": t.complete_date.isoformat() if t.complete_date else None,
        "categories": list(t.categories) if t.categories else [],
    }


def _task_full(t: Task, list_name: str | None = None) -> dict[str, Any]:
    base = _task_to_dict(t, list_name)
    base["body_text"] = t.text_body or (str(t.body) if t.body else "")
    return base


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

    # ---------------- TO DO (read) ----------------

    @mcp.tool(
        description=(
            "List Microsoft To Do lists (the default Tasks list plus any custom lists), "
            "with per-list open and total task counts."
        ),
    )
    def list_todo_lists() -> list[dict]:
        def _act(a: Account) -> list[dict]:
            out = []
            for f in _todo_folders(a):
                open_count = f.filter(is_complete=False).count()
                out.append(
                    {
                        "name": f.name,
                        "is_default": f.id == a.tasks.id,
                        "open": open_count,
                        "total": f.total_count,
                    }
                )
            return out

        return with_account(_act)

    @mcp.tool(
        description=(
            "List tasks in a Microsoft To Do list. list is the list name (omit for the "
            "default Tasks list). include_completed=false hides finished tasks. Newest first."
        ),
    )
    def list_tasks(
        list: str | None = None,
        include_completed: bool = False,
        limit: int = 50,
    ) -> list[dict]:
        def _act(a: Account) -> list[dict]:
            f = _resolve_list(a, list)
            qs = f.all().order_by("-last_modified_time")
            if not include_completed:
                qs = qs.filter(is_complete=False)
            return [_task_to_dict(t, f.name) for t in qs[:limit]]

        return with_account(_act)

    @mcp.tool(description="Fetch a single To Do task by EWS id, including its notes/body.")
    def get_task(task_id: str) -> dict:
        def _act(a: Account) -> dict:
            t = a.tasks.get(id=task_id)
            return _task_full(t)

        return with_account(_act)

    @mcp.tool(
        description=(
            "Search To Do tasks by subject/body substring. Searches the default Tasks list "
            "unless list names a specific list. Omits completed tasks unless include_completed=true."
        ),
    )
    def search_tasks(
        query: str,
        list: str | None = None,
        include_completed: bool = False,
        limit: int = 25,
    ) -> list[dict]:
        def _act(a: Account) -> list[dict]:
            f = _resolve_list(a, list)
            qs = f.filter(Q(subject__icontains=query) | Q(body__icontains=query)).order_by(
                "-last_modified_time"
            )
            if not include_completed:
                qs = qs.filter(is_complete=False)
            return [_task_to_dict(t, f.name) for t in qs[:limit]]

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

    # ---------------- TO DO (write, gated) ----------------

    @mcp.tool(
        description=(
            "Create a Microsoft To Do task. list is the target list name (omit for the "
            "default Tasks list). due is a date 'YYYY-MM-DD'. reminder is an ISO datetime "
            "(local, settings.json timezone). importance is low|normal|high."
        ),
    )
    def create_task(
        title: str,
        list: str | None = None,
        due: str | None = None,
        reminder: str | None = None,
        body: str | None = None,
        importance: str = "normal",
    ) -> dict:
        imp = _IMPORTANCE.get(importance.lower())
        if imp is None:
            raise ValueError("importance must be low|normal|high")

        def _act(a: Account) -> dict:
            f = _resolve_list(a, list)
            t = Task(account=a, folder=f, subject=title, importance=imp)
            if body is not None:
                t.body = body
            if due is not None:
                t.due_date = _parse_due(due)
            if reminder is not None:
                t.reminder_is_set = True
                t.reminder_due_by = _parse_reminder(reminder)
            t.save()
            # Re-fetch: server derives status/percent/is_complete; the saved
            # in-memory object does not reflect them.
            return _task_to_dict(a.tasks.get(id=t.id), f.name)

        return with_account(_act)

    @mcp.tool(
        description=(
            "DESTRUCTIVE: Patch fields on a To Do task (title, due 'YYYY-MM-DD', reminder "
            "ISO datetime, body, importance low|normal|high, status "
            "notStarted|inProgress|completed|waitingOnOthers|deferred)."
        ),
    )
    def update_task(
        task_id: str,
        title: str | None = None,
        due: str | None = None,
        reminder: str | None = None,
        body: str | None = None,
        importance: str | None = None,
        status: str | None = None,
    ) -> dict:
        status_map = {
            "notstarted": "NotStarted",
            "inprogress": "InProgress",
            "completed": "Completed",
            "waitingonothers": "WaitingOnOthers",
            "deferred": "Deferred",
        }
        imp = None
        if importance is not None:
            imp = _IMPORTANCE.get(importance.lower())
            if imp is None:
                raise ValueError("importance must be low|normal|high")
        st = None
        if status is not None:
            st = status_map.get(status.lower())
            if st is None:
                raise ValueError(f"status must be one of {', '.join(status_map)}")

        def _act(a: Account) -> dict:
            t = a.tasks.get(id=task_id)
            if title is not None:
                t.subject = title
            if body is not None:
                t.body = body
            if imp is not None:
                t.importance = imp
            if due is not None:
                t.due_date = _parse_due(due)
            if reminder is not None:
                t.reminder_is_set = True
                t.reminder_due_by = _parse_reminder(reminder)
            if st is not None:
                # Setting status drives percent_complete server-side; only
                # reopening needs an explicit reset (percent is a Decimal field).
                t.status = st
                if st == "NotStarted":
                    t.percent_complete = Decimal(0)
            t.save()
            return _task_to_dict(a.tasks.get(id=task_id))

        return with_account(_act)

    @mcp.tool(description="Mark a To Do task complete. Reversible via update_task status=notStarted.")
    def complete_task(task_id: str) -> dict:
        def _act(a: Account) -> dict:
            t = a.tasks.get(id=task_id)
            t.status = "Completed"  # server auto-sets percent_complete=100
            t.save()
            return _task_to_dict(a.tasks.get(id=task_id))

        return with_account(_act)

    @mcp.tool(
        description="DESTRUCTIVE: Move a To Do task to Deleted Items (recoverable from trash).",
    )
    def delete_task(task_id: str) -> dict:
        def _act(a: Account) -> dict:
            t = a.tasks.get(id=task_id)
            t.move_to_trash()
            return {"deleted": True, "task_id": task_id}

        return with_account(_act)

    @mcp.tool(description="Create a new Microsoft To Do list (a task folder under Tasks).")
    def create_todo_list(name: str) -> dict:
        def _act(a: Account) -> dict:
            f = Folder(parent=a.tasks, name=name)
            f.folder_class = "IPF.Task"
            f.save()
            return {"created": True, "name": name}

        return with_account(_act)
