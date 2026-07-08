"""
Outlook MCP Server
Provides access to Microsoft Outlook via the Windows COM interface.
"""

import asyncio
import pythoncom
import win32com.client
from datetime import datetime, timedelta
from typing import Any

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp import types

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _fmt_dt(value: Any) -> str:
    """
    Format a pywintypes.datetime (or any datetime-like) as a local-time string.
    win32com returns Outlook times already in local time but str() appends a
    misleading '+00:00' offset – using .Format() avoids that.
    """
    try:
        return value.Format("%Y-%m-%d %H:%M:%S")
    except AttributeError:
        # Fallback for plain Python datetime objects
        return str(value)


def _get_outlook() -> Any:
    """Return the running Outlook Application COM object (or start a new one)."""
    try:
        return win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        return win32com.client.Dispatch("Outlook.Application")


def _get_namespace() -> Any:
    outlook = _get_outlook()
    return outlook.GetNamespace("MAPI")


def _safe_get_store_property(store: Any, prop_tag: str) -> Any:
    """Read a MAPI property from a store and return None if unavailable."""
    try:
        return store.PropertyAccessor.GetProperty(prop_tag)
    except Exception:
        return None


def _safe_set_store_property(store: Any, prop_tag: str, value: Any) -> bool:
    """Set a MAPI property on a store. Returns True when successful."""
    try:
        store.PropertyAccessor.SetProperty(prop_tag, value)
        return True
    except Exception:
        return False


def _safe_set_store_text_property(
    store: Any,
    text_prop_tag: str,
    binary_prop_tag: str,
    value: str,
) -> bool:
    """Set store text robustly across profiles exposing string or binary proptags."""
    text_value = str(value)

    if _safe_set_store_property(store, text_prop_tag, text_value):
        return True

    # Some profiles only accept the binary text property (PT_BINARY) for OOF text.
    try:
        encoded = text_value.encode("utf-16-le")
    except Exception:
        encoded = text_value.encode("utf-8", errors="ignore")

    return _safe_set_store_property(store, binary_prop_tag, encoded)


def _decode_store_text(value: Any) -> str:
    """Decode text values that may come back as bytes or strings from COM."""
    if value is None:
        return ""
    if isinstance(value, bytes):
        for encoding in ("utf-16-le", "utf-8", "latin-1"):
            try:
                return value.decode(encoding).rstrip("\x00")
            except Exception:
                continue
        return ""
    return str(value)


def _folder_by_path(namespace: Any, path: str) -> Any:
    """
    Resolve a slash-separated folder path such as 'Inbox' or 'Inbox/Subfolder'.
    The first segment is matched against top-level folders (Inbox, Sent Items …).
    """
    parts = [p.strip() for p in path.split("/") if p.strip()]
    if not parts:
        raise ValueError("Empty folder path")

    # Try well-known folder constants first for the root segment
    WELL_KNOWN = {
        "inbox": 6,
        "sent items": 5,
        "sent": 5,
        "outbox": 4,
        "deleted items": 3,
        "drafts": 16,
        "calendar": 9,
        "contacts": 10,
        "tasks": 13,
        "notes": 12,
        "journal": 11,
    }
    key = parts[0].lower()
    if key in WELL_KNOWN:
        folder = namespace.GetDefaultFolder(WELL_KNOWN[key])
    else:
        # Search in the default store's root
        folder = namespace.Folders.Item(1)
        folder = folder.Folders[parts[0]]

    for part in parts[1:]:
        folder = folder.Folders[part]
    return folder


def _mail_item_to_dict(item: Any, include_body: bool = False) -> dict:
    """Convert an Outlook MailItem to a plain dict."""
    try:
        entry_id = item.EntryID
    except Exception:
        entry_id = ""
    try:
        subject = item.Subject or ""
    except Exception:
        subject = ""
    try:
        sender = item.SenderEmailAddress or item.SenderName or ""
    except Exception:
        sender = ""
    try:
        sender_name = item.SenderName or ""
    except Exception:
        sender_name = ""
    try:
        received = _fmt_dt(item.ReceivedTime)
    except Exception:
        received = ""
    try:
        sent = _fmt_dt(item.SentOn)
    except Exception:
        sent = ""
    try:
        to = item.To or ""
    except Exception:
        to = ""
    try:
        cc = item.CC or ""
    except Exception:
        cc = ""
    try:
        unread = bool(item.UnRead)
    except Exception:
        unread = False
    try:
        size = item.Size
    except Exception:
        size = 0
    try:
        has_attachments = item.Attachments.Count > 0
    except Exception:
        has_attachments = False
    try:
        categories = item.Categories or ""
    except Exception:
        categories = ""

    result = {
        "entry_id": entry_id,
        "subject": subject,
        "sender": sender,
        "sender_name": sender_name,
        "to": to,
        "cc": cc,
        "received": received,
        "sent": sent,
        "unread": unread,
        "size": size,
        "has_attachments": has_attachments,
        "categories": categories,
    }

    if include_body:
        try:
            result["body"] = item.Body or ""
        except Exception:
            result["body"] = ""
        try:
            result["html_body"] = item.HTMLBody or ""
        except Exception:
            result["html_body"] = ""
        try:
            attachments = []
            for i in range(1, item.Attachments.Count + 1):
                att = item.Attachments.Item(i)
                attachments.append({"name": att.FileName, "size": att.Size})
            result["attachments"] = attachments
        except Exception:
            result["attachments"] = []

    return result


def _appointment_to_dict(item: Any) -> dict:
    try:
        entry_id = item.EntryID
    except Exception:
        entry_id = ""
    try:
        subject = item.Subject or ""
    except Exception:
        subject = ""
    try:
        location = item.Location or ""
    except Exception:
        location = ""
    try:
        start = _fmt_dt(item.Start)
    except Exception:
        start = ""
    try:
        end = _fmt_dt(item.End)
    except Exception:
        end = ""
    try:
        duration = item.Duration
    except Exception:
        duration = 0
    try:
        organizer = item.Organizer or ""
    except Exception:
        organizer = ""
    try:
        body = item.Body or ""
    except Exception:
        body = ""
    try:
        all_day = bool(item.AllDayEvent)
    except Exception:
        all_day = False
    try:
        is_recurring = bool(item.IsRecurring)
    except Exception:
        is_recurring = False

    return {
        "entry_id": entry_id,
        "subject": subject,
        "location": location,
        "start": start,
        "end": end,
        "duration_minutes": duration,
        "organizer": organizer,
        "body": body,
        "all_day": all_day,
        "is_recurring": is_recurring,
    }


def _contact_to_dict(item: Any) -> dict:
    fields = {
        "entry_id": "EntryID",
        "full_name": "FullName",
        "first_name": "FirstName",
        "last_name": "LastName",
        "email1": "Email1Address",
        "email2": "Email2Address",
        "email3": "Email3Address",
        "company": "CompanyName",
        "job_title": "JobTitle",
        "mobile": "MobileTelephoneNumber",
        "business_phone": "BusinessTelephoneNumber",
        "home_phone": "HomeTelephoneNumber",
        "business_address": "BusinessAddressStreet",
        "categories": "Categories",
    }
    result = {}
    for key, attr in fields.items():
        try:
            result[key] = getattr(item, attr) or ""
        except Exception:
            result[key] = ""
    return result


# ---------------------------------------------------------------------------
# MCP Server
# ---------------------------------------------------------------------------

server = Server("outlook-mcp")


# ── list_folders ────────────────────────────────────────────────────────────
@server.list_tools()
async def list_tools() -> list[types.Tool]:
    return [
        types.Tool(
            name="list_folders",
            description="List all mail folders (top-level and optionally recursive).",
            inputSchema={
                "type": "object",
                "properties": {
                    "recursive": {
                        "type": "boolean",
                        "description": "Include sub-folders recursively. Default: false.",
                    }
                },
            },
        ),
        types.Tool(
            name="get_emails",
            description=(
                "Retrieve emails from a mail folder. "
                "Returns subject, sender, date, unread flag and a short preview."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "folder": {
                        "type": "string",
                        "description": "Folder path, e.g. 'Inbox' or 'Inbox/Projects'. Default: 'Inbox'.",
                    },
                    "count": {
                        "type": "integer",
                        "description": "Maximum number of emails to return (newest first). Default: 20.",
                    },
                    "unread_only": {
                        "type": "boolean",
                        "description": "Return only unread emails. Default: false.",
                    },
                },
            },
        ),
        types.Tool(
            name="get_email_details",
            description="Get the full details (including body and attachments) of a single email by its EntryID.",
            inputSchema={
                "type": "object",
                "properties": {
                    "entry_id": {
                        "type": "string",
                        "description": "The EntryID of the email (returned by get_emails).",
                    }
                },
                "required": ["entry_id"],
            },
        ),
        types.Tool(
            name="search_emails",
            description="Search emails across a folder using a keyword (subject, body, or sender).",
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "Search keyword or phrase.",
                    },
                    "folder": {
                        "type": "string",
                        "description": "Folder to search in. Default: 'Inbox'.",
                    },
                    "field": {
                        "type": "string",
                        "enum": ["subject", "body", "sender", "all"],
                        "description": "Which field to search. Default: 'all'.",
                    },
                    "count": {
                        "type": "integer",
                        "description": "Maximum results to return. Default: 20.",
                    },
                },
                "required": ["query"],
            },
        ),
        types.Tool(
            name="send_email",
            description="Compose and send a new email.",
            inputSchema={
                "type": "object",
                "properties": {
                    "to": {
                        "type": "string",
                        "description": "Recipient(s), semicolon-separated.",
                    },
                    "subject": {"type": "string", "description": "Email subject."},
                    "body": {"type": "string", "description": "Plain-text body."},
                    "cc": {
                        "type": "string",
                        "description": "CC recipient(s), semicolon-separated.",
                    },
                    "bcc": {
                        "type": "string",
                        "description": "BCC recipient(s), semicolon-separated.",
                    },
                },
                "required": ["to", "subject", "body"],
            },
        ),
        types.Tool(
            name="reply_to_email",
            description="Reply to an existing email.",
            inputSchema={
                "type": "object",
                "properties": {
                    "entry_id": {
                        "type": "string",
                        "description": "EntryID of the email to reply to.",
                    },
                    "body": {"type": "string", "description": "Reply text (prepended to the original)."},
                    "reply_all": {
                        "type": "boolean",
                        "description": "Reply to all recipients. Default: false.",
                    },
                },
                "required": ["entry_id", "body"],
            },
        ),
        types.Tool(
            name="forward_email",
            description="Forward an existing email to new recipients.",
            inputSchema={
                "type": "object",
                "properties": {
                    "entry_id": {
                        "type": "string",
                        "description": "EntryID of the email to forward.",
                    },
                    "to": {
                        "type": "string",
                        "description": "Recipient(s), semicolon-separated.",
                    },
                    "body": {
                        "type": "string",
                        "description": "Optional additional text prepended to the forwarded email.",
                    },
                },
                "required": ["entry_id", "to"],
            },
        ),
        types.Tool(
            name="move_email",
            description="Move an email to a different folder.",
            inputSchema={
                "type": "object",
                "properties": {
                    "entry_id": {
                        "type": "string",
                        "description": "EntryID of the email to move.",
                    },
                    "destination_folder": {
                        "type": "string",
                        "description": "Destination folder path, e.g. 'Inbox/Newsletter'.",
                    },
                },
                "required": ["entry_id", "destination_folder"],
            },
        ),
        types.Tool(
            name="delete_email",
            description="Move an email to the Deleted Items folder.",
            inputSchema={
                "type": "object",
                "properties": {
                    "entry_id": {
                        "type": "string",
                        "description": "EntryID of the email to delete.",
                    }
                },
                "required": ["entry_id"],
            },
        ),
        types.Tool(
            name="mark_email",
            description="Mark an email as read or unread.",
            inputSchema={
                "type": "object",
                "properties": {
                    "entry_id": {
                        "type": "string",
                        "description": "EntryID of the email.",
                    },
                    "unread": {
                        "type": "boolean",
                        "description": "True = mark as unread, False = mark as read.",
                    },
                },
                "required": ["entry_id", "unread"],
            },
        ),
        types.Tool(
            name="get_calendar_events",
            description="Retrieve calendar events within a date range.",
            inputSchema={
                "type": "object",
                "properties": {
                    "start_date": {
                        "type": "string",
                        "description": "Start date in ISO format YYYY-MM-DD. Default: today.",
                    },
                    "end_date": {
                        "type": "string",
                        "description": "End date in ISO format YYYY-MM-DD. Default: 7 days from today.",
                    },
                    "count": {
                        "type": "integer",
                        "description": "Maximum number of events to return. Default: 50.",
                    },
                },
            },
        ),
        types.Tool(
            name="create_calendar_event",
            description="Create a new calendar appointment.",
            inputSchema={
                "type": "object",
                "properties": {
                    "subject": {"type": "string", "description": "Event title."},
                    "start": {
                        "type": "string",
                        "description": "Start datetime in ISO format, e.g. '2026-05-26T10:00:00'.",
                    },
                    "end": {
                        "type": "string",
                        "description": "End datetime in ISO format.",
                    },
                    "location": {"type": "string", "description": "Optional location."},
                    "body": {"type": "string", "description": "Optional description/notes."},
                    "all_day": {
                        "type": "boolean",
                        "description": "Create as an all-day event. Default: false.",
                    },
                    "reminder_minutes": {
                        "type": "integer",
                        "description": "Reminder in minutes before the event. Default: 15.",
                    },
                },
                "required": ["subject", "start", "end"],
            },
        ),
        types.Tool(
            name="get_contacts",
            description="Retrieve contacts from the Outlook address book.",
            inputSchema={
                "type": "object",
                "properties": {
                    "filter": {
                        "type": "string",
                        "description": "Optional name/email substring to filter contacts.",
                    },
                    "count": {
                        "type": "integer",
                        "description": "Maximum number of contacts to return. Default: 50.",
                    },
                },
            },
        ),
        types.Tool(
            name="create_contact",
            description="Create a new contact in the default Contacts folder.",
            inputSchema={
                "type": "object",
                "properties": {
                    "first_name": {"type": "string", "description": "First name."},
                    "last_name": {"type": "string", "description": "Last name."},
                    "email": {"type": "string", "description": "Primary email address."},
                    "company": {"type": "string", "description": "Company name."},
                    "job_title": {"type": "string", "description": "Job title."},
                    "mobile": {"type": "string", "description": "Mobile phone number."},
                    "business_phone": {"type": "string", "description": "Business phone number."},
                },
                "required": ["first_name"],
            },
        ),
        types.Tool(
            name="get_tasks",
            description="Retrieve tasks from the default Tasks folder.",
            inputSchema={
                "type": "object",
                "properties": {
                    "completed": {
                        "type": "boolean",
                        "description": "Include completed tasks. Default: false.",
                    },
                    "count": {
                        "type": "integer",
                        "description": "Maximum number of tasks to return. Default: 50.",
                    },
                },
            },
        ),
        types.Tool(
            name="get_automatic_replies",
            description=(
                "Get automatic replies (Out of Office) status and configured "
                "internal/external reply text for a mailbox store."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "store_index": {
                        "type": "integer",
                        "description": "Mailbox store index (1-based). Default: 1.",
                    }
                },
            },
        ),
        types.Tool(
            name="set_automatic_replies",
            description=(
                "Set automatic replies (Out of Office) status and optionally "
                "internal/external reply text for a mailbox store."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "enabled": {
                        "type": "boolean",
                        "description": "Enable or disable automatic replies.",
                    },
                    "internal_text": {
                        "type": "string",
                        "description": "Optional internal auto-reply text.",
                    },
                    "external_text": {
                        "type": "string",
                        "description": "Optional external auto-reply text.",
                    },
                    "store_index": {
                        "type": "integer",
                        "description": "Mailbox store index (1-based). Default: 1.",
                    },
                },
                "required": ["enabled"],
            },
        ),
    ]


# ── call_tool dispatcher ─────────────────────────────────────────────────────
@server.call_tool()
async def call_tool(name: str, arguments: dict) -> list[types.TextContent]:
    try:
        result = await asyncio.get_event_loop().run_in_executor(
            None, _dispatch, name, arguments
        )
        return [types.TextContent(type="text", text=result)]
    except Exception as exc:
        return [types.TextContent(type="text", text=f"ERROR: {exc}")]


def _dispatch(name: str, args: dict) -> str:
    """Synchronous dispatcher running in a thread-pool executor."""
    pythoncom.CoInitialize()
    try:
        if name == "list_folders":
            return _tool_list_folders(args)
        elif name == "get_emails":
            return _tool_get_emails(args)
        elif name == "get_email_details":
            return _tool_get_email_details(args)
        elif name == "search_emails":
            return _tool_search_emails(args)
        elif name == "send_email":
            return _tool_send_email(args)
        elif name == "reply_to_email":
            return _tool_reply_to_email(args)
        elif name == "forward_email":
            return _tool_forward_email(args)
        elif name == "move_email":
            return _tool_move_email(args)
        elif name == "delete_email":
            return _tool_delete_email(args)
        elif name == "mark_email":
            return _tool_mark_email(args)
        elif name == "get_calendar_events":
            return _tool_get_calendar_events(args)
        elif name == "create_calendar_event":
            return _tool_create_calendar_event(args)
        elif name == "get_contacts":
            return _tool_get_contacts(args)
        elif name == "create_contact":
            return _tool_create_contact(args)
        elif name == "get_tasks":
            return _tool_get_tasks(args)
        elif name == "get_automatic_replies":
            return _tool_get_automatic_replies(args)
        elif name == "set_automatic_replies":
            return _tool_set_automatic_replies(args)
        else:
            return f"Unknown tool: {name}"
    finally:
        pythoncom.CoUninitialize()


# ---------------------------------------------------------------------------
# Tool implementations
# ---------------------------------------------------------------------------

def _collect_folders(folder: Any, recursive: bool, prefix: str = "") -> list[str]:
    name = f"{prefix}/{folder.Name}".lstrip("/")
    results = [name]
    if recursive:
        try:
            for i in range(1, folder.Folders.Count + 1):
                results.extend(_collect_folders(folder.Folders.Item(i), True, name))
        except Exception:
            pass
    return results


def _tool_list_folders(args: dict) -> str:
    recursive = args.get("recursive", False)
    ns = _get_namespace()
    folders: list[str] = []
    for i in range(1, ns.Folders.Count + 1):
        store = ns.Folders.Item(i)
        try:
            for j in range(1, store.Folders.Count + 1):
                folders.extend(_collect_folders(store.Folders.Item(j), recursive))
        except Exception:
            pass
    return "\n".join(folders) if folders else "No folders found."


def _tool_get_emails(args: dict) -> str:
    folder_path = args.get("folder", "Inbox")
    count = int(args.get("count", 20))
    unread_only = bool(args.get("unread_only", False))

    ns = _get_namespace()
    folder = _folder_by_path(ns, folder_path)
    items = folder.Items
    items.Sort("[ReceivedTime]", True)  # newest first

    results = []
    for i in range(1, items.Count + 1):
        if len(results) >= count:
            break
        try:
            item = items.Item(i)
            if item.Class != 43:  # 43 = olMail
                continue
            if unread_only and not item.UnRead:
                continue
            d = _mail_item_to_dict(item)
            results.append(
                f"[{i}] EntryID: {d['entry_id']}\n"
                f"    Subject : {d['subject']}\n"
                f"    From    : {d['sender_name']} <{d['sender']}>\n"
                f"    To      : {d['to']}\n"
                f"    Received: {d['received']}\n"
                f"    Unread  : {d['unread']}  Attachments: {d['has_attachments']}"
            )
        except Exception as e:
            results.append(f"[{i}] Error reading item: {e}")

    return "\n\n".join(results) if results else "No emails found."


def _tool_get_email_details(args: dict) -> str:
    entry_id = args["entry_id"]
    ns = _get_namespace()
    item = ns.GetItemFromID(entry_id)
    d = _mail_item_to_dict(item, include_body=True)
    lines = [
        f"Subject   : {d['subject']}",
        f"From      : {d['sender_name']} <{d['sender']}>",
        f"To        : {d['to']}",
        f"CC        : {d['cc']}",
        f"Received  : {d['received']}",
        f"Sent      : {d['sent']}",
        f"Unread    : {d['unread']}",
        f"Categories: {d['categories']}",
        f"Attachments: {d.get('attachments', [])}",
        "",
        "--- Body ---",
        d.get("body", ""),
    ]
    return "\n".join(lines)


def _tool_search_emails(args: dict) -> str:
    query = args["query"].lower()
    folder_path = args.get("folder", "Inbox")
    field = args.get("field", "all")
    count = int(args.get("count", 20))

    ns = _get_namespace()
    folder = _folder_by_path(ns, folder_path)
    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    results = []
    for i in range(1, items.Count + 1):
        if len(results) >= count:
            break
        try:
            item = items.Item(i)
            if item.Class != 43:
                continue
            subject = (item.Subject or "").lower()
            sender = (item.SenderEmailAddress or item.SenderName or "").lower()
            body = ""
            if field in ("body", "all"):
                try:
                    body = (item.Body or "").lower()
                except Exception:
                    body = ""

            match = False
            if field == "subject" and query in subject:
                match = True
            elif field == "sender" and query in sender:
                match = True
            elif field == "body" and query in body:
                match = True
            elif field == "all" and (query in subject or query in sender or query in body):
                match = True

            if match:
                d = _mail_item_to_dict(item)
                results.append(
                    f"EntryID : {d['entry_id']}\n"
                    f"Subject : {d['subject']}\n"
                    f"From    : {d['sender_name']} <{d['sender']}>\n"
                    f"Received: {d['received']}\n"
                    f"Unread  : {d['unread']}"
                )
        except Exception as e:
            continue

    return "\n\n".join(results) if results else "No matching emails found."


def _tool_send_email(args: dict) -> str:
    outlook = _get_outlook()
    mail = outlook.CreateItem(0)  # 0 = olMailItem
    mail.To = args["to"]
    mail.Subject = args["subject"]
    mail.Body = args["body"]
    if args.get("cc"):
        mail.CC = args["cc"]
    if args.get("bcc"):
        mail.BCC = args["bcc"]
    mail.Send()
    return f"Email sent to {args['to']} with subject '{args['subject']}'."


def _tool_reply_to_email(args: dict) -> str:
    ns = _get_namespace()
    item = ns.GetItemFromID(args["entry_id"])
    reply_all = bool(args.get("reply_all", False))
    reply = item.ReplyAll() if reply_all else item.Reply()
    reply.Body = args["body"] + "\n\n" + reply.Body
    reply.Send()
    return "Reply sent."


def _tool_forward_email(args: dict) -> str:
    ns = _get_namespace()
    item = ns.GetItemFromID(args["entry_id"])
    fwd = item.Forward()
    fwd.To = args["to"]
    if args.get("body"):
        fwd.Body = args["body"] + "\n\n" + fwd.Body
    fwd.Send()
    return f"Email forwarded to {args['to']}."


def _tool_move_email(args: dict) -> str:
    ns = _get_namespace()
    item = ns.GetItemFromID(args["entry_id"])
    dest = _folder_by_path(ns, args["destination_folder"])
    item.Move(dest)
    return f"Email moved to '{args['destination_folder']}'."


def _tool_delete_email(args: dict) -> str:
    ns = _get_namespace()
    item = ns.GetItemFromID(args["entry_id"])
    item.Delete()
    return "Email moved to Deleted Items."


def _tool_mark_email(args: dict) -> str:
    ns = _get_namespace()
    item = ns.GetItemFromID(args["entry_id"])
    item.UnRead = bool(args["unread"])
    item.Save()
    state = "unread" if args["unread"] else "read"
    return f"Email marked as {state}."


def _tool_get_calendar_events(args: dict) -> str:
    today = datetime.now().date()
    start_str = args.get("start_date", str(today))
    end_str = args.get("end_date", str(today + timedelta(days=7)))
    count = int(args.get("count", 50))

    start_dt = datetime.fromisoformat(start_str)
    end_dt = datetime.fromisoformat(end_str) + timedelta(hours=23, minutes=59, seconds=59)

    ns = _get_namespace()
    cal_folder = ns.GetDefaultFolder(9)  # olFolderCalendar
    items = cal_folder.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")

    # Outlook DASL filter for date range
    fmt = "%m/%d/%Y %H:%M %p"
    filter_str = (
        f"[Start] >= '{start_dt.strftime(fmt)}' AND [Start] <= '{end_dt.strftime(fmt)}'"
    )
    restricted = items.Restrict(filter_str)

    results = []
    for i in range(1, restricted.Count + 1):
        if len(results) >= count:
            break
        try:
            appt = restricted.Item(i)
            d = _appointment_to_dict(appt)
            results.append(
                f"EntryID : {d['entry_id']}\n"
                f"Subject : {d['subject']}\n"
                f"Start   : {d['start']}\n"
                f"End     : {d['end']}\n"
                f"Location: {d['location']}\n"
                f"All-day : {d['all_day']}  Recurring: {d['is_recurring']}\n"
                f"Organizer: {d['organizer']}"
            )
        except Exception as e:
            results.append(f"Error reading event: {e}")

    return "\n\n".join(results) if results else "No calendar events found in the given range."


def _tool_create_calendar_event(args: dict) -> str:
    outlook = _get_outlook()
    appt = outlook.CreateItem(1)  # 1 = olAppointmentItem
    appt.Subject = args["subject"]
    start_value = datetime.fromisoformat(args["start"])
    end_value = datetime.fromisoformat(args["end"])
    appt.Start = start_value
    appt.Duration = int((end_value - start_value).total_seconds() // 60)

    if args.get("location"):
        try:
            appt.Location = args["location"]
        except Exception:
            pass
    if args.get("body"):
        try:
            appt.Body = args["body"]
        except Exception:
            pass
    if args.get("all_day"):
        try:
            appt.AllDayEvent = True
        except Exception:
            pass
    try:
        appt.ReminderSet = True
        appt.ReminderMinutesBeforeStart = int(args.get("reminder_minutes", 15))
    except Exception:
        pass
    appt.Save()
    return f"Calendar event '{args['subject']}' created for {start_value.strftime('%Y-%m-%d %H:%M')}."


def _tool_get_contacts(args: dict) -> str:
    filter_str = (args.get("filter") or "").lower()
    count = int(args.get("count", 50))

    ns = _get_namespace()
    contacts_folder = ns.GetDefaultFolder(10)  # olFolderContacts
    items = contacts_folder.Items

    results = []
    for i in range(1, items.Count + 1):
        if len(results) >= count:
            break
        try:
            item = items.Item(i)
            if item.Class != 40:  # 40 = olContact
                continue
            d = _contact_to_dict(item)
            if filter_str and not (
                filter_str in (d["full_name"] or "").lower()
                or filter_str in (d["email1"] or "").lower()
                or filter_str in (d["company"] or "").lower()
            ):
                continue
            results.append(
                f"Name   : {d['full_name']}\n"
                f"Email  : {d['email1']}\n"
                f"Company: {d['company']}  Title: {d['job_title']}\n"
                f"Mobile : {d['mobile']}  Phone: {d['business_phone']}\n"
                f"EntryID: {d['entry_id']}"
            )
        except Exception:
            continue

    return "\n\n".join(results) if results else "No contacts found."


def _tool_create_contact(args: dict) -> str:
    outlook = _get_outlook()
    contact = outlook.CreateItem(2)  # 2 = olContactItem
    contact.FirstName = args.get("first_name", "")
    contact.LastName = args.get("last_name", "")
    if args.get("email"):
        contact.Email1Address = args["email"]
    if args.get("company"):
        contact.CompanyName = args["company"]
    if args.get("job_title"):
        contact.JobTitle = args["job_title"]
    if args.get("mobile"):
        contact.MobileTelephoneNumber = args["mobile"]
    if args.get("business_phone"):
        contact.BusinessTelephoneNumber = args["business_phone"]
    contact.Save()
    full_name = f"{args.get('first_name', '')} {args.get('last_name', '')}".strip()
    return f"Contact '{full_name}' created."


def _tool_get_tasks(args: dict) -> str:
    include_completed = bool(args.get("completed", False))
    count = int(args.get("count", 50))

    ns = _get_namespace()
    tasks_folder = ns.GetDefaultFolder(13)  # olFolderTasks
    items = tasks_folder.Items
    items.Sort("[DueDate]")

    results = []
    for i in range(1, items.Count + 1):
        if len(results) >= count:
            break
        try:
            item = items.Item(i)
            if item.Class != 48:  # 48 = olTask
                continue
            try:
                complete = bool(item.Complete)
            except Exception:
                complete = False
            if not include_completed and complete:
                continue
            try:
                subject = item.Subject or ""
            except Exception:
                subject = ""
            try:
                due = str(item.DueDate)
            except Exception:
                due = ""
            try:
                priority = item.Importance
            except Exception:
                priority = ""
            try:
                entry_id = item.EntryID
            except Exception:
                entry_id = ""
            results.append(
                f"Subject : {subject}\n"
                f"Due     : {due}\n"
                f"Complete: {complete}  Priority: {priority}\n"
                f"EntryID : {entry_id}"
            )
        except Exception:
            continue

    return "\n\n".join(results) if results else "No tasks found."


def _tool_get_automatic_replies(args: dict) -> str:
    store_index = int(args.get("store_index", 1))
    if store_index < 1:
        raise ValueError("store_index must be >= 1")

    ns = _get_namespace()
    if store_index > ns.Stores.Count:
        raise ValueError(f"store_index out of range. Found {ns.Stores.Count} store(s).")

    store = ns.Stores.Item(store_index)

    # MAPI proptags for Out of Office in mailbox store.
    enabled_raw = _safe_get_store_property(store, "http://schemas.microsoft.com/mapi/proptag/0x661D000B")
    internal_raw = _safe_get_store_property(store, "http://schemas.microsoft.com/mapi/proptag/0x661E001F")
    external_raw = _safe_get_store_property(store, "http://schemas.microsoft.com/mapi/proptag/0x661F001F")

    # Fallback: some profiles expose text as binary data.
    if internal_raw is None:
        internal_raw = _safe_get_store_property(store, "http://schemas.microsoft.com/mapi/proptag/0x661E0102")
    if external_raw is None:
        external_raw = _safe_get_store_property(store, "http://schemas.microsoft.com/mapi/proptag/0x661F0102")

    enabled = bool(enabled_raw) if enabled_raw is not None else False
    internal_text = _decode_store_text(internal_raw)
    external_text = _decode_store_text(external_raw)

    if not internal_text:
        internal_text = "(not set)"
    if not external_text:
        external_text = "(not set)"

    lines = [
        f"Store               : {store.DisplayName}",
        f"Store index         : {store_index}",
        f"Automatic replies   : {enabled}",
        "",
        "Internal reply text:",
        internal_text,
        "",
        "External reply text:",
        external_text,
    ]
    return "\n".join(lines)


def _tool_set_automatic_replies(args: dict) -> str:
    store_index = int(args.get("store_index", 1))
    if store_index < 1:
        raise ValueError("store_index must be >= 1")

    if "enabled" not in args:
        raise ValueError("enabled is required")

    ns = _get_namespace()
    if store_index > ns.Stores.Count:
        raise ValueError(f"store_index out of range. Found {ns.Stores.Count} store(s).")

    store = ns.Stores.Item(store_index)
    enabled = bool(args.get("enabled"))
    internal_text = args.get("internal_text")
    external_text = args.get("external_text")

    ok_enabled = _safe_set_store_property(
        store,
        "http://schemas.microsoft.com/mapi/proptag/0x661D000B",
        enabled,
    )

    ok_internal = True
    if internal_text is not None:
        ok_internal = _safe_set_store_text_property(
            store,
            "http://schemas.microsoft.com/mapi/proptag/0x661E001F",
            "http://schemas.microsoft.com/mapi/proptag/0x661E0102",
            str(internal_text),
        )

    ok_external = True
    if external_text is not None:
        ok_external = _safe_set_store_text_property(
            store,
            "http://schemas.microsoft.com/mapi/proptag/0x661F001F",
            "http://schemas.microsoft.com/mapi/proptag/0x661F0102",
            str(external_text),
        )

    if not ok_enabled:
        raise RuntimeError("Could not set automatic replies state for this store.")
    if internal_text is not None and not ok_internal:
        raise RuntimeError("Automatic replies state changed, but internal_text could not be set.")
    if external_text is not None and not ok_external:
        raise RuntimeError("Automatic replies state changed, but external_text could not be set.")

    # Return the current configured values as confirmation.
    return _tool_get_automatic_replies({"store_index": store_index})


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

async def main() -> None:
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options(),
        )


if __name__ == "__main__":
    asyncio.run(main())
