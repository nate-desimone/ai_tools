## @file
# MCP Server for Outlook email access via MAPI COM.
#
# Copyright 2026 Intel Corporation All Rights Reserved.
# SPDX-License-Identifier: BSD-2-Clause-Patent
##
"""
MCP Server for Outlook email access via MAPI COM.

Provides tools for AI agents to read, search, and draft emails in Microsoft
Outlook on Windows systems.

Requires Microsoft Outlook to be installed. The server will start Outlook
automatically if it is not already running.

Response format:
    Every tool returns one Markdown document.

    - The first line is an H1 title naming the result (e.g. "# Emails").
    - Scalar fields are bullets of the form "- **field_name:** value".
      Identifiers, addresses, and paths are in code spans. Booleans are
      true/false. Lists of scalars are nested bullets. A field with no value
      (such as the sender of an unsent draft) is omitted.
    - Lists of records are GitHub-flavored tables whose header cells are the
      field names. A literal "|" in a cell is written "\\|". An empty list is
      the line "_None._".
    - Free-form text (email bodies, attachment contents) is in a fenced code
      block whose fence is longer than any backtick run in the text, so the
      text is everything between the opening and closing fence.
    - A document holding several records (get_thread) has one "##" section
      per record.
    - Failures are raised as errors, which MCP reports as tool errors.

Usage:
    python mcp_email_server.py
"""

import os
import re
import tempfile
import winreg
from datetime import datetime
from pathlib import Path
from typing import Optional

from bs4 import BeautifulSoup
from markdownify import MarkdownConverter
import markdown as _markdown
import pythoncom
import win32com.client
from PIL import Image
import pytesseract

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("Outlook Email")


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _get_outlook():
    """Return (outlook_app, mapi_namespace).

    Connects to an already-running Outlook instance when possible; falls back
    to launching Outlook via COM Dispatch if it is not running.
    """
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    return outlook, namespace


def _resolve_folder(namespace, account_name: str, folder_path: str):
    """Return the Outlook Folder object at *folder_path* within *account_name*.

    *folder_path* is a slash-separated path relative to the account root,
    e.g. "Inbox" or "Inbox/Projects/Active".

    Raises ValueError if the account or any path component is not found.
    """
    root_folder = None
    for store in namespace.Folders:
        if store.Name.lower() == account_name.lower():
            root_folder = store
            break

    if root_folder is None:
        available = [s.Name for s in namespace.Folders]
        raise ValueError(
            f"Account '{account_name}' not found. "
            f"Available accounts: {available}"
        )

    parts = [p for p in folder_path.replace("\\", "/").split("/") if p]
    folder = root_folder
    for part in parts:
        found = None
        for sub in folder.Folders:
            if sub.Name.lower() == part.lower():
                found = sub
                break
        if found is None:
            raise ValueError(
                f"Folder '{part}' not found inside '{folder.Name}'"
            )
        folder = found

    return folder


# MAPI properties holding SMTP addresses. Exchange mail otherwise exposes only
# an X.500 address (/O=EXCHANGELABS/...), and neither property is set on every
# message, so senders are tried against both.
_PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
_PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"


def _mapi_string(obj, tag: str) -> str:
    """Read a MAPI string property from *obj*, or "" when it is unavailable."""
    try:
        return obj.PropertyAccessor.GetProperty(tag) or ""
    except Exception:
        return ""


def _sender_address(item) -> str:
    """Return the sender's SMTP address, or the raw address if it cannot be resolved."""
    try:
        address = item.SenderEmailAddress or ""
        is_exchange = item.SenderEmailType == "EX"
    except Exception:
        return ""
    if not is_exchange:
        return address

    smtp = _mapi_string(item, _PR_SENDER_SMTP_ADDRESS)
    if not smtp:
        try:
            smtp = _mapi_string(item.Sender, _PR_SMTP_ADDRESS)
        except Exception:
            pass
    return smtp or address


def _recipient_address(recipient) -> str:
    """Return a recipient's SMTP address, else its raw address, else its name."""
    smtp = _mapi_string(recipient, _PR_SMTP_ADDRESS)
    if smtp:
        return smtp
    try:
        return recipient.Address or recipient.Name
    except Exception:
        return recipient.Name


def _mail_item_summary(item) -> dict:
    """Return a concise summary dict for a MailItem."""
    try:
        received = item.ReceivedTime
        received_str = received.isoformat() if hasattr(received, "isoformat") else str(received)
    except Exception:
        received_str = None

    try:
        sender_email = _sender_address(item)
    except Exception:
        sender_email = ""

    try:
        conversation_id = item.ConversationID
    except Exception:
        conversation_id = ""

    return {
        "entry_id": item.EntryID,
        "subject": item.Subject,
        "sender_name": item.SenderName,
        "sender_email": sender_email,
        "received_time": received_str,
        "has_attachments": item.Attachments.Count > 0,
        "is_unread": bool(item.UnRead),
        "conversation_id": conversation_id,
    }


# ---------------------------------------------------------------------------
# Markdown rendering
# ---------------------------------------------------------------------------

_NONE = "_None._"


def _inline(value) -> str:
    """Return *value* as a single line of text (line breaks become spaces)."""
    if value is None:
        return ""
    return " ".join(p.strip() for p in str(value).splitlines() if p.strip())


def _scalar(value) -> str:
    """Return a bool as true/false and anything else as a single line."""
    if isinstance(value, bool):
        return "true" if value else "false"
    return _inline(value)


def _lf(text: str) -> str:
    """Return *text* with every line ending as a bare newline."""
    return text.replace("\r\n", "\n").replace("\r", "\n")


def _longest_backtick_run(text: str) -> int:
    return max((len(m) for m in re.findall(r"`+", text)), default=0)


def _code(value) -> str:
    """Return *value* as an inline code span, or "" when it is empty."""
    text = _inline(value)
    if not text:
        return ""
    ticks = "`" * (_longest_backtick_run(text) + 1)
    pad = " " if text.startswith("`") or text.endswith("`") else ""
    return f"{ticks}{pad}{text}{pad}{ticks}"


def _fence(text: str, info: str = "") -> str:
    """Return *text* in a fenced code block that no line of *text* can close."""
    text = _lf(text)
    ticks = "`" * max(3, _longest_backtick_run(text) + 1)
    return f"{ticks}{info}\n{text}\n{ticks}"


def _kv(**fields) -> str:
    """Render *fields* as "- **name:** value" bullets, skipping empty ones.

    List values become nested bullets.
    """
    lines = []
    for name, value in fields.items():
        if isinstance(value, (list, tuple)):
            items = [text for text in map(_inline, value) if text]
            if items:
                lines.append(f"- **{name}:**")
                lines.extend(f"  - {text}" for text in items)
        else:
            text = _scalar(value)
            if text:
                lines.append(f"- **{name}:** {text}")
    return "\n".join(lines)


def _table(columns: list, rows: list) -> str:
    """Render a GitHub-flavored table, or "_None._" when there are no rows."""
    if not rows:
        return _NONE

    def _line(cells) -> str:
        return "| " + " | ".join(_scalar(c).replace("|", "\\|") for c in cells) + " |"

    separator = "|" + "|".join(["---"] * len(columns)) + "|"
    return "\n".join([_line(columns), separator, *(_line(r) for r in rows)])


def _doc(*blocks: str) -> str:
    """Join non-empty Markdown blocks with blank lines."""
    return "\n\n".join(b for b in blocks if b)


_EMAIL_COLUMNS = [
    "received_time", "sender_name", "sender_email", "subject",
    "is_unread", "has_attachments", "entry_id", "conversation_id",
]


def _email_table(summaries: list, with_flags: bool = False) -> str:
    """Render _mail_item_summary dicts as a table, add flag column when `with_flags==true`."""
    columns = list(_EMAIL_COLUMNS)
    if with_flags:
        columns[4:4] = ["flag_status", "flag_request"]
    rows = []
    for s in summaries:
        cells = [
            s["received_time"], s["sender_name"], _code(s["sender_email"]), s["subject"],
        ]
        if with_flags:
            cells += [s["flag_status"], s["flag_request"]]
        cells += [
            s["is_unread"], s["has_attachments"],
            _code(s["entry_id"]), _code(s["conversation_id"]),
        ]
        rows.append(cells)
    return _table(columns, rows)


def _email_block(heading: str, entry: dict, level: int) -> str:
    """Render a full email (summary dict plus body fields) as a section."""
    marks = "#" * level
    return _doc(
        f"{marks} {heading}",
        _kv(
            entry_id=_code(entry["entry_id"]),
            conversation_id=_code(entry["conversation_id"]),
            subject=entry["subject"],
            sender_name=entry["sender_name"],
            sender_email=_code(entry["sender_email"]),
            recipients=[_code(r) for r in entry["recipients"]],
            received_time=entry["received_time"],
            is_unread=entry["is_unread"],
            has_attachments=entry["has_attachments"],
            attachment_count=entry["attachment_count"],
            body_format=entry["body_format"],
        ),
        f"{marks}# Body",
        _fence(entry["body"], entry["body_format"]),
    )


def _draft_result(title: str, entry_id: str, message: str) -> str:
    return _doc(f"# {title}", _kv(entry_id=_code(entry_id), message=message))


_MD_EXTENSIONS = ["fenced_code", "tables"]
_BODY_FONT = "font-family:Aptos,Calibri,sans-serif;font-size:11pt;"
_MONO_FONT = "font-family:Consolas,'Courier New',monospace;font-size:11pt;"
_PRE_STYLE = _MONO_FONT + "background-color:#f5f5f5;padding:8px;"


def _md_to_fragment(md_text: str) -> str:
    """Convert Markdown to a styled HTML fragment (no html/head/body wrappers)."""
    html = _markdown.markdown(md_text, extensions=_MD_EXTENSIONS)
    html = re.sub(r"<code(?=[>\s])", f'<code style="{_MONO_FONT}"', html)
    html = re.sub(r"<pre(?=[>\s])", f'<pre style="{_PRE_STYLE}"', html)
    return f'<div style="{_BODY_FONT}">{html}</div>'


def _markdown_to_html(md_text: str) -> str:
    """Convert Markdown to a complete HTML document."""
    return (
        '<html><head><meta charset="utf-8"></head><body>'
        + _md_to_fragment(md_text)
        + "</body></html>"
    )


def _prepend_html(fragment: str, existing_html: str) -> str:
    """Insert *fragment* right after the opening <body> tag of *existing_html*."""
    match = re.search(r"<body[^>]*>", existing_html, re.IGNORECASE)
    if match:
        pos = match.end()
        return existing_html[:pos] + fragment + existing_html[pos:]
    return fragment + existing_html


# ---------------------------------------------------------------------------
# Email body (HTML to Markdown)
# ---------------------------------------------------------------------------

# It is common for Email HTML to use tables for page layout, which can make it
# difficult to determine if a HTML table is truly tabular data. Markdownify turns
# every table into a Markdown table, leaking nested tables into cells as literal
# pipes. To improve the readability, apply heuristics to attempt to categorize
# tables as layout versus data. Tables judged to be layout are unwrapped into
# plain blocks.

_LAYOUT_ROLES = ("presentation", "none")
_DATA_ROLES = ("table", "grid")
_BLOCK_TAGS = {
    "p", "div", "ul", "ol", "li", "h1", "h2", "h3", "h4", "h5", "h6",
    "pre", "blockquote", "hr", "table",
}
_TABLE_PARTS = ["td", "th", "tr", "tbody", "thead", "tfoot", "caption"]
# A multi-line cell holding more than this many characters is prose, not data.
_PROSE_CELL_CHARS = 160
_DATA_VALUE_RE = re.compile(r"[\d.,/:_-]+")


def _own(table, names):
    """Return the descendants of *table* named in *names*, excluding nested tables'."""
    return [e for e in table.find_all(names) if e.find_parent("table") is table]


def _colspan(cell) -> int:
    try:
        return max(1, int(cell.get("colspan", 1)))
    except ValueError:
        return 1


def _is_multiline(cell) -> bool:
    blocks = [
        e for e in cell.find_all(True)
        if e.name in _BLOCK_TAGS and e.find_parent(["td", "th"]) is cell
    ]
    return len(blocks) > 1 or len(cell.find_all("br")) > 1


def _is_data_table(table) -> bool:
    """Return True when *table* holds tabular data rather than page layout."""
    role = (table.get("role") or "").lower()
    if table.find("table") is not None or role in _LAYOUT_ROLES:
        return False

    rows = [row.find_all(["td", "th"], recursive=False) for row in _own(table, "tr")]
    cols = max((sum(_colspan(c) for c in row) for row in rows), default=0)
    if _own(table, ["th", "thead"]) or table.find("caption") is not None or role in _DATA_ROLES:
        return cols >= 2
    if len(rows) < 2 or cols < 2:
        return False

    cells = [c for row in rows for c in row]
    first = [c.get_text(" ", strip=True) for c in rows[0]]
    # Empty-first-column grids are comment threads, legends, and cards.
    if not first or sum(1 for t in first if t) < 0.5 * len(first):
        return False
    if all(not row[0].get_text(strip=True) for row in rows if row):
        return False

    multiline = sum(1 for c in cells if _is_multiline(c))
    longest = max((len(c.get_text(" ", strip=True)) for c in cells), default=0)
    return not (multiline > 0.25 * len(cells) and longest > _PROSE_CELL_CHARS)


def _looks_like_header(row: list) -> bool:
    """Return False when *row* reads as a record: a "Label:" or mostly IDs and numbers."""
    if row and row[0].endswith(":"):
        return False
    return sum(bool(_DATA_VALUE_RE.fullmatch(t)) for t in row) < max(1, len(row) // 2)


def _unwrap_layout_tables(soup) -> None:
    """Flatten layout tables to divs and give header-less data tables a header row."""
    tables = soup.find_all("table")
    data = {id(t) for t in tables if _is_data_table(t)}
    for table in reversed(tables):  # inner tables first
        if id(table) in data:
            rows = _own(table, "tr")
            if rows and not _own(table, "th"):
                first = rows[0].find_all(["td", "th"], recursive=False)
                if _looks_like_header([c.get_text(" ", strip=True) for c in first]):
                    for cell in first:
                        cell.name = "th"
        else:
            for tag in _own(table, _TABLE_PARTS):
                tag.name = "div"
            table.name = "div"


_PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
# Images at most this many pixels wide or tall are tracking pixels and spacers.
_TRACKING_PIXEL_PX = 2


def _pixels(img, dimension: str):
    """Return the declared "width" or "height" of *img* in pixels, or None."""
    number = r"(\d+(?:\.\d+)?)"
    match = re.fullmatch(rf"\s*{number}\s*(?:px)?\s*", img.get(dimension) or "")
    if not match:
        match = re.search(rf"(?<![-\w]){dimension}\s*:\s*{number}\s*px", img.get("style") or "")
    return float(match.group(1)) if match else None


def _is_decorative_image(img) -> bool:
    """Return True for tracking pixels, spacers, and hidden images."""
    style = re.sub(r"\s+", "", img.get("style") or "").lower()
    if "display:none" in style or "visibility:hidden" in style or img.has_attr("hidden"):
        return True
    sizes = (_pixels(img, "width"), _pixels(img, "height"))
    return any(s is not None and s <= _TRACKING_PIXEL_PX for s in sizes)


def _image_target(src: str, inline_images: dict) -> str:
    """Return a reachable, compact reference for an <img> *src*."""
    src = src.strip()
    lowered = src.lower()
    if lowered.startswith("cid:"):
        cid = lowered[4:]
        name = inline_images.get(cid) or inline_images.get(cid.split("@")[0])
        return f"attachment:{name}" if name else src
    if lowered.startswith("data:"):
        return src.split(";", 1)[0].split(",", 1)[0]
    return src


def _mark_images(soup, inline_images: dict) -> None:
    """Drop decorative images and rewrite the rest's alt and src for output."""
    for img in soup.find_all("img"):
        if not (img.get("src") or "").strip() or _is_decorative_image(img):
            img.decompose()
            continue
        img["alt"] = " ".join((img.get("alt") or img.get("title") or "").split())
        img["src"] = _image_target(img["src"], inline_images)


class _EmailMarkdown(MarkdownConverter):
    """MarkdownConverter that keeps images everywhere, even in table cells and headings."""

    def convert_img(self, el, *args, **kwargs):
        alt = " ".join(re.sub(r"[\[\]|]", "", el.get("alt") or "").split())
        target = (el.get("src") or "").replace("|", "%7C")
        if re.search(r"[\s()]", target):
            target = "<" + target.replace("<", "%3C").replace(">", "%3E") + ">"
        return f"![{alt}]({target})"


def _html_to_markdown(html: str, inline_images: Optional[dict] = None) -> str:
    """Convert an email's HTML body to Markdown.

    Data tables stay tables, layout tables are flattened, and images are kept as
    ![alt](target) except tracking pixels. *inline_images* maps Content-IDs to
    attachment filenames.
    """
    lowered = html.lower()
    if "<table" in lowered or "<img" in lowered:
        soup = BeautifulSoup(html, "html.parser")
        _mark_images(soup, inline_images or {})
        _unwrap_layout_tables(soup)
        html = str(soup)
    return _EmailMarkdown().convert(html).strip()


def _inline_images(item) -> dict:
    """Map each attachment's Content-ID (lowercase) to its filename."""
    images = {}
    for att in item.Attachments:
        cid = _mapi_string(att, _PR_ATTACH_CONTENT_ID)
        if cid:
            images[cid.lower()] = att.FileName
    return images


def _item_body(item) -> tuple:
    """Return (body, body_format) for a MailItem: Markdown from its HTML, else plain text."""
    html = item.HTMLBody
    if not html:
        return item.Body or "", "plain"
    images = _inline_images(item) if "cid:" in html.lower() else {}
    return _html_to_markdown(html, images), "markdown"


# ---------------------------------------------------------------------------
# Query-string parser
# ---------------------------------------------------------------------------

# Tokens recognized in the 'query' parameter of search_emails.
# Supported operators (case-insensitive):
#   from:<value>        → sender
#   subject:<value>     → subject
#   body:<value>        → body
#   after:<YYYY-MM-DD>  → since  (also accepts YYYY/MM/DD)
#   before:<YYYY-MM-DD> → until  (also accepts YYYY/MM/DD)
#   newer_than:<value>  → alias for after:
#   older_than:<value>  → alias for before:
# Remaining unrecognized tokens are joined and used as a body substring filter.
_QUERY_TOKEN_RE = re.compile(
    r'(from|subject|body|after|before|newer_than|older_than):("[^"]*"|\S+)',
    re.IGNORECASE,
)


def _parse_query(
    query: str,
    subject: Optional[str],
    sender: Optional[str],
    body: Optional[str],
    since: Optional[str],
    until: Optional[str],
) -> tuple:
    """Parse a Gmail-style query string and merge with any explicit keyword args.

    Explicit keyword arguments always take precedence over values extracted
    from the query string.  Returns (subject, sender, body, since, until).
    """
    remaining = query
    q_subject = q_sender = q_body = q_since = q_until = None

    for m in _QUERY_TOKEN_RE.finditer(query):
        op = m.group(1).lower()
        val = m.group(2).strip('"')
        # Normalise date separators so both YYYY/MM/DD and YYYY-MM-DD work.
        val_norm = val.replace("/", "-")
        if op == "from":
            q_sender = val
        elif op == "subject":
            q_subject = val
        elif op == "body":
            q_body = val
        elif op in ("after", "newer_than"):
            q_since = val_norm
        elif op in ("before", "older_than"):
            q_until = val_norm
        # Remove the matched token from 'remaining' so leftover text can be
        # used as a plain-text body filter.
        remaining = remaining.replace(m.group(0), "", 1)

    # Any leftover words become a body substring filter.
    leftover = remaining.strip()
    if leftover and q_body is None:
        q_body = leftover

    # Explicit keyword args win over query-parsed values.
    return (
        subject if subject is not None else q_subject,
        sender  if sender  is not None else q_sender,
        body    if body    is not None else q_body,
        since   if since   is not None else q_since,
        until   if until   is not None else q_until,
    )


# ---------------------------------------------------------------------------
# Tools
# ---------------------------------------------------------------------------

@mcp.tool(structured_output=False)
def list_email_accounts() -> str:
    """List all email accounts configured in the local Outlook MAPI profile.

    Returns Markdown: a "# Email Accounts" heading, a "count" bullet, and a
    table with the columns account_name and email_address. Pass account_name
    to the other tools.
    """
    _, namespace = _get_outlook()
    rows = [
        (_code(account.DisplayName), _code(account.SmtpAddress))
        for account in namespace.Accounts
    ]
    return _doc(
        "# Email Accounts",
        _kv(count=len(rows)),
        _table(["account_name", "email_address"], rows),
    )


@mcp.tool(structured_output=False)
def list_folders(account_name: str, recursive: bool = False) -> str:
    """List folders in an Outlook account.

    Args:
        account_name: The display name of the account as returned by
            list_email_accounts (e.g. "John Smith").
        recursive: When True, lists all sub-folders at every depth.
            When False (default), lists only the top-level folders.

    Returns Markdown: a "# Folders" heading, bullets for account_name,
    recursive, and count, and a table with the columns folder_path and
    item_count. folder_path is relative to the account root and can be passed
    directly to the other tools.
    """
    _, namespace = _get_outlook()

    root_folder = None
    for store in namespace.Folders:
        if store.Name.lower() == account_name.lower():
            root_folder = store
            break

    if root_folder is None:
        available = [s.Name for s in namespace.Folders]
        raise ValueError(
            f"Account '{account_name}' not found. "
            f"Available accounts: {available}"
        )

    rows = []

    def _walk(folder, path: str):
        try:
            count = folder.Items.Count
        except Exception:
            count = None
        rows.append((_code(path), count))
        if recursive:
            for sub in folder.Folders:
                _walk(sub, f"{path}/{sub.Name}")

    for top in root_folder.Folders:
        _walk(top, top.Name)

    return _doc(
        "# Folders",
        _kv(
            account_name=_code(account_name),
            recursive=recursive,
            count=len(rows),
        ),
        _table(["folder_path", "item_count"], rows),
    )


@mcp.tool(structured_output=False)
def list_emails(
    account_name: str,
    folder_path: str = "Inbox",
    max_results: int = 20,
    since: Optional[str] = None,
    until: Optional[str] = None,
    unread_only: bool = False,
) -> str:
    """List emails in a folder, sorted newest first.

    Use this tool to browse or page through emails without a text search
    filter. To search by subject, sender, or body text, use search_emails
    instead (which also supports date filtering).

    Args:
        account_name: Display name of the Outlook account.
        folder_path: Slash-separated path from the account root
            (e.g. "Inbox" or "Inbox/Projects"). Default: "Inbox".
        max_results: Maximum number of emails to return (1-100). Default: 20.
        since: Optional ISO-8601 date string (e.g. "2026-01-01"). Only emails
            received on or after this date are returned.
        until: Optional ISO-8601 date string (e.g. "2026-12-31"). Only emails
            received on or before this date are returned.
        unread_only: When True, only unread emails are returned. Default: False.

    Returns Markdown: a "# Emails" heading, bullets for account_name,
    folder_path, and count, and a table with one row per email and the
    columns received_time, sender_name, sender_email, subject, is_unread,
    has_attachments, entry_id, and conversation_id. Pass entry_id to
    get_email and conversation_id to get_thread.
    """
    max_count = max(1, min(max_results, 100))
    _, namespace = _get_outlook()
    folder = _resolve_folder(namespace, account_name, folder_path)

    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    # Build a combined Restrict filter for date range and unread flag.
    filters = []
    if since:
        dt = datetime.fromisoformat(since)
        filters.append(f"[ReceivedTime] >= '{dt.strftime('%m/%d/%Y')}'")
    if until:
        dt = datetime.fromisoformat(until)
        filters.append(f"[ReceivedTime] <= '{dt.strftime('%m/%d/%Y')}'")
    if unread_only:
        filters.append("[UnRead] = True")

    if filters:
        try:
            items = items.Restrict(" AND ".join(filters))
        except Exception:
            pass  # fall through to Python-side filtering below

    results = []
    for item in items:
        if len(results) >= max_count:
            break
        try:
            if item.Class != 43:  # olMail = 43; skip meetings, contacts, etc.
                continue
            summary = _mail_item_summary(item)
            # Python-side fallbacks in case Restrict was not applied.
            if since and summary["received_time"] and summary["received_time"] < since:
                continue
            if until and summary["received_time"] and summary["received_time"][:10] > until[:10]:
                continue
            if unread_only and not summary["is_unread"]:
                continue
            results.append(summary)
        except Exception:
            continue

    return _doc(
        "# Emails",
        _kv(
            account_name=_code(account_name),
            folder_path=_code(folder_path),
            count=len(results),
        ),
        _email_table(results),
    )


@mcp.tool(structured_output=False)
def search_emails(
    account_name: str,
    folder_path: str = "Inbox",
    query: Optional[str] = None,
    subject: Optional[str] = None,
    sender: Optional[str] = None,
    body: Optional[str] = None,
    since: Optional[str] = None,
    until: Optional[str] = None,
    max_results: int = 20,
    flagged: Optional[bool] = None,
    include_completed: Optional[bool] = None,
) -> str:
    """Search emails by subject, sender, body text, and/or date range within a folder.

    You may supply filters using either the 'query' convenience parameter
    (Gmail-style syntax) or the individual named parameters — or both.
    When both are given, the named parameter always overrides the
    corresponding query token.  At least one filter, or flagged, must be provided.

    'query' supports the following Gmail-style operators:
        from:<sender>           — match sender name or email address
        subject:<text>          — match subject line
        body:<text>             — match email body
        after:<YYYY-MM-DD>      — received on or after date (also YYYY/MM/DD)
        before:<YYYY-MM-DD>     — received on or before date
        newer_than:<YYYY-MM-DD> — alias for after:
        older_than:<YYYY-MM-DD> — alias for before:
        <plain text>            — unrecognized tokens become a body filter

    Examples:
        query="after:2026-03-28"
        query="from:alice@example.com subject:budget after:2026-01-01"
        query="project update"  (body substring search)

    Args:
        account_name: Display name of the Outlook account (e.g.
            "tom@example.com"). Required.
        folder_path: Slash-separated path from the account root.
            Default: "Inbox".
        query: Optional Gmail-style query string (see above).
        subject: Substring to match against the email subject
            (case-insensitive). Example: subject="project update".
        sender: Substring to match against the sender name or email address
            (case-insensitive). Example: sender="alice@example.com".
        body: Substring to search for within the email body
            (case-insensitive).
        since: ISO-8601 date string. Only emails received on or after this
            date are returned. Example: since="2026-03-29".
        until: ISO-8601 date string. Only emails received on or before this
            date are returned. Example: until="2026-04-02".
        max_results: Maximum number of results to return (1-100). Default: 20.
        flagged: None returns emails regardless of flag state. True returns only
            flagged emails. False returns only emails that are not flagged.
            "flagged" is a named parameter, not a query token. Default: None.
        include_completed: Whether a completed flag counts as flagged. When
            False, completed flags count as not flagged. Default: None (same as False).

    Returns a Markdown list of emails: a "# Search Results" heading, bullets
    for account_name, folder_path, flagged, and count, and a table with one row
    per match and the columns received_time, sender_name, sender_email, subject,
    is_unread, has_attachments, entry_id, and conversation_id. When flagged is
    True, the table also has flag_status ("marked" or "complete") and
    flag_request columns after subject. Pass entry_id to get_email and
    conversation_id to get_thread. When nothing matches, the table is replaced by
    the line "_None._".
    """
    if query:
        subject, sender, body, since, until = _parse_query(
            query, subject, sender, body, since, until
        )

    if flagged is None and subject is None and sender is None and body is None and since is None and until is None:
        raise ValueError(
            "At least one filter must be provided via 'query', the individual "
            "parameters 'subject', 'sender', 'body', 'since', or 'until', or flagged."
        )

    max_count = max(1, min(max_results, 100))
    _, namespace = _get_outlook()
    folder = _resolve_folder(namespace, account_name, folder_path)

    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    # Build a COM-level Restrict filter to reduce iteration cost.
    # Single quotes in user-supplied strings are escaped to prevent
    # the Restrict call from breaking on values that contain apostrophes.
    restrict_applied = False
    filters = []
    if since:
        dt = datetime.fromisoformat(since)
        filters.append(f"[ReceivedTime] >= '{dt.strftime('%m/%d/%Y')}'")
    if until:
        dt = datetime.fromisoformat(until)
        filters.append(f"[ReceivedTime] <= '{dt.strftime('%m/%d/%Y')}'")
    if subject and not sender and not body:
        escaped_subject = subject.replace("'", "''")
        filters.append(
            f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{escaped_subject}%'"
        )
    if filters:
        try:
            items = items.Restrict(" AND ".join(filters))
            restrict_applied = True
        except Exception:
            pass

    subject_lower = subject.lower() if subject else None
    sender_lower = sender.lower() if sender else None
    body_lower = body.lower() if body else None

    flag_names = {2: "marked", 1: "complete"}
    flagged_codes = {2, 1} if include_completed else {2}

    results = []
    for item in items:
        if len(results) >= max_count:
            break
        try:
            if item.Class != 43:
                continue
            # Restrict("[FlagStatus] = 2") under-matches on some profiles, so check here.
            if flagged is not None and (item.FlagStatus in flagged_codes) != flagged:
                continue
            if not restrict_applied and subject_lower:
                if subject_lower not in (item.Subject or "").lower():
                    continue
            if not restrict_applied and since:
                try:
                    if item.ReceivedTime.isoformat()[:10] < since[:10]:
                        continue
                except Exception:
                    pass
            if not restrict_applied and until:
                try:
                    if item.ReceivedTime.isoformat()[:10] > until[:10]:
                        continue
                except Exception:
                    pass
            if sender_lower:
                name_match = sender_lower in (item.SenderName or "").lower()
                email_match = sender_lower in _sender_address(item).lower()
                if not name_match and not email_match:
                    continue
            if body_lower:
                if body_lower not in (item.Body or "").lower():
                    continue
            summary = _mail_item_summary(item)
            if flagged:
                summary["flag_status"] = flag_names[item.FlagStatus]
                summary["flag_request"] = _inline(item.FlagRequest)
            results.append(summary)
        except Exception:
            continue

    return _doc(
        "# Search Results",
        _kv(
            account_name=_code(account_name),
            folder_path=_code(folder_path),
            flagged=flagged,
            count=len(results),
        ),
        _email_table(results, with_flags=flagged is True),
    )


@mcp.tool(structured_output=False)
def get_email(entry_id: str) -> str:
    """Retrieve the full content of a single email by its entry ID.

    The body is returned as Markdown converted from the HTML source when
    available, preserving hyperlinks, headings, and formatting that
    Outlook's plain-text rendering would otherwise discard. Tables that hold
    data become Markdown tables; tables used only for page layout are
    flattened to plain paragraphs. Falls back to Outlook's plain-text Body if
    no HTML source is present.

    Images appear in place as ![alt](target). The target is an attachment
    ("attachment:image001.png"; pass the filename to read_attachment_image_ocr
    to read text inside it) or an image URL. Tracking pixels are omitted.

    Args:
        entry_id: The EntryID string from a previous list_emails or
            search_emails call.

    Returns Markdown: an "# Email" heading, a bullet list of entry_id,
    conversation_id, subject, sender_name, sender_email, recipients (nested
    bullets), received_time, is_unread, has_attachments, attachment_count,
    and body_format ("markdown" or "plain"), then a "## Body" section whose
    fenced code block holds the body text.
    """
    _, namespace = _get_outlook()
    item = namespace.GetItemFromID(entry_id)

    recipients = [_recipient_address(r) for r in item.Recipients]

    body, body_format = _item_body(item)

    result = _mail_item_summary(item)
    result["body"] = body
    result["body_format"] = body_format
    result["recipients"] = recipients
    result["attachment_count"] = item.Attachments.Count

    return _email_block("Email", result, 1)


@mcp.tool(structured_output=False)
def list_attachments(entry_id: str) -> str:
    """List all attachments on an email.

    Args:
        entry_id: The EntryID of the email containing the attachments.

    Returns Markdown: an "# Attachments" heading, bullets for entry_id and
    count, and a table with the columns filename, size_bytes, and
    attachment_type.

    attachment_type values: 1 = By Value, 2 = By Reference,
    4 = Embedded Item, 5 = OLE.
    """
    _, namespace = _get_outlook()
    item = namespace.GetItemFromID(entry_id)

    rows = [
        (_code(att.FileName), att.Size, att.Type)
        for att in item.Attachments
    ]

    return _doc(
        "# Attachments",
        _kv(entry_id=_code(entry_id), count=len(rows)),
        _table(["filename", "size_bytes", "attachment_type"], rows),
    )


@mcp.tool(structured_output=False)
def download_attachment(entry_id: str, filename: str, output_dir: str) -> str:
    """Download a specific email attachment to a local directory.

    Args:
        entry_id: The EntryID of the email containing the attachment.
        filename: The exact filename of the attachment to download, as
            returned by list_attachments.
        output_dir: Absolute path to the directory where the file will be
            saved. The directory must already exist.

    Returns Markdown: an "# Attachment Downloaded" heading and bullets for
    filename and path (the absolute path of the saved file).
    """
    output_dir_path = Path(output_dir).resolve()
    if not output_dir_path.is_dir():
        raise ValueError(
            f"output_dir does not exist or is not a directory: {output_dir}"
        )

    # Security: resolve final path and confirm it stays within output_dir.
    # This blocks path traversal via filenames containing ".." or separators.
    output_path = (output_dir_path / filename).resolve()
    if output_path.parent != output_dir_path:
        raise ValueError(
            f"Resolved output path '{output_path}' escapes the output directory. "
            "filename must not contain path separators or '..' components."
        )

    _, namespace = _get_outlook()
    item = namespace.GetItemFromID(entry_id)

    for att in item.Attachments:
        if att.FileName == filename:
            att.SaveAsFile(str(output_path))
            return _doc(
                "# Attachment Downloaded",
                _kv(filename=_code(filename), path=_code(output_path)),
            )

    raise ValueError(f"Attachment '{filename}' not found on email '{entry_id}'.")

# File extensions treated as plain-text and returned as-is.
_TEXT_EXTENSIONS = {
    ".txt", ".md", ".rst", ".csv", ".tsv", ".log",
    ".json", ".yaml", ".yml", ".toml", ".ini", ".cfg", ".conf",
    ".xml", ".html", ".htm", ".xhtml",
    ".py", ".pyw", ".js", ".ts", ".jsx", ".tsx",
    ".c", ".h", ".cpp", ".hpp", ".cs", ".java", ".go", ".rs",
    ".sh", ".bash", ".zsh", ".ps1", ".bat", ".cmd",
    ".sql", ".r", ".rb", ".php",
    ".diff", ".patch",
}

# ---------------------------------------------------------------------------
# Tesseract PATH auto-detection via the Windows registry
# ---------------------------------------------------------------------------

def _find_tesseract_in_registry() -> Optional[str]:
    """Search HKLM Uninstall keys for a Tesseract-OCR entry and return its install dir."""
    uninstall_keys = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
    ]
    for key_path in uninstall_keys:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                i = 0
                while True:
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        with winreg.OpenKey(key, subkey_name) as subkey:
                            try:
                                display_name, _ = winreg.QueryValueEx(subkey, "DisplayName")
                                if "Tesseract-OCR" in display_name:
                                    try:
                                        uninstall_str, _ = winreg.QueryValueEx(subkey, "UninstallString")
                                        return str(Path(uninstall_str.strip('"')).parent)
                                    except FileNotFoundError:
                                        pass
                            except FileNotFoundError:
                                pass
                        i += 1
                    except OSError:
                        break
        except OSError:
            continue
    return None


def _ensure_tesseract_on_path() -> None:
    """Add the Tesseract-OCR directory to PATH if tesseract is not already accessible."""
    import shutil
    if shutil.which("tesseract"):
        return
    install_dir = _find_tesseract_in_registry()
    if install_dir and install_dir not in os.environ.get("PATH", ""):
        os.environ["PATH"] = install_dir + os.pathsep + os.environ.get("PATH", "")


_ensure_tesseract_on_path()

# Image file extensions supported by the OCR tool.
_IMAGE_EXTENSIONS = {
    ".png", ".jpg", ".jpeg", ".gif",
    ".bmp", ".tiff", ".tif", ".webp",
}
# Inline images from Teams and Jira are often attached with no file extension.
_PR_ATTACH_MIME_TAG = "http://schemas.microsoft.com/mapi/proptag/0x370E001F"
_IMAGE_MIME_TYPES = {
    "image/png", "image/jpeg", "image/gif",
    "image/bmp", "image/tiff", "image/webp",
}

@mcp.tool(structured_output=False)
def read_attachment_text(
    entry_id: str,
    filename: str,
    encoding: str = "utf-8",
    max_chars: int = 100_000,
) -> str:
    """Read the text content of an email attachment without saving it permanently.

    Extracts the attachment to a secure temporary file, reads its content,
    and deletes the temporary file before returning. This is more convenient
    than download_attachment when you only need to inspect the text — nothing
    is left on disk after the call.

    Supported file types (plain-text formats):
        .txt .md .rst .csv .tsv .log .json .yaml .yml .toml .ini .cfg .conf
        .xml .html .htm .py .js .ts .c .h .cpp .cs .java .go .rs .sh .ps1
        .bat .sql .r .rb .php .diff .patch

    Binary formats (e.g. .docx, .xlsx, .pdf, .png) are not supported; use
    download_attachment to save those to disk for processing by other tools.

    Args:
        entry_id: The EntryID of the email containing the attachment.
        filename: The exact filename of the attachment, as returned by
            list_attachments.
        encoding: Text encoding to use when decoding the file.
            Default: "utf-8". Common alternatives: "latin-1", "cp1252".
        max_chars: Maximum number of characters to return. Content beyond
            this limit is cut off and "truncated" is true.
            Default: 100,000.

    Returns Markdown: an "# Attachment Text" heading, bullets for filename,
    encoding, char_count (the length of the text returned), and truncated,
    then a "## Content" section whose fenced code block holds the file text.
    When truncated, a line after the block states the limit.
    """
    suffix = Path(filename).suffix.lower()
    if suffix not in _TEXT_EXTENSIONS:
        raise ValueError(
            f"'{filename}' has an unsupported extension ('{suffix}'). "
            f"Supported extensions: {', '.join(sorted(_TEXT_EXTENSIONS))}. "
            "Use download_attachment for binary formats."
        )

    _, namespace = _get_outlook()
    item = namespace.GetItemFromID(entry_id)

    target_att = None
    for att in item.Attachments:
        if att.FileName == filename:
            target_att = att
            break

    if target_att is None:
        raise ValueError(f"Attachment '{filename}' not found on email '{entry_id}'.")

    # Save to a temp file, read it, then delete immediately.
    with tempfile.NamedTemporaryFile(
        suffix=suffix, delete=False
    ) as tmp:
        tmp_path = Path(tmp.name)

    try:
        target_att.SaveAsFile(str(tmp_path))
        raw = tmp_path.read_bytes()
    finally:
        try:
            tmp_path.unlink()
        except OSError:
            pass

    content = _lf(raw.decode(encoding, errors="replace"))
    truncated = len(content) > max_chars
    if truncated:
        content = content[:max_chars]

    return _doc(
        "# Attachment Text",
        _kv(
            filename=_code(filename),
            encoding=encoding,
            char_count=len(content),
            truncated=truncated,
        ),
        "## Content",
        _fence(content),
        f"_Truncated at {max_chars:,} characters._" if truncated else "",
    )

@mcp.tool(structured_output=False)
def read_attachment_image_ocr(
    entry_id: str,
    filename: str,
    lang: str = "eng",
) -> str:
    """Extract text from an image email attachment using OCR (Tesseract).

    Downloads the image attachment to a temporary file, runs
    Tesseract OCR on it, deletes the temporary file, and returns the
    extracted text. The agent only ever sees the final text output;
    no file paths or intermediate state are exposed.

    Requires Tesseract to be installed on the system. Download from:
    https://github.com/UB-Mannheim/tesseract/wiki

    Supported image formats:
        .png .jpg .jpeg .gif .bmp .tiff .tif .webp
    An attachment with no file extension is also accepted when its MIME type
    is one of these image types (typical of inline images from Teams or Jira).

    Args:
        entry_id: The EntryID of the email containing the attachment.
        filename: The exact filename of the attachment, as returned by
            list_attachments or shown in an image reference such as
            ![alt](attachment:image001.png) in get_email output.
        lang: Tesseract language code to use for OCR. Default: "eng"
            (English). Use a valid Tesseract language code for other
            languages (e.g. "fra" for French, "deu" for German).

    Returns Markdown: an "# Attachment OCR Text" heading, bullets for
    filename, lang, and char_count, then a "## Text" section whose fenced
    code block holds the OCR-extracted text.
    """
    suffix = Path(filename).suffix.lower()

    _, namespace = _get_outlook()
    item = namespace.GetItemFromID(entry_id)

    target_att = None
    for att in item.Attachments:
        if att.FileName == filename:
            target_att = att
            break

    if target_att is None:
        raise ValueError(f"Attachment '{filename}' not found on email '{entry_id}'.")

    is_image = suffix in _IMAGE_EXTENSIONS or (
        not suffix and _mapi_string(target_att, _PR_ATTACH_MIME_TAG).lower() in _IMAGE_MIME_TYPES
    )
    if not is_image:
        raise ValueError(
            f"'{filename}' has an unsupported extension ('{suffix}'). "
            f"Supported image extensions: {', '.join(sorted(_IMAGE_EXTENSIONS))}."
        )

    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp_path = Path(tmp.name)

    try:
        target_att.SaveAsFile(str(tmp_path))
        with Image.open(tmp_path) as img:
            text = _lf(pytesseract.image_to_string(img, lang=lang))
    finally:
        try:
            tmp_path.unlink()
        except OSError:
            pass

    return _doc(
        "# Attachment OCR Text",
        _kv(filename=_code(filename), lang=lang, char_count=len(text)),
        "## Text",
        _fence(text),
    )


@mcp.tool(structured_output=False)
def get_thread(
    account_name: str,
    conversation_id: str,
    folder_paths: Optional[list] = None,
) -> str:
    """Retrieve all emails in a conversation thread.

    Searches the specified folders for every email sharing the given
    conversation ID, sorted oldest first so the exchange reads
    chronologically.

    Args:
        account_name: Display name of the Outlook account.
        conversation_id: The conversation_id string from a previous
            list_emails or search_emails result.
        folder_paths: List of slash-separated folder paths to search within
            the account. Defaults to ["Inbox", "Sent Items"].

    Returns Markdown: a "# Thread" heading, bullets for account_name,
    conversation_id, and message_count, then one "## Message N" section per
    email, oldest first. Each section has the same bullets and "### Body"
    block as get_email.
    """
    if folder_paths is None:
        folder_paths = ["Inbox", "Sent Items"]

    _, namespace = _get_outlook()

    results = []
    seen_ids = set()

    for folder_path in folder_paths:
        try:
            folder = _resolve_folder(namespace, account_name, folder_path)
        except ValueError:
            continue  # skip folders that don't exist in this account

        items = folder.Items
        try:
            items = items.Restrict(f"[ConversationID] = '{conversation_id}'")
        except Exception:
            pass  # fall through to Python-side filtering

        for item in items:
            try:
                if item.Class != 43:
                    continue
                if item.EntryID in seen_ids:
                    continue
                if item.ConversationID != conversation_id:
                    continue  # Python-side guard if Restrict was not applied
                seen_ids.add(item.EntryID)

                recipients = [_recipient_address(r) for r in item.Recipients]

                body, body_format = _item_body(item)
                entry = _mail_item_summary(item)
                entry["body"] = body
                entry["body_format"] = body_format
                entry["recipients"] = recipients
                entry["attachment_count"] = item.Attachments.Count
                results.append(entry)
            except Exception:
                continue

    # Sort oldest first for chronological reading.
    results.sort(key=lambda e: e["received_time"] or "")

    messages = [
        _email_block(f"Message {n}", entry, 2)
        for n, entry in enumerate(results, 1)
    ]
    return _doc(
        "# Thread",
        _kv(
            account_name=_code(account_name),
            conversation_id=_code(conversation_id),
            message_count=len(results),
        ),
        *(messages or [_NONE]),
    )

@mcp.tool(structured_output=False)
def get_folder_counts(account_name: str, folder_path: str = "Inbox") -> str:
    """Get the total and unread item counts for a folder.

    A lightweight alternative to list_emails when only counts are needed,
    e.g. for inbox triage or status summaries.

    Args:
        account_name: Display name of the Outlook account.
        folder_path: Slash-separated path from the account root.
            Default: "Inbox".

    Returns Markdown: a "# Folder Counts" heading and bullets for
    folder_path, total_count, and unread_count.
    """
    _, namespace = _get_outlook()
    folder = _resolve_folder(namespace, account_name, folder_path)

    return _doc(
        "# Folder Counts",
        _kv(
            folder_path=_code(folder_path),
            total_count=folder.Items.Count,
            unread_count=folder.UnReadItemCount,
        ),
    )

@mcp.tool(structured_output=False)
def create_draft(
    to: list,
    subject: str,
    body: str,
    cc: Optional[list] = None,
    bcc: Optional[list] = None,
) -> str:
    """Create and save an email draft in Outlook.

    The draft is saved to the Drafts folder and is NOT sent.

    Args:
        to: List of recipient email addresses or display names.
        subject: Subject line of the email.
        body: Markdown-formatted body of the email. Standard text is rendered in
            Aptos 11pt; code spans and fenced code blocks use Consolas 11pt.
        cc: Optional list of CC recipient addresses or display names.
        bcc: Optional list of BCC recipient addresses or display names.

    Returns Markdown: a "# Draft Created" heading and bullets for the
    entry_id of the saved draft and a confirmation message.
    """
    cc = cc or []
    bcc = bcc or []

    outlook, _ = _get_outlook()
    mail = outlook.CreateItem(0)  # 0 = olMailItem

    mail.Subject = subject
    mail.HTMLBody = _markdown_to_html(body)

    for address in to:
        r = mail.Recipients.Add(address)
        r.Type = 1  # olTo

    for address in cc:
        r = mail.Recipients.Add(address)
        r.Type = 2  # olCC

    for address in bcc:
        r = mail.Recipients.Add(address)
        r.Type = 3  # olBCC

    mail.Save()

    # Open the draft in an Outlook inspector window so the user can see it.
    inspector = mail.GetInspector
    inspector.Display(False)  # False = non-modal, does not block

    # Trigger a sync so the draft propagates to Exchange immediately.
    _, namespace = _get_outlook()
    namespace.SendAndReceive(False)

    return _draft_result(
        "Draft Created", mail.EntryID, "Draft saved and opened in Outlook."
    )

@mcp.tool(structured_output=False)
def create_reply(
    entry_id: str,
    body: str,
    reply_all: bool = False,
) -> str:
    """Create a reply to an existing email and open it for review in Outlook.

    The reply is pre-populated with the correct To/CC recipients, subject
    (Re: ...), and quoted original message, exactly as Outlook's own Reply
    button would produce. It is saved as a draft and opened in an Outlook
    compose window for the user to review and send — it is NOT sent
    automatically.

    Args:
        entry_id: The EntryID of the email to reply to, as returned by
            list_emails, search_emails, or get_email.
        body: Markdown-formatted body to insert above the quoted original.
            Standard text is rendered in Aptos 11pt; code spans and fenced
            code blocks use Consolas 11pt.
        reply_all: When True, replies to all recipients (Reply All).
            When False (default), replies only to the sender.

    Returns Markdown: a "# Reply Draft Created" heading and bullets for the
    entry_id of the saved draft reply and a confirmation message.
    """
    _, namespace = _get_outlook()
    original = namespace.GetItemFromID(entry_id)

    if reply_all:
        reply = original.ReplyAll()
    else:
        reply = original.Reply()

    reply.HTMLBody = _prepend_html(_md_to_fragment(body), reply.HTMLBody)

    reply.Save()

    inspector = reply.GetInspector
    inspector.Display(False)  # Non-modal — user reviews and sends manually

    namespace.SendAndReceive(False)

    return _draft_result(
        "Reply Draft Created",
        reply.EntryID,
        "Reply draft opened in Outlook. Review and click Send when ready.",
    )

@mcp.tool(structured_output=False)
def create_forward(
    entry_id: str,
    to: list,
    body: str,
    cc: Optional[list] = None,
    bcc: Optional[list] = None,
) -> str:
    """Create a forward draft for an existing email and open it in Outlook.

    The forward is pre-populated with the original message body (Fw: ...),
    exactly as Outlook's own Forward button would produce. It is saved as a
    draft and opened in an Outlook compose window for the user to review and
    send — it is NOT sent automatically.

    Args:
        entry_id: The EntryID of the email to forward, as returned by
            list_emails, search_emails, or get_email.
        to: List of recipient email addresses or display names.
        body: Markdown-formatted message to insert above the forwarded original.
            Standard text is rendered in Aptos 11pt; code spans and fenced
            code blocks use Consolas 11pt.
        cc: Optional list of CC recipient addresses or display names.
        bcc: Optional list of BCC recipient addresses or display names.

    Returns Markdown: a "# Forward Draft Created" heading and bullets for
    the entry_id of the saved forward draft and a confirmation message.
    """
    cc = cc or []
    bcc = bcc or []

    _, namespace = _get_outlook()
    original = namespace.GetItemFromID(entry_id)
    fwd = original.Forward()

    fwd.HTMLBody = _prepend_html(_md_to_fragment(body), fwd.HTMLBody)

    for address in to:
        r = fwd.Recipients.Add(address)
        r.Type = 1  # olTo

    for address in cc:
        r = fwd.Recipients.Add(address)
        r.Type = 2  # olCC

    for address in bcc:
        r = fwd.Recipients.Add(address)
        r.Type = 3  # olBCC

    fwd.Save()

    inspector = fwd.GetInspector
    inspector.Display(False)  # Non-modal — user reviews and sends manually

    namespace.SendAndReceive(False)

    return _draft_result(
        "Forward Draft Created",
        fwd.EntryID,
        "Forward draft opened in Outlook. Review and click Send when ready.",
    )

if __name__ == "__main__":
    mcp.run()
