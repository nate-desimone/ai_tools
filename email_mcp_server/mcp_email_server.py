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

Usage:
    python mcp_email_server.py
"""

import json
import re
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Optional

import html2text
import markdown as _markdown
import pythoncom
import win32com.client

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


def _mail_item_summary(item) -> dict:
    """Return a concise summary dict for a MailItem."""
    try:
        received = item.ReceivedTime
        received_str = received.isoformat() if hasattr(received, "isoformat") else str(received)
    except Exception:
        received_str = None

    try:
        sender_email = item.SenderEmailAddress
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
# Remaining unrecognised tokens are joined and used as a body substring filter.
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

@mcp.tool()
def list_email_accounts() -> str:
    """List all email accounts configured in the local Outlook MAPI profile.

    Returns a newline-separated list of accounts in the format
    "DisplayName <email>".
    """
    _, namespace = _get_outlook()
    lines = []
    for account in namespace.Accounts:
        display = account.DisplayName or ""
        email = account.SmtpAddress or ""
        if email:
            lines.append(f"{display} <{email}>")
        else:
            lines.append(display)

    if not lines:
        return "No accounts found in the Outlook MAPI profile."
    return "\n".join(lines)


@mcp.tool()
def list_folders(account_name: str, recursive: bool = False) -> str:
    """List folders in an Outlook account.

    Args:
        account_name: The display name of the account as returned by
            list_email_accounts (e.g. "John Smith").
        recursive: When True, lists all sub-folders at every depth.
            When False (default), lists only the top-level folders.

    Returns a text tree showing each folder name and its item count.
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

    lines = []

    def _walk(folder, indent: int):
        try:
            count = folder.Items.Count
        except Exception:
            count = "?"
        lines.append("  " * indent + f"{folder.Name}  ({count} items)")
        if recursive:
            for sub in folder.Folders:
                _walk(sub, indent + 1)

    for top in root_folder.Folders:
        _walk(top, 0)

    return "\n".join(lines) if lines else "No folders found."


@mcp.tool()
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

    Returns a JSON array of email summary objects, each containing:
        entry_id, subject, sender_name, sender_email, received_time,
        has_attachments, is_unread, conversation_id.
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

    return json.dumps(results, indent=2)


@mcp.tool()
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
) -> str:
    """Search emails by subject, sender, body text, and/or date range within a folder.

    You may supply filters using either the 'query' convenience parameter
    (Gmail-style syntax) or the individual named parameters — or both.
    When both are given, the named parameter always overrides the
    corresponding query token.  At least one filter must be provided.

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

    Returns a JSON array of matching email summary objects, each containing:
        entry_id, subject, sender_name, sender_email, received_time,
        has_attachments, is_unread, conversation_id.
    Returns an empty array if no emails match.
    """
    if query:
        subject, sender, body, since, until = _parse_query(
            query, subject, sender, body, since, until
        )

    if subject is None and sender is None and body is None and since is None and until is None:
        raise ValueError(
            "At least one filter must be provided via 'query' or the individual "
            "parameters 'subject', 'sender', 'body', 'since', or 'until'."
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

    results = []
    for item in items:
        if len(results) >= max_count:
            break
        try:
            if item.Class != 43:
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
                email_match = sender_lower in (item.SenderEmailAddress or "").lower()
                if not name_match and not email_match:
                    continue
            if body_lower:
                if body_lower not in (item.Body or "").lower():
                    continue
            results.append(_mail_item_summary(item))
        except Exception:
            continue

    return json.dumps(results, indent=2)


@mcp.tool()
def get_email(entry_id: str) -> str:
    """Retrieve the full content of a single email by its entry ID.

    The body is returned as Markdown converted from the HTML source when
    available, preserving hyperlinks, tables, headings, and formatting that
    Outlook's plain-text rendering would otherwise discard. Falls back to
    Outlook's plain-text Body if no HTML source is present.

    Args:
        entry_id: The EntryID string from a previous list_emails or
            search_emails call.

    Returns a JSON object containing:
        entry_id, subject, sender_name, sender_email, received_time,
        has_attachments, body (Markdown), body_format ("markdown" or "plain"),
        recipients (list of addresses), attachment_count.
    """
    _, namespace = _get_outlook()
    item = namespace.GetItemFromID(entry_id)

    recipients = []
    for r in item.Recipients:
        try:
            recipients.append(r.Address or r.Name)
        except Exception:
            recipients.append(r.Name)

    html_body = item.HTMLBody
    if html_body:
        converter = html2text.HTML2Text()
        converter.ignore_images = True   # image tags add noise, not value
        converter.body_width = 0         # disable line wrapping
        body = converter.handle(html_body).strip()
        body_format = "markdown"
    else:
        body = item.Body or ""
        body_format = "plain"

    result = _mail_item_summary(item)
    result["body"] = body
    result["body_format"] = body_format
    result["recipients"] = recipients
    result["attachment_count"] = item.Attachments.Count

    return json.dumps(result, indent=2)


@mcp.tool()
def list_attachments(entry_id: str) -> str:
    """List all attachments on an email.

    Args:
        entry_id: The EntryID of the email containing the attachments.

    Returns a JSON array of attachment objects, each containing:
        filename, size_bytes, attachment_type.

    attachment_type values: 1 = By Value, 2 = By Reference,
    4 = Embedded Item, 5 = OLE.
    """
    _, namespace = _get_outlook()
    item = namespace.GetItemFromID(entry_id)

    results = []
    for att in item.Attachments:
        results.append({
            "filename": att.FileName,
            "size_bytes": att.Size,
            "attachment_type": att.Type,
        })

    return json.dumps(results, indent=2)


@mcp.tool()
def download_attachment(entry_id: str, filename: str, output_dir: str) -> str:
    """Download a specific email attachment to a local directory.

    Args:
        entry_id: The EntryID of the email containing the attachment.
        filename: The exact filename of the attachment to download, as
            returned by list_attachments.
        output_dir: Absolute path to the directory where the file will be
            saved. The directory must already exist.

    Returns the absolute path of the saved file.
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
            return str(output_path)

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

@mcp.tool()
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
            this limit is truncated and a notice is appended.
            Default: 100,000.

    Returns a JSON object containing:
        filename, encoding, char_count, truncated (bool), content (str).
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

    content = raw.decode(encoding, errors="replace")
    truncated = len(content) > max_chars
    if truncated:
        content = content[:max_chars] + f"\n\n[... truncated at {max_chars:,} characters ...]"

    return json.dumps({
        "filename": filename,
        "encoding": encoding,
        "char_count": len(content),
        "truncated": truncated,
        "content": content,
    }, indent=2)

@mcp.tool()
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

    Returns a JSON array of full email objects (same fields as get_email)
    sorted from oldest to newest.
    """
    if folder_paths is None:
        folder_paths = ["Inbox", "Sent Items"]

    _, namespace = _get_outlook()

    def _get_body(item):
        html_body = item.HTMLBody
        if html_body:
            converter = html2text.HTML2Text()
            converter.ignore_images = True
            converter.body_width = 0
            return converter.handle(html_body).strip(), "markdown"
        return item.Body or "", "plain"

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

                recipients = []
                for r in item.Recipients:
                    try:
                        recipients.append(r.Address or r.Name)
                    except Exception:
                        recipients.append(r.Name)

                body, body_format = _get_body(item)
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

    return json.dumps(results, indent=2)

@mcp.tool()
def get_folder_counts(account_name: str, folder_path: str = "Inbox") -> str:
    """Get the total and unread item counts for a folder.

    A lightweight alternative to list_emails when only counts are needed,
    e.g. for inbox triage or status summaries.

    Args:
        account_name: Display name of the Outlook account.
        folder_path: Slash-separated path from the account root.
            Default: "Inbox".

    Returns a JSON object containing:
        folder_path, total_count, unread_count.
    """
    _, namespace = _get_outlook()
    folder = _resolve_folder(namespace, account_name, folder_path)

    return json.dumps({
        "folder_path": folder_path,
        "total_count": folder.Items.Count,
        "unread_count": folder.UnReadItemCount,
    }, indent=2)

@mcp.tool()
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

    Returns a JSON object containing the entry_id of the saved draft and a
    confirmation message.
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

    return json.dumps({
        "entry_id": mail.EntryID,
        "message": "Draft saved and opened in Outlook.",
    }, indent=2)

@mcp.tool()
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

    Returns a JSON object containing the entry_id of the saved draft reply
    and a confirmation message.
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

    return json.dumps({
        "entry_id": reply.EntryID,
        "message": (
            "Reply draft opened in Outlook. Review and click Send when ready."
        ),
    }, indent=2)

@mcp.tool()
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

    Returns a JSON object containing the entry_id of the saved forward draft
    and a confirmation message.
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

    return json.dumps({
        "entry_id": fwd.EntryID,
        "message": (
            "Forward draft opened in Outlook. Review and click Send when ready."
        ),
    }, indent=2)

if __name__ == "__main__":
    mcp.run()
