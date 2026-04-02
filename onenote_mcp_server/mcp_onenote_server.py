## @file
# MCP Server for OneNote access via COM automation.
#
# Copyright 2026 Intel Corporation All Rights Reserved.
# SPDX-License-Identifier: BSD-2-Clause-Patent
##
"""
MCP Server for Microsoft OneNote access via COM automation.

Provides tools for AI agents to read, search, and edit OneNote notebooks,
sections, and pages on Windows systems using the classic OneNote 2016/2019
desktop application.

Requires Microsoft OneNote 2016 or 2019 (classic desktop) to be installed.
The server will start OneNote automatically if it is not already running.

Usage:
    python mcp_onenote_server.py
"""

import html as _html_lib
import json
import re
import time
from xml.etree import ElementTree as ET

import comtypes.client
import html2text
import markdown as _markdown

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("OneNote")


# ---------------------------------------------------------------------------
# COM API constants
# ---------------------------------------------------------------------------

# XML namespace used by OneNote 2016/2019
_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"

# HierarchyScope enum (values verified from the OneNote 2016 type library)
_HS_SELF = 0           # The node itself only
_HS_CHILDREN = 1       # Immediate children
_HS_NOTEBOOKS = 2      # All notebooks (used when root is "")
_HS_SECTIONS = 3       # All sections and section groups
_HS_PAGES = 4          # All pages

# CreateFileType enum (used with OpenHierarchy to create sections)
_CFT_SECTION = 3

# NewPageStyle enum (used with CreateNewPage)
_NPS_BLANK_WITH_TITLE = 1

# PageInfo enum (used with GetPageContent; 0 = basic text content)
_PI_BASIC = 0

# Characters not allowed in OneNote section names (Windows file-system chars)
_INVALID_SECTION_CHARS = frozenset(r'\/:*?"<>|')

# OneNote 2016/2019 type library GUID
_ONENOTE_TYPELIB_GUID = '{0EA692EE-BB50-4E3C-AEF0-356D91732725}'

# COM retry settings for RPC_E_CALL_REJECTED / RPC_E_SERVERCALL_RETRYLATER.
# Office apps temporarily reject COM calls when the UI thread is busy.
_RPC_E_CALL_REJECTED       = -2147418111  # 0x80010001
_RPC_E_SERVERCALL_RETRYLATER = -2147417846  # 0x8001010A
_COM_MAX_RETRIES = 10
_COM_RETRY_DELAY = 0.1  # seconds

# Cached comtypes module (generated from typelib on first call)
_onenote_mod = None


class _ComRetryProxy:
    """Transparent proxy around a comtypes COM object.

    Wraps every method call with a retry loop that sleeps and retries when
    Office returns RPC_E_CALL_REJECTED or RPC_E_SERVERCALL_RETRYLATER.  These
    errors occur when the Office UI thread is momentarily busy.
    """

    def __init__(self, obj):
        object.__setattr__(self, '_obj', obj)

    def __getattr__(self, name):
        attr = getattr(object.__getattribute__(self, '_obj'), name)
        if not callable(attr):
            return attr

        def _retrying(*args, **kwargs):
            import comtypes
            for attempt in range(_COM_MAX_RETRIES):
                try:
                    return attr(*args, **kwargs)
                except comtypes.COMError as exc:
                    if exc.hresult in (
                        _RPC_E_CALL_REJECTED,
                        _RPC_E_SERVERCALL_RETRYLATER,
                    ) and attempt < _COM_MAX_RETRIES - 1:
                        time.sleep(_COM_RETRY_DELAY)
                        continue
                    raise

        return _retrying


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _get_onenote():
    """Return the OneNote IApplication COM object via comtypes.

    comtypes generates vtable bindings from the OneNote type library, which
    is required because IApplication is a TKIND_INTERFACE (pure vtable)
    type that cannot be called via IDispatch::Invoke.

    Creates a new COM reference on each call.  For a single-instance app
    like OneNote, CoCreateInstance routes to the running instance.
    """
    global _onenote_mod
    if _onenote_mod is None:
        _onenote_mod = comtypes.client.GetModule((_ONENOTE_TYPELIB_GUID, 1, 1))
    obj = comtypes.client.CreateObject(
        'OneNote.Application',
        interface=_onenote_mod.IApplication,
    )
    return _ComRetryProxy(obj)


def _tag(local: str) -> str:
    """Return the fully-qualified ElementTree tag for a OneNote XML element."""
    return f"{{{_NS}}}{local}"


def _parse_xml(xml_str: str) -> ET.Element:
    """Parse a OneNote XML string, stripping any leading BOM."""
    if xml_str and xml_str[0] == "\ufeff":
        xml_str = xml_str[1:]
    ET.register_namespace("one", _NS)
    return ET.fromstring(xml_str)


def _sections_recursive(element: ET.Element, parent_id: str) -> list:
    """Walk a hierarchy XML element and return a flat list of section dicts.

    Recurses into SectionGroup elements. Skips the built-in Recycle Bin
    section group.
    """
    results = []
    for child in element:
        local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if local == "Section":
            results.append({
                "id": child.get("ID", ""),
                "name": child.get("name", ""),
                "type": "Section",
                "parent_id": parent_id,
                "last_modified_time": child.get("lastModifiedTime", ""),
                "color": child.get("color", ""),
                "is_unread": child.get("isUnread", "false") == "true",
                "is_locked": child.get("locked", "false") == "true",
            })
        elif local == "SectionGroup":
            # Skip the built-in OneNote Recycle Bin
            if child.get("isRecycleBin") == "true":
                continue
            group_id = child.get("ID", "")
            results.append({
                "id": group_id,
                "name": child.get("name", ""),
                "type": "SectionGroup",
                "parent_id": parent_id,
                "last_modified_time": child.get("lastModifiedTime", ""),
                "color": "",
                "is_unread": child.get("isUnread", "false") == "true",
                "is_locked": False,
            })
            results.extend(_sections_recursive(child, group_id))
    return results


def _extract_search_results(root: ET.Element, max_results: int) -> list:
    """Recursively walk FindPages result XML and return a flat list of page dicts."""
    results = []

    def _walk(element, notebook_id="", notebook_name="",
              section_id="", section_name=""):
        if len(results) >= max_results:
            return
        local = element.tag.split("}")[-1] if "}" in element.tag else element.tag
        if local == "Notebooks":
            for child in element:
                _walk(child, notebook_id, notebook_name, section_id, section_name)
        elif local == "Notebook":
            nb_id = element.get("ID", "")
            nb_name = element.get("name", "")
            for child in element:
                _walk(child, nb_id, nb_name, section_id, section_name)
        elif local == "SectionGroup":
            for child in element:
                _walk(child, notebook_id, notebook_name, section_id, section_name)
        elif local == "Section":
            sec_id = element.get("ID", "")
            sec_name = element.get("name", "")
            for child in element:
                _walk(child, notebook_id, notebook_name, sec_id, sec_name)
        elif local == "Page":
            results.append({
                "page_id": element.get("ID", ""),
                "name": element.get("name", ""),
                "date_time": element.get("dateTime", ""),
                "last_modified_time": element.get("lastModifiedTime", ""),
                "section_id": section_id,
                "section_name": section_name,
                "notebook_id": notebook_id,
                "notebook_name": notebook_name,
            })

    _walk(root)
    return results


# ---------------------------------------------------------------------------
# Markdown <-> OneNote XML conversion
# ---------------------------------------------------------------------------

_MD_EXTENSIONS = ["fenced_code", "tables"]
_H2T = html2text.HTML2Text()
_H2T.ignore_images = True
_H2T.body_width = 0
_H2T.strong_mark = "**"
_H2T.ul_item_mark = "-"
# Use asterisks for italic so round-trip markdown uses *text* not _text_
try:
    _H2T.emphasis_mark = "*"
except AttributeError:
    pass  # older html2text versions don't expose this

# ---------------------------------------------------------------------------
# Read-path helpers: OneNote XML → Markdown
# ---------------------------------------------------------------------------

def _extract_font_size(style: str) -> float:
    """Return font-size in pt from a OneNote OE style attribute, or 0."""
    m = re.search(r"font-size:([\d.]+)pt", style)
    return float(m.group(1)) if m else 0.0


def _is_consolas_oe(html: str) -> bool:
    """True if the entire OE text is a Consolas-font span (= code line).

    NOTE: OneNote moves font-family to the OE style attribute when storing;
    this function is kept for future use but detection is done via OE style.
    """
    stripped = html.strip()
    return bool(re.match(r"<span\s[^>]*font-family:Consolas", stripped))


def _is_blockquote_oe(html: str) -> bool:
    """True if the entire OE text is an italic+gray span (= blockquote).

    NOTE: OneNote moves color to the OE style attribute when storing;
    this function is kept for future use but detection is done via OE style.
    """
    stripped = html.strip()
    return bool(re.match(
        r"<span\s[^>]*font-style:italic[^>]*color:#555555", stripped
    ))


def _inline_html_to_md(html_frag: str) -> str:
    """Convert inline HTML with CSS style spans to Markdown.

    html2text does not understand `style='font-weight:bold'`.  We
    pre-convert style spans to semantic HTML that html2text handles:
      font-weight:bold   → <strong>
      font-style:italic  → <em>
      font-family:Consolas → <code>
    """
    frag = html_frag
    # Normalise multi-line tag attributes that OneNote emits
    for _ in range(3):
        frag = re.sub(r"(<[^>]+)\n", r"\1 ", frag)

    # Blockquote styling span — strip the decoration, keep content
    frag = re.sub(
        r"<span\s[^>]*font-style:italic[^>]*color:#555555[^>]*>(.*?)</span>",
        r"\1", frag, flags=re.DOTALL,
    )
    # Bold
    frag = re.sub(
        r"<span\s[^>]*font-weight:bold[^>]*>(.*?)</span>",
        r"<strong>\1</strong>", frag, flags=re.DOTALL,
    )
    # Italic
    frag = re.sub(
        r"<span\s[^>]*font-style:italic[^>]*>(.*?)</span>",
        r"<em>\1</em>", frag, flags=re.DOTALL,
    )
    # Consolas → inline code
    frag = re.sub(
        r"<span\s[^>]*font-family:Consolas[^>]*>(.*?)</span>",
        r"<code>\1</code>", frag, flags=re.DOTALL,
    )
    # Remaining spans — just unwrap
    frag = re.sub(r"<span[^>]*>(.*?)</span>", r"\1", frag, flags=re.DOTALL)

    md = _H2T.handle(frag).strip()
    # html2text leaves HTML entities unescaped in some edge cases
    md = _html_lib.unescape(md)
    return md


def _table_el_to_md(table_el) -> str:
    """Convert a <one:Table> ElementTree element to a Markdown table."""
    has_header = table_el.get("hasHeaderRow", "false") == "true"
    rows = []
    for row_el in table_el.findall(_tag("Row")):
        cells = []
        for cell_el in row_el.findall(_tag("Cell")):
            parts = []
            for t_el in cell_el.iter(_tag("T")):
                raw = (t_el.text or "").strip()
                if raw:
                    parts.append(_inline_html_to_md(raw))
            cells.append(" ".join(p for p in parts if p))
        rows.append(cells)

    if not rows:
        return ""

    n_cols = max(len(r) for r in rows)
    rows = [r + [""] * (n_cols - len(r)) for r in rows]

    md_rows = ["| " + " | ".join(rows[0]) + " |"]
    if has_header:
        md_rows.append("| " + " | ".join(["---"] * n_cols) + " |")
    for row in rows[1:]:
        md_rows.append("| " + " | ".join(row) + " |")
    return "\n".join(md_rows)


def _page_xml_to_markdown(xml_str: str) -> str:
    """Convert a OneNote page XML string to a Markdown string.

    Reads the page title and the text content of all Outline elements.
    Recursively traverses nested <one:OEChildren> to reconstruct indented
    lists at any depth.  Consecutive Consolas-font OEs become fenced code
    blocks.  Tables are reconstructed as GFM pipe tables.
    """
    root = _parse_xml(xml_str)
    # output_parts: top-level blocks separated by blank lines
    output_parts: list = []
    # list_accum: consecutive list/blockquote lines joined by single newlines
    list_accum: list = []
    code_lines: list = []

    def _flush_code() -> None:
        if code_lines:
            list_accum.append("```\n" + "\n".join(code_lines) + "\n```")
            code_lines.clear()

    def _flush_list() -> None:
        _flush_code()
        if list_accum:
            output_parts.append("\n".join(list_accum))
            list_accum.clear()

    def _add_block(text: str) -> None:
        _flush_list()
        output_parts.append(text)

    def _process_oec(oec_el, depth: int) -> None:
        indent = "  " * depth  # 2 spaces/level; _normalise_list_indent converts to 4 on write-back
        for oe in oec_el.findall(_tag("OE")):
            # Table
            table_el = oe.find(_tag("Table"))
            if table_el is not None:
                _flush_list()
                output_parts.append(_table_el_to_md(table_el))
                child_oec = oe.find(_tag("OEChildren"))
                if child_oec is not None:
                    _process_oec(child_oec, depth)
                continue
            # Image / Ink
            if oe.find(_tag("Image")) is not None:
                _add_block(f"{indent}[Image]")
                continue
            if oe.find(_tag("InkNode")) is not None:
                _add_block(f"{indent}[Ink]")
                continue
            # Gather text CDATA
            combined = "".join(t.text or "" for t in oe.findall(_tag("T")))
            oe_style = oe.get("style", "")
            child_oec = oe.find(_tag("OEChildren"))
            # Code line: font-family:Consolas on the OE style
            if "font-family:Consolas" in oe_style:
                _flush_list()  # don't mix list_accum with code
                inner = re.sub(r"<[^>]+>", "", combined)
                inner = _html_lib.unescape(inner)
                if inner in ("\u00a0", "\xa0"):
                    inner = ""
                code_lines.append(indent + inner)
                if child_oec is not None:
                    _process_oec(child_oec, depth + 1)
                continue
            # Non-code: flush any pending code block
            _flush_code()
            # Blockquote: color:#555555 on the OE style
            if "color:#555555" in oe_style:
                inner = re.sub(r"<[^>]+>", "", combined)
                inner = _html_lib.unescape(inner).strip()
                if inner:
                    list_accum.append(f"{indent}> {inner}")
                if child_oec is not None:
                    _process_oec(child_oec, depth + 1)
                continue
            # Empty OE — still recurse into children at same depth
            if not combined.strip():
                if child_oec is not None:
                    _process_oec(child_oec, depth)
                continue
            text = _inline_html_to_md(combined)
            if not text.strip():
                if child_oec is not None:
                    _process_oec(child_oec, depth)
                continue
            font_size = _extract_font_size(oe_style)
            list_el = oe.find(_tag("List"))
            # Headings — only at depth 0, no list marker
            if list_el is None and depth == 0:
                text_clean = re.sub(r"^\*\*(.*)\*\*$", r"\1", text.strip())
                if font_size >= 20:
                    _add_block(f"# {text_clean}")
                    if child_oec is not None:
                        _process_oec(child_oec, depth)
                    continue
                if font_size >= 17:
                    _add_block(f"## {text_clean}")
                    if child_oec is not None:
                        _process_oec(child_oec, depth)
                    continue
                if font_size >= 14:
                    _add_block(f"### {text_clean}")
                    if child_oec is not None:
                        _process_oec(child_oec, depth)
                    continue
                if font_size >= 12:
                    _add_block(f"#### {text_clean}")
                    if child_oec is not None:
                        _process_oec(child_oec, depth)
                    continue
            # List items
            if list_el is not None:
                if list_el.find(_tag("Bullet")) is not None:
                    list_accum.append(f"{indent}- {text}")
                    if child_oec is not None:
                        _process_oec(child_oec, depth + 1)
                    continue
                num_el = list_el.find(_tag("Number"))
                if num_el is not None:
                    num = num_el.get("text", "").rstrip(".")
                    list_accum.append(f"{indent}{num}. {text}")
                    if child_oec is not None:
                        _process_oec(child_oec, depth + 1)
                    continue
            # Plain paragraph
            _flush_list()
            output_parts.append(f"{indent}{text}")
            if child_oec is not None:
                _process_oec(child_oec, depth)

    # --- Title ---
    title_el = root.find(_tag("Title"))
    if title_el is not None:
        for t_el in title_el.iter(_tag("T")):
            raw = (t_el.text or "").strip()
            if raw:
                text = _inline_html_to_md(raw)
                if text:
                    output_parts.append(f"# {text}")
                    break
    if not output_parts:
        name = root.get("name", "")
        if name:
            output_parts.append(f"# {name}")

    # --- Content ---
    for outline in root.findall(_tag("Outline")):
        for oec in outline.findall(_tag("OEChildren")):
            _process_oec(oec, 0)
    _flush_list()

    return "\n\n".join(output_parts)


def _html_table_to_onenote(block: str) -> str:
    """Convert an HTML <table> block to a native <one:OE><one:Table> element.

    Parses <thead>/<tbody>/<tr>/<th>/<td> with regex and builds the OneNote
    XML table schema.  The first row is marked as a header if <th> elements
    or a <thead> section is present.
    """
    # Detect header: <thead> section or any <th> element
    has_header = bool(re.search(r"<thead[\s>]|<th[\s>]", block, re.IGNORECASE))

    # Collect all rows as lists of cell-content strings
    rows = []
    for row_match in re.finditer(r"<tr[^>]*>(.*?)</tr>", block, re.DOTALL | re.IGNORECASE):
        cells = re.findall(r"<(?:td|th)[^>]*>(.*?)</(?:td|th)>",
                           row_match.group(1), re.DOTALL | re.IGNORECASE)
        # Strip inner tags from cell content to get plain+inline HTML
        cells = [c.strip() for c in cells]
        if cells:
            rows.append(cells)

    if not rows:
        return ""

    header_attr = ' hasHeaderRow="true"' if has_header else ""
    row_parts = []
    for row in rows:
        cell_parts = "".join(
            f"<one:Cell><one:OEChildren><one:OE>"
            f"<one:T><![CDATA[{cell.replace(']]>', ']]]]><![CDATA[>')}]]></one:T>"
            f"</one:OE></one:OEChildren></one:Cell>"
            for cell in row
        )
        row_parts.append(f"<one:Row>{cell_parts}</one:Row>")

    return (
        f"<one:OE>"
        f"<one:Table{header_attr}>"
        + "".join(row_parts)
        + "</one:Table>"
        "</one:OE>"
    )


def _markdown_to_oes(md_text: str) -> list:
    """Convert a Markdown string to a list of OneNote OE XML element strings.

    Each top-level HTML block produced by the Markdown parser becomes one
    or more <one:OE> elements.  CDATA-unsafe sequences are escaped.

    Handles the following conversions required by OneNote's XML schema:
      - <ul>/<li>     → <one:OE><one:List><one:Bullet .../> per item
      - <ol>/<li>     → <one:OE><one:List><one:Number .../> per item
      - <table>       → <one:OE><one:Table> native schema
      - <blockquote>  → inner content extracted; outer tags stripped
      - other blocks  → passed as CDATA (inline HTML allowed)

    Strategy: pre-stash ALL container types that can contain <p> (ul, ol,
    blockquote, table) before splitting.  This prevents the split regex from
    firing on nested <p> tags inside those containers, which would produce
    broken XML fragments.  After splitting on the remaining top-level tags,
    the stashed blocks are expanded in-place.
    """
    html = _markdown.markdown(md_text.strip(), extensions=_MD_EXTENSIONS)

    # Normalise 2-space-indented list items to 4-space so that the markdown
    # library correctly nests them.  Agents often emit 2-space indentation;
    # Python's markdown library requires 4 spaces per indent level.
    # We only touch lines that start with spaces followed by a list marker.
    def _normalise_list_indent(md: str) -> str:
        lines = md.split("\n")
        out = []
        for line in lines:
            stripped = line.lstrip(" ")
            n_spaces = len(line) - len(stripped)
            if n_spaces > 0 and stripped and stripped[0] in ("-", "*", "+") or (
                n_spaces > 0 and stripped and stripped[0].isdigit() and ". " in stripped[:5]
            ):
                # Convert each 2-space indent unit to 4 spaces
                units = n_spaces // 2
                remainder = n_spaces % 2
                line = " " * (units * 4 + remainder) + stripped
            out.append(line)
        return "\n".join(out)

    md_normalised = _normalise_list_indent(md_text.strip())
    html = _markdown.markdown(md_normalised, extensions=_MD_EXTENSIONS)

    # --- Phase 1: stash all container elements ---
    stash: dict = {}

    def _stash(tag: str) -> "Callable[[re.Match], str]":
        def _replace(m: re.Match) -> str:
            token = f"\x00{tag.upper()}{len(stash)}\x00"
            stash[token] = (tag, m.group(0))
            return token
        return _replace

    # Tags that cannot nest — simple non-greedy match is fine.
    for tag in ("blockquote", "pre", "table"):
        html = re.sub(
            rf"<{tag}[^>]*>.*?</{tag}>", _stash(tag), html, flags=re.DOTALL
        )

    # Tags that CAN nest (ul, ol): use balanced open/close counting so that
    # the outer <ul> is stashed as a whole, not just the innermost one.
    def _stash_balanced(html: str, tag: str) -> str:
        open_re = re.compile(rf"<{tag}(?:\s[^>]*)?>", re.IGNORECASE)
        close_re = re.compile(rf"</{tag}>", re.IGNORECASE)
        # Collect spans of all top-level blocks (right-to-left to preserve offsets).
        spans = []
        pos = 0
        while pos < len(html):
            m_open = open_re.search(html, pos)
            if m_open is None:
                break
            depth = 1
            scan = m_open.end()
            while scan < len(html) and depth > 0:
                m_o = open_re.search(html, scan)
                m_c = close_re.search(html, scan)
                if m_c is None:
                    break
                if m_o is not None and m_o.start() < m_c.start():
                    depth += 1
                    scan = m_o.end()
                else:
                    depth -= 1
                    scan = m_c.end()
                    if depth == 0:
                        spans.append((m_open.start(), scan))
                        pos = scan
                        break
            else:
                break
        for start, end in reversed(spans):
            block = html[start:end]
            token = f"\x00{tag.upper()}{len(stash)}\x00"
            stash[token] = (tag, block)
            html = html[:start] + token + html[end:]
        return html

    for tag in ("ul", "ol"):
        html = _stash_balanced(html, tag)

    # --- Phase 2: split remaining top-level block tags ---
    # Only <p>, <h1-6>, <hr> remain after stashing; all are top-level now.
    blocks = re.split(
        r"(?=<(?:p|h[1-6]|hr)[\s>/])", html
    )

    oes = []

    def _process_token(token: str) -> None:
        """Expand a stash token into one or more OE strings."""
        tag, raw = stash[token]

        if tag in ("ul", "ol"):
            # Parse the HTML with ElementTree to support nested lists.
            # Nested <ul>/<ol> inside <li> become <one:OEChildren> blocks.
            def _oes_from_list_el(list_el, is_bullet: bool) -> list:
                result = []
                for li_el in list_el:
                    if li_el.tag != "li":
                        continue
                    # Collect inline text content, stop at any nested list.
                    text_parts = [li_el.text or ""]
                    nested_list_el = None
                    for child in li_el:
                        if child.tag in ("ul", "ol"):
                            nested_list_el = child
                            break
                        # Serialize inline element back to HTML
                        text_parts.append(ET.tostring(child, encoding="unicode"))
                        if child.tail:
                            text_parts.append(child.tail)
                    item_text = "".join(text_parts)
                    item_text = re.sub(r"</?p[^>]*>", "", item_text).strip()
                    item_text = re.sub(r"\n+", " ", item_text).strip()
                    cdata = item_text.replace("]]>", "]]]]><![CDATA[>")
                    if nested_list_el is not None:
                        sub_is_bullet = nested_list_el.tag == "ul"
                        sub_oes = _oes_from_list_el(nested_list_el, sub_is_bullet)
                        children_xml = "".join(sub_oes)
                        children_block = f"<one:OEChildren>{children_xml}</one:OEChildren>"
                    else:
                        children_block = ""
                    if is_bullet:
                        result.append(
                            "<one:OE>"
                            "<one:List><one:Bullet bullet=\"2\" fontSize=\"11.0\"/>"
                            "</one:List>"
                            f"<one:T><![CDATA[{cdata}]]></one:T>"
                            f"{children_block}"
                            "</one:OE>"
                        )
                    else:
                        result.append(
                            "<one:OE>"
                            "<one:List>"
                            "<one:Number numberSequence=\"0\" numberFormat=\"##.\""
                            " fontSize=\"11.0\"/>"
                            "</one:List>"
                            f"<one:T><![CDATA[{cdata}]]></one:T>"
                            f"{children_block}"
                            "</one:OE>"
                        )
                return result
            try:
                list_root = ET.fromstring(f"<root>{raw}</root>")
                list_el = list_root.find(tag)
                if list_el is not None:
                    oes.extend(_oes_from_list_el(list_el, tag == "ul"))
                    return
            except ET.ParseError:
                pass  # fall through to simple flat extraction below
            # Fallback: flat extraction (no nesting)
            bullet = tag == "ul"
            for item in re.findall(r"<li[^>]*>(.*?)</li>", raw, re.DOTALL):
                item = re.sub(r"</?p[^>]*>", "", item).strip()
                item = re.sub(r"\n+", " ", item).strip()
                cdata = item.replace("]]>", "]]]]><![CDATA[>")
                if bullet:
                    oes.append(
                        "<one:OE>"
                        "<one:List><one:Bullet bullet=\"2\" fontSize=\"11.0\"/>"
                        "</one:List>"
                        f"<one:T><![CDATA[{cdata}]]></one:T>"
                        "</one:OE>"
                    )
                else:
                    oes.append(
                        "<one:OE>"
                        "<one:List>"
                        "<one:Number numberSequence=\"0\" numberFormat=\"##.\""
                        " fontSize=\"11.0\"/>"
                        "</one:List>"
                        f"<one:T><![CDATA[{cdata}]]></one:T>"
                        "</one:OE>"
                    )

        elif tag == "pre":
            # Extract code text, preserving line breaks as separate OEs.
            # Strip the <pre> and optional <code> wrapper tags, then
            # HTML-unescape entities (markdown encodes < > & inside code).
            inner = re.sub(r"<pre[^>]*>", "", raw)
            inner = re.sub(r"</pre>", "", inner)
            inner = re.sub(r"<code[^>]*>", "", inner)
            inner = re.sub(r"</code>", "", inner)
            inner = _html_lib.unescape(inner).rstrip("\n")
            lines = inner.split("\n")
            for line in lines:
                display = line if line else "\u00a0"  # nbsp keeps blank lines
                cdata = display.replace("]]>", "]]]]><![CDATA[>")
                oes.append(
                    "<one:OE>"
                    "<one:T>"
                    "<![CDATA["
                    "<span style='font-family:Consolas;font-size:10.0pt'>"
                    f"{cdata}"
                    "</span>"
                    "]]>"
                    "</one:T>"
                    "</one:OE>"
                )

        elif tag == "blockquote":
            inner = re.sub(r"</?blockquote[^>]*>", "", raw)
            inner = re.sub(r"</?p[^>]*>", "", inner)
            inner = re.sub(r"\n+", " ", inner).strip()
            cdata = inner.replace("]]>", "]]]]><![CDATA[>")
            oes.append(
                "<one:OE>"
                "<one:T>"
                "<![CDATA["
                "<span style='font-style:italic;color:#555555'>"
                f"{cdata}"
                "</span>"
                "]]>"
                "</one:T>"
                "</one:OE>"
            )

        elif tag == "table":
            table_oe = _html_table_to_onenote(raw)
            if table_oe:
                oes.append(table_oe)

    # --- Phase 3: walk split blocks and handle tokens + plain blocks ---
    for block in blocks:
        # A block may be a token, a plain HTML block, or mix of tokens + text.
        # Tokens are always produced by phase 1 and look like \x00TAG0\x00.
        if not block.strip():
            continue

        # Check if the entire (stripped) block is a single token.
        stripped = block.strip()
        if re.fullmatch(r"\x00[A-Z]+\d+\x00", stripped):
            _process_token(stripped)
            continue

        # Mixed: may contain tokens interleaved with plain text fragments.
        # Process each part left-to-right.
        parts = re.split(r"(\x00[A-Z]+\d+\x00)", block)
        for part in parts:
            part = part.strip()
            if not part:
                continue
            if re.fullmatch(r"\x00[A-Z]+\d+\x00", part):
                _process_token(part)
            else:
                # Plain block: collapse newlines, pass as CDATA.
                flat = re.sub(r"\n+", " ", part).strip()
                if flat:
                    cdata = flat.replace("]]>", "]]]]><![CDATA[>")
                    oes.append(
                        f"<one:OE><one:T>"
                        f"<![CDATA[{cdata}]]>"
                        f"</one:T></one:OE>"
                    )

    # Fallback: entire HTML as one block.
    if not oes and html.strip():
        cdata = re.sub(r"\n+", " ", html).strip()
        cdata = cdata.replace("]]>", "]]]]><![CDATA[>")
        oes.append(
            f"<one:OE><one:T><![CDATA[{cdata}]]></one:T></one:OE>"
        )

    return oes


def _build_page_xml(page_id: str, title: str, oes: list) -> str:
    """Build a complete OneNote page XML string for use with UpdatePageContent.

    Includes a Title element and a single Outline block with the supplied OEs.
    """
    title_cdata = title.replace("]]>", "]]]]><![CDATA[>")
    indented_oes = "\n        ".join(oes)
    return (
        f'<?xml version="1.0"?>\n'
        f'<one:Page xmlns:one="{_NS}" ID="{page_id}">\n'
        f"  <one:Title>\n"
        f"    <one:OE>\n"
        f"      <one:T><![CDATA[{title_cdata}]]></one:T>\n"
        f"    </one:OE>\n"
        f"  </one:Title>\n"
        f"  <one:Outline>\n"
        f"    <one:OEChildren>\n"
        f"      {indented_oes}\n"
        f"    </one:OEChildren>\n"
        f"  </one:Outline>\n"
        f"</one:Page>"
    )


def _build_append_xml(page_id: str, oes: list) -> str:
    """Build a minimal page XML that adds a new Outline block.

    When passed to UpdatePageContent, OneNote inserts this as a new content
    block alongside any existing content on the page (because no Outline ID
    is set, OneNote treats it as a new element to be inserted).
    """
    indented_oes = "\n        ".join(oes)
    return (
        f'<?xml version="1.0"?>\n'
        f'<one:Page xmlns:one="{_NS}" ID="{page_id}">\n'
        f"  <one:Outline>\n"
        f"    <one:OEChildren>\n"
        f"      {indented_oes}\n"
        f"    </one:OEChildren>\n"
        f"  </one:Outline>\n"
        f"</one:Page>"
    )


def _build_replace_xml(existing_xml: str, oes: list) -> str:
    """Build a page XML that replaces the Outline content of an existing page.

    Fetches the Outline objectIDs from existing_xml so that UpdatePageContent
    performs an in-place replacement (matched by ID) rather than appending.
    The first Outline's OEChildren are replaced with *oes*; any additional
    Outlines are emptied to consolidate all content into a single block.
    """
    root = _parse_xml(existing_xml)
    page_id = root.get("ID", "")
    outlines = root.findall(_tag("Outline"))

    if not outlines:
        # Page has no Outlines yet — fall back to append behaviour
        return _build_append_xml(page_id, oes)

    indented_oes = "\n        ".join(oes)

    outline_parts = []
    first = True
    for outline in outlines:
        oid = outline.get("objectID", "")
        id_attr = f' objectID="{oid}"' if oid else ""
        if first:
            outline_parts.append(
                f"  <one:Outline{id_attr}>\n"
                f"    <one:OEChildren>\n"
                f"      {indented_oes}\n"
                f"    </one:OEChildren>\n"
                f"  </one:Outline>"
            )
            first = False
        else:
            # Clear additional outlines so the page has one unified content block
            outline_parts.append(
                f"  <one:Outline{id_attr}>\n"
                f"    <one:OEChildren>\n"
                f"      <one:OE><one:T><![CDATA[]]></one:T></one:OE>\n"
                f"    </one:OEChildren>\n"
                f"  </one:Outline>"
            )

    outlines_xml = "\n".join(outline_parts)
    return (
        f'<?xml version="1.0"?>\n'
        f'<one:Page xmlns:one="{_NS}" ID="{page_id}">\n'
        f"{outlines_xml}\n"
        f"</one:Page>"
    )


# ---------------------------------------------------------------------------
# Tools
# ---------------------------------------------------------------------------

@mcp.tool()
def list_notebooks() -> str:
    """List all OneNote notebooks currently open in the local OneNote application.

    Returns a JSON array of notebook objects, each containing:
      - id: the unique OneNote object ID (use this in other tool calls)
      - name: display name
      - path: local or UNC path to the notebook folder
      - last_modified_time: ISO 8601 timestamp
      - color: notebook color code (may be empty)
      - is_unread: whether the notebook has unread content
    """
    onenote = _get_onenote()
    xml = onenote.GetHierarchy("", _HS_CHILDREN)
    root = _parse_xml(xml)

    notebooks = []
    for nb in root.findall(_tag("Notebook")):
        notebooks.append({
            "id": nb.get("ID", ""),
            "name": nb.get("name", ""),
            "path": nb.get("path", ""),
            "last_modified_time": nb.get("lastModifiedTime", ""),
            "color": nb.get("color", ""),
            "is_unread": nb.get("isUnread", "false") == "true",
        })

    if not notebooks:
        return "No open notebooks found in OneNote."
    return json.dumps(notebooks, indent=2)


@mcp.tool()
def list_sections(notebook_id: str) -> str:
    """List all sections and section groups in a OneNote notebook.

    Args:
        notebook_id: The ID of the notebook, obtained from list_notebooks.

    Returns a JSON array of section objects, each containing:
      - id: unique OneNote object ID (use this in other tool calls)
      - name: display name
      - type: "Section" or "SectionGroup"
      - parent_id: ID of the parent notebook or section group
      - last_modified_time: ISO 8601 timestamp
      - color: section color code (may be empty)
      - is_unread: whether the section has unread content
      - is_locked: whether the section is password-protected

    The list is ordered depth-first so that each section group immediately
    precedes its children.
    """
    onenote = _get_onenote()
    xml = onenote.GetHierarchy(notebook_id, _HS_SECTIONS)
    root = _parse_xml(xml)
    sections = _sections_recursive(root, notebook_id)

    if not sections:
        return "No sections found in this notebook."
    return json.dumps(sections, indent=2)


@mcp.tool()
def list_pages(section_id: str) -> str:
    """List all pages in a OneNote section.

    Args:
        section_id: The ID of the section, obtained from list_sections.

    Returns a JSON array of page objects, each containing:
      - id: unique OneNote object ID (use this in other tool calls)
      - name: page title
      - date_time: page creation time (ISO 8601)
      - last_modified_time: ISO 8601 timestamp
      - page_level: indentation level (1 = top-level, 2–3 = sub-page)
      - is_unread: whether the page has been marked as unread
    """
    onenote = _get_onenote()
    xml = onenote.GetHierarchy(section_id, _HS_PAGES)
    root = _parse_xml(xml)

    pages = []
    for page in root.iter(_tag("Page")):
        pages.append({
            "id": page.get("ID", ""),
            "name": page.get("name", ""),
            "date_time": page.get("dateTime", ""),
            "last_modified_time": page.get("lastModifiedTime", ""),
            "page_level": int(page.get("pageLevel", "1")),
            "is_unread": page.get("isUnread", "false") == "true",
        })

    if not pages:
        return "No pages found in this section."
    return json.dumps(pages, indent=2)


@mcp.tool()
def get_page_content(page_id: str) -> str:
    """Get the full content of a OneNote page as Markdown.

    Retrieves basic text content (images and ink are noted as placeholders).
    HTML formatting inside text runs is converted to Markdown.

    Args:
        page_id: The ID of the page, obtained from list_pages.

    Returns a string with an HTML comment metadata header followed by the
    page content in Markdown format.
    """
    onenote = _get_onenote()
    xml = onenote.GetPageContent(page_id, _PI_BASIC)
    root = _parse_xml(xml)

    metadata = {
        "page_id": root.get("ID", page_id),
        "name": root.get("name", ""),
        "date_time": root.get("dateTime", ""),
        "last_modified_time": root.get("lastModifiedTime", ""),
        "page_level": root.get("pageLevel", "1"),
    }
    header = "\n".join(f"<!-- {k}: {v} -->" for k, v in metadata.items())
    body = _page_xml_to_markdown(xml)
    return f"{header}\n\n{body}"


@mcp.tool()
def search_notes(
    query: str,
    notebook_id: str = "",
    max_results: int = 50,
) -> str:
    """Search for pages matching a query string across OneNote notebooks.

    Uses the OneNote full-text search index; the same index that powers the
    in-application search bar.

    Args:
        query: The text to search for.
        notebook_id: Optional. When supplied, restricts the search to the
            specified notebook. Leave blank (default) to search all notebooks.
        max_results: Maximum number of results to return (1–200, default 50).

    Returns a JSON array of matching page summaries, each containing:
      - page_id, name, date_time, last_modified_time
      - section_id, section_name
      - notebook_id, notebook_name
    """
    if not query.strip():
        raise ValueError("query must not be empty")
    max_results = max(1, min(max_results, 200))

    onenote = _get_onenote()
    # FindPages([in] startNodeID, [in] searchString, [out] xmlOut,
    #           [in, opt] includeUnindexed, [in, opt] display,
    #           [in, opt] hsScope)
    # In win32com dispatch the [out] parameter is returned; omit it here.
    xml = onenote.FindPages(notebook_id, query)
    root = _parse_xml(xml)

    results = _extract_search_results(root, max_results)
    if not results:
        return json.dumps({"message": f"No pages found matching '{query}'.", "results": []}, indent=2)
    return json.dumps(results, indent=2)


@mcp.tool()
def create_section(parent_id: str, section_name: str) -> str:
    """Create a new section in a OneNote notebook or section group.

    Args:
        parent_id: The ID of the target notebook or section group, obtained
            from list_notebooks or list_sections.
        section_name: Display name for the new section. Must not contain
            any of the characters: \\ / : * ? \" < > |

    Returns a JSON object with the new section's id and name.
    """
    bad = _INVALID_SECTION_CHARS & set(section_name)
    if bad:
        raise ValueError(
            f"section_name contains invalid characters: {sorted(bad)}"
        )
    if not section_name.strip():
        raise ValueError("section_name must not be blank")

    onenote = _get_onenote()
    # OpenHierarchy(bstrPath, bstrRelativeToObjectID, cftIfNotExist)
    # bstrPath must include the .one extension so OneNote creates the file.
    # The return value is the new section's object ID.
    path = section_name if section_name.lower().endswith(".one") else section_name + ".one"
    new_section_id = onenote.OpenHierarchy(path, parent_id, _CFT_SECTION)
    return json.dumps({"id": new_section_id, "name": section_name}, indent=2)


@mcp.tool()
def create_page(
    section_id: str,
    title: str,
    content: str = "",
) -> str:
    """Create a new page in a OneNote section.

    First creates a blank titled page via the COM API, then populates it
    with the supplied title and Markdown content.

    Args:
        section_id: The ID of the section, obtained from list_sections.
        title: The page title.
        content: Optional body content in Markdown format. Leave blank to
            create a page with a title only.

    Returns a JSON object with the new page's id and title.
    """
    if not title.strip():
        raise ValueError("title must not be blank")

    onenote = _get_onenote()
    # CreateNewPage([in] sectionID, [out] pageID, [in, opt] newPageStyle)
    page_id = onenote.CreateNewPage(section_id, _NPS_BLANK_WITH_TITLE)

    oes = _markdown_to_oes(content) if content.strip() else [
        "<one:OE><one:T><![CDATA[]]></one:T></one:OE>"
    ]
    page_xml = _build_page_xml(page_id, title, oes)
    # Pass datetime(1, 1, 1) as the expected-last-modified date to skip
    # conflict detection (equivalent to the COM automation "zero date").
    # Pass 0.0 as dateExpectedLastModified (COM DATE float); this skips
    # conflict detection, which is the correct behaviour for an agent tool.
    onenote.UpdatePageContent(page_xml, 0.0)

    return json.dumps({"id": page_id, "title": title}, indent=2)


@mcp.tool()
def update_page(
    page_id: str,
    content: str,
    mode: str = "append",
) -> str:
    """Update the content of an existing OneNote page.

    Args:
        page_id: The ID of the page to update, obtained from list_pages.
        content: The content to write to the page, in Markdown format.
        mode: How to handle existing content:
            - "append" (default): adds new content below existing content
              by inserting a new Outline block.
            - "replace": replaces all existing Outline content with the
              new content, consolidating the page into a single block.

    Returns a JSON object confirming the page_id and applied mode.
    """
    if mode not in ("append", "replace"):
        raise ValueError(f"mode must be 'append' or 'replace', got '{mode!r}'")
    if not content.strip():
        raise ValueError("content must not be empty")

    oes = _markdown_to_oes(content)
    if not oes:
        raise ValueError("content_markdown produced no renderable content")

    onenote = _get_onenote()

    if mode == "append":
        page_xml = _build_append_xml(page_id, oes)
    else:
        existing_xml = onenote.GetPageContent(page_id, _PI_BASIC)
        page_xml = _build_replace_xml(existing_xml, oes)

    onenote.UpdatePageContent(page_xml, 0.0)
    return json.dumps({"page_id": page_id, "mode": mode, "status": "ok"}, indent=2)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run()
