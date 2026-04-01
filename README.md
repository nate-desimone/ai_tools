# AI Tools

A collection of miscellaneous AI helper tools.

## Tools

### GuidMcpServer

An [MCP](https://modelcontextprotocol.io/) server that provides GUID/UUID generation and conversion tools for AI agents. Built with the [FastMCP](https://github.com/modelcontextprotocol/python-sdk) Python SDK.

#### Features

- **generate_guid** - Generate one or more random UUIDs (v4)
- **generate_guid_v1** - Generate a time-based UUID (v1)
- **generate_guid_v5** - Generate a deterministic UUID (v5, SHA-1 based) from a namespace and name
- **convert_guid_format** - Convert an existing GUID between formats

Supported output formats:

| Format | Example |
|---|---|
| `standard` | `a1b2c3d4-e5f6-7890-abcd-ef1234567890` |
| `uppercase` | `A1B2C3D4-E5F6-7890-ABCD-EF1234567890` |
| `no-hyphens` | `a1b2c3d4e5f67890abcdef1234567890` |
| `braces` | `{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}` |
| `uefi-struct` | `{ 0xA1B2C3D4, 0xE5F6, 0x7890, { 0xAB, 0xCD, ... }}` |

#### Setup

```bash
cd guid_mcp_server
pip install -r requirements.txt
```

#### Usage

Run as a stdio MCP server:

```bash
python mcp_guid_server.py
```

##### VS Code (workspace level)

Add the following to your workspace `settings.json`:

```json
"mcp": {
    "servers": {
        "guid": {
            "command": "python",
            "args": ["${workspaceFolder}/guid_mcp_server/mcp_guid_server.py"]
        }
    }
}
```

##### VS Code (system level)

To make the MCP server available across all workspaces, copy `mcp_guid_server.py` to a permanent location (e.g. `C:\Users\<user name>\AppData\Roaming\Code\User\mcp\guid\`) and add the following to your global `mcp.json`, located at `C:\Users\<user name>\AppData\Roaming\Code\User\mcp.json`:

```json
{
    "servers": {
        "guid": {
            "command": "python3",
            "args": ["C:\\Users\\<user name>\\AppData\\Roaming\\Code\\User\\mcp\\guid\\mcp_guid_server.py"]
        }
    }
}
```

---

### EmailMcpServer

An [MCP](https://modelcontextprotocol.io/) server that gives AI agents read, search, and drafting access to Microsoft Outlook email via the MAPI COM interface. Built with the [FastMCP](https://github.com/modelcontextprotocol/python-sdk) Python SDK and `pywin32`. Requires Outlook to be installed locally.

> **Windows only.** Depends on the Win32 COM automation layer provided by Outlook.

#### Features

**Reading & browsing**
- **list_email_accounts** - List all accounts in the local Outlook MAPI profile
- **list_folders** - List folders (optionally recursive) for an account with item counts
- **get_folder_counts** - Get total and unread item counts for a folder (lightweight, no email iteration)
- **list_emails** - List emails in a folder, newest first; supports `since`/`until` date range and `unread_only` filter
- **search_emails** - Search emails by subject, sender, and/or body text within a folder
- **get_email** - Retrieve the full content of an email by entry ID (body returned as Markdown)
- **get_thread** - Retrieve all messages in a conversation thread, sorted oldest-first

**Attachments**
- **list_attachments** - List all attachments on an email
- **read_attachment_text** - Read the text content of a plain-text attachment directly (no permanent file saved); supports .txt, .csv, .py, .md, .json, .yaml, .xml, .html, .sh, .ps1, and more
- **download_attachment** - Download a specific attachment to a local directory

**Drafting** *(all tools save a draft and open it in Outlook for review — nothing is sent automatically)*
- **create_draft** - Compose a new email draft
- **create_reply** - Create a reply to an existing email (Reply or Reply All)
- **create_forward** - Create a forward draft for an existing email

#### Setup

```bash
cd email_mcp_server
pip install -r requirements.txt
```

#### Usage

Run as a stdio MCP server:

```bash
python mcp_email_server.py
```

##### VS Code (workspace level)

Add the following to your workspace `settings.json`:

```json
"mcp": {
    "servers": {
        "email": {
            "command": "python",
            "args": ["${workspaceFolder}/email_mcp_server/mcp_email_server.py"]
        }
    }
}
```

##### VS Code (system level)

To make the MCP server available across all workspaces, copy `mcp_email_server.py` to a permanent location (e.g. `C:\Users\<user name>\AppData\Roaming\Code\User\mcp\email\`) and add the following to your global `mcp.json`, located at `C:\Users\<user name>\AppData\Roaming\Code\User\mcp.json`:

```json
{
    "servers": {
        "email": {
            "command": "python3",
            "args": ["C:\\Users\\<user name>\\AppData\\Roaming\\Code\\User\\mcp\\email\\mcp_email_server.py"]
        }
    }
}
```

---

### OneNoteMcpServer

An [MCP](https://modelcontextprotocol.io/) server that gives AI agents read, search, and write access to Microsoft OneNote notebooks via the OneNote COM automation interface. Built with the [FastMCP](https://github.com/modelcontextprotocol/python-sdk) Python SDK and `comtypes`. Requires OneNote 2016 or 2019 (classic desktop) to be installed locally.

> **Windows only.** Depends on the OneNote 2016/2019 COM type library (`IApplication` vtable interface).

#### Features

**Browsing**
- **list_notebooks** - List all OneNote notebooks currently open in the local OneNote application
- **list_sections** - List all sections and section groups in a notebook (recursive, depth-first)
- **list_pages** - List all pages in a section

**Reading**
- **get_page_content** - Retrieve the full content of a page as clean Markdown, suitable for direct consumption by AI agents. Handles headings, nested bullet and numbered lists (up to any depth), tables, code blocks, blockquotes, bold/italic inline formatting, and hyperlinks.
- **search_notes** - Full-text search across all notebooks (or a specific notebook) using the OneNote search index

**Writing**
- **create_section** - Create a new section inside a notebook or section group
- **create_page** - Create a new page in a section with a title and Markdown body content
- **update_page** - Update an existing page by appending new Markdown content or replacing all existing content

#### Setup

```bash
cd onenote_mcp_server
pip install -r requirements.txt
```

#### Usage

Run as a stdio MCP server:

```bash
python mcp_onenote_server.py
```

##### VS Code (workspace level)

Add the following to your workspace `settings.json`:

```json
"mcp": {
    "servers": {
        "onenote": {
            "command": "python",
            "args": ["${workspaceFolder}/onenote_mcp_server/mcp_onenote_server.py"]
        }
    }
}
```

##### VS Code (system level)

To make the MCP server available across all workspaces, copy `mcp_onenote_server.py` to a permanent location (e.g. `C:\Users\<user name>\AppData\Roaming\Code\User\mcp\onenote\`) and add the following to your global `mcp.json`, located at `C:\Users\<user name>\AppData\Roaming\Code\User\mcp.json`:

```json
{
    "servers": {
        "onenote": {
            "command": "python3",
            "args": ["C:\\Users\\<user name>\\AppData\\Roaming\\Code\\User\\mcp\\onenote\\mcp_onenote_server.py"]
        }
    }
}
```
