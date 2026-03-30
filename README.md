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
