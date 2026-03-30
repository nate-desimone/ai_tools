## @file
# MCP Server for GUID/UUID generation.
#
# Copyright 2026 Intel Corporation All Rights Reserved.
# SPDX-License-Identifier: BSD-2-Clause-Patent
##
"""
MCP Server for GUID/UUID generation.

Provides tools for AI agents to generate and convert GUIDs in multiple formats,
including UEFI EFI_GUID C struct initializer syntax.

Usage:
    python mcp_guid_server.py
"""

import uuid
from typing import Literal

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("GUID Generator")

FormatType = Literal["standard", "uppercase", "no-hyphens", "braces", "uefi-struct"]

NAMESPACE_MAP = {
    "dns": uuid.NAMESPACE_DNS,
    "url": uuid.NAMESPACE_URL,
    "oid": uuid.NAMESPACE_OID,
    "x500": uuid.NAMESPACE_X500,
}


def format_guid(guid: uuid.UUID, fmt: FormatType = "standard") -> str:
    """Format a UUID object into the requested string representation."""
    if fmt == "standard":
        return str(guid)
    elif fmt == "uppercase":
        return str(guid).upper()
    elif fmt == "no-hyphens":
        return guid.hex
    elif fmt == "braces":
        return "{" + str(guid).upper() + "}"
    elif fmt == "uefi-struct":
        # EFI_GUID format: { 0xXXXXXXXX, 0xXXXX, 0xXXXX, { 0xXX, 0xXX, ... }}
        # Fields: Data1 (32-bit), Data2 (16-bit), Data3 (16-bit), Data4 (8 bytes)
        b = guid.bytes
        # uuid.bytes is big-endian for the first three fields
        data1 = int.from_bytes(b[0:4], "big")
        data2 = int.from_bytes(b[4:6], "big")
        data3 = int.from_bytes(b[6:8], "big")
        data4 = b[8:16]
        data4_str = ", ".join(f"0x{byte:02X}" for byte in data4)
        return f"{{ 0x{data1:08X}, 0x{data2:04X}, 0x{data3:04X}, {{ {data4_str} }}}}"
    else:
        return str(guid)


@mcp.tool()
def generate_guid(
    count: int = 1,
    format: FormatType = "standard",
) -> str:
    """Generate one or more random GUIDs (UUID v4).

    Args:
        count: Number of GUIDs to generate (1-100).
        format: Output format. One of:
            - "standard": lowercase with hyphens (e.g. a1b2c3d4-e5f6-7890-abcd-ef1234567890)
            - "uppercase": uppercase with hyphens
            - "no-hyphens": 32 hex characters, no separators
            - "braces": uppercase with hyphens, wrapped in curly braces (Windows registry style)
            - "uefi-struct": C struct initializer for EFI_GUID (e.g. { 0xA1B2C3D4, 0xE5F6, 0x7890, { 0xAB, 0xCD, ... }})
    """
    count = max(1, min(count, 100))
    guids = [format_guid(uuid.uuid4(), format) for _ in range(count)]
    return "\n".join(guids)


@mcp.tool()
def generate_guid_v1(
    format: FormatType = "standard",
) -> str:
    """Generate a time-based GUID (UUID v1).

    The generated UUID encodes the current timestamp and node (MAC address or random).

    Args:
        format: Output format (see generate_guid for format descriptions).
    """
    return format_guid(uuid.uuid1(), format)


@mcp.tool()
def generate_guid_v5(
    namespace: str,
    name: str,
    format: FormatType = "standard",
) -> str:
    """Generate a deterministic GUID (UUID v5, SHA-1 based).

    Given the same namespace and name, this always produces the same GUID.

    Args:
        namespace: One of "dns", "url", "oid", "x500", or an arbitrary UUID string to use as the namespace.
        name: The name to hash within the namespace.
        format: Output format (see generate_guid for format descriptions).
    """
    ns_key = namespace.lower()
    if ns_key in NAMESPACE_MAP:
        ns_uuid = NAMESPACE_MAP[ns_key]
    else:
        ns_uuid = uuid.UUID(namespace)
    return format_guid(uuid.uuid5(ns_uuid, name), format)


@mcp.tool()
def convert_guid_format(
    guid: str,
    target_format: FormatType = "standard",
) -> str:
    """Convert an existing GUID string to a different format.

    Accepts any standard UUID string representation (with or without hyphens/braces)
    or a UEFI struct format string.

    Args:
        guid: The GUID string to convert.
        target_format: Desired output format (see generate_guid for format descriptions).
    """
    cleaned = guid.strip().strip("{}")
    # Handle UEFI struct format input: extract hex values
    if "0x" in cleaned:
        hex_parts = []
        for part in cleaned.replace("{", "").replace("}", "").split(","):
            part = part.strip()
            if part.startswith("0x") or part.startswith("0X"):
                hex_parts.append(part[2:])
        if len(hex_parts) == 11:
            # Data1 (4 bytes) + Data2 (2 bytes) + Data3 (2 bytes) + Data4 (8 bytes)
            hex_str = (
                hex_parts[0].zfill(8)
                + hex_parts[1].zfill(4)
                + hex_parts[2].zfill(4)
                + "".join(p.zfill(2) for p in hex_parts[3:])
            )
            parsed = uuid.UUID(hex_str)
        else:
            parsed = uuid.UUID(cleaned)
    else:
        parsed = uuid.UUID(cleaned)
    return format_guid(parsed, target_format)


if __name__ == "__main__":
    mcp.run(transport="stdio")
