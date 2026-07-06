# outlook-mcp-rs

A single-binary Rust MCP server controlling Microsoft Outlook desktop via the
Win32 COM API. This is the Rust counterpart to
[outlook-mcp](https://github.com/adamkopelman/outlook-mcp) (Python):
same tools, same behavior, distributed as a standalone Windows `.exe` with no
Python or Rust toolchain required to run it.

## Install

Download the latest `outlook-mcp-rs.exe` from
[Releases](https://github.com/adamkopelman/outlook-mcp-rs/releases) and point
your MCP client at it directly — no install step.

## Requirements

- Windows
- Classic Outlook desktop installed and signed in

## Development

See `TESTING.md` for how to run the unit test suite and the local live-Outlook
system tests.
