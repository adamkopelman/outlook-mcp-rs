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

## Available tools

20 MCP tools, grouped by category:

**Email**
- `list_folders` — list mail folders (name, path, item counts)
- `list_emails` — find emails in a folder with an optional text query and filters (sender, category, date range, attachments, flagged, importance)
- `get_email` — get the full body and attachment list of one email by id
- `send_email` — send a new email immediately
- `create_draft` — create a draft email without sending it
- `reply_email` — reply to an existing email
- `update_email` — change an existing email: move to a folder, mark read/unread, flag (follow_up/complete/clear), add/remove categories, set importance
- `delete_email` — delete an email

**Calendar**
- `list_events` — list calendar events in a date range (default: next 7 days)
- `get_event` — get the full details of one calendar event by id
- `create_event` — create a new calendar event
- `respond_to_meeting` — respond to a meeting invite (accept, decline, or tentative)

**Attachments**
- `list_attachments` — list an email's attachments (filename and size)
- `save_attachments` — save an email's attachments to disk

**Tasks**
- `list_tasks` — list Outlook tasks
- `create_task` — create a new Outlook task
- `update_task` — change an existing task: mark complete/reopen, subject, body, due_date, start_date, importance, add/remove categories, percent_complete, reminder_time

**Notes**
- `list_notes` — list Outlook notes
- `get_note` — get the full body of one note by id
- `create_note` — create a new Outlook note

## Development

See `TESTING.md` for how to run the unit test suite and the local live-Outlook
system tests.
