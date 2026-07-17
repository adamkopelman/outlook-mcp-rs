# outlook-mcp-rs

A single-binary [Model Context Protocol](https://modelcontextprotocol.io) server that
gives AI assistants control of the **classic Microsoft Outlook desktop app** on Windows —
email, calendar, tasks, and notes — by driving Outlook's native Win32 COM automation API.

No Python, no Node, no cloud/Graph API, and no runtime toolchain: you ship one Windows
`.exe`, point your MCP client at it, and it talks to the copy of Outlook already running
on the machine, signed in as you.

## Highlights

- **One self-contained binary.** A single `.exe` with no runtime dependencies — nothing to
  install, no interpreter, no service to host.
- **Runs against local Outlook, as you.** It drives the desktop app's own COM automation, so
  it inherits your existing session, accounts, and shared-folder permissions. There are no
  tokens to manage and no separate authentication step — if you can see it in Outlook, so can
  the server.
- **26 tools across five areas** — email, calendar, attachments, tasks, and notes (full list
  below).
- **Deliberate about side effects.** The handful of tools that actually send mail or meeting
  responses are explicit and opt-in, and the test suite is built so nothing is delivered by
  accident (see [Safety and side effects](#safety-and-side-effects)).

## Requirements

- Windows
- Classic Outlook desktop installed, running, and signed in to a mailbox

> **Note:** this drives the *classic* desktop client's COM object model. The "new Outlook"
> preview and Outlook on the web expose no such interface and are not supported.

## Install

Download the latest `outlook-mcp-rs.exe` from
[Releases](https://github.com/adamkopelman/outlook-mcp-rs/releases). It's a standalone
executable — there's no install step and no runtime dependencies to add.

## Configure your MCP client

By default, `outlook-mcp-rs` speaks MCP over stdio and takes no arguments — point any
local MCP-capable client at the executable's path. (To connect from another machine on
your network, see [Remote / network mode](#remote--network-mode-connect-from-another-machine)
below.)

For example, in Claude Desktop's `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\path\\to\\outlook-mcp-rs.exe"
    }
  }
}
```

Restart the client after editing its config. The server connects to whatever Outlook is
already running, and the 26 tools below become available.

## Remote / network mode (connect from another machine)

By default the server speaks MCP over **stdio**, which requires the MCP client
to launch the binary as a local child process — so the client must run on the
same Windows machine as Outlook.

If your MCP client runs on a **different machine on the same LAN** (for example,
Claude on a Linux box talking to Outlook on your Windows PC), run the server in
**network mode** instead. It listens on a TCP port using MCP Streamable HTTP,
and the remote client connects by URL.

### On the Windows machine (where Outlook runs)

Start the server in HTTP mode with a shared secret token:

```
outlook-mcp-rs.exe --http --port 8080 --token YOUR_SECRET
```

It prints `outlook-mcp-rs listening on http://0.0.0.0:8080/mcp` and serves the
same 26 tools as stdio mode. Outlook must be running and signed in, as usual.

Find this machine's name (the client connects to it):

```
hostname
```

Allow the port through Windows Firewall, scoped to your local subnet (adjust
the port and subnet to match your network):

```
netsh advfirewall firewall add rule name="outlook-mcp-rs" dir=in action=allow ^
  protocol=TCP localport=8080 remoteip=192.168.1.0/24
```

### On the client machine (e.g. Ubuntu running Claude Code)

Point the client at `http://<WINDOWS-HOSTNAME>:8080/mcp`, sending the token as a
bearer header. With Claude Code:

```
claude mcp add --transport http outlook http://WINDOWS-PC:8080/mcp \
  --header "Authorization: Bearer YOUR_SECRET"
```

Replace `WINDOWS-PC` with the `hostname` from above (or the machine's LAN IP).
The Outlook tools then appear in that client.

### Options

| Flag | Meaning |
|---|---|
| `--http` | Enable network (Streamable HTTP) mode instead of stdio. |
| `--port <PORT>` | Listen on `0.0.0.0:<PORT>`. Mutually exclusive with `--bind`. |
| `--bind <ADDR>` | Listen on an exact socket address, e.g. `127.0.0.1:8080`. |
| `--token <SECRET>` | Require `Authorization: Bearer <SECRET>` on every request. |

### Security

- **Set a token.** Without `--token`, network mode accepts any request that
  reaches the port — anyone on the network could read or send your mail. The
  token is optional only to make first-run testing easy; use it for any real
  setup.
- **Scope the firewall** to the specific hosts/subnet that need access, as
  shown above.
- **Traffic is plain HTTP** (no TLS in this version). Keep it on a trusted LAN;
  if you need encryption across untrusted networks, front it with a reverse
  proxy or tunnel that terminates TLS.

## Available tools

26 MCP tools, grouped by category:

**Email**
- `list_folders` — list mail folders (name, path, item counts)
- `list_emails` — find emails in a folder with an optional text query (matches subject, sender, and body) and filters (sender, category, date range, attachments, flagged, importance)
- `get_email` — get the full body and attachment list of one email by id
- `send_email` — send a new email immediately
- `create_draft` — create a draft email without sending it
- `reply_email` — reply to an email, optionally to all recipients, optionally as a draft
- `update_email` — change an existing email: move to a folder, mark read/unread, flag (follow_up/complete/clear), add/remove categories, set importance
- `delete_email` — delete an email (moves it to Deleted Items)

**Calendar**
- `list_events` — list/search calendar events by date range, text (subject/location), category, show_as, your response, or attendees; view meetings-only or all-day; or open another person's shared calendar with `calendar_of`
- `get_event` — get the full details of one calendar event by id
- `create_event` — create a calendar event; supports two tiers of attendees, categories, `show_as`, and recurrence, with `send` controlling whether invites actually go out
- `update_event` — change an existing event (subject, times, location, body, attendees, reminder, recurrence…); optionally notify attendees
- `respond_to_meeting` — respond to a meeting invite (accept, decline, or tentative)
- `delete_event` — delete/cancel an event (moves it to Deleted Items); optionally send a cancellation to attendees
- `check_availability` — check free/busy for one or more people over a time window; returns each person's per-slot status plus the windows where everyone is free

**Attachments**
- `list_attachments` — list an email's attachments (filename and size)
- `save_attachments` — save an email's attachments to a local directory

**Tasks**
- `list_tasks` — list Outlook tasks (filter by category, importance, or a text query matching subject or body)
- `create_task` — create a new Outlook task
- `update_task` — change an existing task: mark complete/reopen, subject, body, due_date, start_date, importance, add/remove categories, percent_complete, reminder_time
- `delete_task` — delete a task (moves it to Deleted Items)

**Notes**
- `list_notes` — list Outlook notes (filter by category or a text query on the body)
- `get_note` — get the full body of one note by id
- `create_note` — create a new Outlook note (optional categories, color)
- `update_note` — change an existing note: body, add/remove categories, color
- `delete_note` — delete a note (moves it to Deleted Items)

## How it works

The server links directly against the Windows COM / OLE Automation APIs (via the
[`windows`](https://crates.io/crates/windows) crate) and drives the same
`Outlook.Application` object model that Outlook VBA macros use. Every request runs against
the live desktop client, so it sees exactly the folders, accounts, and shared calendars
your Outlook session already has access to — no mailbox data leaves the machine except what
your MCP client chooses to send upstream.

## Building from source

Building requires a [Rust toolchain](https://rustup.rs) (2024 edition) on Windows:

```
cargo build --release
```

The binary is produced at `target/release/outlook-mcp-rs.exe`.

## Safety and side effects

Most tools are read-only or reversible — deletes move items to Deleted Items rather than
destroying them. A few have real, outbound effects and are explicit in the tool call:
`send_email` delivers mail; `respond_to_meeting` notifies an organizer; and
`create_event`, `update_event`, and `delete_event` notify attendees when their send/update/
cancellation flag is set. [`TESTING.md`](TESTING.md) spells out exactly which behaviors are
covered by automated tests versus verified by hand precisely because they send real mail.

## Development

See [`TESTING.md`](TESTING.md) for how to run the unit test suite and the local
live-Outlook system tests.

## License

MIT — see [`LICENSE`](LICENSE).
