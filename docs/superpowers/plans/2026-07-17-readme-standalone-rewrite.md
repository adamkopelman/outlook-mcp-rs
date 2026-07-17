# Standalone README Rewrite Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Rewrite `README.md` so it (1) stands entirely on its own — no framing as a "counterpart" to the Python [outlook-mcp](https://github.com/adamkopelman/outlook-mcp) repo — and (2) actually documents the project a new reader needs: what it is, how to wire it into an MCP client, how to build it, what's safe vs. what sends real mail, and the *complete, correct* tool list.

**Architecture:** The current `README.md` opens by defining the project *relative to* another repo ("This is the Rust counterpart to outlook-mcp (Python): same tools, same behavior") — so a reader can't understand it without leaving to a second project. It also undersells and under-documents the actual surface: it claims "20 MCP tools" but `src/server.rs` exposes **26** (`#[tool(...)]` count), missing `update_event`, `delete_event`, `check_availability`, and `delete_task` entirely. There is no MCP-client configuration example anywhere (the single most important thing a user needs), no build-from-source instructions, and no statement of which tools have real outbound side effects (a gap `TESTING.md` already documents in detail but the README never surfaces). This plan replaces the intro, adds the missing sections, and regenerates the tool list from `src/server.rs` as the source of truth.

**Tech Stack:** Markdown only. No code changes. The tool list is verified against `src/server.rs`'s `#[tool(description = ...)]` attributes; the config snippet against `src/main.rs` (stdio transport, no args, no env).

## Global Constraints

- **Source of truth for the tool list is `src/server.rs`** — every bullet must correspond to a real `#[tool(...)]` method, and the stated count must equal the number of `#[tool(` attributes (26). Do not copy the stale "20" or the old category counts.
- **No dependency on the Python `outlook-mcp` repo for comprehension.** The rewritten README must make sense with zero knowledge of that project. A single, optional "prior art / see also" mention is acceptable, but nothing in the intro, install, or usage path may require it.
- **Don't invent behavior.** Tool one-liners are condensed from the real `#[tool(description=...)]` strings and the existing README; config/build facts come from `src/main.rs` and `Cargo.toml`. No speculative flags, platforms, or install methods.
- **Keep the existing `## Development` → `TESTING.md` handoff.** Testing detail lives in `TESTING.md`; the README links to it rather than duplicating it.
- **Windows-only framing stays accurate:** classic Outlook desktop, COM automation, `.exe`. Don't imply macOS/Linux/new-Outlook support.

---

### Task 1: Rewrite README.md as a standalone document

Replace the whole file. The new structure: title + standalone tagline → Highlights → Requirements → Install → Configure your MCP client → Available tools (all 26) → How it works → Building from source → Safety and side effects → Development → License.

**Files:**
- Rewrite: `README.md` (entire file)

**Interfaces:**
- Consumes (as source of truth, read-only): `src/server.rs` (`#[tool(...)]` set), `src/main.rs` (stdio, no args/env), `Cargo.toml` (edition 2024, name `outlook-mcp-rs`), `LICENSE` (MIT, Adam Kopelman 2026).
- Produces: nothing consumed by later tasks; this is the only content task.

- [ ] **Step 1: Replace `README.md` with the new content**

Write `README.md` in full as:

````markdown
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

`outlook-mcp-rs` speaks MCP over stdio, takes no arguments, and needs no environment
variables. Point any MCP-capable client at the executable's path.

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
````

- [ ] **Step 2: Verify the tool list is complete and the count is right**

Cross-check every bullet against the real tool set and confirm the count:

```bash
grep -c '#\[tool(' src/server.rs        # expect 26
grep -oP '(?<=pub async fn )\w+' src/server.rs | sort > /tmp/impl_tools.txt
```

Then confirm each of the 26 method names appears as a bullet in `README.md` and no bullet
names a tool that isn't in `src/server.rs`. The five categories must total 26
(Email 8, Calendar 7, Attachments 2, Tasks 4, Notes 5).

- [ ] **Step 3: Sanity-check the rewrite for the two acceptance criteria**

1. **Disconnected:** `grep -n 'outlook-mcp' README.md` returns nothing pointing at the Python
   repo as a prerequisite — the file must read as self-contained. (A match on the project's
   *own* name `outlook-mcp-rs` is fine; a "counterpart to outlook-mcp (Python)" framing is not.)
2. **Renders:** headings, the JSON code fence, and the `#safety-and-side-effects` anchor link
   all resolve.

- [ ] **Step 4: Commit**

```bash
git add README.md docs/superpowers/plans/2026-07-17-readme-standalone-rewrite.md
git commit -m "Rewrite README as a standalone document"
```

---

## Self-Review Notes

- **Spec coverage:** "disconnect it from the outlook-mcp repo" → new intro defines the project
  on its own terms; Task 1 Step 3 asserts no Python-repo prerequisite remains. "more stuff that
  would be good" → adds Highlights, an MCP-client **Configure** section (the biggest gap — a
  copy-pasteable JSON snippet), How it works, Building from source, Safety and side effects, and
  a License section, and corrects the tool list from a stale 20 to the real 26.
- **Accuracy:** the tool list is regenerated from `src/server.rs` (the `#[tool]` source of
  truth) rather than edited in place, and Step 2 verifies the count; the config snippet matches
  `src/main.rs` (stdio, no args, no env).
- **No placeholders:** Step 1 contains the complete literal README content, not a sketch.
- **Scope:** README-only; no code, no behavior change, so nothing downstream depends on it.
