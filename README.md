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
- **30 tools across five areas** — email, calendar, attachments, tasks, and notes (full list
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
already running, and the 30 tools below become available.

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
same 30 tools as stdio mode. Outlook must be running and signed in, as usual.

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

30 MCP tools, grouped by category:

**Email**
- `list_folders` — list mail folders (name, path, item counts)
- `list_emails` — find emails in a folder with an optional text `query` (subject, sender, and body; see [List tool conventions](#list-tool-conventions); non-ASCII queries such as Hebrew fall back to a client-side scan when Outlook's search finds nothing) and filters: sender via `from`, recipient via `to` (any To/CC name or address), `category`, `item_type` (e.g. `"email"` to skip meeting invites and bounces), `importance`, `flag` (`follow_up`/`complete`/`clear`), attachments, unread, `received_after`/`received_before`; newest first, `count` default 10, max 200, page with `offset`
- `get_email` — get one email by id, or several in one call; see [Reading items](#reading-items) for `include`, `body_format`, `max_body_chars`, `output_dir` and `resolve_inline_images`
- `send_email` — send a new email immediately (to/cc/bcc, `body` or `html_body` — or `body_file`/`html_body_file` for large bodies — file attachments, inline images, categories, importance)
- `create_draft` — create a draft email without sending it (same options as `send_email`)
- `reply_email` — reply to an email, optionally to all recipients, optionally as a draft (same body, inline image, category and importance options)
- `update_email` — change an existing email: move to a folder, mark read/unread, flag (follow_up/complete/clear), add/remove categories, set importance
- `update_draft` — edit an unsent draft by `email_id` (`draft_id` is a deprecated alias): subject, body/HTML body (or `body_file`/`html_body_file`), To/CC/BCC (each replaces that line; `[]` clears it), append attachments or inline images, add/remove categories, set importance; saves, and sends only with `send=true`
- `delete_email` — delete an email (moves it to Deleted Items), or hard-delete it with `permanent=true` (like shift+delete; **irreversible**, not recoverable from Deleted Items)
- `empty_deleted_items` — **permanently** delete everything in Deleted Items (items and subfolders); **irreversible**, refuses unless `confirm=true`. On Exchange/Microsoft 365, retention policy may still keep items in Recoverable Items

**Calendar**
- `list_events` — list/search calendar events by start (`start_after`/`start_before`, default today plus 7 days), text (subject/location), category, show_as, your response, or attendees; view meetings-only or all-day; or open another person's shared calendar with `calendar_of`; `count` default and max 250, page with `offset`
- `get_event` — get the full details of one calendar event by id, or several in one call (same output options as `get_email`, plain-text body only)
- `create_event` — create a calendar event; supports two tiers of attendees, categories, `show_as`, and recurrence, with `send` controlling whether invites actually go out; `body` or `body_file`
- `update_event` — change an existing event (subject, times, location, body or `body_file`, attendees, reminder, recurrence…); optionally notify attendees
- `respond_to_meeting` — respond to a meeting invite (accept, decline, or tentative)
- `delete_event` — delete/cancel an event (moves it to Deleted Items), or hard-delete it with `permanent=true` (**irreversible**); for a meeting you organize, `send_cancellation` (default true) notifies attendees first, independent of `permanent`. A recurring event's id (including an occurrence's) names the whole series, so the whole series is deleted
- `check_availability` — check free/busy for one or more people over a time window; returns each person's per-slot status plus the windows where everyone is free

**Attachments**
- `list_attachments` — list an email's attachments with metadata: index, filename, size, type (file/link/item/ole), Content-ID (for `cid:` references in HTML bodies), MIME type, hidden flag, and `is_inline` (inline `cid:` content vs. a standalone attachment); accepts a list of email ids
- `save_attachments` — save an email's attachments to a local directory (each result carries the same metadata plus `saved_to`/`status`). `inline: false` saves only regular attachments and skips inline images (`inline: true` the reverse), so re-sending an email read with `resolve_inline_images` doesn't attach its images twice
- `get_inline_image` — fetch an attachment by Content-ID (e.g. an inline `cid:` image) as a base64 data URI (up to 10 MB); optional `context_lines` (max 50) also returns `context`, the plain-text lines just before the image's first `cid:` reference in the HTML body (`""` if it isn't referenced); `content_ids` fetches several images of one email in one call, and `output_dir` writes the image files to disk (`data_file`) instead of returning base64

**Tasks**
- `list_tasks` — list Outlook tasks (filter by category, importance, due date via `due_after`/`due_before`, or a text query matching subject or body); `count` default and max 500, page with `offset`
- `get_task` — get the full details of one task by id, or several in one call: body, start date, completion date, percent complete, reminder, created/modified (same output options as `get_event`)
- `create_task` — create a new Outlook task (`body` or `body_file`)
- `update_task` — change an existing task: mark complete/reopen, subject, body (or `body_file`), due_date, start_date, importance, add/remove categories, percent_complete, reminder_time
- `delete_task` — delete a task (moves it to Deleted Items), or hard-delete it with `permanent=true` (**irreversible**)

**Notes**
- `list_notes` — list Outlook notes (filter by category, creation date via `created_after`/`created_before`, or a text query on the body); `count` default and max 500, page with `offset`
- `get_note` — get the full body of one note by id, or several in one call (same output options as `get_event`)
- `create_note` — create a new Outlook note from `body` or `body_file` (optional categories, color)
- `update_note` — change an existing note: body (or `body_file`), add/remove categories, color
- `delete_note` — delete a note (moves it to Deleted Items), or hard-delete it with `permanent=true` (**irreversible**)

### List tool conventions

`list_emails`, `list_events`, `list_tasks` and `list_notes` share these rules:

- **Filters combine with AND.** A string filter (`from`, `to`, `category`, `attendees`, `show_as`, `my_response`, `importance`, `item_type`, `flag`) takes one value or a list, and a list matches **any** of its values: `from: ["Person A", "Person B"]`.
- **`query`** is a list of terms that must all match, case-insensitive and in any language:
  - `"quoted phrase"` keeps words together, in order;
  - `*` matches any run of characters: `status*report`;
  - `field:term` limits a term to one field: `subject:invoice`, `from:"Ada Lovelace"`. Fields: `subject`, `from`, `to`, `body` (emails); `subject`, `location`, `organizer`, `attendees` (events); `subject`, `body` (tasks and notes). Any other `word:` is plain text, so `10:30` still matches literally.
- **Date ranges** are `<field>_after` / `<field>_before`, both inclusive: `received_*` (emails), `start_*` (events; `start_date`/`end_date` still work), `due_*` (tasks), `created_*` (notes). A bare date in `*_before` includes that whole day.
- **Paging:** `count` and `offset` on every list tool. Caps: emails 200 (default 10), events 250, tasks 500, notes 500 (default = cap). Call again with `offset += count` until a page comes back short.
- **Deprecated** (still accepted): `since_days` (use `received_after: "-14d"`), `flagged` (use `flag`), `high_importance` (use `importance: "high"`).

### Dates

Every date parameter on every tool (list filters, `create_event`/`update_event` `start`/`end`, recurrence `until`, task `due_date`/`start_date`/`reminder_time`, `check_availability`) accepts the same forms, in local time:

| Form | Examples | Meaning |
|---|---|---|
| ISO date or datetime | `2026-06-10`, `2026-06-10T14:30`, `2026-06-10 14:30:00` | Always year-month-day |
| Keyword | `now`, `today`, `yesterday`, `tomorrow`, `start_of_week`, `end_of_week`, `start_of_month`, `end_of_month`, `start_of_year`, `end_of_year` | Day keywords and `start_of_*` are midnight; `end_of_*` is 23:59:59 on the last day. Weeks start on the first day of the week in the Windows user's regional settings (e.g. Monday, Sunday or Saturday; Monday only if it can't be read) |
| Offset from now | `-14d`, `+3h`, `-2w`, `-30m`, `+1mo`, `-1y` | Units: `m` minutes, `h` hours, `d` days, `w` weeks, `mo` months, `y` years |
| Keyword plus offsets | `today-1d`, `start_of_week-1w`, `tomorrow+9h` | Applied left to right |

Known issue: on a day-first Windows locale, date *filters* can still return the wrong range (issue #1); the grammar above is not affected.

### Reading items

`get_email`, `get_event`, `get_note` and `get_task` share the same output controls:

- **Several ids at once:** pass the id parameter (`email_id`, `event_id`, `note_id`, `task_id`;
  plural aliases such as `email_ids` work too) as a list. The result is then a list in the
  same order, and an id that fails becomes `{"id": ..., "error": ...}` instead of failing the
  whole call. A single id (not a list) returns exactly the object it always did.
  `list_attachments` (a list of email ids) and `get_inline_image` (`content_ids`) batch the
  same way.
- **`include`** picks the optional, heavy fields. `get_email`: `"body"`, `"html_body"`,
  `"attachments"`, `"meeting"` (default `["body", "attachments", "meeting"]`, today's output).
  The others: `"body"` (default `["body"]`). `[]` returns metadata only.
- **`body_format`**: `"text"` (default) or `"html"`. On `get_email`, `"html"` also returns
  `html_body`; `prefer_html: true` is the deprecated spelling of the same thing. Events,
  notes and tasks only have plain-text bodies.
- **`max_body_chars`** (default 100,000, 1,000–5,000,000) cuts each inline body. Every
  returned body reports `*_truncated` and `*_length` (the full length in characters):
  `body_truncated`/`body_length`, `html_truncated`/`html_length`.
- **`output_dir`** writes each body in full (never truncated) to a file in that directory
  (created if missing) and returns its absolute path as `body_file` / `html_body_file`
  instead of the text. File names are derived from the item id, so reading the same item
  again overwrites them.
- **`resolve_inline_images`** (`get_email` only, implies `html_body`) replaces every
  `cid:` reference in `html_body` with the image's base64 `data:` URI, so the HTML is
  self-contained. Content-ID matching ignores case and `<>`; references to unknown
  Content-IDs, to images over 10 MB, or to unreadable attachments stay as `cid:` and are
  listed in `inline_images_unresolved` (`inline_images_resolved` counts the replaced ones).
  The resolved HTML is usually large, so combine it with `output_dir`.

### Body inputs

Every tool that writes content follows one convention:

- Mail tools (`send_email`, `create_draft`, `reply_email`, `update_draft`) take the body as exactly one of
  `body` (plain text), `html_body`, `body_file` or `html_body_file` (`update_draft`: at most one, since the
  body is optional there). The `*_file` forms take a local path to a UTF-8 file; use them for large bodies,
  which also avoids escaping big HTML inside JSON.
- The old `html: true` flag still works on `send_email`/`create_draft`/`reply_email` but is deprecated:
  `html: true` with `body`/`body_file` means HTML. `html: false` together with `html_body` is an error.
- Events, tasks and notes have a plain-text `body` with a `body_file` sibling (not both).
- `categories` (create tools) or `add_categories`/`remove_categories` (update tools), and `importance` on
  the mail tools.

All inputs, including files and images, are read and validated before anything is created, changed or sent.

### Inline images

Images inside an HTML body are sent as real Content-ID attachments instead of base64 inside the HTML
string. There are two ways, on every mail-writing tool (`send_email`, `create_draft`, `reply_email`,
`update_draft`):

- **`data:` URIs in the HTML are converted automatically.** Each `data:image/...` URI used as an attribute
  value (`<img src="data:image/png;base64,...">`, quoted or not) or in a CSS `url(...)` becomes a hidden
  inline attachment, and the reference is rewritten to `cid:img-<hash>`. Base64 and percent-encoded
  payloads are accepted; the same image used twice is attached once. Non-image `data:` URIs (for example
  `data:text/plain,...`) are left untouched, and an image URI that isn't valid base64 is an error.
- **`inline_images`** lists images explicitly. Each entry has a `content_id` and exactly one of a local
  `path` or `data_base64` (a `data:image/png;base64,` prefix and whitespace are accepted). `filename` and
  `mime_type` are optional; the MIME type is otherwise guessed from the file name or the image bytes.
  Reference each image in the HTML as `<img src="cid:CONTENT_ID">`. Requires an HTML body.

```json
{
  "to": ["ada@example.com"], "subject": "Q3",
  "html_body": "<p>Results:</p><img src=\"cid:chart\"><p>Logo:</p><img src=\"data:image/png;base64,iVBORw0...\">",
  "inline_images": [{"content_id": "chart", "path": "C:/reports/q3.png"}]
}
```

On `update_draft`, an `inline_images` entry replaces an existing attachment with the same Content-ID, and
a converted `data:` image is not added again if the draft already has it (its Content-ID is a hash of the
image). Base64 data is written to a temp file for `Attachments.Add` and deleted afterwards. Images are
marked hidden, but some Outlook versions may still list them as regular attachments.

### Large requests

Neither transport has a small size limit: stdio has none, and the HTTP transport accepts request bodies up
to 64 MiB and answers a larger one with an explicit `413 payload too large`. A compressed (`Content-Encoding`)
request is refused with `415`. For large bodies, prefer `body_file` / `html_body_file`: the content never
has to pass through the tool call at all.

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

Most tools are read-only or reversible. By default, deletes move items to Deleted Items
rather than destroying them. Permanent deletes must always be asked for explicitly:

- `delete_email`, `delete_event`, `delete_task` and `delete_note` with `permanent=true`
  hard-delete one item (like Shift+Delete; not recoverable from Deleted Items, though
  Exchange/Microsoft 365 retention may still keep it in Recoverable Items);
- `empty_deleted_items` with `confirm=true` destroys everything in Deleted Items.

A few tools have real, outbound effects, and these are also explicit in the tool call:

- `send_email` delivers mail, as do `reply_email` (unless `send=false`) and `update_draft` with `send=true`;
- `respond_to_meeting` notifies an organizer;
- `create_event`, `update_event` and `delete_event` notify attendees when their send/update/cancellation flag is set.

`update_draft` edits a draft and only sends it when `send=true` is passed. [`TESTING.md`](TESTING.md) spells out exactly which behaviors are
covered by automated tests versus verified by hand precisely because they send real mail.

## Troubleshooting

### Hebrew or other non-ASCII text shows up as `????` or as gibberish

The server's output is UTF-8 end to end. Outlook text stays UTF-16 until it is
decoded to a Rust `String`, serialized as raw UTF-8 JSON (no `\uXXXX` escapes), and
written unchanged to stdout or the HTTP response. In network mode every JSON and SSE
response also says `charset=utf-8` in its `Content-Type`. The server never converts
text to a Windows code page, so garbled text means something else re-decoded it.
The shape of the garbage tells you what:

| You see (for `מייל שיקוף`) | What happened | Where to look |
|---|---|---|
| `???? ?????` | Text was squeezed into a code page that has no Hebrew | A console or wrapper that isn't UTF-8 |
| `×ž×™×™×œ ×©×™×§×•×£` | UTF-8 bytes were read as latin1 / windows-1252 | The client's decoding: an HTTP client guessing ISO-8859-1 for a response without a charset (versions before the `charset=utf-8` fix), or a script reading the server's output with the system code page |
| `îééì ùé÷åó` | Hebrew was encoded as **windows-1255** and then read as latin1 | Something between Outlook and you converted the text to the Windows "ANSI" code page (see below) |

The last row is the one where a `latin1 → windows-1255` re-decode "fixes" the
text. The server cannot produce it, because it never emits windows-1255 bytes. Two
things can:

1. **A wrapper around the executable.** If the MCP client starts the server
   through a script instead of running `outlook-mcp-rs.exe` directly, that
   script may decode the server's UTF-8 output with the system code page and
   re-encode it. Common culprits:
   - a **PowerShell** pipeline or `powershell -Command "... outlook-mcp-rs.exe ..."`
     launcher. PowerShell reads a native program's output using
     `[Console]::OutputEncoding` and writes it out again using `$OutputEncoding` or
     the console's code page, neither of which is UTF-8 by default on Windows
     PowerShell 5.1;
   - a **Python** wrapper or client using `subprocess` with `text=True` but no
     `encoding="utf-8"` (it then uses the ANSI code page, windows-1255 on a Hebrew
     system);
   - saving results with Windows PowerShell 5.1's `Set-Content` / `Out-File`, which
     do not write UTF-8 by default.

   The same wrappers also break Hebrew *arguments*, so a Hebrew search finds nothing.
2. **The message itself.** Mail sent with a missing or wrong charset label can be
   stored by Outlook with the wrong code page. Then the mailbox already holds the
   garbled text, Outlook shows the same garbage when you open the message, and the
   server reports it faithfully. To check, run the read-only live diagnostic
   `cargo test --test live_outlook -- --ignored inbox_text_from_com_has_no_cp1255_mojibake`
   on the Outlook machine. It fails, naming the emails, if COM itself returns this
   kind of text.

To fix it:

- Point the client's `command` straight at `outlook-mcp-rs.exe`, as in
  [Configure your MCP client](#configure-your-mcp-client), with no `cmd`,
  PowerShell or Python launcher in between.
- If you need a wrapper, make it UTF-8 in both directions. In PowerShell:
  `$OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)`.
  In Python: `subprocess.Popen(..., encoding="utf-8")`, or set `PYTHONUTF8=1`.
- To view output in a console, switch it to UTF-8 first: `chcp 65001` (cmd) or the
  PowerShell line above. Or inspect it in a UTF-8 viewer.
- On a remote (`--http`) setup, use a client that honours the `charset` in
  `Content-Type`; any MCP SDK client does.

## Skills for AI assistants

The [`skills/`](skills) folder has two [Agent Skills](https://docs.claude.com/en/docs/agents-and-tools/agent-skills/overview)
that teach an assistant (such as Claude Code) to use these tools well:

- [`using-outlook-mcp`](skills/using-outlook-mcp/SKILL.md) covers:
  - which calls send mail or destroy data, and when to confirm first;
  - how ids behave;
  - paging, truncated bodies, drafts and inline images;
  - a quick reference for every tool.
- [`troubleshooting-outlook-mcp`](skills/troubleshooting-outlook-mcp/SKILL.md) maps each
  error message to its cause and fix, and covers results that look wrong without an error.

To use them with Claude Code, copy both folders into `~/.claude/skills/` (for every
project) or into a project's `.claude/skills/`.

## Development

See [`TESTING.md`](TESTING.md) for how to run the unit test suite and the local
live-Outlook system tests. Maintainers system-testing a change against a real mailbox
can follow the `live-outlook-system-test` skill in
[`.claude/skills/`](.claude/skills/live-outlook-system-test/SKILL.md).

## License

MIT — see [`LICENSE`](LICENSE).
