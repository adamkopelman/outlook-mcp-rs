---
name: using-outlook-mcp
description: Use when working with the user's Microsoft Outlook through the outlook-mcp-rs MCP tools (list_emails, send_email, create_draft, update_draft, list_events, create_event, list_tasks, get_task, list_notes, check_availability, get_inline_image, …) — reading, searching, sending or replying to mail, editing drafts, handling attachments and inline images, scheduling or answering meetings, or managing tasks and notes.
---

# Using the Outlook MCP tools

## Overview

outlook-mcp-rs drives the user's real, running Outlook desktop app. Every call acts on
live mail. A send, invite, reply or cancellation reaches real people and can't be
undone. Items are addressed by opaque ids that the list and create tools return.

## Actions that reach other people

These default to **sending**. Show the user what will go out and get a yes first,
unless they already asked for that exact send in this conversation:

| Tool | Sends by default via | Quiet alternative |
|---|---|---|
| `send_email` | always | `create_draft` |
| `reply_email` | `send` (default true) | `send: false` saves a draft |
| `create_event` with attendees | `send` (default true) | `send: false` saves unsent |
| `update_event` on a meeting | `send_update` (default true) | `send_update: false` |
| `delete_event` you organize | `send_cancellation` (default true) | `send_cancellation: false` |
| `respond_to_meeting` | `send` (default true) | `send: false` |
| `update_draft` | only with `send: true` | omit `send` (default false) |

When the user's intent is ambiguous ("write to Dana about…"), make a draft and say so.
To revise a draft before it goes out, use `update_draft`. It saves the draft and sends it only
with `send: true`, which needs the same go-ahead as `send_email`.

## Actions that destroy data

| Tool | Effect |
|---|---|
| `delete_email` (default) | Moves the email to Deleted Items, where it can be recovered |
| `delete_email` with `permanent: true` | Hard delete, like Shift+Delete. **Can't be recovered from Deleted Items** |
| `empty_deleted_items` with `confirm: true` | Permanently deletes **everything** in Deleted Items, including the user's own items |

Use `permanent` or `empty_deleted_items` only when the user explicitly asks for a
permanent delete. Never set `confirm: true` on your own initiative.

## Ids

- Take ids only from a tool result (`list_*`, `create_*`). Never build or guess one.
- Moving an email (`update_email` with `move_to`) or deleting it **changes its id**.
  From then on, use the `id` in that call's result, or find the item again.

## Quick reference

- **Dates (every tool):** local time, as ISO (`2026-06-10`, `2026-06-10T14:30`), a keyword
  (`today`, `yesterday`, `tomorrow`, `now`, `start_of_week`/`end_of_week`, `start_of_month`/`end_of_month`,
  `start_of_year`/`end_of_year`; weeks start on the first day of the week in the Windows user's
  regional settings, e.g. Monday or Sunday), an offset from now (`-14d`, `+3h`, `-2w`;
  units `m h d w mo y`), or a keyword plus offsets (`start_of_week-1w`). A bare date in a
  `*_before` filter includes that whole day.
- **All `list_*` tools:**
  - Filters AND together. String filters take one value or a list (matching any): `from: ["A", "B"]`.
  - `query`: terms must all match (any case, any language); `"quoted phrase"`; `*` wildcard;
    `field:term` scopes (`subject:`, `from:`, `to:`, `body:` for email; `subject:`, `location:`,
    `organizer:`, `attendees:` for events; `subject:`, `body:` for tasks and notes).
  - Date ranges are `<field>_after`/`<field>_before`: `received_*`, `start_*`, `due_*`, `created_*`.
  - Paging: `count` + `offset` everywhere (caps: emails 200, events 250, tasks/notes 500).
- **`list_emails`:**
  - `folder` defaults to `inbox`. Accepts `inbox`, `sent`, `drafts`, `deleted`, `outbox`, or a path like `Inbox/Receipts` (see `list_folders`).
  - Results are newest first. `count` defaults to 10, max 200.
  - **To page:** call again with `offset += count` until fewer than `count` come back.
  - `query` searches subject, sender and body. It works for Hebrew and other non-Latin text, but a non-ASCII query may be slower, because the server scans up to 2,000 items itself when Outlook's own search finds nothing.
  - Filters:
    - `from` matches the sender; `to` matches any To/CC recipient name or address;
    - `item_type: "email"` skips meeting invites, bounces and read receipts (values as `get_email` reports them);
    - `importance` (`low`/`normal`/`high`), `flag` (`follow_up`/`complete`/`clear`), `category`, `unread_only`, `has_attachments`, `received_after`/`received_before` (e.g. `"-14d"`).
    - `since_days`, `flagged` and `high_importance` still work but are deprecated.
- **`get_email` / `get_event` / `get_note` / `get_task`:**
  - Pass a **list of ids** to read many items in one call. The result is a list in the same order; an id that fails comes back as `{"id", "error"}` while the rest succeed.
  - `include` picks the heavy fields: on `get_email` any of `body`, `html_body`, `attachments`, `meeting` (default: `body`, `attachments`, `meeting`); on the others just `body`. `include: []` gives metadata only, which is the cheap way to scan many items.
  - `body_format: "html"` (`get_email` only) adds `html_body`. `prefer_html` is the old name.
  - Bodies are cut at `max_body_chars` (default 100,000, up to 5,000,000). If `body_truncated` (or `html_truncated`) is true, call again with `max_body_chars` ≥ `body_length` (or `html_length`), or use `output_dir`.
  - `output_dir` writes each body in full to a file and returns `body_file` / `html_body_file` instead. Use it for big HTML mail, then read the file.
  - `resolve_inline_images: true` on `get_email` returns `html_body` with every `cid:` image replaced by a `data:` URI. Unresolved ones are listed in `inline_images_unresolved`. The result is large: pair it with `output_dir`.
- **Bodies (all tools that write content):**
  - Mail tools take exactly one of `body` (plain text), `html_body`, `body_file` or `html_body_file`. Events, tasks and notes take `body` or `body_file`.
  - The `*_file` forms are local paths to UTF-8 files. Use them for long or image-heavy bodies: write the HTML to a file and pass its path instead of a huge JSON string.
  - `html: true` (old flag) still works but is deprecated; use `html_body`.
  - Mail tools also take `categories` (or `add_categories`/`remove_categories` on `update_draft`) and `importance` (`low`/`normal`/`high`).
- **`update_draft`:**
  - Pass the draft's id as `email_id` (`draft_id` still works).
  - Only works on unsent drafts. It refuses received or sent mail.
  - `subject` and the body replace the current value.
  - `to`, `cc` and `bcc` replace the whole line, and `[]` clears it.
  - `attachments` and `inline_images` are added to the existing ones.
  - `send: true` sends the draft after the changes (and alone, sends it as is).
- **Attachments:**
  - `list_attachments` gives each attachment's `type`, `content_id`, `mime_type`, `hidden`, and `is_inline` (an image shown inside the HTML body, as opposed to a regular attachment). Pass a list of email ids to get `[{id, attachments}]` in one call.
  - Give `save_attachments` an absolute `save_dir` (or `~/...`). A relative path resolves against the server's working directory, not the user's.
  - To re-send an email with its attachments, read it with `get_email` + `resolve_inline_images` (images travel inside the HTML) and save the rest with `save_attachments` + `inline: false`. Saving everything and re-attaching it would add each inline image a second time as a regular attachment.
- **Inline images:**
  - **Writing:** on `send_email`, `create_draft`, `reply_email` and `update_draft`, `data:` image URIs in `html_body` (e.g. `<img src="data:image/png;base64,...">`) become real inline attachments automatically. Or pass `inline_images: [{content_id, path | data_base64}]` and reference each as `<img src="cid:CONTENT_ID">` (needs an HTML body).
  - **Reading:** `get_inline_image(email_id, content_id)` returns the image as a `data:` URI (max 10 MB). Add `context_lines` (max 50) to also get the text just before the image. `content_ids: [...]` fetches several at once; `output_dir` saves them as files (`data_file`).
  - **Whole email at once:** `get_email` with `resolve_inline_images: true` (and `output_dir`) gives self-contained HTML in one call.
- **`list_events`:**
  - `start_after`/`start_before` bound the event start; the default window is today plus 7 days. `count` default and max 250, page with `offset`.
  - Recurring events are expanded.
  - `calendar_of` opens someone's shared calendar.
- **`check_availability`:**
  - Shows free/busy only, never event details, in `interval_minutes` slots (default 30).
  - `common_free` lists the windows where everyone resolved is free.
- **Tasks and notes:**
  - `list_tasks` hides completed tasks unless `include_completed` is set; `due_after`/`due_before` filter by due date (tasks with no due date never match).
  - `list_notes` filters by `created_after`/`created_before`.
  - `get_task` returns one task's body, start/completion dates, percent complete and reminder.
  - `update_task` with `mark_complete` completes or reopens a task.
  - Deleting a task, note or event moves it to Deleted Items, where it can be recovered.

## Common mistakes

| Mistake | Fix |
|---|---|
| Replying with defaults when the user only wanted text to review | `send: false`, then report the draft |
| Reusing an email id after `move_to` or delete | Use the id returned by the move, or find the item again |
| Treating the first page as all the mail | Page with `offset` until a page comes back short |
| Quoting a cut-off email as complete | Check `body_truncated` and refetch with a larger `max_body_chars`, or use `output_dir` |
| Calling `get_email` once per id in a loop | Pass the ids as a list; use `include: []` when only headers are needed |
| Editing a draft by deleting it and creating a new one | `update_draft` keeps the same draft |
| Using `permanent: true` for an ordinary "delete this" | Omit it; the default is recoverable |
| Treating an empty result as "nothing there" when the call looked wrong | See troubleshooting-outlook-mcp |
