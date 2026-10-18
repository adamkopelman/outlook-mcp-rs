# Changelog

Notable changes to outlook-mcp-rs. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project uses [semantic versioning](https://semver.org/).

**Every pull request adds its own entry under [Unreleased]**, in the matching group (Added / Changed / Deprecated / Removed / Fixed), ending with a link to the PR. A release moves those entries under a new version heading.

## [Unreleased]

## [0.4.0] - 2026-10-18

### Added

- `list_emails`: `offset` paging; the `count` cap is raised to 200 ([#14](https://github.com/adamkopelman/outlook-mcp-rs/pull/14)).
- `delete_email` with `permanent: true` hard-deletes an email, and the new `empty_deleted_items` tool empties Deleted Items (`confirm: true` required) ([#16](https://github.com/adamkopelman/outlook-mcp-rs/pull/16)).
- `list_emails`: a `to` filter matching any To/CC recipient by name or address ([#17](https://github.com/adamkopelman/outlook-mcp-rs/pull/17)).
- Attachment metadata (Content-ID, MIME type, attachment type, hidden flag) from `list_attachments` and `save_attachments` ([#18](https://github.com/adamkopelman/outlook-mcp-rs/pull/18)).
- `update_draft` tool: edit an unsent draft's subject, body, HTML body, To/CC/BCC and attachments ([#19](https://github.com/adamkopelman/outlook-mcp-rs/pull/19)).
- `get_email`: `body_truncated` / `html_truncated` flags and a `max_body_chars` limit ([#20](https://github.com/adamkopelman/outlook-mcp-rs/pull/20)).
- Hebrew round-trip tests over the full COM → JSON → stdio path ([#21](https://github.com/adamkopelman/outlook-mcp-rs/pull/21)).
- Inline Content-ID images on `send_email` and `create_draft` (`inline_images`) ([#22](https://github.com/adamkopelman/outlook-mcp-rs/pull/22)).
- `is_inline` on attachments: inline `cid:` content vs. a regular attachment ([#23](https://github.com/adamkopelman/outlook-mcp-rs/pull/23)).
- `get_inline_image` tool: an attachment by Content-ID as a base64 `data:` URI ([#24](https://github.com/adamkopelman/outlook-mcp-rs/pull/24)), with optional `context_lines` (the text before the image) ([#25](https://github.com/adamkopelman/outlook-mcp-rs/pull/25)).
- A system test for PRs #14–#25, and guidance for testing a batch of PRs ([#26](https://github.com/adamkopelman/outlook-mcp-rs/pull/26)).
- Live system tests for a big calendar, real inbound mail and invites, and edge cases, plus an HTML review page of every system test's input, expected and actual output (`docs/system-tests.html`) ([#58](https://github.com/adamkopelman/outlook-mcp-rs/pull/58)).
- User-facing skills (`using-outlook-mcp`, `troubleshooting-outlook-mcp`), and README / CONTRIBUTING brought up to date ([#27](https://github.com/adamkopelman/outlook-mcp-rs/pull/27)).
- This changelog; every PR adds its entry under Unreleased ([#56](https://github.com/adamkopelman/outlook-mcp-rs/pull/56)).
- Content-writing tools share one input convention: exactly one of `body` / `html_body` / `body_file` / `html_body_file` on the mail tools, `body_file` on event, task and note tools; `categories` / `importance` on `send_email`, `create_draft`, `reply_email` and `update_draft`; `inline_images` on `reply_email` and `update_draft`; `update_draft` can `send`. ([#37](https://github.com/adamkopelman/outlook-mcp-rs/pull/37))
- `data:` image URIs in an HTML body become inline attachments on every mail-writing tool ([#37](https://github.com/adamkopelman/outlook-mcp-rs/pull/37)).
- Read tools share one set of output controls: batch ids on `get_email`, `get_event`, `get_note`, `get_task`, `list_attachments` and `get_inline_image` (per-item errors); `include`, `body_format`, `max_body_chars` with `*_truncated` / `*_length`, and `output_dir` on every body-reading tool ([#38](https://github.com/adamkopelman/outlook-mcp-rs/pull/38)).
- `get_task` tool; `get_email` `resolve_inline_images` (`cid:` → `data:`); `save_attachments` `inline` filter ([#38](https://github.com/adamkopelman/outlook-mcp-rs/pull/38)).
- One date grammar for every date parameter (ISO, keywords such as `start_of_week`, offsets such as `-14d`), and `<field>_after` / `<field>_before` range filters on every list tool ([#39](https://github.com/adamkopelman/outlook-mcp-rs/pull/39)).
- Shared `list_*` conventions: one `query` syntax (phrases, wildcards, `field:` scopes), list-valued filters, enum filters that mirror output fields (`item_type`, `importance`, `flag`), and `count` / `offset` paging on every list tool ([#39](https://github.com/adamkopelman/outlook-mcp-rs/pull/39)).
- A Hebrew round-trip test of every text field of every tool over the wire, a live mojibake diagnostic, and README troubleshooting for garbled text ([#40](https://github.com/adamkopelman/outlook-mcp-rs/pull/40)).
- `delete_event`, `delete_task` and `delete_note` take `permanent`, like `delete_email` ([#41](https://github.com/adamkopelman/outlook-mcp-rs/pull/41)).
- Every `list_*` tool takes `fields` / `compact` (a per-tool preset) to return less per row, and `output_dir` to write the rows to a JSON Lines file (`{rows_file, count}`) ([#49](https://github.com/adamkopelman/outlook-mcp-rs/pull/49)).
- Events report `user_is_required` / `user_is_optional` (matched by address and through distribution lists nested up to 10 levels, within a per-call member and time budget; `null` when unknown), `required_count` / `optional_count` and `is_canceled`, with an `is_canceled` filter ([#49](https://github.com/adamkopelman/outlook-mcp-rs/pull/49)).
- Recurring meetings: every occurrence has an `occurrence_id`; `get_event`, `respond_to_meeting`, `update_event` and `delete_event` act on just that occurrence with it ([#50](https://github.com/adamkopelman/outlook-mcp-rs/pull/50)).
- Batch writes: every `update_*` / `delete_*` tool and `respond_to_meeting` take a list of ids (the same change to each); `update_*` tools also take `updates: [...]` (a different change per item); per-item results and errors ([#51](https://github.com/adamkopelman/outlook-mcp-rs/pull/51)).
- Every `update_*` result has `undo`, the update that reverts it (pass a batch's `undo` objects as `updates` to roll it back), and `not_undoable` for what it can't revert ([#51](https://github.com/adamkopelman/outlook-mcp-rs/pull/51)).
- `require_no_attendees` on `create_event`, `update_event` and `delete_event`: the server refuses, changing nothing, to touch a meeting with attendees or add any ([#53](https://github.com/adamkopelman/outlook-mcp-rs/pull/53)).
- MCP safety hints (read-only, destructive, open-world) on every tool, and a calendar-autopilot Claude Code hook (`hooks/calendar_autopilot.py`) that auto-approves reads, silent meeting responses and guarded event changes and asks for everything else ([#53](https://github.com/adamkopelman/outlook-mcp-rs/pull/53)).
- A Claude Code plugin: `/plugin marketplace add adamkopelman/outlook-mcp-rs` installs the MCP server, both skills and the calendar-autopilot hook (off until its option is turned on) ([#55](https://github.com/adamkopelman/outlook-mcp-rs/pull/55)).

### Changed

- HTTP mode answers a request body over 64 MiB with `413` and a compressed body with `415`, instead of a generic error ([#37](https://github.com/adamkopelman/outlook-mcp-rs/pull/37)).
- A multi-word unquoted `query` matches all words in any order; a bare date in `*_before` includes that whole day; tasks and notes page at 500 ([#39](https://github.com/adamkopelman/outlook-mcp-rs/pull/39)).
- The calendar tools share one notify flag, `notify`, on `create_event`, `respond_to_meeting`, `update_event` and `delete_event` ([#53](https://github.com/adamkopelman/outlook-mcp-rs/pull/53)).
- `reminder_minutes: -1` on `create_event` / `update_event` turns the reminder off (0 still reminds at the start), so the `undo` of a reminder change can restore "no reminder" ([#58](https://github.com/adamkopelman/outlook-mcp-rs/pull/58)).

### Deprecated

- `send_email` / `create_draft` `html: bool` (use `html_body`), and `update_draft`'s `draft_id` (use `email_id`) ([#37](https://github.com/adamkopelman/outlook-mcp-rs/pull/37)).
- `get_email` `prefer_html` (use `body_format: "html"`) ([#38](https://github.com/adamkopelman/outlook-mcp-rs/pull/38)).
- `list_emails` `since_days`, `flagged`, `high_importance`, and `list_events` `start_date` / `end_date` ([#39](https://github.com/adamkopelman/outlook-mcp-rs/pull/39)).
- `send` on `create_event` / `respond_to_meeting`, `send_update` on `update_event`, `send_cancellation` on `delete_event` (use `notify`) ([#53](https://github.com/adamkopelman/outlook-mcp-rs/pull/53)).

### Fixed

- `list_emails` search with Hebrew or other non-ASCII text returned nothing; it now falls back to a client-side match ([#15](https://github.com/adamkopelman/outlook-mcp-rs/pull/15)).
- HTTP mode declares `charset=utf-8` on its responses, so clients no longer decode Hebrew as Latin-1 ([#40](https://github.com/adamkopelman/outlook-mcp-rs/pull/40)).
- Ids after a write: every tool that changes an item and leaves it in place returns its current id. `respond_to_meeting` returned none, though accepting gives the item a new one; `update_event`, `update_task` and `update_note` echoed the id they were given ([#50](https://github.com/adamkopelman/outlook-mcp-rs/pull/50)).
- Date filters on day-first Windows locales returned the wrong range: bounds are now written so no regional format can misread them, and items are filtered exactly on their real dates ([#42](https://github.com/adamkopelman/outlook-mcp-rs/pull/42)).
- `list_events` dropped recurring occurrences on the first day of a window; `check_availability` shifted every slot by the window's start time and reported "working elsewhere" as busy; batch writes returned stale ids ([#58](https://github.com/adamkopelman/outlook-mcp-rs/pull/58)).
- `check_availability` is exact on daylight-saving days: slots are real elapsed time, so a 23-hour day has 46 half-hour slots and a 25-hour day 50, each at its true local time ([#58](https://github.com/adamkopelman/outlook-mcp-rs/pull/58)).
- `create_event` / `update_event` refuse a start or end in the hour skipped by daylight saving time. It used to fail and leave a broken item that made every `list_events` over that day fail ([#58](https://github.com/adamkopelman/outlook-mcp-rs/pull/58)).
- `save_attachments` numbers attachments whose names collide (`a.txt`, `a (2).txt`) instead of overwriting the first one ([#58](https://github.com/adamkopelman/outlook-mcp-rs/pull/58)).
- Category names containing a comma are refused on every tool that writes categories: Outlook stored them as two categories ([#58](https://github.com/adamkopelman/outlook-mcp-rs/pull/58)).
- A yearly series starting on February 29 can be created ([#58](https://github.com/adamkopelman/outlook-mcp-rs/pull/58)).
- An all-day event with start = end lasts one day instead of zero, and the `undo` of an `all_day` change restores the event's times ([#58](https://github.com/adamkopelman/outlook-mcp-rs/pull/58)).
- `get_event`, `get_task`, `get_note` and their update and delete tools refuse an id of another kind of item; a task id given to `get_event` came back as an event starting 1899-12-30. `respond_to_meeting` says plainly when the id isn't a meeting invite ([#58](https://github.com/adamkopelman/outlook-mcp-rs/pull/58)).
- `check_availability` refuses `interval_minutes` below 1 (0 was read as 1) ([#58](https://github.com/adamkopelman/outlook-mcp-rs/pull/58)).

## [0.3.0] - 2026-07-18

### Added

- Optional network mode: Streamable HTTP transport with bearer-token auth (`--http`), alongside the default stdio mode.
- CONTRIBUTING.md, issue and pull request templates, and a standalone README.

## [0.2.0] - 2026-07-17

### Added

- `check_availability`: free/busy for several people and the windows when everyone is free.
- Task tools: `list_tasks` filters (category, importance, query), `update_task` (replaces `complete_task`, now `mark_complete`), `delete_task`, and more `create_task` fields.
- Note tools: `list_notes` filters, `update_note`, `delete_note`, `modified` on `get_note`, and categories/color on `create_note`.
- `list_tasks` `query` matches the task body, not only the subject.

## [0.1.1] - 2026-07-07

### Fixed

- Detail tools tolerate items that lack optional properties.

## [0.1.0] - 2026-07-07

### Added

- First release: email, calendar, attachment, task and note tools over stdio, driving classic Outlook through COM, with a fake client for tests, live-Outlook system tests, and CI release builds.

[Unreleased]: https://github.com/adamkopelman/outlook-mcp-rs/compare/v0.4.0...HEAD
[0.4.0]: https://github.com/adamkopelman/outlook-mcp-rs/compare/v0.3.0...v0.4.0
[0.3.0]: https://github.com/adamkopelman/outlook-mcp-rs/compare/v0.2.0...v0.3.0
[0.2.0]: https://github.com/adamkopelman/outlook-mcp-rs/compare/v0.1.1...v0.2.0
[0.1.1]: https://github.com/adamkopelman/outlook-mcp-rs/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/adamkopelman/outlook-mcp-rs/releases/tag/v0.1.0
