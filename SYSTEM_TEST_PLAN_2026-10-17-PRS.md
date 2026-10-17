# System test plan: open PRs #37–#56 merged together, 2026-10-17

## Purpose

Live-verify the new and changed behaviour of all 12 open PRs, **merged
together** on the integration branch `systest/all-prs-2026-10-17`, against the
real running Outlook mailbox, before landing them. This follows
`.claude/skills/live-outlook-system-test/SKILL.md` ("Testing a batch of PRs").

The user asked for the PRs to be merged once the system test passes, so this
plan was not signed off in advance. The results doc lists the judgement calls.

The PRs:

- **One linear stack:** #56 → #37 → #38 → #39 → #40 → #41 → #49 → #50 → #51 → #53 → #55. Each branch contains the previous one.
- **#42:** independent, based on #56.

None of them had been run on Windows, so the COM paths here are executing for the first time.

## Mechanism

**Through the real MCP tool layer.** This is a first for these system tests: the earlier runs called `WindowsOutlookClient` directly. `tests/system_test_prs_2026_10_17.rs` builds `OutlookMcpServer` around a real `WindowsOutlookClient`, serves it in-process, and sends every call as a JSON-RPC `tools/call` over the newline-delimited stdio codec rmcp uses (the same `WireClient` idea as `tests/tools.rs`). Every argument goes through:

- serde;
- the deprecated aliases;
- `OneOrMany`;
- the server-layer batching (`updates`, id lists), `undo`, `require_no_attendees`, `fields` / `compact` / `output_dir`;
- the JSON result.

Every response must also be strict UTF-8.

    cargo test --test system_test_prs_2026_10_17 -- --ignored --nocapture

One `#[ignore]`d, non-panicking test. Every step records pass/fail in a log, cleanup always runs, and the only `assert!` is the last line.

**Second pass:** the PRs' own 28 new `#[ignore]`d tests in `tests/live_outlook.rs`, plus `permanent_delete_of_draft_skips_deleted_items` as a regression check on #16.

## Accounts and real-world traffic

- **Mailbox:** `adamkopelman@outlook.com`, the only account in the profile. Regional format en-US (`M/d/yyyy`, week starts Sunday).
- **One real send (W4):** `update_draft` with `send: true` on a draft addressed to the mailbox itself. It is the self-loop address used by every earlier system test, and the success path of #37's new `send` has no other live test.
- **No other mail is sent.**
  - Every other email item is a draft to `nobody@example.invalid`.
  - Every event is a personal appointment, except the one meeting in S3. That meeting has the attendee `nobody@example.invalid`, is created with `notify: false` (saved, never sent), and is deleted with `notify: false`.
  - The `require_no_attendees` refusal tests also pass `notify: false`, so even a broken refusal cannot send anything.

## Tagging / cleanup

- **Subject tag:** `[outlook-mcp-rs systest 1017]` plus a per-run token `r<unix-secs>`.
- **Dates:** events and tasks use dates in 2099.
- **Cleanup, in order:**
  1. Every id the test created is deleted through the tools with `permanent: true`.
  2. A sweep over Drafts, Inbox, Sent Items and Deleted Items (`list_emails` with the run token), and over events, tasks and notes (the `list_*` tools), repeated until a pass finds nothing. The sweep deletes permanently too.
- **Ground truth:** a raw PowerShell COM count of Inbox, Drafts, Sent, Deleted Items, Outbox, Calendar, Tasks and Notes, compared with the baseline taken before the first run: 48 / 12 / 213 / 53 / 0 / 19 / 4 / 7.
- **`empty_deleted_items(confirm=true)` is not run.** Deleted Items holds the user's own items.
- **Categories:** the mailbox's default palette (Red / Blue / Green Category).

## Seed data (phase 0)

| label | kind | details |
|---|---|---|
| TA | task | due 2099-03-10, high, Red Category, Hebrew body |
| TB | task | due 2099-03-20 09:00, normal, Blue Category |
| TC | task | no due date, low |
| NA | note | "… alpha note" |
| NB | note | Hebrew text |
| EA | event | 2099-04-13 10:00–10:30, busy, Red Category |
| EB | event | 2099-04-14 11:00–11:30, free, Blue Category |
| EX | event | 2099-03-04 09:00 (day ≤ 12: misread bait) |
| EY | event | 2099-04-03 09:00 (EX with day and month swapped) |
| ER | event | daily series, 2099-04-20 08:00, 5 occurrences |
| DA | draft | subject "… draft alpha beta" |
| DB | draft | subject "… draft beta alpha" (same words, reversed) |

## Tests (ID → expected)

### #37: content-writing tools (W)

- **W1** `create_draft` with `html_body_file` (Hebrew plus a `data:image/png` URI), `importance: high` and `categories`:
  - `get_email` HTML contains `cid:img-`, not `data:`;
  - `list_attachments` has one hidden `img-*` image whose bytes decode to the PNG;
  - importance and categories are set.
- **W2** `create_task`, `create_note` and `create_event` from a Hebrew `body_file` → `get_task` / `get_note` / `get_event` body contains the exact text.
- **W3** `body` + `body_file` together → error, and nothing is created.
- **W4** `update_draft` with `send: true` on a self-addressed draft → `status: sent`, and the copy arrives in Inbox within 120 s.
- **W5** deprecated forms: `create_draft` with `html: true` + `body` stores HTML; `update_draft` with `draft_id` works.

### #38: read tools (R)

- **R1** `get_task` with `[TA, bogus]` → a list in order, the second item is `{id, error}`; `include: []` → no body.
- **R2** `get_event` with `[EA, EB]` → 2 results; `max_body_chars: 10` on W2's event → `body_truncated: true` and a `body_length`.
- **R3** `get_note(NB, output_dir)` → `body_file` exists, and its contents equal the body.
- **R4** `get_email(body_format: "text", prefer_html: true)` → error naming both.

### #39: list conventions (L)

- **L1** `list_tasks` with `due_after 2099-03-01, due_before 2099-03-15` → {TA}; with `due_before 2099-03-20` (a bare date covers the whole day) → {TA, TB}.
- **L2** `importance: ["high","low"]` → {TA, TC}; `category: ["Red Category","Blue Category"]` → {TA, TB}.
- **L3** `list_events` over 2099-04-01..30 with `show_as: "free"` → {EB}; `category: ["Red Category"]` → {EA}.
- **L4** Hebrew `query` on `list_notes` → {NB} (client-side match).
- **L5** `list_notes` with `created_after: "-1h"` → contains NA and NB.
- **L6** errors: `start_after` later than `start_before`; `importance: "urgent"`; `item_type: "fax"`; `since_days` contradicting `received_after`.
- **L7** deprecated aliases: `list_events` `start_date` / `end_date`; `list_emails` `since_days: 1` + `flagged: false` run.
- **L8** wildcard on Drafts: `subject:alpha*beta` → {DA} and not DB. This checks the DASL fix plus the client-side order check.
- **L9** paging: `list_tasks` with `count: 1` at `offset` 0 / 1 / 2 → disjoint pages whose union is {TA, TB, TC}.

### #40: UTF-8 (U)

- **U1** the Hebrew subjects and bodies of TA, NB and W2's items come back over the wire byte-exact. The `WireClient` rejects any response that isn't UTF-8.

### #41: permanent delete (D)

- **D1** `delete_task` with ids `[D-task1, D-task2]` and `permanent` → a list of 2 `{permanent: true}`; neither id is found in Deleted Items or Drafts afterwards.
- **D2** `delete_note` / `delete_event` with `permanent` → the same checks.

### #42: locale-proof date filters (Q)

- **Q1** 2099-03-01..2099-03-10 → contains EX, not EY.
- **Q2** 2099-04-01..2099-04-05 → contains EY, not EX.
- **Q3** `start_after: "2099-04-13T10:30"` excludes EA (10:00), and `start_before: "2099-04-14"` includes EB (11:00).

### #49: output controls (O)

- **O1** `list_tasks` with `fields: [subject, due_date]` → each row has exactly those keys; an unknown field → error listing the valid ones.
- **O2** `list_events` with `compact: true` → rows include `id`, `subject` and `start`, and not `required_attendees`; with `output_dir` → `{rows_file, count}`, where the file has `count` JSON lines.
- **O3** EA (personal) → `user_is_required: true`, `required_count: 0`.
- **O4** `list_folders` with `compact` and `list_attachments` with `fields` run.

### #50: ids after a write, occurrences (I)

- **I1** the `id` returned by `update_task` / `update_note` / `update_event` opens the item with `get_*`.
- **I2** ER over its 5 days → 5 rows, each with a distinct `occurrence_id`; `get_event(occurrence_id)` → that start.
- **I3** `update_event` on occurrence 2's subject → only that row changes; `update_event` with `occurrence_id` + `recurrence` → error.
- **I4** `delete_event(occurrence 3)` → 4 rows remain.

### #51: batch writes and undo (B)

- **B1** `update_task` with `[TA, TB]` and `add_categories: [Green Category]` → 2 results, each with `undo`; passing the undos as `updates` → the categories are restored.
- **B2** `update_event` with `updates` holding different changes for EA and EB, plus a bogus id → per-item results, and the bogus id becomes `{id, error}`.
- **B3** `update_email(DA, flag follow_up, importance high)`, then its `undo` → flag and importance restored.
- **B4** `update_note(NA, color pink)`, then its `undo` → previous color.
- **B5** 101 ids → error, and nothing changes.

### #53: calendar safety (S)

- **S1** `create_event` with an attendee, `require_no_attendees: true` and `notify: false` → error; no event with that subject exists.
- **S2** `update_event(EA, require_no_attendees)` → ok; adding an attendee with `require_no_attendees` → error, and EA is still not a meeting.
- **S3** a meeting saved with `notify: false`, then `delete_event(require_no_attendees)` → refused, and the meeting still exists; `delete_event(notify: false, permanent)` → gone.
- **S4** `tools/list`: the `readOnlyHint` set is exactly the `list_*`, `get_*` and `check_availability` tools; every tool that sends or notifies is `openWorldHint`.
- **S5** `update_event(EA, send_update: false)` is accepted as the alias of `notify`.

### #55: plugin (P)

Not in the Rust test. The plugin's hook launcher was run manually under Git Bash with real hook events, with the mode on and off; see the results doc.

### Regression smoke (X)

- **X1** `list_folders`, `list_emails` (inbox, count 5) and `check_availability` for the mailbox itself over 2099-04-13 run, and the rows have their usual keys.

## Explicit skips

- **`respond_to_meeting`**, including #53's `notify: false` and #50's new id: answering needs an invite from an independently controlled second mailbox. Answering a real third party's invite to test it is off-limits. #53 auto-approves `notify: false`, so this is flagged as an open risk.
- **Day-first locales for #42:** changing the user's Windows Region and restarting Outlook is disruptive. This run covers en-US only, plus the swapped-date bait (Q1/Q2).
- **`empty_deleted_items(confirm=true)`:** never run live.
- **The plugin installed in a real Claude Code session:** needs a session restart.
