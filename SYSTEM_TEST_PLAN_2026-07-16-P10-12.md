# System test plan — Plans 10–12 (check_availability, Tasks, Notes), 2026-07-16

## Purpose

Live-verify every tool built across Plans 10–12 (the three plans shipped
after the earlier Plans 1–9 system test), against the real, running Outlook
mailbox `adamkopelman@outlook.com`. Per `.claude/skills/live-outlook-system-test/SKILL.md`.

## Scope

- **Plan 10:** `check_availability`.
- **Plan 11 (Tasks):** `list_tasks` (filters), `create_task` (+ additions), `update_task`, `delete_task`.
- **Plan 12 (Notes):** `list_notes` (filters), `get_note` (+ modified), `create_note` (+ additions), `update_note`, `delete_note`.

Not in scope: Plans 1–9 (already system-tested 2026-07-16, see
`SYSTEM_TEST_PLAN_2026-07-16.md`/`SYSTEM_TEST_RESULTS_2026-07-16.md`).

## Mechanism

Same as the earlier system test: direct `WindowsOutlookClient` calls via a
new `#[ignore]`d test in `tests/system_test_p10_12.rs`, not actual MCP tool
calls (this session's MCP tool schemas were fixed at session start and can't
pick up the newly-rebuilt binary without a restart).

## Real-world traffic

None. Every tool in this scope (`check_availability`, all Task/Note CRUD) is
either read-only or self-contained (no email sent, no invite sent, nothing
delivered to any third party). No external authorization needed beyond what
was already given for Plans 1–9.

## Tagging / cleanup

All test data tagged `[outlook-mcp-rs systest P10-12]`, far-future dates
(`2099-xx-xx`) for anything date-bearing. Every created task/note is deleted
by its own test via `delete_task`/`delete_note`.

## Seed data

Not needed at this scale (a handful of tools, not a full filter-matrix sweep
like the Plans 1–9 test) — each test seeds its own 1–3 items inline and
asserts against them directly.

## Tests

### C1. `check_availability` — own mailbox
Call with `people: [self address]`, a 2-hour far-future window,
`interval_minutes: 30`. Expect: `resolved: true`, 4 slots, each a valid
status word, `common_free` non-empty (nothing scheduled that far out).

### T1. `list_tasks` filters
Create 2 tasks: one high-importance with a category and a distinctive
subject word, one plain. Filter by `category`, `importance`, and `query`
separately; confirm each isolates the right task. Clean up both via
`delete_task`.

### T2. `create_task` additions
Create a task with `categories`, `start_date`, `reminder_time`. Confirm via
`list_tasks` that it exists (categories echoed back). Clean up.

### T3. `update_task` — mark complete, reopen, field edits
Create a task. `update_task` with `mark_complete: true`; confirm via
`list_tasks(include_completed: true)` it's complete. `update_task` with
`mark_complete: false`; confirm reopened. One more `update_task` call
editing `subject`/`importance`/`add_categories` together; confirm `changed`
lists all three. Clean up.

### T4. `delete_task`
Create a task, delete it, assert `status == "deleted"`.

### N1. `list_notes` filters
Create 2 notes: one with a category and a distinctive body word, one plain.
Filter by `category` and `query` separately; confirm each isolates the right
note. Clean up both via `delete_note`.

### N2. `create_note` additions + `get_note` modified
Create a note with `categories` and `color`. `get_note` it; confirm
categories present. Capture `modified`, then `update_note` its body, then
`get_note` again; confirm `modified` is non-decreasing and body changed
(mirrors the skill's "baseline before action" lesson). Clean up.

### N3. `update_note` — categories and color
Create a note. `update_note` with `add_categories` and `color`; confirm
`changed` lists both, and a follow-up `get_note` shows the category.
Then `remove_categories` it back off; confirm removed. Clean up.

### N4. `delete_note`
Create a note, delete it, assert `status == "deleted"`.

## Cleanup checklist

- [ ] Every task/note created above deleted by its own test.
- [ ] Final raw-COM sweep of Tasks/Notes folders confirming no
      `[outlook-mcp-rs systest P10-12]`-tagged item remains.

## Reporting

Results doc `SYSTEM_TEST_RESULTS_2026-07-16-P10-12.md`: pass/fail per test
id, root-caused writeup for any failure, confirmed cleanup state.
