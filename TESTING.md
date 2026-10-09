# Testing outlook-mcp-rs

## Unit tests

```
cargo test
```

Runs everything except the live suite (`tests/live_outlook.rs`, all
`#[ignore]`d) — no Outlook required, safe to run anywhere, and what CI runs
on every push.

## Live system tests

These exercise the real `WindowsOutlookClient` against Outlook actually
running on your machine. Preconditions:

- Windows, with classic Outlook desktop installed
- Outlook is open and signed in to a normal mailbox
- You're comfortable with a handful of test items (a draft, a task, a note,
  a calendar event, each clearly named "outlook-mcp-rs live test ...") being
  created in that mailbox — every live test cleans up after itself
  (calendar events via `delete_event`, since Plan 8).

Run them with:

```
cargo test --test live_outlook -- --ignored
```

## Full system tests (one run across many tools)

Besides the per-feature live tests above, there are larger system tests. Each
one seeds tagged test data, checks many tools in a single run, and always
cleans up. Like the live tests, they are `#[ignore]`d and need a running
Outlook:

| Test | Covers | Plan / results |
|---|---|---|
| `cargo test --test system_test -- --ignored --nocapture` | Plans 1–9 (email + calendar) | `SYSTEM_TEST_PLAN_2026-07-16.md`, `SYSTEM_TEST_RESULTS_2026-07-16.md` |
| `cargo test --test system_test_p10_12 -- --ignored --nocapture` | Plans 10–12 (availability, tasks, notes) | `SYSTEM_TEST_PLAN_2026-07-16-P10-12.md`, `SYSTEM_TEST_RESULTS_2026-07-16-P10-12.md` |
| `cargo test --test system_test_prs -- --ignored --nocapture` | PRs #14–#25: `list_emails` offset/`to`/non-ASCII search, `get_email` truncation flags, `update_draft`, permanent delete, attachment metadata, inline images, `get_inline_image` | `SYSTEM_TEST_PLAN_2026-10-03-PRS.md`, `SYSTEM_TEST_RESULTS_2026-10-03-PRS.md` |

`system_test_prs` sends **one** real email, a self-loop to the mailbox
owner's own address (`SELF_ADDR` in the test; change it before running on
another mailbox). Everything else is a draft addressed to
`nobody@example.invalid`. It never calls `empty_deleted_items(confirm=true)`.

How these are planned, run, and root-caused is described in the
`live-outlook-system-test` skill (`.claude/skills/live-outlook-system-test/SKILL.md`).

## Manual-only tests (not automated at all)

`send_email`, `respond_to_meeting`, and `create_event` with `send: true` have
real, unrecoverable side effects (an actually-delivered email; an actual
meeting response sent to an organizer; an actual meeting invite sent to real
attendees) and are not covered by any automated test. To verify them by hand
before a release:

1. Pick a test recipient you control (e.g. a second mailbox of your own).
2. Call `send_email` with that recipient and a clearly-marked test subject;
   confirm it arrives.
3. Find (or create) a meeting invite in your test mailbox and call
   `respond_to_meeting` with `response: "tentative"`; confirm the organizer
   sees a tentative response.
4. Call `create_event` with `send: true`, real attendees in `required_attendees`
   and/or `optional_attendees`, and your own email as a recipient; confirm the
   invite arrives in their mailbox. (With `send: false`, the event is saved
   without sending, so the attendee addresses are not required to be real —
   this is covered by the automated test `create_event_with_tiers_categories_and_show_as`.)

5. Call `update_event` on a meeting you organize with `send_update: true` and
   real attendees; confirm they receive the update email. Call `delete_event`
   on a meeting you organize with `send_cancellation: true`; confirm they
   receive the cancellation. (The automated live test
   `update_event_edits_fields_and_manages_attendees` uses placeholder
   attendee addresses with `send_update: false`, so nothing is ever
   delivered — this is why real-recipient delivery still needs a manual check.)

`update_draft` with `send: true` actually sends the draft, so its success
path is manual-only: create a draft addressed to a recipient you control, call
`update_draft` with `send: true` (optionally with a change such as a new
`subject`), and confirm it arrives and leaves Drafts. The automated live test
`update_draft_send_without_recipients_saves_and_refuses` covers only the
refusal path (a draft with no recipients is saved, not sent).

The content-writing conventions (`html_body`/`*_file` inputs, `data:` URI
images converted to inline attachments, `inline_images` on `update_draft` and
`reply_email`, categories and importance on drafts) are covered by drafts-only
live tests:
`cargo test --test live_outlook -- --ignored create_draft_turns_data_uri_into_inline_attachment_with_metadata update_draft_adds_inline_images_categories_and_importance_without_duplicates update_draft_send_without_recipients_saves_and_refuses reply_email_draft_with_data_uri_image`.
`reply_email_draft_with_data_uri_image` replies (as an unsent draft, then
deleted) to the newest Inbox item, and skips if the Inbox is empty.

`update_email`'s `flag` field (`follow_up`/`complete`/`clear`) is also
manual-only. The automated live test (`update_email_applies_state_then_moves`)
exercises `mark_read`, `add_categories`, `importance`, and `move_to` against a
disposable draft, but not `flag`: `MarkAsTask` is only valid on items that have
been *sent or received*, and Outlook rejects it on a draft. To verify by hand:

4. Pick a received email in your test mailbox and call `update_email` with
   `flag: "follow_up"`; confirm a follow-up flag appears. Repeat with
   `"complete"` (flag shows complete) and `"clear"` (flag removed).

`list_events` with `calendar_of` pointing to **another user's calendar**
requires that user to have granted you calendar-sharing permission; this
setup cannot be automated in a test suite. The automated live test
(`list_events_calendar_of_self_opens_own_calendar`) exercises the
recipient-resolve + GetSharedDefaultFolder path by opening your own
calendar; to verify cross-user sharing works, call `list_events` with a
colleague's email address in `calendar_of` (one who has shared their
calendar with you) and confirm it returns their events without error.

`list_emails`'s `query` filter matching real email body text (not just
subject/sender) is covered by the live suite:
`cargo test --test live_outlook -- --ignored list_emails_query_matches_real_body_text`.

`list_tasks` filters (`category`, `importance`, `query`, `include_completed`),
`create_task`'s additions (`categories`, `start_date`, `reminder_time`), and
`update_task`/`delete_task` (which retired the standalone `complete_task`
tool — `mark_complete: true`/`false` on `update_task` now covers completing
*and* reopening a task) are covered by the live suite:
`cargo test --test live_outlook -- --ignored list_tasks_filters_and_create_task_additions_round_trip update_task_marks_complete_then_reopens delete_task_removes_it`.
`list_tasks`'s `query` filter matching real task body text (not just
subject) is covered separately by:
`cargo test --test live_outlook -- --ignored list_tasks_query_matches_real_body_text`.

`list_notes` filters (`category`, `query` — the latter matching the note's
real body text, not just a derived subject), `create_note`'s additions
(`categories`, `color`), `get_note`'s `modified` field, and
`update_note`/`delete_note` are covered by the live suite:
`cargo test --test live_outlook -- --ignored list_notes_filters_and_create_note_additions_round_trip get_note_includes_modified_after_update update_note_manages_categories_and_color delete_note_removes_it`.

The shared `list_*` conventions (date grammar on `received_*`, `item_type`,
`flag`, scoped/phrase/wildcard `query`, and `count`/`offset` paging on events,
tasks and notes) are covered by read-only live tests:
`cargo test --test live_outlook -- --ignored list_emails_relative_dates_bound_received_time list_emails_item_type_filter_matches_get_email list_emails_flag_states_partition_the_folder list_emails_subject_scope_and_phrase_find_a_recent_subject list_events_tasks_notes_pages_tile_without_overlap`.
On a day-first Windows locale the date test can still fail because of issue #1
(`jet_datetime`), which is independent of the grammar.

`check_availability`'s single-mailbox path (resolving your own address and
reading its free/busy slots, plus the graceful-failure path for an address
that can't provide free/busy data) is covered by the live suite:
`cargo test --test live_outlook -- --ignored check_availability`. Checking
a real second person's free/busy (someone outside this mailbox) is a manual
check, since Outlook must actually have published free/busy for that
account, which can't be arranged from an automated test. To verify by hand,
call `check_availability` with a colleague's email address in `people` and
confirm their slots reflect their real calendar.
