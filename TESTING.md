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

`list_tasks` filters (`category`, `importance`, `query`, `include_completed`),
`create_task`'s additions (`categories`, `start_date`, `reminder_time`), and
`update_task`/`delete_task` (which retired the standalone `complete_task`
tool — `mark_complete: true`/`false` on `update_task` now covers completing
*and* reopening a task) are covered by the live suite:
`cargo test --test live_outlook -- --ignored list_tasks_filters_and_create_task_additions_round_trip update_task_marks_complete_then_reopens delete_task_removes_it`.

`list_notes` filters (`category`, `query` — the latter matching the note's
real body text, not just a derived subject), `create_note`'s additions
(`categories`, `color`), `get_note`'s `modified` field, and
`update_note`/`delete_note` are covered by the live suite:
`cargo test --test live_outlook -- --ignored list_notes_filters_and_create_note_additions_round_trip get_note_includes_modified_after_update update_note_manages_categories_and_color delete_note_removes_it`.

`check_availability`'s single-mailbox path (resolving your own address and
reading its free/busy slots, plus the graceful-failure path for an address
that can't provide free/busy data) is covered by the live suite:
`cargo test --test live_outlook -- --ignored check_availability`. Checking
a real second person's free/busy (someone outside this mailbox) is a manual
check, since Outlook must actually have published free/busy for that
account, which can't be arranged from an automated test. To verify by hand,
call `check_availability` with a colleague's email address in `people` and
confirm their slots reflect their real calendar.
