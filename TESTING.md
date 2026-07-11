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
  created in that mailbox — most are cleaned up automatically, but the
  calendar event test currently requires manual deletion afterward (there's
  no `delete_event` tool; see `tests/live_outlook.rs` for why).

Run them with:

```
cargo test --test live_outlook -- --ignored
```

## Manual-only tests (not automated at all)

`send_email` and `respond_to_meeting` have real, unrecoverable side effects
(an actually-delivered email; an actual meeting response sent to an
organizer) and are not covered by any automated test. To verify them by
hand before a release:

1. Pick a test recipient you control (e.g. a second mailbox of your own).
2. Call `send_email` with that recipient and a clearly-marked test subject;
   confirm it arrives.
3. Find (or create) a meeting invite in your test mailbox and call
   `respond_to_meeting` with `response: "tentative"`; confirm the organizer
   sees a tentative response.

`update_email`'s `flag` field (`follow_up`/`complete`/`clear`) is also
manual-only. The automated live test (`update_email_applies_state_then_moves`)
exercises `mark_read`, `add_categories`, `importance`, and `move_to` against a
disposable draft, but not `flag`: `MarkAsTask` is only valid on items that have
been *sent or received*, and Outlook rejects it on a draft. To verify by hand:

4. Pick a received email in your test mailbox and call `update_email` with
   `flag: "follow_up"`; confirm a follow-up flag appears. Repeat with
   `"complete"` (flag shows complete) and `"clear"` (flag removed).
