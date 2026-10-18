# System test plan: edge cases, 2026-10-18

## Purpose

A review of `docs/system-tests.html` against the code listed edge cases that the earlier suites never touch: colliding attachment names, commas in category names, DASL wildcard characters in queries, all-day and DST boundaries, odd availability intervals, recurrence edge cases, permanent deletes outside the happy path, paging under mutation, non-default folders and `calendar_of`, attendee no-ops, size limits, flag round trips, and stale or malformed ids. Each test below names the hypothesis it checks. The hypotheses come from reading the code and are unconfirmed; the point of this suite is to confirm or refute them live.

## Mechanism

`tests/system_test_edge.rs`, through the real MCP tool layer (`tests/common`: `OutlookMcpServer` over a real `WindowsOutlookClient`, JSON-RPC over the stdio codec), like `system_test_big_calendar`. One `#[ignore]`d, non-panicking test; cleanup always runs.

    cargo test --test system_test_edge -- --ignored --nocapture

A test whose hypothesis is confirmed records `[FAIL]` with the observed behaviour. This suite does not fix the bugs it finds; they are listed in the results doc for a separate fix.

## Traffic

- One self-loop email to the mailbox owner (`SELF`, as in `system_test_prs_2026_10_17`), used for `reply_email` with `reply_all` (G3) and the received-mail tests. The reply-all also goes only to `SELF`.
- Everything else is a draft to `nobody@example.invalid`, a personal appointment, or an unsent meeting saved with `notify: false` to `example.invalid` attendees.

## Tagging / cleanup

- **Tag:** `[outlook-mcp-rs systest edge]` plus a run token. Calendar items are in 2099; DST tests use the 2099 transition dates of the machine's own time zone.
- **Cleanup:** delete every created id permanently, then sweep Inbox, Drafts, Sent Items, Deleted Items, Archive, the test subfolder, Calendar, Tasks and Notes by the run's exact tag, looping until a pass finds nothing. Remove the test subfolder and the local `save_attachments` directory. Confirm with a raw COM count.

## Seed data

- **Mail:** drafts whose subjects contain `50% off`, `a_b`, `axb` (negative for `a_b`), `[draft] x`, `O'Brien`, `say "hi"`, `back\slash`, `star*lit`. One self-loop email received in the Inbox.
- **Attachments draft:** two attachments named `a.txt` with different bytes, `a:b.txt`, `a_b.txt`, `CON.txt`, and one named `...`.
- **Events/tasks:** items with `categories: ["Smith, John", "Red Category"]`.
- **Folder:** an Inbox subfolder `systest edge תיקייה` (space and Hebrew in the name).

## Tests (ID → expected)

### Attachments with clashing names
- **AT1** `save_attachments` on the attachments draft to a fresh directory → six distinct files, each with the exact bytes of its attachment (the two `a.txt` are not merged). Hypothesis: `client.rs:2848` joins `safe_filename(name)` with no de-duplication, so the second `a.txt` overwrites the first while both are reported saved.
- **AT2** the same run, `a:b.txt` and `a_b.txt` → two distinct files. Hypothesis: `safe_filename` (`com.rs:25`) maps `:` to `_`, so they collide.
- **AT3** `CON.txt` and `...` → each saved under a usable name, or a clear per-file error; never a reported save of a file that is not on disk.

### Categories containing a comma
- **CT1** `create_event` and `create_task` with `categories: ["Smith, John", "Red Category"]`, then `get_event` / `get_task` → exactly those 2 categories. Hypothesis: `parse_categories` (`com.rs:37`) splits on `,`, so the name comes back as `Smith` and `John`.
- **CT2** `list_events` and `list_tasks` with `category: "Smith, John"` → the seeded item.
- **CT3** `update_event` / `update_task` with `remove_categories: ["Smith, John"]` → only `Red Category` is left.

### Wildcard and quote characters in queries
- **Q1** `list_emails` `query: "50% off"` → exactly the `50% off` draft. Hypothesis: `text_query.rs:150` escapes only `'`, so `%` is a wildcard.
- **Q2** `query: a_b` → the `a_b` draft, not `axb`. Hypothesis: `_` is a single-character wildcard in `LIKE`.
- **Q3** `query` for `[draft] x`, `O'Brien`, `say "hi"` (quoted phrase), `back\slash` → exactly one match each.
- **Q4** `query: "star*lit"` as a quoted phrase → the `star*lit` draft only (the `*` is literal inside quotes, or the behaviour is documented).
- **Q5** `query: "subject:"` (field with an empty value) and a bare `*` → a clear error or a documented result, never a crash.
- **Q6** `list_emails` with `from: "o'brien"` and `to: "50%"` → no DASL error and no match-everything result. Hypothesis: the `from`/`to` paths (`client.rs:1786`, `1863`) have no wildcard handling.
- **Q7** `list_events` `query: "50% off"` on an event with that subject, and an event `axb` → only the `50% off` event.

### All-day and multi-day boundaries
- **AD1** `create_event` `all_day: true`, `start = end = 2099-07-01T00:00` → `get_event` shows an all-day event on 07-01 only, with the end at 07-02 00:00 (exclusive).
- **AD2** `all_day: true` with a timed `10:00–11:00` range → normalised to the whole day, or a clear error. Hypothesis: `client.rs:2426` sets `Start`/`End` before `AllDayEvent`.
- **AD3** a 3-day all-day event 07-10..07-12 → listed on 07-10, 07-11 and 07-12 single-day windows; not on 07-09 or 07-13.
- **AD4** `update_event(all_day: false)` then `all_day: true`, and the `undo` of each → the start, end and flag return to their earlier values. Hypothesis: `client.rs:2580` toggles the flag with no time normalisation.
- **AD5** an all-day daily series of 5 occurrences → 5 all-day rows, one per day.

### Daylight saving time
- **DST1** find the local zone's 2099 spring-forward and fall-back dates (`[TimeZoneInfo]::Local.GetAdjustmentRules()`; if the zone has no DST, record every DST test as `[PASS] ... skipped: zone has no DST`). Create an event that starts in the skipped hour and one in the repeated hour → `get_event` returns the times as Outlook stored them; the result doc records them.
- **DST2** a daily series across each transition → one occurrence per day, distinct `occurrence_id`s, the same local wall-clock time on both sides.
- **DST3** `check_availability` for yourself in 30-minute slots over each transition day → the slot count matches the day's real length (23 or 25 hours) or the documented behaviour, and the seeded event's slots are busy at the right times. Hypothesis: the free/busy parsing assumes 1440 minutes a day.
- **DST4** `list_events` over each transition day → the seeded events, once each.

### `check_availability` intervals and windows
- **AV1** `interval_minutes` 0, 7, 45 and 1440 over a window holding a known busy event → 0 is rejected or clamped as documented; 7 and 45 give `ceil(window / interval)` slots with the busy slots in the right place; 1440 gives one slot per day. Hypothesis: `client.rs:2744` only clamps to a minimum of 1, and intervals that do not divide a day may index the free/busy string wrongly.
- **AV2** a window 23:30 → 00:30 across midnight, and a 3-day window starting at 23:30 → correct slot count and statuses on both sides of midnight.
- **AV3** a window more than 40 days out → correct statuses, or a clear error if it is past the free/busy horizon.
- **AV4** two people, yourself and `nobody@example.invalid` → your slots as seeded; the unresolved person reported as unresolved, not as free.

### Recurrence edge cases
- **RC1** weekly on Monday and Wednesday with `interval: 2` → occurrences only on those days of every other week.
- **RC2** monthly on `day_of_month: 31` from 2099-01-31 for 4 occurrences → the dates Outlook produces (31 Jan, end of Feb, 31 Mar, …) are recorded and none is skipped or duplicated.
- **RC3** yearly starting 2096-02-29 for 3 occurrences → the dates recorded; no error.
- **RC4** `until` before `start` → a clear error; nothing is created. `occurrences: 0` and `occurrences: -1` → a clear error. Hypothesis: `apply_recurrence` (`client.rs:330`) has no guard for these.
- **RC5** `update_event` switching a weekly series to monthly → the new pattern reads back; occurrences follow it.
- **RC6** `list_events` over a range that crosses the series' end → occurrences stop at the end date.

### Permanent delete edge cases
- **PD1** `delete_email(permanent: true)` on a tagged draft that is already in Deleted Items → it is gone from Deleted Items.
- **PD2** `permanent: true` twice on the same id → the second call returns a clear "not found" error, not a success.
- **PD3** permanent delete of a draft and of an event whose subjects contain `'` → both gone, no leftovers in Deleted Items or Drafts.
- **PD4** `delete_event(permanent: true)` with the `occurrence_id` of one occurrence → the occurrence is gone; the result does not claim a permanent delete if the code ignores the flag. Hypothesis: the occurrence branch (`client.rs:2712`) ignores `permanent` silently.
- **PD5** permanent delete of an email in the test subfolder → gone from the subfolder and from Deleted Items.
- **PD6** `delete_event` with a stale id (a permanently deleted event) → a clear error.
- **PD7** record the number of same-class items in Deleted Items at the start. If it is over 500, permanently delete one tagged event and check it is purged. Hypothesis: `soft_delete_then_purge` scans at most 500 candidates. Do not add items to Deleted Items to reach 500; if the count is lower, record `[PASS] ... not reachable: N items`.

### Deleting a meeting you organize
- **OD1** an unsent meeting (`notify: false`, `example.invalid` attendee) deleted with `send_cancellation: false`, `permanent: false` → nothing in Sent Items or the Outbox; the item in Deleted Items, and `list_events` with `is_canceled: true` does not show it in the calendar.
- **OD2** delete one occurrence of an organized series the same way → the series stays, the occurrence is gone, and the remaining occurrences are not marked canceled. Hypothesis: `client.rs:2707` writes `MeetingStatus = Canceled` before deleting.

### Paging under mutation
- **PG1** 60 tagged drafts; page `list_emails` in 25s; between pages 1 and 2 delete 5 items from page 1 → no error; the shift matches offset semantics (5 rows skipped), recorded.
- **PG2** the same two identical calls back to back with no change → identical rows.
- **PG3** the same for `list_events` over 60 tagged events.

### Non-default folders and `calendar_of`
- **FO1** `update_email(move_to: <subfolder>)` on a tagged draft, then `list_emails(folder: <subfolder>)` → the draft is there under a working id; `get_email` opens it.
- **FO2** move it back to Drafts → found in Drafts, no longer in the subfolder.
- **FO3** `list_folders` → the subfolder is listed under the Inbox with its Hebrew name exact.
- **FO4** `list_events(calendar_of: <SELF>)` over a day with a seeded event → the event is listed.
- **FO5** `list_events(calendar_of: "nobody@example.invalid")` → a clear error.

### Attendee edge cases
- **AE1** an unsent meeting; `update_event(add_attendees: [x@example.invalid, x@example.invalid], notify: false)` → one recipient for `x`.
- **AE2** add an attendee who is already present, in different case (`X@Example.Invalid`) → still one recipient.
- **AE3** `remove_attendees` for someone not on the meeting → a clear message or a documented no-op; the attendees are unchanged.

### Size limits
- **L1** a draft with a 1 MB body; `get_email` with `max_body_chars` 999, 1000 and 5,000,001 → each clamped or rejected as the tool text documents; 1000 gives exactly 1000 characters and a truncation flag.
- **L2** a 300-character subject → reads back exactly.
- **L3** a draft with a 12 MB inline image; `get_inline_image` → a clear "too large" error (the limit is 10 MB), not a crash or a truncated image.
- **L4** `get_email` with 201 ids → the batch-limit error. With 200 ids (tagged drafts) → 200 results.

### Flags and read state
- **FL1** the received self-loop email: `update_email` flag `complete` → `clear` → `follow_up`; after each, `list_emails` with the matching `flag` finds it and the others do not.
- **FL2** `mark_unread` → `list_emails(unread_only: true)` finds it; mark read → it does not.
- **FL3** with the flagged email not the newest in the Inbox, `list_emails(flag: follow_up, count: 1)` → returns it. Hypothesis: the flag filter (`client.rs:1795`) runs after `count`, so the newest unflagged item uses up the count and the result is empty.

### Stale and malformed ids
- **ID1** `get_email`, `update_email`, `delete_email` with an empty id → a clear error each.
- **ID2** the same three tools with the id of a permanently deleted item → a clear error each.
- **ID3** a task id passed to `get_event`, and an event id to `get_task` → a clear error, never the wrong item.
- **ID4** `get_event` with an `occurrence_id` whose date is not in the series → a clear error.
- **ID5** `respond_to_meeting` with a draft id → a clear error.

### Gaps in earlier suites
- **G1** `get_email` with `include: ["meeting"]` on an organized unsent meeting's item, and on a plain draft → meeting fields for the first; none (or a documented empty value) for the second.
- **G2** `list_folders` with nested folders → the subfolder appears with its parent path.
- **G3** `reply_email(reply_all: true)` on the self-loop email → the reply is in Sent Items addressed only to `SELF`.
- **G4** `create_event` with `reminder_minutes: 15`, then `get_event` → 15; update to 0 → no reminder.
- **G5** `list_emails(received_before: <self-loop received time + 1 min>)` → includes it; a minute before → excludes it.
- **G6** `list_events` with `importance: high` on one high and one normal seeded event → only the high one.
- **G7** `list_events` with `attendees` and `attendee_role: optional` on a meeting with one required and one optional `example.invalid` attendee → it matches the optional attendee, not the required one.
- **G8** `update_email(move_to: Archive)` on a tagged draft, then `list_emails(folder: Archive)` → found there (A10 in the 07-16 plan never checked the target folder).
- **G9** an item attachment (an email attached to a draft) → `get_email` lists it with its type; `save_attachments` saves it or reports it as unsupported.

### Cleanup
- **CLEAN** delete by id permanently, sweep every touched folder by tag until empty, remove the subfolder and the local directory → nothing of the run is left.

## Skipped

- `empty_deleted_items(confirm: true)`: it empties the whole Deleted Items folder, including the user's own items. It is never run live without the user's explicit permission. Only the `confirm: false` path is tested, in earlier suites.
- PD7 is not forced by filling Deleted Items with 500 items.
