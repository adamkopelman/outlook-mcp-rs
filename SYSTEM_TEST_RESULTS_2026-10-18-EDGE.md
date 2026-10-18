# System test results: edge cases, 2026-10-18

- **Plan:** `SYSTEM_TEST_PLAN_2026-10-18-EDGE.md`.
- **Test:** `tests/system_test_edge.rs`, run with `cargo test --test system_test_edge -- --ignored --nocapture`.
- **Branch:** `systest/big-calendar-2026-10-17` @ `8d6a331`, plus the uncommitted edge suite. Nothing in `src/` was changed.
- **Mailbox:** `adamkopelman@outlook.com` (Outlook.com). Time zone: Israel Standard Time (2099 transitions: 03-27 02:00 → 03:00 and 10-25 02:00 → 01:00).
- **Mechanism:** the real MCP tool layer (see `tests/common`). Raw Outlook COM from PowerShell is used only for what the tools don't expose: creating and removing the test folders, attaching an item to a draft, reading reminders and the raw `Categories` string, and the final tag count.
- **Traffic:** one self-loop email and its reply-all, both only to the mailbox itself. The reply-all was saved as a draft first, and sent only after its recipients were checked. Everything else was a draft to `nobody@example.invalid`, a personal appointment, or an unsent meeting saved with `notify: false`.

## After the fixes (PR #58)

**85/85 pass** on the fixed code (log on the review page). `system_test_big_calendar` (34/34) and `system_test_p10_12` (13/13) were re-run live on the fixed code too, and the offline tests pass.

| Bug | Fix | Now |
|---|---|---|
| 1. `save_attachments` name collisions | names are made unique within a call: `a.txt`, `a (2).txt` (`unique_file_name`) | AT1, AT2 pass |
| 2. Commas in category names | refused, with a clear error, by every tool that writes categories (`check_category_names`) | CT1–CT3 (re-planned) pass |
| 3. Free/busy on DST days | slots are counted in real time from local midnight and labelled in local time | DST3a: 46 slots, DST3b: 50, busy at the right times |
| 4. Event in the skipped DST hour | `create_event` / `update_event` refuse a start or end that doesn't exist on the local clock (`check_local_time_exists`), before anything is created | DST1a, DST1b refused; DST1d: the day lists, nothing left behind |
| 5. Yearly from Feb 29 | `MonthOfYear` / `DayOfMonth` are written only when they differ from what Outlook derived | RC3: 2096-02-29, 2097-02-28, 2098-02-28 |
| 6. Undo of `all_day` | the undo records start and end too; clearing `all_day` is applied before explicit times | AD4 passes both ways |
| 7. Zero-length all-day | an empty all-day range becomes the single day of its start (`all_day_end`) | AD1 passes |
| 8. Wrong item type | event, task and note tools check the item's class (`check_item_class`) | ID3: "This id is a task, not a calendar event. Use an id from list_events." |
| AV1a | `interval_minutes` below 1 is refused | passes |
| G4b | `reminder_minutes: -1` turns the reminder off (the test now uses -1) | passes |
| ID5 | `respond_to_meeting` names what the id is | "This id is an email, not a meeting invite…" |
| L2 | not a bug: the test now expects Outlook's 255-character limit | passes |

Two test-only changes made the free/busy checks wait for what they test: Outlook.com published one AV1 day before the other in one run, and the DST days late in another. AV1's wait now requires both days, and DST3 waits until its own day shows busy (it took 0 s in the final run).

## Outcome of the first run (before the fixes)

**68/85 result lines pass.** Every failure below was root-caused with a zero-write probe (unsaved items, discarded) or with the run's own raw COM reads.

| Kind | Count | Ids |
|---|---|---|
| Confirmed product bugs | 13 lines, 8 bugs | AT1, AT2, CT1, CT2, CT3, AD1, AD4, DST1a, DST1d, DST3a, DST3b, RC3, ID3 |
| Minor (validation or error text) | 3 | AV1a, G4b, ID5 |
| Outlook behaviour, not a code bug | 1 | L2 |

The first run was 57/81. Its extra failures were my own test code:
- the attachments draft included `CON.txt`, which Outlook refuses to attach, so the AT tests never ran;
- L1 didn't allow for the truncation marker;
- G3 compared against a display name Outlook wraps in quotes, and that skipped FL3;
- G1 asked for meeting data that only a sent meeting has.

All four are fixed in the test, and run 2 is the log on the review page.

## Results per id

| Id | Result | What happened |
|---|---|---|
| SEED | PASS | Subfolder `… systest edge תיקייה` + nested folder, self-loop sent, 11 query drafts, the attachments draft, 7 free/busy events |
| AT1 | **FAIL** | Two `a.txt`: both reported `saved`, both to the same path; the first one's bytes are lost |
| AT2 | **FAIL** | `a_b.txt` and `.a_b.txt` (stand-in for `a:b.txt`): both saved to `a_b.txt`; the first is overwritten |
| AT3 | PASS | Every reported save exists on disk. `CON.txt`, `...` and `a:b.txt` can't be put on an attachment here (see judgement calls) |
| CT1 | **FAIL** | `["Smith, John", "Red Category"]` reads back as `Smith`, `John`, `Red Category` (event and task); raw COM `"Smith, John, Red Category"` |
| CT2 | **FAIL** | `category: "Smith, John"` finds neither the event nor the task |
| CT3 | **FAIL** | `remove_categories: ["Smith, John"]` removes nothing |
| Q1 | PASS | `"50% off"` → only the `50% off` draft, not `50 percent off` |
| Q2 | PASS | `a_b` → only `a_b`, not `axb` |
| Q3 | PASS | `"[draft] x"`, `O'Brien`, `say "hi"`, `back\slash` → exactly one each |
| Q4 | PASS | `"star*lit"` → `star*lit` and `star-lit` (`*` is a documented wildcard, also inside quotes) |
| Q5 | PASS | `subject:` and a bare `*` act as no text filter (all 12 run drafts), no error |
| Q6 | PASS | `from: o'brien`, `to: 50%`, `from: %`, `to: %`, `to: _` → no error, no rows |
| Q7 | PASS | `list_events` `"50% off"` → only that event |
| AD1 | **FAIL** | all_day with start = end = 07-01 00:00 → a zero-length all-day event ending 07-01 00:00 |
| AD2 | PASS | all_day with 10:00–11:00 → normalised to 07-02 00:00 .. 07-03 00:00 |
| AD3a | PASS | 07-10..07-12 stored as 07-10 00:00 .. 07-13 00:00; listed on 07-10, not 07-09 or 07-13 |
| AD3b | PASS | Not listed on 07-11/07-12: documented, `list_events` bounds the event start (judgement call) |
| AD4 | **FAIL** | all-day → timed → undo is exact; timed 10:00–11:00 → all-day → undo gives **08:00–08:30** |
| AD5 | PASS | All-day daily ×5 → five all-day rows 07-25..07-29 |
| DST2 | PASS | Daily 09:00 across both transitions → 09:00 every day, 5 distinct occurrence ids each |
| DST4 | PASS | Each transition day lists its seeded events once |
| DST1a | **FAIL** | 02:30–03:00 on 03-27 (skipped hour) → create fails: "The end date you entered occurs before the start date" |
| DST1b | PASS | 02:30–04:00 → stored 03:30–05:00 |
| DST1c | PASS | 01:30–02:00 on 10-25 (repeated hour) → stored 01:30–02:00 |
| DST1d | **FAIL** | After DST1a, a broken item titled DST1-skipped is in the Calendar, and `list_events` over 03-27 fails with the same error |
| AV1a | **FAIL** | `interval_minutes: 0` silently gives 240 one-minute slots (undocumented) |
| AV1b | PASS | Interval 7 over 28 h: 241 slots on the midnight grid, busy exactly around 09:00–10:00 on both days |
| AV1c | PASS | Interval 45: 38 slots, busy 09:00 and 09:45 on both days |
| AV1d | PASS | Interval 1440: one slot per day (busy, busy, free) |
| AV2 | PASS | 23:30 → 00:30: busy, out_of_office; 3-day window from 23:30: 144 slots, exact |
| AV3 | PASS | 2099-12-01: exact statuses |
| AV4 | PASS | You exact; `nobody@example.invalid` resolved false with no slots |
| DST3a | **FAIL** | 03-27: busy at 08:00, 09:00, 09:30 for events at 09:00 and 10:00–11:00 (an hour early); 48 slots for a 23-hour day |
| DST3b | **FAIL** | 10-25: busy at 10:00, 11:00, 11:30 for events at 09:00 and 10:00–11:00 (an hour late); 48 slots for a 25-hour day |
| RC1 | PASS | Mon + Wed every 2 weeks ×6 → exactly those dates |
| RC2 | PASS | Monthly on day 31 ×4 → Jan 31, Feb 28, Mar 31, Apr 30 |
| RC3 | **FAIL** | Yearly from 2096-02-29 → "The recurrence pattern is not valid" |
| RC4 | PASS | `until` before start, `occurrences` 0 and −1 → errors (Outlook's own wording); nothing created |
| RC5 | PASS | Weekly → monthly on day 1: the pattern reads back monthly, occurrences 06-01, 07-01, 08-01 |
| RC6 | PASS | Daily ×5 listed across its end → exactly 5 |
| PD1 | PASS | Soft delete, then permanent from Deleted Items → gone |
| PD2 | PASS | Second permanent delete → "Item not found" |
| PD3 | PASS | `O'Neil` draft and event deleted permanently; nothing left in Deleted Items, Drafts or Calendar |
| PD4 | PASS | Occurrence delete with `permanent` → only that occurrence gone; the result says permanent has no effect (documented) |
| PD5 | PASS | Mail in the Hebrew subfolder deleted permanently → gone from the subfolder and Deleted Items |
| PD6 | PASS | Stale event id → "Item not found" |
| PD7 | PASS | Not reachable: the largest same-class group in Deleted Items is 15 (needs over 500) |
| OD1 | PASS | Quiet soft delete of an unsent meeting: nothing in Sent Items or Outbox, item in Deleted Items, not in the calendar |
| OD2 | PASS | Quiet delete of one occurrence: 2 left, neither canceled, series not canceled, nothing sent |
| PG1 | PASS | 60 drafts; deleting 5 of page 1 shifts page 2 by 5 (offset semantics) |
| PG2 | PASS | Identical calls → identical rows |
| PG3 | PASS | Same for 60 events |
| FO1 | PASS | `move_to: Inbox/<Hebrew subfolder>` → listed there under the new id; `get_email` opens it |
| FO2 | PASS | Moved back to Drafts → only in Drafts |
| FO3 | PASS | `list_folders` shows the subfolder under Inbox, Hebrew name exact |
| G2 | PASS | The nested folder appears with its full parent path |
| FO4 | PASS | `calendar_of` yourself → the event |
| FO5 | PASS | `calendar_of: nobody@example.invalid` → a clear error |
| AE1 | PASS | Adding `x@example.invalid` twice in one call → one recipient |
| AE2 | PASS | Adding `X@Example.Invalid` again → still one |
| AE3 | PASS | Removing someone not on the meeting → no error, attendees unchanged (the result still lists `remove_attendees` as changed) |
| L1 | PASS | 1 MB body: 999 and 1000 → 1000 characters plus the truncation marker, flagged; 5,000,001 → the whole body |
| L2 | **FAIL** | A 300-character subject is stored as 255 characters (Outlook's limit) |
| L3 | PASS | 12 MB inline image → the "exceeds the 10 MB limit" error |
| L4 | PASS | 200 ids → 200 results; 201 → 201 (no read-batch limit exists or is documented) |
| FL1 | PASS | complete → clear → follow_up: each found only by its own `flag` value |
| FL2 | PASS | `unread_only` follows mark unread / read |
| G5 | PASS | `received_before` received + 1 min includes it, − 1 min excludes it |
| G3 | PASS | Reply-all draft addressed only to you; sent; one reply in Sent Items to you only |
| FL3 | PASS | With the unflagged reply newest, `flag: follow_up, count: 1` returns the flagged original |
| G1a | PASS | Not applicable as planned (see judgement calls): on the meeting's calendar item `get_email` returns item_type `other`, is_meeting false |
| G1b | PASS | Plain draft with `include: ["meeting"]` → no meeting block |
| G4a | PASS | `reminder_minutes: 15` → raw COM ReminderSet True, 15 (`get_event` has no reminder field and documents none) |
| G4b | **FAIL** | `reminder_minutes: 0` → raw COM ReminderSet True, 0: a reminder at the start time, not "no reminder" |
| G6 | PASS | Not applicable: `list_events` has no `importance` filter, and its docs don't claim one |
| G7 | PASS | `attendees` + `attendee_role` match the optional and required tiers exactly |
| G8 | PASS | `move_to: Archive` → found by `list_emails(folder: Archive)` |
| G9 | PASS | An attached item: `list_attachments` type `item`, `get_email` lists it, `save_attachments` saves the `.msg` |
| ID1 | PASS | Empty id → "Invalid item id" from get, update and delete |
| ID2 | PASS | Hard-deleted ids → errors from get, update, delete and `get_event` |
| ID3 | **FAIL** | A task id to `get_event` returns the task as an event (start 1899-12-30); an event id to `get_task` returns it as a task |
| ID4 | PASS | An occurrence id after the series end, or off the pattern's time → "Occurrence not found" |
| ID5 | **FAIL** | `respond_to_meeting` with a draft id → "Outlook error: Unknown name." (an error, but not a clear one) |
| CLEAN | PASS | 89 emails, 96 events and 1 task deleted by id, 5 more swept; folders removed; raw COM count of the run's tag: 0 |

## Confirmed bugs (all fixed in PR #58, see above)

### 1. `save_attachments` overwrites attachments whose safe names collide (AT1, AT2)

- **Observed:** two attachments named `a.txt` were both reported `saved` with the same `saved_to`. The file on disk holds the second one's bytes. The same happens to `a_b.txt` and `.a_b.txt`.
- **Root cause:** `client.rs:2848` builds the target as `dir.join(safe_filename(&info.filename))`, with no check for a name already used in this call or already on disk.
  - `safe_filename` (`com.rs:25`) maps `\/:*?"<>|` and control characters to `_` and trims leading and trailing dots and spaces. So `a:b.txt` and `a_b.txt`, or `.a_b.txt` and `a_b.txt`, become the same name.
- **Confirmed by:** the run itself, which compares the bytes on disk with the source files.

### 2. A category name with a comma is split into two (CT1–CT3)

- **Observed:** `["Smith, John", "Red Category"]` reads back as three categories. The `category` filter and `remove_categories` then can't find `Smith, John`.
- **Root cause:** Outlook keeps categories as one comma-separated string; the raw COM read gives `"Smith, John, Red Category"`.
  - `join_categories` (`com.rs:46`, via `set_item_categories`, `com.rs:519`) writes the name unchanged, and `parse_categories` (`com.rs:37`) splits on `,` when reading back.
  - A comma can't be stored in a category name at all, so the fix is to reject it (or warn), not to parse differently.

### 3. `check_availability` mislabels slots on DST days (DST3a, DST3b)

- **Observed:** on 2099-03-27 the 10:00–11:00 event shows as busy at 09:00–10:00, and the 09:00 event at 08:00. On 2099-10-25 they show an hour late. Both days get 48 half-hour slots, for days of 23 and 25 hours.
- **Root cause:** `Recipient.FreeBusy` returns one character per real elapsed interval from local midnight.
  - A diagnostic read the raw string for 03-27: the 09:00–09:15 event is at index 16 and the 10:00–11:00 event at indices 18–19, which is 8 and 9 elapsed hours after midnight.
  - `parse_freebusy_slots` (`mod.rs:762-766`) labels index i as `midnight + i × interval` in naive local time. That is only right on a 24-hour day.
- This is the same function as the 2026-10-17 midnight fix, but a different defect.

### 4. An event in the skipped DST hour fails, and leaves a broken item that breaks `list_events` (DST1a, DST1d)

- **Observed:** `create_event` 02:30–03:00 on 2099-03-27 fails with "The end date you entered occurs before the start date".
  - Right after, raw COM finds an item with that subject in the **Calendar**, and every `list_events` window covering 03-27 fails with the same error.
  - By cleanup time Outlook had moved the item to **Sync Issues/Local Failures**. There, `Delete()`, `Move()` and every property write fail with the same error. Only `Items.Remove(index)` removes it, and the suite's cleanup now does that.
- **Root cause:** Outlook resolves the nonexistent local 02:30 as 03:30. A zero-write probe of an unsaved item showed StartUTC 00:30 and EndUTC 00:00, so the end is before the start and Save fails.
  - `create_event` (`client.rs:2424-2425`) passes the naive local time through without checking that it exists.
  - On failure it doesn't discard the half-built item, so it stays in the folder.
  - With an end past the gap (02:30–04:00, DST1b) the event saves as 03:30–05:00.

### 5. A yearly series starting on Feb 29 can't be created (RC3)

- **Observed:** yearly ×3 from 2096-02-29 → "The recurrence pattern is not valid".
- **Root cause:** a zero-write probe showed that after `RecurrenceType = yearly`, Outlook has already derived MonthOfYear 2 and DayOfMonth 29 from the start.
  - Any explicit write of `MonthOfYear` (or of `DayOfMonth` 29) is then rejected.
  - Leaving both alone gives a valid series (2096-02-29, 2097-02-28, 2098-02-28).
  - `apply_recurrence` always writes both (`client.rs:353-354`).
  - Other dates are unaffected (`live_outlook.rs`' yearly test on 03-10 passes; the probe on 02-28 is fine).

### 6. `undo` of an `all_day` change doesn't restore the times (AD4)

- **Observed:** 10:00–11:00 → `all_day: true` → its undo gives **08:00–08:30**.
- **Root cause:** the undo records only the old flag (`client.rs:2581`). Outlook rewrites Start/End whenever `AllDayEvent` changes: a probe that cleared the flag on an all-day item got 08:00–08:30.
  - So the undo has to carry the old start and end as well, as the start/end branch does.
  - The other direction round-trips only because setting the flag again snaps back to whole days.

### 7. `all_day` with start = end creates a zero-length all-day event (AD1)

- **Observed:** start = end = 07-01 00:00 with `all_day: true` → stored 07-01 00:00 .. 07-01 00:00.
- **Root cause:** `create_event` sets Start, End, then AllDayEvent (`client.rs:2424-2427`) and never extends the end. A probe shows Outlook accepts Duration 0 in either order.
  - A timed range is normalised by Outlook itself (AD2), which is why only the empty range goes wrong.

### 8. `get_event` / `get_task` return an item of the wrong type (ID3)

- **Observed:** a task id to `get_event` returns the task as an event starting 1899-12-30 (Outlook's zero date). An event id to `get_task` returns it as a `not_started` task.
- **Root cause:** `read_event` (`client.rs:1369`, via `open_event`) and `read_task` (`client.rs:1400`) open any id with `GetItemFromID` and never check the item's class. The summaries fall back to defaults for missing properties.

### Minor

- **AV1a:** `interval_minutes: 0` is silently treated as 1 (`client.rs:2744`, `.max(1)`) and gives 240 one-minute slots. The tool doesn't document this; reject non-positive intervals instead.
- **G4b:** there is no way to turn a reminder off. `reminder_minutes` always sets `ReminderSet = true` (`client.rs:2592`, and in `create_event`), so 0 means "remind at the start", not "no reminder".
- **ID5:** `respond_to_meeting` on a draft id calls `Respond` on a mail item and surfaces "Outlook error: Unknown name. (HRESULT 0x80020006)" (`client.rs:2509`). It is an error, but it doesn't say the id is not a meeting.

## Outlook behaviour (not code bugs)

- **L2:** subjects are cut to 255 characters. A probe of an unsaved item already reads back 255. `create_draft` still echoes the requested 300-character subject in its result.
- **Free/busy in the repeated hour:** a diagnostic event at 01:30–02:00 on 2099-10-25 doesn't show as busy in the raw `FreeBusy` string at all. `get_event` reads it back fine.
- **`CON.txt`:** Windows 11 allows creating it, but `Attachments.Add` refuses it ("Value does not fall within the expected range").
- **DASL wildcards:** `%`, `_` and `[…]` in a query or in `from`/`to` are taken literally on this Outlook.com store (Q1–Q3, Q6). The review's wildcard hypotheses are refuted.
- **Clearing all-day sets 08:00–08:30.** Outlook does this itself (see bug 6).

## Judgement calls

- **AT2 / AT3 use stand-ins.**
  - No Windows file can be named `a:b.txt` or `...`.
  - MAPI refuses a `FileName` change through `PropertyAccessor` (access denied).
  - An NTFS alternate data stream path is rejected by `Attachments.Add`.
  - `Namespace.OpenSharedItem` on an `.eml` with such names hung (the probe process was killed; it created nothing).

  So AT2 uses `.a_b.txt`, whose leading dot `safe_filename` trims: the same collision. AT3 records the unreachable names as a limitation, and checks only that every reported save is on disk. `CON.txt` is unreachable because Outlook won't attach it.
- **AD3b is a pass.** The plan expected the 3-day event on its middle days. `list_events` documents that `start_after`/`start_before` bound the event *start*, and the big-calendar suite (BC7) asserts the same, so it is not listed there.
- **DST1 became four steps.** DST1a (the plan's skipped-hour event) can't be created, so DST1b records what Outlook stores for an event that starts in the gap but ends past it. DST1d records the side effect. DST2 and DST4 run before DST1a, so the broken item doesn't fail them.
- **G1a is not applicable.** Only sending a meeting creates a meeting-request message, and this suite never sends one. `get_email`'s `is_meeting`/`meeting` describe such a message, so on the calendar item they are false/absent. It is recorded as a pass with the observed values; whether `get_email` should reject an event id is a question, not a confirmed bug.
- **L4 is not applicable as planned.** `get_email` has no read-batch limit and documents none, so 201 ids simply return 201 results. 200 ids are 55 tagged drafts repeated; creating 200 drafts would only add traffic.
- **G6 is not applicable.** `list_events` has no `importance` parameter, and its docs don't claim one.
- **G3 sent in two steps.** The reply-all was saved with `send: false`, its recipients were checked to be only the mailbox itself, and then it was sent with `update_draft(send: true)`. `reply_email`'s direct send path is covered by older suites.
- **AV3:** every window in this suite is more than 40 days out (2099), so AV3 is one more far-future window (2099-12-01).
- **Raw COM writes** were only to this run's tagged items: creating and deleting the two test folders, attaching the G9 inner draft to the G9 outer draft, and removing the stranded Local Failures item.

## Cleanup

- **Each run:**
  - delete every created id permanently;
  - wait for both inbound copies;
  - sweep Drafts, Inbox, Sent Items, Deleted Items, Outbox, Archive, both test folders, the calendar (2096–2100), Tasks and Notes by the run's exact tag until a pass finds nothing;
  - raw-COM-delete anything tagged that the tools can't reach (the stranded DST1a item, via `Items.Remove`);
  - remove both folders (from the Inbox and then from Deleted Items) and the local directory (attachments, the 12 MB image, the 1 MB body file).
- **Run 1** left three things, all removed by hand with raw COM, matching only that run's tag:
  - the subfolder in Deleted Items (that version of the folder-removal script failed);
  - the reply-all, which arrived after the sweep;
  - the stranded DST1a item.

  The cleanup was then fixed, and **run 2 cleaned up after itself completely**.
- **Final state** (independent raw COM count over every folder):
  - 0 items or folders tagged `[outlook-mcp-rs systest edge]`;
  - Deleted Items holds 47 items, the same as before the first run;
  - Inbox has only its original subfolder;
  - no `outlook-mcp-rs-systest-edge-*` or probe directories are left in `%TEMP%`.
- **Older debris, left alone:** 204 items tagged `[outlook-mcp-rs systest` from earlier suites: 200 in Sent Items and 4 "Undeliverable: [outlook-mcp-rs systest] seed …" in Sync Issues/Local Failures. They predate this run, and `empty_deleted_items` was never called.
