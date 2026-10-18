# System test results: big calendar, inbound mail and meetings, and every older suite (2026-10-17)

- **Plans:** `SYSTEM_TEST_PLAN_2026-10-17-BIG-CALENDAR.md` and `SYSTEM_TEST_PLAN_2026-10-17-INBOUND.md`.
- **Branch:** `systest/big-calendar-2026-10-17`, from `main` @ `5ec8971` (PRs #37–#56 merged).
- **Review page:** `docs/system-tests.html`. It shows every system test's input, expected and actual output. Regenerate it with `python scripts/system_test_report.py <logs-dir> docs/system-tests.html`.

## Outcome (final code)

| Suite | Result |
|---|---|
| `system_test_big_calendar` (new) | **34/34** |
| `system_test_inbound` (new, with real Gmail mail and Google Calendar invites) | **22/22** |
| `system_test_prs_2026_10_17` | 63/63 |
| `system_test_prs` | 43/43 |
| `system_test` (Plans 1–9) | 66/66 |
| `system_test_p10_12` | 13/13 |
| `live_outlook.rs` (all of it) | **71/71** |
| offline suite (CI's command) | lib 328, tools 153, http 6 |
| hook tests | pass |

Every run leaves the Outlook mailbox exactly at its baseline (raw COM count of every folder). The external side is cleaned up too: the Google Calendar test events are deleted without notifications, and the Gmail test threads are trashed.

## Real bugs found and fixed

1. **`list_events` dropped recurring occurrences on a window's first day (#42).**
   - **Found by:** big calendar, EXC (a one-day window found nothing).
   - **Root cause:** a raw COM probe shows that under `IncludeRecurrences`, Outlook applies a date-only `Restrict` bound up to a day late to recurring occurrences. `[3/15, 3/16)` returned the 03-16 07:00 occurrence and skipped the 03-15 one. Single appointments, a series' first occurrence, and bounds with an explicit time are all exact.
   - **Why it only bit sometimes:** #42's window had no margin when the range started on an "unambiguous" date (day ≥ 13, or day = month). That includes "today" for the second half of every month.
   - **Fix:** two days of margin on each side of the window; the exact bounds are still applied per item.
   - **Test:** a new live test fails without the fix (`[]` instead of the 15th's occurrence).
2. **`check_availability` shifted every slot by the window's start time (since Plan 10).**
   - **Found by:** big calendar, A1. An 08:00 window reported the 08:00 events at 16:00.
   - **Root cause:** `Recipient.FreeBusy`'s string always starts at 00:00 of the start date. A probe put a 10:00 appointment at the 10:00 index for starts of 00:00, 08:00 and 13:00 alike. The parser read the string as starting at `start`.
   - **Fix:** slots are aligned to midnight, and only those overlapping `[start, end)` are returned.
   - **Test:** a new live test fails without the fix (all slots free).
   - **Why it wasn't caught:** `system_test_p10_12`'s check uses an empty window, where a shift is invisible.
3. **`check_availability` reported "working elsewhere" as busy.** A probe saved appointments with `BusyStatus` 0–4 and read them back through free/busy as `"01235"`. The free/busy string uses '5' for working elsewhere, and '5' fell into the "busy" catch-all.
4. **Batch writes returned stale ids (#51 × #50).**
   - **Found by:** inbound, M4. A batch accept of real invites returned ids that "cannot be found".
   - **Root cause:** `merge_id` kept the id each item was called with over the result's own current id. So any batch write that changes an id (accepting a meeting, `move_to`) returned the old id, while the same row's `undo` used the new one.
   - **Fix:** the result's own `id` now wins; the input id is only a fallback.
   - **Test:** a new tool test fails without the fix. Verified live by M4.
5. **`update_email_applies_state_then_moves` had never passed.**
   - **Found by:** the first full `live_outlook.rs` run.
   - **Root cause:** the test read `importance`, which email results never had, and `["summary"]`, which serialization flattens.
   - **Fix:** it now reads importance through the `importance` filter (with a negative case) and the rest through the typed fields.

## The open #53 risk is closed

`respond_to_meeting` with `notify: false` was tested on real Google Calendar invites, three runs in all, and checked from the organizer's side:

- **`notify: false`**, whether a single accept, a batch accept, accepting via the meeting request, or declining one occurrence: Google still shows **needsAction**, and Gmail receives **nothing**.
- **`notify: true`** (the positive control): Google shows **tentative**, and the response with its Hebrew comment arrives in Gmail.

So the calendar-autopilot hook's assumption holds.

## Test-suite fixes

- **The older system tests left about 75 items per run** in Deleted Items and Sent Items. Their cleanups were soft deletes and they had no Sent Items sweep.
  - They now delete permanently.
  - `system_test` sweeps its tagged sent and received copies, received since the run started, until 30 s pass with nothing new. That's needed because meeting invites, updates and cancellations reach Sent Items seconds after the deletes that queue them; a single sweep missed them twice.
  - The sweep retries deletes that Outlook.com throttles ("still being processed by the server").
- **`system_test` B8 left an orphan event in the external Google Calendar on every run.** It sent an invite and then removed the attendee quietly. Both steps are now quiet (four orphans were removed).
- **`live_outlook.rs` left 27 items and a note per run.** Its cleanups are now permanent. The three tests that check a soft delete then purge their item using a run-unique subject.

## Findings that are not code bugs

- **A Hebrew `text/plain` attachment without a charset arrives as `?`s.** Raw COM shows Outlook already stores it that way. The same text as `text/plain; charset=UTF-8` or `application/octet-stream` arrives byte-exact. `save_attachments` is byte-exact.
- **Gmail removes `<img src="cid:…">` references that don't match the Content-ID it assigns.** The image then arrives as a plain attachment with a Content-ID, which the tools handle correctly.
- **Appointment bodies come back with `" \r\n"` or `"\r\n\r\n"` line breaks** (Outlook stores them as RTF). Free/busy takes up to about a minute to publish after a write.
- **Row size: ids are about 700 B of a 1.2 KB event row** (the EntryID plus the full StoreID, and again in `occurrence_id`). So `compact` saves only about 15% on a week like #43's. Shortening ids would help context size more than any field preset.
- **Email results have no `importance` or flag field.** They can be set, but only read back through the filters.

## Judgement calls, including two mistakes

- **I asked before using Google Calendar** (approved). Gmail was authorized in the request.
- **Mistake 1:** two of my cleanup sweeps also removed **older test leftovers with the same subject**, against the skill's rule:
  - one stale "outlook-mcp-rs update_email live test" draft (its modification time had been bumped into my sweep window);
  - six 01/10 leftovers of the P8/P11/P12 delete probes.

  All were test debris, not user data, and should be recoverable via Deleted Items → "Recover items deleted from this folder". The sweep scripts now match on creation time, and the tests use run-unique subjects.
- **Not covered:** invites from Exchange or a distribution list (#44's expansion path), and a day-first Windows region for #42.
