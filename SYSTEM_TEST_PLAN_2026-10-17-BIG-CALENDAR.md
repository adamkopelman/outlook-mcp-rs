# System test plan: a big calendar, 2026-10-17

## Purpose

Many of the recent issues were about calendars too big for one answer: #34 (`list_events` hard-capped at 250 rows, no paging), #43 (a week of ~105 events overflowing the context), #44 (attendee expansion that must stay bounded), #46 (bulk writes). This test fills a calendar with hundreds of events and checks every cap, page and filter against exact ground truth computed from the seed data.

## Mechanism

`tests/system_test_big_calendar.rs`, through the real MCP tool layer (`tests/common`: `OutlookMcpServer` over a real `WindowsOutlookClient`, JSON-RPC over the stdio codec). One `#[ignore]`d, non-panicking test; cleanup always runs.

    cargo test --test system_test_big_calendar -- --ignored --nocapture

## Traffic

Nothing is sent. Every event is a personal appointment, except 42 meetings in the dense week, saved with `notify: false` to `example.invalid` attendees.

## Tagging / cleanup

- **Tag:** `[outlook-mcp-rs systest bigcal]` plus a run token; all events are in 2097–2099.
- **Cleanup:** every created id is deleted permanently in batches of 100. Then a sweep over 2097-11..2100-01, page by page until empty (the calendar holds more than one page, so a single capped query would miss items). Then a check that nothing reached Deleted Items, Drafts, Sent Items or the Outbox.

## Seed data

- **March 2098:** 10 events a day, 08:00–12:30, 310 in all. Show-as rotates busy / free / tentative / out of office / working elsewhere. Categories rotate Red / Blue / Green / none. Every 10th event is Hebrew: subject `פגישת צוות dd`, location `חדר ישיבות`, category `צוות`.
- **Dense week (2098-03-09..15):** 6 more per day at 14:00–16:30, as unsent meetings with 2 required and 1 optional attendee (42 events). The week then has 121 rows.
- **All-day:** 7 events, on March 1, 6, 11, 16, 21, 26 and 31.
- **Slack bait (#42):** 60 events on Feb 17–28 and 60 on Apr 1–12. They lie inside the wider `Restrict` window but outside the requested range.
- **Edges:** 02-28 23:30, 03-01 00:00, 03-31 23:59, 04-01 00:00, and a 4-day event spanning 02-27..03-03.
- **Series:** a daily series of 400 occurrences from 2097-12-01 07:00, and a Hebrew weekly series `סדרה שבועית` of 104 Mondays from 2098-01-06.
- **Expected March rows: 397.**

## Tests (ID → expected)

### Setup
- **SEED** create the ~484 events and 2 series above through `create_event` → every create succeeds.
- **EXC** move the daily series' 2098-03-15 occurrence to 19:00 (`update_event` with its `occurrence_id`, after finding it with a one-day `list_events` window) → succeeds.

### Caps and paging
- **BC1** `list_events` over March with `count: 1000` → exactly 250 rows (the cap), sorted by start, all of them March rows.
- **BC2** `list_events` over March with no `count` → 250 rows (the default is the cap).
- **BC3** pages of 250 at `offset` 0 / 250 / 500 → page sizes [250, 147]; their union equals the 397 expected rows exactly; every row id is distinct; sorted.
- **BC4** pages of 37 until a short page → 11 pages, the last one partial; the union is exact; sorted.
- **BC5** `offset` past the end → `[]`; `offset: -7` → the same page as `offset: 0`.
- **BC6** slack bait: no row outside March → first row 03-01 00:00, last row 03-31 23:59.
- **BC7** edges → 03-01 00:00 and 03-31 23:59 are in; 02-28 23:30, 04-01 00:00 and the 02-27..03-03 span (it starts before the range) are out.

### Recurrence expansion
- **R1** the daily series over March → 31 rows with 31 distinct `occurrence_id`s, including the one moved to 19:00.
- **R2** the daily series over 2097-12-01..2099-01-31 in pages of 250 → pages [250, 150], all 400 occurrences distinct, the last one 2099-01-04.
- **R3** `offset: 390, count: 10` → exactly the last 10 occurrences; `offset: 400` → `[]`.

### The dense week (#43)
- **W1** `list_events` 03-09..03-15 → exactly the 121 expected rows (more than #43's ~105).
- **W2** the same week as full rows, `compact: true`, and `fields: [id, start]` → the same row count; byte size full > compact > id+start (sizes recorded).
- **W3** `output_dir` → `{rows_file, count}`; the file has `count` JSON lines, the same rows, and the Hebrew subjects as UTF-8.
- **W4** the 42 meetings → `required_count: 2`, `optional_count: 1`, `user_is_required: true` (you organize them); personal rows → `user_is_required: true`, `required_count: 0`.

### Filters before paging (pages of 25)
- **F1** `category: "Red Category"` → exactly the 93 red events.
- **F2** `show_as: "free"` → exactly the 63 free events.
- **F3** `show_as: ["out_of_office", "working_elsewhere"]` → exactly the 112 such events.
- **F4** `all_day: true` → exactly the 7 all-day events.
- **F5** `meetings_only: true` → exactly the 42 meetings.
- **F6** `query: m05-` → exactly the 9 day-5 events.

### Hebrew on the big calendar
- **H1** `query: פגישת צוות` → exactly the 31 Hebrew events; subject, location and category byte-exact.
- **H2** `query: location:ישיבות` → the same 31.
- **H3** `query: subject:פגישת*צוות` → the same 31; the reversed `subject:צוות*פגישת` → 0 (the wildcard keeps order).
- **H4** `category: צוות` in pages of 25 → the same 31.
- **H5** the Hebrew weekly series over 2098 in pages of 20 → exactly its 52 Mondays of 2098, distinct ids, exact subject.
- **H6** `update_event` with the 31 Hebrew ids, `add_categories: [קטגוריה נוספת]` → 31 tagged; the batch's `undo` objects as `updates` → 0 tagged.

### Batch writes at the limit (#46)
- **B1** `update_event` with 100 ids, `add_categories: [Orange Category]` → 100 results without errors and 100 orange rows; the 100 `undo` objects in one call → 0 orange.
- **B2** 101 ids → error "acts on 1 to 100 items; got 101"; nothing changes.
- **B3** `delete_event` with 100 bait ids, `permanent: true` → 100 results with `permanent: true`; 20 bait events left; none in Deleted Items, no junk drafts.

### Availability and speed
- **A1** `check_availability` for yourself on 2098-03-10 08:00–18:00 in 30-minute slots → all 20 statuses exactly as seeded (busy, free, tentative, out of office, working elsewhere…), polling until free/busy is published.
- **A2** a 250-row March page with user roles, the same without roles (`fields`), and the whole year 2098 → each under 60 s (timings recorded).

### Cleanup
- **CLEAN** delete by id in batches of 100, then sweep → nothing left; no stray mail items.
