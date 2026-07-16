# System test plan — Plans 1–9 (email + calendar), 2026-07-16

## Purpose

Live-verify every email and calendar tool built across Plans 1–9 of the v2
feature build, against the real, running Outlook mailbox
`adamkopelman@outlook.com`. This is a retro record: every action below will
actually be executed after this document is reviewed, in order, and the
result of each will be reported.

## Mechanism

The MCP server registered for *this* Claude Code session was the old Python
implementation when this session started; it has now been swapped to the new
Rust release binary, but a running session cannot pick up a new MCP tool
list without restarting. So these tests are **not** invoked as MCP tool
calls in this conversation — they call the exact same code the MCP tool
layer calls, one layer down, directly against `WindowsOutlookClient` in
`src/outlook/client.rs` (the file the user asked to test against). This is
the identical mechanism already used and proven for Plan 9's live recurrence
tests: a `#[ignore]`d test in `tests/live_outlook.rs` (or a new sibling file)
that builds a real `WindowsOutlookClient`, calls its methods, and asserts on
the real returned data. `server.rs`'s tool methods are thin wrappers with no
logic of their own (parameter struct → domain struct → client call → JSON) —
calling `client.rs` directly exercises the same COM code path an MCP tool
call would.

Next real Claude Code session (once restarted) can additionally re-run a
handful of these as actual MCP tool calls through the new binary for
end-to-end confidence in the wire protocol itself; that is out of scope for
*this* session since it's not possible here.

## Accounts used

- **Self / organizer mailbox:** `adamkopelman@outlook.com` (the live Outlook
  profile automated by this project).
- **External test recipient:** `adamkopelman2@gmail.com` (the user's own
  second address, explicitly authorized for test emails/invites).

All test data is tagged with the subject prefix `[outlook-mcp-rs systest]`
and dated 2026-07-16 so it's trivially greppable/cleanable, and every
calendar event uses far-future dates (2099-xx-xx) so it can never collide
with a real appointment. Every write this plan performs is cleaned up
(deleted) at the end of its own test, except the two "land in the inbox and
stay long enough to be read back" emails, which are cleaned up in a final
pass after they're verified.

## Scope note

Tasks/Notes (Plans 11–12) are not built yet, so `create_task`/`list_tasks`/
`complete_task`/`create_note`/`list_notes`/`get_note` are out of scope here.
`respond_to_meeting` is scoped out for a documented reason (below) — everything
else across Plans 1–9 is covered.

### Category note

Outlook's `Categories` property (`src/outlook/com.rs:239-248`) is a plain
semicolon-joined string set via COM — it is **not** validated against the
mailbox's Master Category List. Test categories like `"systest-work"` that
don't already exist in this Outlook profile will be accepted and will match
correctly in filters; they just won't render with a color swatch in the
Outlook UI (cosmetic only, not a functional gap).

---

## Part 0 — Seed data

**Ordering note:** A1 (`list_folders`) runs first, before any seeding, so
S5's destination folder is picked from the real folder list rather than
guessed.

The inbox/calendar are close to empty, so the filter tests in A3/B1 below
would mostly prove "the call didn't error," not "the filter actually
selects the right subset." This phase creates a small, deliberately varied
batch of real data first, so every filter has a real positive and negative
case to check against. Every item is tagged `[outlook-mcp-rs systest]` and
cleaned up in the final cleanup pass along with everything else.

### S1–S8: seed emails
All sent via `send_email(to: ["adamkopelman@outlook.com"], ...)` (self-loop,
no external traffic), then adjusted with `update_email` where noted, and
confirmed to have landed via `list_emails` before adjustment:

Categories use the mailbox's real default palette (Red/Blue/Green/Orange/
Purple/Yellow Category) instead of made-up names, so this also proves real
category names round-trip correctly, not just arbitrary strings.

| id | subject suffix | importance | flag (after send) | category (after send) | read state | folder |
|----|----|----|----|----|----|----|
| S1 | seed urgent | high | follow_up | Red Category | unread | Inbox |
| S2 | seed work read | normal | (none) | Blue Category | **read** (`mark_read: true`) | Inbox |
| S3 | seed personal | normal | (none) | Green Category | unread | Inbox |
| S4 | seed docs+attachment | normal | (none) | Purple Category | unread | Inbox (has a real attachment, small scratch `.txt`) |
| S5 | seed archived | normal | (none) | Blue Category | unread | moved to a non-Inbox folder (`update_email(move_to: ...)`, exact folder chosen from the real `list_folders` result in A1) |
| S6 | seed completed | normal | complete | Red Category | unread | Inbox |
| S7 | seed low importance | low | (none) | Green Category | unread | Inbox |
| S8 | seed plain | normal | (none) | (none) | unread | Inbox — the "no filters match" negative control |

### S9–S14: seed calendar events
All `create_event`, no attendees (personal appointments — no external
traffic), dated 2099-06-01 through 2099-06-06 (a range untouched by Part
B's 2099-05-xx events, so filter tests can target either range cleanly):

| id | date | show_as | all_day | category |
|----|----|----|----|----|
| S9 | 2099-06-01 | busy | false | Blue Category |
| S10 | 2099-06-02 | free | false | Green Category |
| S11 | 2099-06-03 | tentative | false | Blue Category |
| S12 | 2099-06-04 | out_of_office | false | Red Category |
| S13 | 2099-06-05 | working_elsewhere | false | (none) |
| S14 | 2099-06-06 | busy | **true** (all-day) | Green Category |

Once S1–S14 exist, A3 and B1 below run their filters against this known
set and assert the exact expected subject list each filter should return,
not just "call succeeded."

---

## Part A — Email tools

### A1. `list_folders`
**Action:** Call with no arguments.
**Expected:** Returns a list including at least "Inbox", "Sent Items",
"Deleted Items", "Drafts", each with a name, path, and item count ≥ 0.

### A2. `list_emails` — defaults
**Action:** Run *after* S1–S8 exist. Call with no filters (folder defaults
to `"inbox"`, count 10).
**Expected:** Returns the seeded inbox emails (S1–S4, S6–S8; S5 was moved
out of the inbox), most-recent-first, real subjects/from/received/categories.

### A3. `list_emails` — filters, against the known S1–S8 seed set
Each call is read-only and checked against the exact expected subject set:

| filter | expected match |
|----|----|
| `unread_only: true` | S1, S3, S4, S5\*, S6, S7, S8 (everything except S2, which was explicitly marked read) — \*S5 only if its folder is also queried; the Inbox-only call excludes it since it was moved out |
| `flagged: true` | S1 (follow_up) and S6 (complete) — both count as flagged; S2/S3/S4/S7/S8 do not |
| `high_importance: true` | S1 only |
| `category: "Red Category"` | S1, S6 |
| `category: "Blue Category"` | S2 (and S5 if querying its destination folder) |
| `since_days: 1` | all of S1–S8 (all created today, 2026-07-16) |
| `query: "docs"` (from S4's subject) | S4 only |
| a filter combination, e.g. `category: "Green Category", unread_only: true` | S3, S7 (both unread + Green Category); NOT S8 (no category) |

**Expected:** Every row's actual result set matches the expected subject
list exactly — not just "call succeeded." Any mismatch is a real finding.

### A4. `get_email`
**Action:** Take one real email id from A2's results and call `get_email`
with `prefer_html: false`, then again with `prefer_html: true`.
**Expected:** Full body text returned both times (HTML variant may differ in
formatting/tags); no error.

### A5. `send_email` — to external test address
**Action:** `send_email(to: ["adamkopelman2@gmail.com"], subject: "[outlook-mcp-rs systest] send_email external", body: "Automated system test — Plans 1-9 live verification, 2026-07-16.", html: false)`.
**Expected:** Success status returned (e.g. `{"status": "sent"}` or similar);
the item leaves the Outbox. (Delivery/read confirmation on the Gmail side is
out of reach from this session — success here means Outlook accepted and sent it.)

### A6. `send_email` — to self, then read it back
**Action:** `send_email(to: ["adamkopelman@outlook.com"], subject: "[outlook-mcp-rs systest] send_email self-loop", body: "Self-loop test.", html: false)`.
Then `list_emails(query: "systest self-loop")` to find it once it lands in
the inbox, then `get_email` on that id to confirm the body round-trips.
**Expected:** The self-sent email appears in the inbox with matching subject
and body. This is the "send to yourself and confirm you can get it back"
check the user asked for.

### A7. `create_draft`
**Action:** `create_draft(to: ["adamkopelman2@gmail.com"], subject: "[outlook-mcp-rs systest] draft probe", body: "Draft, never sent.")`.
**Expected:** Success, item created in Drafts folder, NOT sent (no delivery).
Verify via `list_emails(folder: "drafts", query: "systest draft probe")`.
Then delete it via `delete_email` (draft cleanup).

### A8. `reply_email`
**Action:** Reply to the self-loop email from A6:
`reply_email(email_id: <A6 id>, body: "Reply body.", reply_all: false, send: true)`.
Then confirm a reply landed back in the inbox (same mailbox both ways) via
`list_emails`.
**Expected:** Reply sent successfully; new item appears with "RE:" subject
prefix.

### A9. `update_email` — full field sweep
**Action:** On the A6 self-loop email (still present), run a sequence of
`update_email` calls: (1) `mark_read: true` → then `mark_read: false`;
(2) `flag: "follow_up"` → confirm → `flag: "clear"`; (3)
`add_categories: ["Orange Category"]` → confirm via `get_email`/`list_emails` →
`remove_categories: ["Orange Category"]`; (4) `importance: "high"`.
**Expected:** Each update reports the field(s) it changed in its `changed`
list; a follow-up `get_email`/`list_emails(category: "Orange Category")` call
confirms categories were actually applied before being removed.

### A10. `update_email` — move_to
**Action:** `update_email(email_id: <A6 id>, move_to: "Archive")` (or another
existing non-default folder from A1's list if "Archive" doesn't exist —
confirm from A1's output first). Then `list_emails(folder: "<that folder>")`
to confirm it arrived, noting the **new** id `move_to` returns (moving
changes the EntryID).
**Expected:** Email now appears in the destination folder under its new id;
no longer in the inbox.

### A11. `delete_email`
**Action:** Delete every test email created above (A5's sent copy from Sent
Items if desired, A6/A9/A10's self-loop email by its latest id, A8's reply).
**Expected:** Each reports `{"status": "deleted", ...}`; items move to
Deleted Items (not permanently purged — matches documented behavior).

### A12. `list_attachments` / `save_attachments`
**Action:** `send_email` a new test message to self with a real small
attachment (pick an existing small file on disk, e.g. a short `.txt` scratch
file created for this purpose) — subject `[outlook-mcp-rs systest] attachment probe`.
Once it lands in the inbox (`list_emails`), call `list_attachments` on it,
then `save_attachments(save_dir: "<scratchpad dir>")`.
**Expected:** `list_attachments` reports the filename/size; `save_attachments`
writes the file to the given directory and reports the saved path; the
saved file's content matches the original. Clean up: delete the email
(`delete_email`) and the saved copy on disk.

---

## Part B — Calendar tools

### B1. `list_events` — defaults, then filters against the known S9–S14 seed set
**Action:** First, `list_events(start_date: "2026-07-16", end_date: "2026-08-15")`
with no other filters, to confirm real near-term events on the calendar (if
any) are returned without error — no fixed expectation here since this is
real pre-existing data, not seeded.

Then, run against the seeded range `start_date: "2099-06-01", end_date: "2099-06-10"`:

| filter | expected match (from S9–S14) |
|----|----|
| (none, just the date range) | S9, S10, S11, S12, S13, S14 (all 6) |
| `all_day: true` | S14 only |
| `all_day: false` | S9, S10, S11, S12, S13 |
| `show_as: "busy"` | S9, S14 |
| `show_as: "tentative"` | S11 |
| `show_as: "out_of_office"` | S12 |
| `show_as: "working_elsewhere"` | S13 |
| `category: "Blue Category"` | S9, S11 |
| `category: "Green Category"` | S10, S14 |
| `meetings_only: true` | none — S9–S14 are all personal appointments, no attendees |

**Expected:** Every row's actual result set matches exactly. Any mismatch is
a real finding, not a "the mailbox happened to be different" excuse — this
is fully controlled seed data.

### B2. `create_event` — personal appointment
**Action:** `create_event(subject: "[outlook-mcp-rs systest] personal appt", start: "2099-05-01T09:00", end: "2099-05-01T09:30", show_as: "busy", categories: ["Purple Category"])` — no attendees.
**Expected:** Created, `status: "saved"` (non-meeting). `get_event` confirms
subject/times/show_as/categories.

### B3. `get_event`
**Action:** `get_event` on B2's id.
**Expected:** Full detail returned: subject, start/end, show_as="busy",
categories includes "Purple Category", `recurrence: null` (one-off event).

### B4. `create_event` — real meeting invite to the external address
**Action:** `create_event(subject: "[outlook-mcp-rs systest] meeting invite", start: "2099-05-02T14:00", end: "2099-05-02T14:30", required_attendees: ["adamkopelman2@gmail.com"], send: true)`.
**Expected:** `status: "meeting_sent"`; a real invite email is sent to
`adamkopelman2@gmail.com`. `get_event` confirms `required_attendees` includes
that address and `is_recurring: false`.

### B5. `create_event` — meeting saved (not sent)
**Action:** Same as B4 but `send: false`, different subject/time
(`2099-05-03T14:00`).
**Expected:** `status: "meeting_saved"` — no invite email sent, item exists
only in the organizer's own calendar. Used later to test `update_event`'s
`send_update` without spamming the external inbox further.

### B6. `list_events` — confirm B2/B4/B5 show up
**Action:** `list_events(start_date: "2099-05-01", end_date: "2099-05-04")`.
**Expected:** All three events present; `list_events(meetings_only: true)`
over the same range returns only B4 and B5 (not B2).

### B7. `update_event` — field edits
**Action:** On B2 (the personal appointment): change `subject`, `location`,
`body`, `reminder_minutes`, `show_as: "tentative"`, `add_categories: ["Yellow Category"]`
→ confirm via `get_event` → `remove_categories: ["Yellow Category"]`.
**Expected:** Each reports the changed fields; `get_event` reflects every
change; final state has only `"Purple Category"` in categories.

### B8. `update_event` — attendee management (converts personal → meeting)
**Action:** On B2 (still a personal appointment): `update_event(event_id: <B2>, add_required_attendees: ["adamkopelman2@gmail.com"], send_update: true)`.
**Expected:** B2 becomes a meeting; invite email sent to the external
address (2nd real email this plan sends there, after B4). `get_event`
confirms the attendee is present. Then
`update_event(event_id: <B2>, remove_attendees: ["adamkopelman2@gmail.com"], send_update: false)`
to revert it to personal quietly (no extra email).

### B9. `update_event` — recurrence set/replace and clear
**Not repeated here.** This exact path (`recurrence` set, pattern replace,
`clear_recurrence`) was already live-verified end-to-end in Plan 9, all 4
scenarios passing (weekly/monthly+until/yearly+no-end/update+clear) —
re-running it would be redundant. Skipped by design, not an oversight.

### B10. `respond_to_meeting` — SKIPPED
**Reason:** This tool responds to a meeting invite *received* by this
mailbox. Every meeting this plan creates is organized *by* this mailbox, so
there is no inbound invite in `adamkopelman@outlook.com` to respond to
without a second, independently-controlled mailbox (the external address is
a plain Gmail account with no COM/MCP access from this environment).
Responding to a real pre-existing invite from an actual third party in the
live inbox would send a real accept/decline notification to that person,
which is not something this test plan will do without explicit separate
authorization. This mirrors the project's existing documented precedent
(`TESTING.md`) of leaving irreversible, real-person-facing actions as
manual-only/out-of-scope for automated live tests.

### B11. `delete_event`
**Action:** Delete every calendar event created above: B2 (now reverted to
personal), B4, B5.
- B2: `delete_event(event_id: <B2>, send_cancellation: false)` — never was a
  meeting at delete time (reverted in B8), cancellation flag irrelevant.
- B4: `delete_event(event_id: <B4>, send_cancellation: true)` — this
  organizes a real external invite, so a real cancellation notice will be
  sent to `adamkopelman2@gmail.com`, which is correct and expected.
- B5: `delete_event(event_id: <B5>, send_cancellation: false)` — never sent
  an invite (`send: false` in B5), so no cancellation email should go out;
  quiet delete.
**Expected:** All three report `{"status": "deleted", ...}`; calendar is
clean afterward (confirmed by a final `list_events` over the test date range
returning empty).

---

## Cleanup checklist (run at the end regardless of individual test outcomes)

- [ ] Delete all `[outlook-mcp-rs systest]`-tagged emails from every folder
      they ended up in (Inbox, Sent Items, Drafts, Archive/other, Deleted
      Items left as deleted is fine — no need to purge) — this includes the
      8 seed emails (S1–S8) plus everything created in Part A.
- [ ] Delete all `[outlook-mcp-rs systest]`-tagged calendar events — this
      includes the 6 seed events (S9–S14) plus everything created in Part B.
- [ ] Delete the scratch attachment file saved to disk in A12.
- [ ] Final `list_events`/`list_emails(query: "systest")` sweep confirming
      nothing test-tagged remains live in the mailbox.

## Reporting

After execution, a results summary will be produced showing, per test ID
above: pass/fail, and for any failure, the exact error and whether it's a
newly-discovered bug (root-caused, not just noted) or a test-environment
issue (matching the same rigor applied throughout Plan 9's bug fixes this
session).
