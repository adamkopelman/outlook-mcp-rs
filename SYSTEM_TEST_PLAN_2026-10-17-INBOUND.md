# System test plan: real external mail and meeting invites, 2026-10-17

## Purpose

Every earlier system test made its own data: drafts, self-sent mail, its own appointments. This one tests what the mailbox *receives* from somebody else:

- mail from an external Gmail account, in Hebrew, with inline images and attachments;
- real meeting invites from Google Calendar, answered through `respond_to_meeting`.

`respond_to_meeting` was never live-tested before. #53's autopilot hook auto-approves it with `notify: false`, on the assumption that nothing reaches the organizer; this test checks that assumption from the organizer's side.

## Mechanism

Two parts:

1. **Connector step** (an agent with Gmail and Google Calendar connectors, or a person):
   - send G1–G3 from `adamkopelman2@gmail.com` to `adamkopelman@outlook.com`;
   - create the Google Calendar invites M1–M7 (M4 is two invites) to the same address, all in 2099;
   - tag everything `[outlook-mcp-rs systest inbound] r<token>`.
2. **`tests/system_test_inbound.rs`**, through the real MCP tool layer (`tests/common`):

       SYSTEST_INBOUND_RUN=r<token> cargo test --test system_test_inbound -- --ignored --nocapture

   It waits up to 5 minutes for delivery, runs the checks, prints `VERIFY` lines for the Google side, and deletes everything of the run on the Outlook side.

Then the connector step checks the `VERIFY` lines in Gmail and Google Calendar, deletes the Google events with `notificationLevel: NONE`, and trashes the test threads.

## Accounts and traffic

The user authorized both:

- **Gmail ↔ Outlook:** between the user's own two accounts.
- **Google Calendar invites:** sent from the user's Google account.

Outlook sends four messages to Gmail: one reply, two emails, and one tentative response (M2). Everything else stays in Outlook.

## Connector-step inputs

- **G1:** subject `… G1 שלום מג'ימייל`; plain body `גוף ההודעה בעברית: תפוזכתום` plus an English line. The Hebrew word appears only in the body.
- **G2:** HTML `<p>תמונה מוטמעת</p><img src="cid:pixel.png">`, plus a 1×1 PNG attached as inline.
- **G3:** four attachments:
  - `notes.txt`;
  - `מסמך.txt` (`text/plain`);
  - `מסמך-utf8.txt` (`text/plain; charset=UTF-8`);
  - `מסמך.bin` (`application/octet-stream`).

  The last three carry the same Hebrew text.
- **M1–M7:** invites on 2099-06-01..07:
  - M3 repeats weekly, 6 times;
  - M5 has the mailbox as an optional attendee;
  - M7 has a Hebrew title, location `תל אביב, חדר 3` and description `סדר יום: תכנון רבעוני`.

## Tests (ID → expected)

### Setup
- **ARRIVE** poll `list_emails` (from Gmail) and `list_events` (June 2099) → G1–G3 and all 13 invite rows arrive (M3 expands to 6).

### Reading external mail, in Hebrew
- **G1** `list_emails` + `get_email` on G1 → subject byte-exact Hebrew; the sender is the Gmail address; `item_type: email`; the body contains the Hebrew text.
- **G2** `list_emails` with the Hebrew word that appears only in G1's body → exactly G1.
- **G3** G2's image:
  - Gmail gives it its own Content-ID (`ii_…`) and strips the unmatched `<img>` from the HTML;
  - expected: `get_inline_image` by that Content-ID returns the exact PNG bytes; `is_inline: false`; `resolve_inline_images` resolves 0 and leaves 0 unresolved.
- **G4** `list_attachments` + `save_attachments` on G3 → the Hebrew file names are intact. The UTF-8 and octet-stream copies are byte-exact. The charset-less `text/plain` copy arrives as `?`s (upstream: raw COM shows Outlook already stores it that way).
- **G6** `update_email` on G1–G3 in one call: `mark_read`, `flag: follow_up` and `add_categories: [דחוף]` → 3 match the category and flag filters; the `undo` objects as `updates` → 0 tagged, 0 flagged.
- **G7** `received_after: -6h` + `to` + `item_type: email` → all 3; `received_before: -6h` → none.
- **G8** `send_email` to yourself, HTML with a `data:` image → it arrives with one inline image (`is_inline: true`); `resolve_inline_images` resolves 1.
- **G5** `save_attachments` on G8's copy: `inline: false` → the error "No attachments matched…" and 0 files; `inline: true` → 1 file.

### Sending Hebrew back to Gmail
- **O1** `reply_email` to G1, Hebrew HTML with a `data:` image, `send: true` → `sent`. VERIFY: the reply is in G1's Gmail thread with the Hebrew text and one inline image.
- **O2** `send_email` to Gmail: Hebrew subject and body, attachment `דוח.txt`, `importance: high` → `sent`. VERIFY: it arrives with the Hebrew attachment name.
- **O3** `create_draft`, then `update_draft` with `send: true` and a Hebrew subject and body → `sent`. VERIFY: it arrives.

### Real meeting invites
- **M0** the invites → `is_meeting: true`, organized by the Gmail account.
- **M1** `respond_to_meeting` accept, `notify: false` → `accept_saved`; the returned id opens with `my_response: accepted`. VERIFY: Google still shows needsAction, and Gmail has no response.
- **M2** tentative with a Hebrew comment, `notify: true` → `tentative_sent`. VERIFY: Google shows tentative, and Gmail gets the response with the comment.
- **M3** decline occurrence 3 of 6 by its `occurrence_id`, `notify: false` → 5 occurrences left, the declined one gone. VERIFY: no response on Google.
- **M4** batch accept of M4a and M4b, `notify: false` → 2 results whose ids open accepted M4a/M4b.
- **M5** the optional invite → `user_is_optional: true`, `user_is_required: false`; M1 → required.
- **M6** `respond_to_meeting` with the Inbox meeting request's id → the returned id is the calendar item (M6, accepted).
- **M7** the Hebrew invite → subject exact; location starts with `תל אביב, חדר 3` (Google appends the Meet link); the body has the Hebrew agenda.
- **M8** `update_event` and `delete_event` with `require_no_attendees: true` on a real invite → both refused; nothing changes.

### Cleanup
- **CLEAN** delete the run's calendar items quietly, then sweep Inbox, Sent Items, Drafts and Deleted Items until nothing of the run turns up for 30 s → nothing left.

## Not covered

- Invites from Exchange or a distribution list (the #44 expansion path). Google invites name the mailbox directly.
- A charset-less `text/plain` attachment in Hebrew: transport already loses it, before this code.
