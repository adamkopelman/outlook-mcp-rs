---
name: troubleshooting-outlook-mcp
description: Use when an outlook-mcp-rs tool returns an error ("Outlook error: … (HRESULT 0x…)", "Item not found", "Folder not found", "Could not resolve", "only unsent drafts can be edited", "Content-ID … not found", "inline_images requires html=true") or a result that looks wrong — empty calendar or inbox that shouldn't be empty, events or emails from the wrong dates, a cut-off email body, fewer emails than expected.
---

# Troubleshooting the Outlook MCP tools

## Overview

Every tool call goes through COM to the Outlook desktop app on the user's Windows
machine. Most failures come from that app's state, the id used, or how a date was
read, not from the request itself. Diagnose from the exact error text before retrying.
Never blindly retry a sending call (`send_email`, `reply_email`, …): a "failed" call
may still have sent.

## Error → cause → fix

| Error text contains | Cause | Fix |
|---|---|---|
| `Item not found — it may have been moved or deleted` | Stale id: the item moved (ids change on move and delete) or was deleted | Find it again with `list_emails`/`list_events` and use the fresh id |
| `Folder not found: "…" (no subfolder named …)` | Typo in the path | Run `list_folders` and copy the exact `path` |
| `Could not resolve "…" to a person` | The name or email isn't in the address book | Ask the user for the full email address |
| `Could not open "…"'s calendar` | That person hasn't shared their calendar | Use `check_availability` instead (free/busy needs no sharing) |
| `Invalid … expected ISO format` | The date isn't in ISO format | Use `2026-06-10` or `2026-06-10T14:30` |
| `attachment not found: …` / `inline image not found: …` | The local path doesn't exist (nothing was sent) | Fix the absolute path and retry |
| `only unsent drafts can be edited` | `update_draft` was called on received or sent mail | Use `reply_email`, or `create_draft` for a new message |
| `pass either 'body' or 'html_body', not both` / `update_draft needs at least one of…` | Invalid `update_draft` arguments (nothing was changed) | Send exactly one body field, and at least one change |
| `inline_images requires html=true` | `inline_images` used with a plain-text body | Set `html: true` and reference each image as `<img src="cid:ID">` |
| `duplicate inline image content_id` / `data_base64 is not valid base64` | Bad `inline_images` entry (nothing was created) | Give each image a unique `content_id` and valid base64, or use a `path` |
| `Content-ID '…' not found. Available Content-IDs: …` | Wrong `content_id` for `get_inline_image` | Use one of the listed ids. The `cid:` prefix and letter case don't matter |
| `no attachments with a Content-ID` | The email has no inline images | `list_attachments` shows its regular attachments |
| `exceeds the 10 MB limit for get_inline_image` | The attachment is too big to return inline | Use `save_attachments` to write it to disk |
| `empty_deleted_items permanently deletes EVERYTHING` | `confirm` was not true | Only pass `confirm: true` if the user explicitly asked to empty Deleted Items |
| `Could not apply the date filter: Outlook rejected every date format tried` | Outlook couldn't parse the date filter in this Windows regional format | Report it with the user's Region settings (short date format, calendar). Meanwhile, fetch without dates and filter by date yourself |
| `Outlook misread the date filter under this Windows regional format` | Outlook returned items far outside the requested dates, so results could be incomplete | Same as above |
| `HRESULT 0x80040154` (class not registered) | Classic Outlook isn't installed; the "new Outlook" app has no COM | Tell the user the tools need classic Outlook |
| `HRESULT 0x80080005` (server execution failed) | Outlook isn't running or is hung, or only one of Outlook and the server runs as administrator | Ask the user to (re)start Outlook normally, with the same elevation as the MCP client |
| `HRESULT 0x80010001` (call rejected) | Outlook is busy or showing a dialog box | Ask the user to close the dialog, then retry once |

## Results that look wrong (no error)

- **Calendar or inbox is empty, or items come from the wrong dates**, especially on a
  day-first Windows locale (en-GB, en-IL, de-DE, …).
  - **Cause:** versions before the fix for issue #1 sent date filters in US month/day
    order, so Outlook could misread them. Current versions can't misread them: they
    return either the right items or an error (below).
  - **Confirm:** compare a date-filtered result with an unfiltered one.
  - **Workaround on an old version:** fetch without the date filter and filter by date yourself.
- **Fewer emails than expected:**
  - `list_emails` returns one page: `count` defaults to 10, max 200. Page with `offset`.
  - The mail may also have been filed into a subfolder by a rule; check `list_folders`.
- **A Hebrew (or other non-Latin) search finds nothing:** the server falls back to
  scanning the 2,000 newest items that match the other filters. Older mail needs a date
  filter or a narrower folder.
- **Email body ends with `[... truncated at N characters]`:** `body_truncated` is
  true. Call `get_email` again with `max_body_chars` ≥ `body_length`.
- **Only 250 events:** that's the hard cap. Use a smaller date window.
- **`list_tasks` misses a task:** completed tasks are hidden unless `include_completed` is set.
- **Non-ASCII text shows as `????`:** the server's output is UTF-8. The console
  displaying it isn't (see the README's troubleshooting section).

## Don't

- Don't report "you have no meetings/emails" from an empty result if it contradicts
  what the user expects, until you've ruled out the date and paging causes above.
- Don't loop on a COM error. Retry once after the user acts, then report the exact text.
