---
name: troubleshooting-outlook-mcp
description: Use when an outlook-mcp-rs tool returns an error ("Outlook error: … (HRESULT 0x…)", "Item not found", "Folder not found", "Could not resolve", "only unsent drafts can be edited", "Content-ID … not found", "inline_images requires an HTML body", "payload too large") or a result that looks wrong — empty calendar or inbox that shouldn't be empty, events or emails from the wrong dates, a cut-off email body, fewer emails than expected.
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
| `Invalid <param> "…": … use an ISO local date/datetime like …` | The date isn't in a supported form | Use ISO (`2026-06-10`, `2026-06-10T14:30`), a keyword (`today`, `start_of_week`) or an offset (`-14d`) |
| `… is later than …: nothing can match` | A `*_after` bound is after its `*_before` bound | Swap or fix the two dates |
| `Invalid importance` / `Invalid flag` / `Invalid item_type` / `Invalid show_as` / `Invalid my_response` | Unknown filter value | Use one of the values the message lists |
| ``pass either `received_after` or `since_days`, not both`` (also `flag`/`flagged`, `importance`/`high_importance`) | A deprecated filter contradicts its replacement | Drop the deprecated one |
| `attachment not found: …` / `inline image not found: …` | The local path doesn't exist (nothing was sent) | Fix the absolute path and retry |
| `only unsent drafts can be edited` | `update_draft` was called on received or sent mail | Use `reply_email`, or `create_draft` for a new message |
| ``pass either `body` or `html_body`, not both`` (or any two of `body`, `html_body`, `body_file`, `html_body_file`) | More than one body source (nothing was changed) | Pass exactly one |
| `… needs a body: pass one of …` | A mail tool or `create_note` got no body at all | Pass `body`, `html_body`, `body_file` or `html_body_file` |
| `` `html: false` contradicts `html_body` `` | The deprecated `html` flag disagrees with `html_body` | Drop `html` |
| `body_file: could not read …` / `… is not valid UTF-8 text` | A `*_file` path is wrong or the file isn't UTF-8 text (nothing was changed) | Use an absolute path to a UTF-8 file |
| `update_draft needs at least one of…` | `update_draft` with nothing to change and no `send` | Pass at least one change, or `send: true` |
| `the draft has no recipients, so it was saved but not sent` | `update_draft` with `send: true` on a draft with no To/CC/BCC | Add recipients (`to`/`cc`/`bcc`) in the same call |
| `inline_images requires an HTML body` | `inline_images` used with a plain-text body | Use `html_body` and reference each image as `<img src="cid:ID">` |
| `duplicate inline image content_id` / `data_base64 is not valid base64` | Bad `inline_images` entry (nothing was created) | Give each image a unique `content_id` and valid base64, or use a `path` |
| `html_body: a image/… data: URI is not valid base64` / `has no data` / `unterminated` | A `data:` image in the HTML is broken (nothing was created) | Fix or remove that `<img src="data:...">` |
| `payload too large` (HTTP 413) | The HTTP request body is over 64 MiB | Write the body to a file and pass `body_file` / `html_body_file` |
| A large tool call fails while its JSON is being parsed, with no server error | The client cut the arguments short (often the model's output limit when it writes a very long body or base64 inline) | Write the content to a file and pass `html_body_file`/`body_file` |
| `Content-ID '…' not found. Available Content-IDs: …` | Wrong `content_id` for `get_inline_image` | Use one of the listed ids. The `cid:` prefix and letter case don't matter |
| `no attachments with a Content-ID` | The email has no inline images | `list_attachments` shows its regular attachments |
| `exceeds the 10 MB limit for get_inline_image` | The attachment is too big to return inline | Use `save_attachments` to write it to disk |
| `unknown include field "…"; valid fields are: …` | `include` names a field that tool doesn't have | Use one of the listed fields (`html_body`, `attachments` and `meeting` exist only on `get_email`) |
| `this item only has a plain-text body; body_format must be "text"` | `body_format: "html"` on `get_event`/`get_note`/`get_task` | Drop `body_format` |
| `` `prefer_html` (deprecated) is … but `body_format` is …`` | Both were passed and disagree | Pass only `body_format` |
| ``pass either `content_id` or `content_ids`, not both`` | `get_inline_image` got both | Use `content_id` for one image, `content_ids` for several |
| `output_dir "…" exists and is not a directory` / `Could not create output_dir` | Bad `output_dir` path | Give an absolute directory path the server can write to |
| `{"id": …, "error": …}` inside a list result | That one id failed in a batch read; the others are fine | Handle it per item (often a stale id: find the item again) |
| `empty_deleted_items permanently deletes EVERYTHING` | `confirm` was not true | Only pass `confirm: true` if the user explicitly asked to empty Deleted Items |
| `HRESULT 0x80040154` (class not registered) | Classic Outlook isn't installed; the "new Outlook" app has no COM | Tell the user the tools need classic Outlook |
| `HRESULT 0x80080005` (server execution failed) | Outlook isn't running or is hung, or only one of Outlook and the server runs as administrator | Ask the user to (re)start Outlook normally, with the same elevation as the MCP client |
| `HRESULT 0x80010001` (call rejected) | Outlook is busy or showing a dialog box | Ask the user to close the dialog, then retry once |

## Results that look wrong (no error)

- **Calendar or inbox is empty, or items come from the wrong dates**, especially on a
  day-first Windows locale (en-GB, en-IL, de-DE, …).
  - **Cause:** this version sends date filters in US month/day order, so Outlook can misread them.
  - **Confirm:** compare a date-filtered result with an unfiltered one.
  - **Workaround:** until a locale-aware build ships, fetch without the date filter and filter by date yourself.
- **Fewer emails than expected:**
  - `list_emails` returns one page: `count` defaults to 10, max 200. Page with `offset`.
  - A multi-word `query` requires every word (in any order); quote it (`"weekly update"`) for an exact phrase, or drop words.
  - The mail may also have been filed into a subfolder by a rule; check `list_folders`.
- **A Hebrew (or other non-Latin) search finds nothing:** the server falls back to
  scanning the 2,000 newest items that match the other filters. Older mail needs a date
  filter or a narrower folder.
- **Email body ends with `[... truncated at N characters]`:** `body_truncated` is
  true. Call `get_email` again with `max_body_chars` ≥ `body_length`, or pass
  `output_dir` to get the full body as a file.
- **`resolve_inline_images` left some `cid:` references:** they are listed in
  `inline_images_unresolved`: no attachment has that Content-ID, the image is over 10 MB,
  or it couldn't be read. Try `get_inline_image` on one to see the exact error.
- **Only 250 events:** that's the page cap. Page with `offset`, or use a smaller date window.
- **Only 500 tasks or notes:** that's the page cap. Page with `offset`.
- **`list_tasks` misses a task:** completed tasks are hidden unless `include_completed` is set.
- **Non-ASCII text shows as `????` or gibberish:** the server's output is always UTF-8
  and never converted to a Windows code page. Gibberish like `îééì` (Hebrew as
  windows-1255 read as latin1) or `×ž×™×™×œ` (UTF-8 read as latin1) comes from a
  console, wrapper script or client decoding it with the wrong code page, or from a
  mail stored with the wrong charset (then Outlook shows it garbled too). Don't try
  to repair it yourself; point the user to the README's troubleshooting section.

## Don't

- Don't report "you have no meetings/emails" from an empty result if it contradicts
  what the user expects, until you've ruled out the date and paging causes above.
- Don't loop on a COM error. Retry once after the user acts, then report the exact text.
