# System test results: open PRs #37–#56 merged together, 2026-10-17

- **Plan:** `SYSTEM_TEST_PLAN_2026-10-17-PRS.md`.
- **Branch:** `systest/all-prs-2026-10-17` (worktree `.worktrees/systest-1017`), based on `origin/main` @ `26d6570`.
- **Test:** `tests/system_test_prs_2026_10_17.rs`.
- **Mailbox:** `adamkopelman@outlook.com` (Outlook.com), en-US regional format.

## Outcome

- **System test: 63/63 PASS** on the final tree. Every call went through the real MCP tool layer: JSON-RPC `tools/call` over the stdio codec, to `OutlookMcpServer` over a real `WindowsOutlookClient`.
  - The first run was 59/63. All four failures were wrong expectations in my own test (details below).
- **The PRs' own new live tests: 28/28 PASS** (27 new ones plus #16's `permanent_delete_of_draft_skips_deleted_items`).
  - The first run was 22/27. The failures were **two real bugs**, in #41 and #39. Both are fixed.
- **Offline suite (CI's command, `cargo test --all -- --skip live_outlook`) passes:** lib 324, tools 152, http_transport 6.
- **Hook tests (`python -m unittest discover hooks`): 16/16** after a Windows portability fix to one test (#55).
- **Mailbox afterwards:** identical to the baseline in every folder (raw COM, see Cleanup).

## Per-PR status

| PR | Topic | Merge | Live result | Change made before landing |
|---|---|---|---|---|
| #56 | CHANGELOG | fast-forward | n/a (docs) | — |
| #37 | write tools, `*_file`, `data:` images, `update_draft send` | fast-forward | W1–W5 PASS (incl. one real send), 4 PR tests PASS | test cleanup made permanent (reply draft was untaggable debris) |
| #38 | batch reads, include/output_dir, `get_task` | fast-forward | R1–R4 PASS, 6 PR tests PASS | — |
| #39 | date grammar, list conventions, paging | fast-forward | L1–L9 PASS; **PR test FAILED → fixed** | **yes: `*` wildcard on `list_emails` (below)** |
| #40 | UTF-8 hardening | fast-forward | U1 PASS, mojibake diagnostic PASS | — |
| #41 | permanent delete on events/tasks/notes | fast-forward | **4 PR tests FAILED → fixed**; D1, D2 PASS | **yes: permanent delete on Outlook.com (below)** |
| #49 | fields/compact/output_dir, user role | fast-forward | O1–O4 PASS, PR test PASS | — |
| #50 | current ids, occurrence ids | fast-forward | I1–I4 PASS, PR test PASS | test cleanup made permanent |
| #51 | batch writes + undo | fast-forward | B1–B5 PASS, PR test PASS | test cleanup made permanent |
| #53 | `require_no_attendees`, `notify`, annotations, hook | fast-forward | S1–S5 PASS | — (open risk: `respond_to_meeting notify:false`, below) |
| #55 | Claude Code plugin | fast-forward | launcher checked live under Git Bash | **yes: one hook test was not runnable on Windows** |
| #42 | locale-proof date filters | **conflict with #39** (client.rs ×5, server.rs ×2, mod.rs, CHANGELOG, skill, live tests) | Q1–Q3 PASS, 3 PR tests PASS | merge resolution (below) |

"Fast-forward": the stack is linear and every PR contains its predecessor, so #56 → #55 fast-forward onto `main`.

## Real bugs found and fixed

### #41: permanent delete of events, tasks and notes left them in Deleted Items, plus a junk draft

- **Symptom:**
  - `permanent_delete_{task,note,event}_skips_deleted_items` and `permanent_delete_of_unsent_meeting…` all failed with "a permanently deleted item must not remain in Deleted Items".
  - A raw COM sweep afterwards also found a **mail draft** with each probe's subject in Deleted Items, which nothing had asked for.
- **Root cause:**
  - On this Outlook.com store, `Move(Deleted Items)` on an appointment, task or note does move the item. It also **creates a junk mail draft with the same subject in Drafts, and returns that draft** (Class 43, `Parent` = Drafts) instead of the moved item.
  - #41's shared `delete_item` reused #16's mail fix: re-open the moved item by EntryID and `Delete()` it. So it soft-deleted the junk draft and never touched the real item.
- **How it was confirmed:** raw PowerShell COM probes, outside this codebase.
  - For every one of task, note and appointment: Move returned Class 43 with parent Drafts, the Drafts count went up by one, and the real item sat in Deleted Items.
  - A plain soft `Delete()` creates no junk draft.
  - Deleting the junk object, or `GetItemFromID(moved.EntryID)`, both left the real item behind.
- **Finding the item again after a soft delete:**
  - `PR_SEARCH_KEY` survives the delete. So did a named property I stamped on the item, but a DASL `Restrict` on that named property found nothing on this store.
  - Finding it by search key among same-`MessageClass` items, newest first, hit on the first candidate for mail, appointment, task and note (about 40 ms each).
- **Fix:** commit "Fix permanent delete of events, tasks and notes on Outlook.com (PR #41)".
  - A pure `permanent_delete_plan` (unit-tested) picks one of three paths: `Direct` (already in Deleted Items), `MoveThenDelete` (mail, #16's live-verified path, unchanged) or `SoftDeleteThenPurge` (everything else).
  - The search key is read **before** the delete, so if it can't be read nothing is deleted.
  - The live helper now also asserts that **no draft is left in Drafts**.
- **Verification:**
  - The four PR tests pass, including the new Drafts check.
  - #16's mail test still passes.
  - System tests D1 (a batch of 2 tasks) and D2 (a note and an event) pass, with nothing in Deleted Items or Drafts.

### #39: `list_emails` `*` wildcard found nothing on Outlook.com

- **Symptom:** `list_emails_subject_scope_and_phrase_find_a_recent_subject` failed: `subject:Upcoming*changes` didn't find "Upcoming changes to your Microsoft 365 subscription".
- **Root cause:** confirmed with a zero-write raw COM probe. This store's DASL `LIKE` only honours `%` at the **start and end** of a pattern, and silently returns 0 with no error otherwise:

  | pattern | matches |
  |---|---|
  | `'%Upcoming changes%'` | 1 |
  | `'%Upcoming%changes%'` | 0 |
  | `'%Microsoft%subscription%'` | 0 |

  #39 turned `*` into an inner `%`.
- **Fix:** commit "Make list_emails' * wildcard work on Outlook.com's DASL (PR #39)".
  - `to_dasl` splits a term on `*` into pieces that must all occur in the same column: `(col LIKE '%a%' AND col LIKE '%b%') OR …`. That is a store-safe superset.
  - `list_emails` then re-checks those hits client-side with `TextQuery::matches`, which keeps the order. It only reads the already-narrowed hits.
  - A term that is only `*` adds no clause, where it would otherwise have produced an unparseable `( OR )`.
  - Events, tasks and notes match client-side already, so they were unaffected.
- **Verification:**
  - The PR test passes.
  - System test L8: `subject:alpha*beta` finds "draft alpha beta" and **not** "draft beta alpha".
  - Unit tests cover the new DASL shape, leading/trailing `*`, and a `*`-only term.

### #55: a hook test could not run on Windows

`test_launcher_asks_when_no_python_is_found` spawned `"/bin/sh"`, which native Windows Python can't resolve, so it raised `FileNotFoundError`. The fix resolves `sh` with `shutil.which` before the test empties `PATH`.

The launcher itself was then checked under Git Bash, the way Claude Code runs plugin hooks on Windows, with real hook events and `CLAUDE_PLUGIN_ROOT` / `CLAUDE_PLUGIN_OPTION_CALENDAR_AUTOPILOT` set:

| hook event | decision |
|---|---|
| `list_events` | allow |
| `respond_to_meeting` with `notify:false` | allow |
| `respond_to_meeting` (notifying) | ask |
| `send_email` | ask |
| `update_event` with `require_no_attendees` | allow |
| any event, mode off | no output (normal prompts) |

On this machine `python3` is the Microsoft Store stub, which exits 49. The launcher correctly skips it and uses `python`.

## Cross-PR integration: #42 vs #39

#39 resolves every date through its new grammar (`DateRange`), but still wrote the `Restrict` strings with `jet_datetime`. #42 replaces those strings with an unambiguous date-only window plus exact per-item bounds. Git conflicted in 5 hunks of `client.rs`.

**Resolution (commit "Merge #42 …"):**

- **Kept from #39:**
  - the parsing;
  - `since_days`, already mapped to `received_after` in the server layer, so #42's own `since_days` / `parse_dt` code is dropped;
  - `item_type` / `flag` / query / `to` filtering;
  - event paging with `offset` / `count`;
  - `EventRows` / user role.
- **Taken from #42:**
  - the bounds go into `DateFilter` / `restrict_dates`;
  - the exact-range check runs per item, first, before #39's filters and before `offset`;
  - on `list_emails` the scan stops at the first item older than the range;
  - `Scan` replaces `Option`.
- **#42's three live tests** are ported to the merged API:
  - `start_after` / `start_before`;
  - `EmailQuery` defaults;
  - `received_after: "-30d"` instead of `since_days`;
  - permanent cleanup.
- **Docs:** README, TESTING.md and a live-test comment said issue #1 was still open. They are updated.

## Wrong expectations in my own test (first run, corrected)

- **W1b:** I expected `get_email` to return `importance`. Email results have no importance field at all (that predates these PRs). The test now reads it back through #39's `importance` filter.
- **W2-event:** I expected the event body to equal the file text. Outlook re-renders an appointment's body (stored as RTF): line breaks come back as `" \r\n"` after an ASCII line and `"\r\n\r\n"` after a non-ASCII one. A raw COM probe with no project code reproduces it. Task and note bodies round-trip exactly. The test now compares words.
- **R2b:** I passed `max_body_chars: 10`. The documented range is 1,000–5,000,000, so 10 is clamped up. The test now uses a 3,000-character body with 1,000.
- **B3:** I flagged a draft. Outlook refuses: "Draft items cannot be marked". The test now uses W4's received self-loop copy and reads flag and importance through the filters.

## Not tested live (by design)

- **`respond_to_meeting`:** answering needs an invite from a second, independently controlled mailbox. **This leaves an open risk for #53.** The autopilot hook auto-approves `respond_to_meeting` with `notify: false`, assuming `Respond(…, fNoUI=True)` + `Save()` reaches nobody, and the #53 author notes that Microsoft's docs contradict each other on this. Verify it once with a real test invite before enabling the autopilot. #50's new id after accepting is untested for the same reason.
- **Day-first locales (#42):** only en-US was run. Q1/Q2 seed a date and its day/month swap (2099-03-04 / 2099-04-03) and each range returned only the right one. A run under en-GB/en-IL/de-DE needs the Windows Region changed and Outlook restarted.
- **`empty_deleted_items(confirm=true)`:** never run live.
- **The plugin installed in a real Claude Code session (#55):** not done here.

## Judgement calls (autonomous run)

- **Mechanism:** through the tool layer in-process, not the client directly. It covers the server-layer features (batching, `undo`, `require_no_attendees`, output controls, aliases), which a client-only test can't reach.
- **One real send (W4):** a self-loop, to the address every earlier system test used.
- **Live-test cleanup:** six of the PRs' live tests switched from soft to permanent delete, so the suite stops leaving debris. One of them (#37's reply test) left an **untagged** `RE: <real subject>` draft in Deleted Items, which I removed after confirming it was the unsent reply that test created 1 hour earlier.
- **Scope of the #41 fix:** mail keeps #16's Move path. Every other class now takes `SoftDeleteThenPurge`. That includes meeting requests and reports deleted through `delete_email`, which used the Move path before. That path is item-class agnostic and was verified live for mail, appointment, task and note, but **not** for meeting requests or reports.
- **`clippy --fix` reverted:** it also rewrote three library files I hadn't touched, so I reverted those and kept only its fixes to the new test. My changes add no clippy warnings.

## Cleanup

- Every run deleted what it created, through the tools, with `permanent: true`, then swept by run token until a pass found nothing.
- **Debris from the first live run, all from this session, all removed:**
  - left by the #41 bug: 25 tagged items in Deleted Items, including 7 junk drafts;
  - left by my raw COM probes: 4 drafts in Drafts;
  - the one untagged reply draft above.
- **Final raw-COM state, identical to the baseline taken before the first run:**

  | Inbox | Drafts | Sent | Deleted Items | Outbox | Calendar | Tasks | Notes |
  |---|---|---|---|---|---|---|---|
  | 48 | 12 | 213 | 53 | 0 | 19 | 4 | 7 |

  0 tagged items in any of them.

## Other observations (not failures)

- **Importance can be set but not read** on email: `create_draft` / `update_email` set it, and no result field reports it; only the `importance` filter reveals it. The same goes for an email's flag. Worth adding both to `get_email` / `list_emails` rows.
- **Search-key collisions:** a copy of an item shares its `PR_SEARCH_KEY`. If a copy is already in Deleted Items, `SoftDeleteThenPurge` deletes the newest same-class match, which is the item just deleted in practice (it was the first candidate every time).
- **Issue #31 (mojibake)** stays open: #40 hardens HTTP mode and adds a diagnostic, but doesn't claim to fix it. The diagnostic passes on this mailbox, so the garbling, where it happens, comes from after the server.
