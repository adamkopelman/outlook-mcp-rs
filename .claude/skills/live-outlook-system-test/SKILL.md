---
name: live-outlook-system-test
description: Use when live-verifying outlook-mcp-rs tools against the real, running Outlook mailbox — after shipping a plan/feature that touches OutlookClient, or to sanity-check the whole tool surface end-to-end with real data.
---

# Live Outlook System Test

Live-test outlook-mcp-rs tools against the real, running Outlook mailbox with real data, real assertions, and real cleanup — not against `FakeOutlookClient`. This is heavier and slower than `tests/live_outlook.rs`'s per-feature `#[ignore]`d tests; use it for a *system*-level pass across many tools at once (e.g. "verify everything Plans 1-9 built"), not for a single function's regression test (that's what `tests/live_outlook.rs` is for).

**Core principle:** write the plan doc first, get it reviewed, execute as one non-panicking test that always cleans up, root-cause every failure for real, never leave test data behind.

## When to Use

- After a plan ships and you want end-to-end confidence beyond unit/fake-client tests.
- When the user asks to "test everything" / "system test" against their real mailbox.
- Periodically, as a health check on the live COM surface (Outlook API behavior drifts under real accounts in ways `FakeOutlookClient` can never catch).
- Before landing a batch of open PRs — see "Testing a batch of PRs" below.

Not for: a single function's regression coverage (add a normal `#[ignore]`d test to `tests/live_outlook.rs` instead).

## Process

### 1. Scope and write the plan doc first

Read `src/server.rs`'s `#[tool_router]` block to enumerate every tool in scope. Decide what's covered (e.g. "email + calendar, Plans 1-9") and what's explicitly out of scope (e.g. tools for a not-yet-shipped plan, or a tool that needs a second mailbox to test safely).

Write a plan doc (`SYSTEM_TEST_PLAN_<date>.md` at the repo root) **before writing any test code**, and get the user's sign-off before executing (unless they told you to run autonomously — then write the plan anyway, keep to previously-authorized addresses only, and call out every judgement call in the results doc) — they may want different seed data, different categories, a different scope. See `SYSTEM_TEST_PLAN_2026-07-16.md` (git history) as a worked example. The doc must have:

- **Purpose / Mechanism** — state plainly whether this runs through the actual MCP tool layer or calls `WindowsOutlookClient` directly. An already-running Claude Code session can't pick up a newly-registered MCP binary without restarting, so don't use the session's own MCP tools. Prefer serving the tool layer **in-process** instead: `OutlookMcpServer::new(Arc::new(WindowsOutlookClient::new()))`, called with JSON-RPC `tools/call` over an in-memory duplex pipe with rmcp's stdio codec (see `Wire` in `tests/system_test_prs_2026_10_17.rs`). Since 2026-10 the server layer has real logic (batching, `undo`, `require_no_attendees`, `fields`/`compact`/`output_dir`, deprecated aliases) that a direct `WindowsOutlookClient` call skips. Call the client directly only for something the tool layer doesn't expose.
- **Accounts used** — the real mailbox address, and any external test-recipient address the user has explicitly authorized for real sends/invites. Never send real mail/invites to an address you weren't explicitly told is safe to use.
- **A tagging convention** — a unique, greppable subject prefix (e.g. `[outlook-mcp-rs systest]`) plus the date, and far-future calendar dates (e.g. `2099-xx-xx`) so nothing can collide with real data. This is what makes cleanup safe and mechanical.
- **A seed-data phase** — if the mailbox is close to empty, filter tests against it prove nothing ("call didn't error" isn't "filter is correct"). Seed a small, deliberately varied batch first (different categories, flags, read-state, folders, show_as, all_day, etc.) so every filter has a real positive *and* negative case, and assert the exact expected result set, not just success. Use the mailbox's real category names (check with the user what they are — there's no tool to list them; Outlook's `Categories` property is freeform and not validated against the Master Category List, so made-up names also work, but real ones prove more).
- **One test per tool/behavior**, each with a concrete expected result — not "should work," but the actual expected value/status/field.
- **An explicit skip list with rationale** for anything not safely automatable (e.g. `respond_to_meeting` needs a second, independently-controlled mailbox to receive an invite into — don't respond to a real third party's real invite to work around this).
- **A cleanup checklist.**

### 2. Implement as one non-panicking test

One new `#[ignore]`d test function in `tests/system_test.rs` (or a dated sibling), following `tests/live_outlook.rs`'s house style (`fn client() -> WindowsOutlookClient { WindowsOutlookClient::new() }`). Run via `cargo test --test system_test -- --ignored --nocapture`.

**Critical:** this must not panic mid-run. Every step's outcome goes into a results log (`Vec<(&str, bool, String)>`), never an `assert!`/`.expect()` that would unwind past cleanup. Print progress as you go (you'll run with `--nocapture`). Structure:
1. Every step: `match`/`if let Err`, record pass/fail + reason, keep going.
2. Cleanup runs unconditionally at the end, regardless of what failed above.
3. A final summary table, printed last.
4. A real `assert!` only as the literal last line, after cleanup has already run, so the process exit code reflects overall success.

### 3. Write assertions that can actually fail

Before asserting a field changed as a result of an action, ask: could this field already hold a matching value for a reason unrelated to the action under test? If so, the assertion proves nothing — capture a baseline *before* the action and assert against that baseline afterward, not just "populated"/"non-null" after.

Real example: a test asserting `get_note` reports `modified` (`LastModificationTime`) after `update_note` checked only `modified.is_some()` post-update. But the item's own `create_note` call already does a `Save()`, which populates `LastModificationTime` on creation — so the assertion passed identically even with the `update_note` call deleted from the test entirely. It looked like a real regression test but could never fail from the thing it claimed to test. Caught by a task reviewer, not by running it. Fix: read the field once before the action, assert non-decreasing after (a strict "must be later" check risks flaking on coarse/same-second timestamp resolution — pair it with a second assertion, like the actually-edited content, that doesn't depend on timestamp granularity at all).

### 4. Execute, then root-cause every failure for real

Never hand-wave a live failure as "probably flaky." For each failure:
- Reproduce it with a **cheap, isolated, zero-write diagnostic** before spending another full run on it — a tiny throwaway `#[ignore]`d test hitting just the suspect call, or a raw PowerShell COM probe (`New-Object -ComObject Outlook.Application`) that bypasses this project's code entirely. This is how you tell a real code bug from an environment/account issue, and it's much cheaper than another 3-minute full run that sends more real mail.
- If it looks like a transient timing issue, prove it: retry the *exact same* property read on the *exact same* already-obtained COM object several times with delays. If it fails 100% of retries even minutes later, it's not timing — it's structural (e.g. Plan 9's system-test found that items yielded by `GetFirst`/`GetNext` after `Restrict`+`IncludeRecurrences` carry a `.Parent` whose `.StoreID` never resolves, deterministically — a real object-model gap, not sync lag).
- Delete all temporary diagnostic instrumentation (`eprintln!`s, throwaway test files) before committing — verify with `git status`/`git diff` that only the intended fix remains.

### 5. Known Outlook/COM gremlins to check for before blaming your code

- **"The operation failed" / every write fails, reads still work:** often an Outlook licensing/activation hiccup (window title may show "(Unlicensed Product)"). Confirm outside your code with a raw PowerShell `CreateItem`+`Save()` probe. Fix: **ask the user** before restarting Outlook (`$outlook.Quit()` then relaunch) — this is a real, moderately disruptive action (closes any open compose windows) and needs fresh authorization each time, not just because it was approved once earlier in the session.
- **"Unknown name" (`0x80020006` / `DISP_E_UNKNOWNNAME`) on property access:** can be genuinely transient (Cached Exchange Mode sync lag under heavy write load) or structural (see above). Distinguish before "fixing" with a retry — a bounded retry band-aids a transient cause but does nothing for a structural one, and you'll waste a verification cycle finding that out the hard way.
- **Cleanup silently falls behind after many repeated runs in one session:** if a cleanup sweep finds *what to delete* via a `count`-capped list query (e.g. `count: 25`), and the mailbox has accumulated more test debris than that cap from earlier incomplete runs, cleanup only ever touches the newest N and the rest compounds — run over run — until it looks like a mysterious, escalating "empty results" bug in an unrelated code path. If you've run the same live suite many times in one session and start seeing inexplicable empty-result failures, check the *raw* item count via COM before assuming a code regression; a capped cleanup sweep is a prime suspect. Recommended fix: give cleanup sweeps a much higher/uncapped limit than regular tool-facing queries.
- **Outlook.com: `Move(Deleted Items)` of an appointment, task or note returns a junk mail draft.** The item does move, but Outlook also creates a mail draft with the same subject in Drafts and returns *that* (Class 43, `Parent` = Drafts). Anything done to the returned object hits the junk draft. Mail items move normally. Check a raw COM probe's `Class`/`Parent` before trusting a `Move` result (2026-10-17, PR #41).
- **Outlook.com: DASL `LIKE` ignores an inner `%`.** `'%a b%'` works, but `'%a%b%'` silently matches nothing, with no error. Leading and trailing `%` are fine (2026-10-17, PR #39).
- **Outlook's own re-rendering is not a bug:** an appointment body comes back with `" \r\n"` or `"\r\n\r\n"` line breaks (it is stored as RTF); `get_email` has no `importance`/flag field (read them through the filters); drafts can't be flagged; `max_body_chars` is clamped to ≥1,000. Check these before filing a failure against the code.
- **A COM restart doesn't fix everything:** if identical failures persist across an Outlook restart, don't conclude "must be a code bug" — first rule out accumulated mailbox state (see above) with a raw COM item count, since a restart clears process-level state but not mailbox contents.

### 6. Cleanup discipline

- Sweep every folder you touched (Inbox, Drafts, Sent Items, Archive, Calendar, wherever items may have moved to), matching only your session's exact tag.
- Loop the sweep-and-delete until a pass finds 0 matches — a single pass over a live COM collection while deleting from it can skip items (index shifting mid-iteration).
- **Never touch anything that doesn't match your tag.** If you find other test debris (from older sessions, other tools, or ambiguous-looking real data), stop and ask the user rather than guessing — don't fold unrelated cleanup into your task.
- Confirm final state via a raw COM count/subject sweep, independent of this project's own (possibly-buggy) list_* code — that's your ground truth, not a self-reported "cleanup: PASS."

### 7. Fixing what you find

This skill covers testing and root-causing, not hasty inline fixes. Once a real code bug is confirmed and root-caused, fix it with the same rigor as any other change to this codebase: TDD where feasible, live re-verification of the specific broken behavior, and independent review of the diff (see `superpowers:subagent-driven-development`) before considering it done — a live-COM bug fix has real blast radius and deserves the same scrutiny as any other shipped change.

### 8. Report

Write a results doc (`SYSTEM_TEST_RESULTS_<date>.md`) with: a pass/fail table per test id, a root-caused writeup for every failure (not just "flaky"), confirmed final cleanup state, and any non-code findings worth flagging separately (e.g. a real mail-delivery bounce discovered along the way, unrelated to the tools themselves).

## Testing a batch of PRs

When the ask is "system-test all the open PRs", test the **merged** result once rather than each branch separately — the interesting bugs are the cross-PR ones.

1. **Isolate.** Create a worktree on a throwaway integration branch from the PRs' base (`git worktree add -b systest/all-prs-<date> .worktrees/systest-all-prs origin/main`). Never merge in the user's main checkout — it may hold uncommitted WIP. Fetch PRs by number (`git fetch origin pull/N/head:pr-N`), since fork branches aren't on `origin`.
2. **Merge stacked PRs in dependency order** (a base PR before the PRs built on it; check with `git merge-base --is-ancestor pr-A pr-B`), then the independent ones. After **each** merge run `cargo build --tests` so a break is attributed to the PR that caused it.
3. **Expect cross-PR fallout that git won't flag as a conflict**, and record each one as a finding (it must be fixed on one of the PRs before they land in any order):
   - a new field on a shared struct (`EmailQuery.offset`, `EmailQuery.to`) or a new trait-method argument (`delete_email(.., permanent)`, `get_email(.., max_body_chars)`, `create_draft(.., inline_images)`) breaks every struct literal / call site *another* PR added in `tests/`;
   - two PRs introducing the **same type name** for different things (2026-10-03: #22's `InlineImage` input vs #24's `InlineImage` result), which surfaces only as confusing type errors once both are in;
   - the same dependency added twice in `Cargo.toml` (duplicate key);
   - both PRs editing the same README tool-list line or the "N tools" count, which must become the merged total;
   - two PRs both rewriting the same hot function (`list_emails`'s filter loop): merge by hand and re-read the result, since a mechanical union can silently drop one PR's behaviour.
4. Scope the plan doc to the **new and changed behaviour per PR** (one test id prefix per PR, so the results table maps straight back to PRs), plus a short regression smoke of the core tools the PRs touched.
5. Report per PR: merged cleanly / conflict resolved / needs a change before it can land, with the fix commit on the integration branch. Don't push to anyone's PR branch or comment on GitHub without the user asking.
6. **Landing** (only when the user asks): land the PRs one at a time, in the same order you merged them. Before each one, merge the *real* current `main` into that PR's branch, resolving to the tree you already verified on the integration branch. Put PR-specific fixes (a bug the test found, a rename that avoids a clash) in their own commits on that PR, not hidden inside a merge commit. Order-independent fixes (e.g. a bug in one PR, a rename) can be pushed before anything is merged; conflict resolutions can't, because they depend on what is already in `main`. Merging to `main` and pushing to a contributor's branch are both outward-facing, so they need explicit authorization each time — if a permission check blocks them, stop and hand the prepared branches to the user rather than routing around it.
7. Keep the system test in its own PR, opened after (or explicitly dependent on) the feature PRs, since it won't compile on `main` until they land.

## Destructive and content-sensitive tools

- **Permanent delete** (`delete_email(permanent=true)`): only ever on items your test created and tagged. Verify afterwards that the item is in neither its folder nor Deleted Items. Use it in cleanup sweeps for systest items only, so cleanup stops adding to Deleted Items.
- **`empty_deleted_items(confirm=true)` is never run live** unless the user explicitly says the Deleted Items folder may be emptied — it destroys every item there, including the user's own. Live-test only the `confirm=false` refusal path (and that it leaves the folder count unchanged); leave the rest to the fake-client tests.
- **Non-ASCII (Hebrew etc.):** round-trip through COM and assert exact string equality, not `contains` on a lossy rendering. For search, seed an item whose Hebrew word appears **only** in the body (not the subject) as well as one with it in the subject, so both the DASL path and the client-side fallback are exercised, plus a Hebrew negative case.
- **Inline images / attachments:** build test images yourself (a 1×1 PNG in base64 is enough) and attach to **drafts** addressed to `nobody@example.invalid`; you never need to send mail to test attachment metadata, `is_inline`, `get_inline_image`, or `context_lines`. Assert the decoded bytes equal the bytes you put in, not just that a data URI came back.

## Worked example

The PR-batch run of 2026-10-17 (PRs #37–#56: a linear stack of 11 plus one independent PR) is the reference for the in-process tool-layer mechanism and for two Outlook.com gremlins (above): `SYSTEM_TEST_PLAN_2026-10-17-PRS.md`, `tests/system_test_prs_2026_10_17.rs`, `SYSTEM_TEST_RESULTS_2026-10-17-PRS.md`. Its first live pass was the PRs' **own** new live tests (they had never run on Windows), which found both real bugs before the system test itself ran.

The PR-batch run (2026-10-03, PRs #14–#25) is the reference for "Testing a batch of PRs": `docs/superpowers/plans/2026-10-03-systest-all-open-prs.md`, `SYSTEM_TEST_PLAN_2026-10-03-PRS.md`, `tests/system_test_prs.rs`, `SYSTEM_TEST_RESULTS_2026-10-03-PRS.md`.

Plans 1-9's live system test (2026-07-16) is the reference implementation of this whole process: `SYSTEM_TEST_PLAN_2026-07-16.md`, `tests/system_test.rs`, `SYSTEM_TEST_RESULTS_2026-07-16.md`, and the fix commit `05904c5` (with its report at `.superpowers/sdd/systest-findings-fixes-report.md`) — including a real example of the "original hypothesis was wrong, re-investigate" pattern (Finding 3 was first assumed to be transient sync lag; live instrumentation proved it was a structural `Parent.StoreID` gap instead) and the "cleanup fell behind after many runs" gremlin (42 stray inbox items, 304 stray drafts, traced to a capped cleanup query — not a code regression).
