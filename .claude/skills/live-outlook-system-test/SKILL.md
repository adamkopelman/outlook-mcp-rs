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

Not for: a single function's regression coverage (add a normal `#[ignore]`d test to `tests/live_outlook.rs` instead).

## Process

### 1. Scope and write the plan doc first

Read `src/server.rs`'s `#[tool_router]` block to enumerate every tool in scope. Decide what's covered (e.g. "email + calendar, Plans 1-9") and what's explicitly out of scope (e.g. tools for a not-yet-shipped plan, or a tool that needs a second mailbox to test safely).

Write a plan doc (`SYSTEM_TEST_PLAN_<date>.md` at the repo root) **before writing any test code**, and get the user's sign-off before executing — they may want different seed data, different categories, a different scope. See `SYSTEM_TEST_PLAN_2026-07-16.md` (git history) as a worked example. The doc must have:

- **Purpose / Mechanism** — state plainly whether this runs through the actual MCP tool layer or calls `WindowsOutlookClient` directly. In practice it's almost always direct: an already-running Claude Code session can't pick up a newly-registered MCP binary without restarting, so route through `src/outlook/client.rs` directly (same code the tool layer calls one layer down — `server.rs`'s tool methods are thin wrappers with no logic of their own).
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

### 3. Execute, then root-cause every failure for real

Never hand-wave a live failure as "probably flaky." For each failure:
- Reproduce it with a **cheap, isolated, zero-write diagnostic** before spending another full run on it — a tiny throwaway `#[ignore]`d test hitting just the suspect call, or a raw PowerShell COM probe (`New-Object -ComObject Outlook.Application`) that bypasses this project's code entirely. This is how you tell a real code bug from an environment/account issue, and it's much cheaper than another 3-minute full run that sends more real mail.
- If it looks like a transient timing issue, prove it: retry the *exact same* property read on the *exact same* already-obtained COM object several times with delays. If it fails 100% of retries even minutes later, it's not timing — it's structural (e.g. Plan 9's system-test found that items yielded by `GetFirst`/`GetNext` after `Restrict`+`IncludeRecurrences` carry a `.Parent` whose `.StoreID` never resolves, deterministically — a real object-model gap, not sync lag).
- Delete all temporary diagnostic instrumentation (`eprintln!`s, throwaway test files) before committing — verify with `git status`/`git diff` that only the intended fix remains.

### 4. Known Outlook/COM gremlins to check for before blaming your code

- **"The operation failed" / every write fails, reads still work:** often an Outlook licensing/activation hiccup (window title may show "(Unlicensed Product)"). Confirm outside your code with a raw PowerShell `CreateItem`+`Save()` probe. Fix: **ask the user** before restarting Outlook (`$outlook.Quit()` then relaunch) — this is a real, moderately disruptive action (closes any open compose windows) and needs fresh authorization each time, not just because it was approved once earlier in the session.
- **"Unknown name" (`0x80020006` / `DISP_E_UNKNOWNNAME`) on property access:** can be genuinely transient (Cached Exchange Mode sync lag under heavy write load) or structural (see above). Distinguish before "fixing" with a retry — a bounded retry band-aids a transient cause but does nothing for a structural one, and you'll waste a verification cycle finding that out the hard way.
- **Cleanup silently falls behind after many repeated runs in one session:** if a cleanup sweep finds *what to delete* via a `count`-capped list query (e.g. `count: 25`), and the mailbox has accumulated more test debris than that cap from earlier incomplete runs, cleanup only ever touches the newest N and the rest compounds — run over run — until it looks like a mysterious, escalating "empty results" bug in an unrelated code path. If you've run the same live suite many times in one session and start seeing inexplicable empty-result failures, check the *raw* item count via COM before assuming a code regression; a capped cleanup sweep is a prime suspect. Recommended fix: give cleanup sweeps a much higher/uncapped limit than regular tool-facing queries.
- **A COM restart doesn't fix everything:** if identical failures persist across an Outlook restart, don't conclude "must be a code bug" — first rule out accumulated mailbox state (see above) with a raw COM item count, since a restart clears process-level state but not mailbox contents.

### 5. Cleanup discipline

- Sweep every folder you touched (Inbox, Drafts, Sent Items, Archive, Calendar, wherever items may have moved to), matching only your session's exact tag.
- Loop the sweep-and-delete until a pass finds 0 matches — a single pass over a live COM collection while deleting from it can skip items (index shifting mid-iteration).
- **Never touch anything that doesn't match your tag.** If you find other test debris (from older sessions, other tools, or ambiguous-looking real data), stop and ask the user rather than guessing — don't fold unrelated cleanup into your task.
- Confirm final state via a raw COM count/subject sweep, independent of this project's own (possibly-buggy) list_* code — that's your ground truth, not a self-reported "cleanup: PASS."

### 6. Fixing what you find

This skill covers testing and root-causing, not hasty inline fixes. Once a real code bug is confirmed and root-caused, fix it with the same rigor as any other change to this codebase: TDD where feasible, live re-verification of the specific broken behavior, and independent review of the diff (see `superpowers:subagent-driven-development`) before considering it done — a live-COM bug fix has real blast radius and deserves the same scrutiny as any other shipped change.

### 7. Report

Write a results doc (`SYSTEM_TEST_RESULTS_<date>.md`) with: a pass/fail table per test id, a root-caused writeup for every failure (not just "flaky"), confirmed final cleanup state, and any non-code findings worth flagging separately (e.g. a real mail-delivery bounce discovered along the way, unrelated to the tools themselves).

## Worked example

Plans 1-9's live system test (2026-07-16) is the reference implementation of this whole process: `SYSTEM_TEST_PLAN_2026-07-16.md`, `tests/system_test.rs`, `SYSTEM_TEST_RESULTS_2026-07-16.md`, and the fix commit `05904c5` (with its report at `.superpowers/sdd/systest-findings-fixes-report.md`) — including a real example of the "original hypothesis was wrong, re-investigate" pattern (Finding 3 was first assumed to be transient sync lag; live instrumentation proved it was a structural `Parent.StoreID` gap instead) and the "cleanup fell behind after many runs" gremlin (42 stray inbox items, 304 stray drafts, traced to a capped cleanup query — not a code regression).
