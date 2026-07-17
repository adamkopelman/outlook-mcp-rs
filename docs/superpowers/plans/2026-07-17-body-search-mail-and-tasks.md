# Body-Text Search for Mail and Tasks Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make `list_tasks`'s `query` filter match against a task's real body text (not just its subject), and prove — with a real, positive-match live test — that `list_emails`'s `query` filter already does the same for emails.

**Architecture:** `list_emails` already runs a server-side DASL `@SQL` restrict that includes `"urn:schemas:httpmail:textdescription"` (the email body) alongside subject/sender (`src/outlook/client.rs:757-766`), but no test — unit or live — has ever asserted a positive match on body-only content; the existing live test (`list_emails_query_filter_narrows_results`) only asserts a nonsense query returns `<=` all results, which would pass even if body search were silently broken. `list_tasks` has no body search at all: `task_matches` (`src/outlook/client.rs:461-480`) only checks `summary.subject`, and `TaskQuery.query`'s own doc comment says why — *"TaskSummary has no body field to match"* (`src/outlook/mod.rs:110`). This plan (1) adds a real positive-match live test that proves email body search works today, and (2) brings task search up to the same standard as `list_notes`, which already reads a note's real `Body` property per-item and matches `query` against it (`note_matches`, `src/outlook/client.rs:486-497`, and `list_notes`, `src/outlook/client.rs:1798-1824`) — the exact pattern this plan copies for tasks.

**Tech Stack:** Rust, Windows COM (via the existing `windows`/COM helper functions already used throughout `src/outlook/client.rs`), `cargo test` for unit tests, `cargo test --test live_outlook -- --ignored` for live-Outlook regression tests.

## Global Constraints

- No new dependencies — this reuses `get_property`, `variant_to_string`, and the existing filter-function pattern (`task_matches`/`note_matches`/`event_matches`) already in `src/outlook/client.rs`.
- Live tests are `#[ignore]`d, live in `tests/live_outlook.rs`, and follow the house style already there: a `[outlook-mcp-rs ... live]`-tagged subject for greppability, and unconditional cleanup (`delete_email`/`delete_task`) even when an assertion could fail — call the cleanup delete *before* the `assert!`, not after, so a failing assertion still leaves the mailbox clean (see `TESTING.md`).
- Matching stays case-insensitive and substring-based, consistent with every other `_matches` function in this file (`event_matches`, `note_matches`, existing `task_matches`).

---

### Task 1: Prove `list_emails` already matches on real body content (live test) + document it

Today, nothing proves the DASL body clause actually works — only that a nonsense query doesn't over-match. This task adds a live test with a genuine positive assertion, and updates the README line that currently undersells this capability.

**Files:**
- Modify: `tests/live_outlook.rs` (add a new test after `list_emails_query_filter_narrows_results`, currently ending at line 289)
- Modify: `README.md:26`
- Modify: `TESTING.md` (add a short paragraph documenting the new live test, after the `list_events`/`calendar_of` paragraph ending at line 76, before the `list_tasks` paragraph at line 78)

**Interfaces:**
- Consumes: `WindowsOutlookClient::create_draft`, `::list_emails(EmailQuery)`, `::delete_email` — all existing, unchanged signatures.
- Produces: nothing consumed by later tasks (independent of Task 2/3).

- [ ] **Step 1: Write the live test**

Add to `tests/live_outlook.rs`, directly after `list_emails_query_filter_narrows_results` (after line 289):

```rust
#[test]
#[ignore]
fn list_emails_query_matches_real_body_text() {
    let c = client();
    let token = "zzbodytoken8842";
    let created = c.create_draft(
        vec!["nobody@example.invalid".to_string()],
        "[outlook-mcp-rs body-search live] draft probe".to_string(),
        format!("this draft's body contains {token} and the subject does not"),
        None, None, false, None,
    ).expect("create_draft should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let found = c.list_emails(EmailQuery {
        query: Some(token.to_string()), folder: "drafts".into(), count: 25,
        unread_only: false, from: None, category: None, received_after: None,
        received_before: None, since_days: None, has_attachments: None,
        flagged: false, high_importance: false,
    }).expect("list_emails query should succeed");

    c.delete_email(id.clone()).expect("cleanup: delete the draft");

    assert!(
        found.iter().any(|e| e.id == id),
        "list_emails query {token:?} should find a draft whose ONLY occurrence \
         of that token is in the body, proving the existing @SQL textdescription \
         clause matches body content and not just subject/sender"
    );
}
```

- [ ] **Step 2: Run it against the real, running Outlook**

Run: `cargo test --test live_outlook -- --ignored list_emails_query_matches_real_body_text --nocapture`

Expected: `test result: ok. 1 passed` — this confirms the existing DASL body clause genuinely works. If it instead **fails** (the draft isn't found), stop here: that's a real, separate bug in the existing `@SQL` restrict (possibly the same class of DASL-Restrict unreliability already documented for `flagged` at `src/outlook/client.rs:783-791`), and needs its own root-cause pass with `superpowers:systematic-debugging` before continuing — do not proceed to Task 2 on top of an unverified assumption, and do not attempt a silent workaround here.

- [ ] **Step 3: Document the (now-proven) behavior in README**

In `README.md`, change line 26 from:

```markdown
- `list_emails` — find emails in a folder with an optional text query and filters (sender, category, date range, attachments, flagged, importance)
```

to:

```markdown
- `list_emails` — find emails in a folder with an optional text query (matches subject, sender, and body) and filters (category, date range, attachments, flagged, importance)
```

- [ ] **Step 4: Document the new live test in TESTING.md**

In `TESTING.md`, insert this new paragraph directly after the `list_events`/`calendar_of` paragraph (which ends at line 76, right before the `list_tasks` paragraph currently at line 78):

```markdown
`list_emails`'s `query` filter matching real email body text (not just
subject/sender) is covered by the live suite:
`cargo test --test live_outlook -- --ignored list_emails_query_matches_real_body_text`.
```

- [ ] **Step 5: Commit**

```bash
git add tests/live_outlook.rs README.md TESTING.md
git commit -m "Add live proof that list_emails query matches real body text"
```

---

### Task 2: Make `list_tasks`'s `query` match a task's body, not just its subject

This is the actual feature gap. Follows the exact pattern `note_matches`/`list_notes` already use: read the item's real `Body` property once per item while iterating, and pass it into the matcher alongside the summary.

**Files:**
- Modify: `src/outlook/client.rs:461-480` (`task_matches`)
- Modify: `src/outlook/client.rs:1623-1650` (`list_tasks`)
- Modify: `src/outlook/mod.rs:99-121` (`TaskQuery`/`NoteQuery` doc comments)
- Modify: `README.md:45`
- Test: new `task_filter_tests` module appended to `src/outlook/client.rs` (after the existing `event_filter_tests` module, which currently ends at line 2342)

**Interfaces:**
- Consumes: `TaskSummary` (`src/outlook/types.rs:150-158`, unchanged), `TaskQuery` (`src/outlook/mod.rs:106-111`, unchanged fields — only the doc comment on `query` changes).
- Produces: `fn task_matches(body: &str, summary: &TaskSummary, q: &TaskQuery) -> bool` — the new signature (body is now the *first* parameter, matching `note_matches`'s parameter order). No other task in this plan calls `task_matches` directly, but note this signature for anyone extending task filtering later.

- [ ] **Step 1: Write the failing unit tests**

Append to `src/outlook/client.rs`, after the closing `}` of `event_filter_tests` (after line 2342):

```rust
#[cfg(test)]
mod task_filter_tests {
    use super::*;

    fn base() -> TaskSummary {
        TaskSummary {
            id: "test-id|store-id".to_string(),
            subject: "Quarterly Report".to_string(),
            due_date: Some("2026-06-10T00:00:00".to_string()),
            complete: false,
            status: "not_started".to_string(),
            importance: "normal".to_string(),
            categories: vec!["Work".to_string()],
        }
    }

    #[test]
    fn empty_query_matches_any_summary() {
        let summary = base();
        let query = TaskQuery::default();
        assert!(task_matches("some body text", &summary, &query));
    }

    #[test]
    fn query_substring_matches_subject() {
        let summary = base();
        let query = TaskQuery { query: Some("quarterly".to_string()), ..Default::default() };
        assert!(task_matches("unrelated body", &summary, &query));
    }

    #[test]
    fn query_substring_matches_body() {
        let summary = base();
        let query = TaskQuery { query: Some("budget numbers".to_string()), ..Default::default() };
        assert!(task_matches("here are the budget numbers for review", &summary, &query));
    }

    #[test]
    fn query_substring_no_match_in_subject_or_body() {
        let summary = base();
        let query = TaskQuery { query: Some("nonexistent".to_string()), ..Default::default() };
        assert!(!task_matches("also nothing here", &summary, &query));
    }

    #[test]
    fn query_substring_case_insensitive_against_body() {
        let summary = base();
        let query = TaskQuery { query: Some("BUDGET NUMBERS".to_string()), ..Default::default() };
        assert!(task_matches("here are the budget numbers", &summary, &query));
    }

    #[test]
    fn empty_query_string_is_noop() {
        let summary = base();
        let query = TaskQuery { query: Some("".to_string()), ..Default::default() };
        assert!(task_matches("anything", &summary, &query));
    }

    #[test]
    fn category_filter_still_applies_alongside_body_query() {
        let summary = base();
        let query = TaskQuery {
            query: Some("budget".to_string()),
            category: Some("Personal".to_string()),
            ..Default::default()
        };
        assert!(!task_matches("budget numbers", &summary, &query));
    }

    #[test]
    fn importance_filter_still_applies_alongside_body_query() {
        let summary = base();
        let query = TaskQuery {
            query: Some("budget".to_string()),
            importance: Some("high".to_string()),
            ..Default::default()
        };
        assert!(!task_matches("budget numbers", &summary, &query));
    }
}
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `cargo test task_filter_tests`

Expected: a compile error — `task_matches` currently takes 2 arguments (`summary`, `q`), not 3. Something like:
```
error[E0061]: this function takes 2 arguments but 3 arguments were supplied
```

- [ ] **Step 3: Update `task_matches` to accept and search the body**

In `src/outlook/client.rs`, replace (lines 459-480):

```rust
/// Client-side filter for `list_tasks`'s `category`/`importance`/`query`.
/// `include_completed` is applied earlier via `Restrict`, not here.
fn task_matches(summary: &TaskSummary, q: &TaskQuery) -> bool {
    if let Some(query) = q.query.as_deref().filter(|s| !s.is_empty()) {
        let needle = query.to_lowercase();
        if !summary.subject.to_lowercase().contains(&needle) {
            return false;
        }
    }
    if let Some(cat) = q.category.as_deref().filter(|s| !s.is_empty()) {
        let want = cat.to_lowercase();
        if !summary.categories.iter().any(|c| c.to_lowercase() == want) {
            return false;
        }
    }
    if let Some(imp) = q.importance.as_deref().filter(|s| !s.is_empty()) {
        if !summary.importance.eq_ignore_ascii_case(imp) {
            return false;
        }
    }
    true
}
```

with:

```rust
/// Client-side filter for `list_tasks`'s `category`/`importance`/`query`.
/// `include_completed` is applied earlier via `Restrict`, not here. `query`
/// matches either the subject or the real task body — `body` is the task's
/// real, untruncated body text, read once per item by the caller (see
/// `list_tasks` below), the same pattern `note_matches` uses for notes.
fn task_matches(body: &str, summary: &TaskSummary, q: &TaskQuery) -> bool {
    if let Some(query) = q.query.as_deref().filter(|s| !s.is_empty()) {
        let needle = query.to_lowercase();
        if !summary.subject.to_lowercase().contains(&needle)
            && !body.to_lowercase().contains(&needle)
        {
            return false;
        }
    }
    if let Some(cat) = q.category.as_deref().filter(|s| !s.is_empty()) {
        let want = cat.to_lowercase();
        if !summary.categories.iter().any(|c| c.to_lowercase() == want) {
            return false;
        }
    }
    if let Some(imp) = q.importance.as_deref().filter(|s| !s.is_empty()) {
        if !summary.importance.eq_ignore_ascii_case(imp) {
            return false;
        }
    }
    true
}
```

- [ ] **Step 4: Update `list_tasks` to read the real body and pass it in**

In `src/outlook/client.rs`, inside `fn list_tasks` (lines 1623-1650), replace:

```rust
            let count = variant_to_i32(&get_property(&items, "Count")?).unwrap_or(0);
            let mut results = Vec::new();
            for i in 1..=count {
                let item = to_disp(call_method(&items, "Item", &mut [variant_from_i32(i)])?)?;
                let summary = task_summary(&item)?;
                if task_matches(&summary, &q) {
                    results.push(summary);
                }
            }
            Ok(results)
```

with:

```rust
            let count = variant_to_i32(&get_property(&items, "Count")?).unwrap_or(0);
            let mut results = Vec::new();
            for i in 1..=count {
                let item = to_disp(call_method(&items, "Item", &mut [variant_from_i32(i)])?)?;
                let summary = task_summary(&item)?;
                // Read the real body directly for query matching — `task_summary`
                // doesn't expose the body at all, so this is a second, deliberate
                // property read (same pattern `list_notes` already uses).
                let body = variant_to_string(&get_property(&item, "Body").unwrap_or_default());
                if task_matches(&body, &summary, &q) {
                    results.push(summary);
                }
            }
            Ok(results)
```

- [ ] **Step 5: Run the unit tests to verify they pass**

Run: `cargo test task_filter_tests`

Expected: `test result: ok. 8 passed`

- [ ] **Step 6: Update the `TaskQuery`/`NoteQuery` doc comments**

In `src/outlook/mod.rs`, replace (lines 99-121):

```rust
/// All filters for `list_tasks`. Every field is optional except
/// `include_completed`; supplying several ANDs them. `include_completed`
/// drives a server-side `Restrict`; the rest filter the streamed tasks
/// client-side (there's no established DASL text-search path for the Tasks
/// folder in this codebase, unlike email's `@SQL` queries — same approach
/// `EventQuery`'s `query`/`category` already use).
#[derive(Debug, Clone, Default)]
pub struct TaskQuery {
    pub include_completed: bool,
    pub category: Option<String>,
    pub importance: Option<String>,
    pub query: Option<String>, // text match on subject (TaskSummary has no body field to match)
}

/// All filters for `list_notes`. Both fields optional; supplying both ANDs
/// them. Unlike `TaskQuery`'s `query` (subject-only, since tasks have a
/// separate subject), a note's *only* content is its body — `note_matches`
/// reads the real body text to match `query`, not just the derived subject.
#[derive(Debug, Clone, Default)]
pub struct NoteQuery {
    pub category: Option<String>,
    pub query: Option<String>,
}
```

with:

```rust
/// All filters for `list_tasks`. Every field is optional except
/// `include_completed`; supplying several ANDs them. `include_completed`
/// drives a server-side `Restrict`; the rest filter the streamed tasks
/// client-side (there's no established DASL text-search path for the Tasks
/// folder in this codebase, unlike email's `@SQL` queries — same approach
/// `EventQuery`'s `query`/`category` already use). `query` matches either
/// the subject or the real task body, read per-item — same as `NoteQuery`'s
/// `query` below.
#[derive(Debug, Clone, Default)]
pub struct TaskQuery {
    pub include_completed: bool,
    pub category: Option<String>,
    pub importance: Option<String>,
    pub query: Option<String>, // text match on subject OR body
}

/// All filters for `list_notes`. Both fields optional; supplying both ANDs
/// them. A note's *only* content is its body (it has no separate subject),
/// so `note_matches` reads the real body text to match `query` — the same
/// approach `TaskQuery`'s `query` above now uses alongside its subject.
#[derive(Debug, Clone, Default)]
pub struct NoteQuery {
    pub category: Option<String>,
    pub query: Option<String>,
}
```

- [ ] **Step 7: Update README**

In `README.md`, change line 45 from:

```markdown
- `list_tasks` — list Outlook tasks
```

to:

```markdown
- `list_tasks` — list Outlook tasks (filter by category, importance, or a text query matching subject or body)
```

- [ ] **Step 8: Run the full unit suite**

Run: `cargo test`

Expected: all tests pass, including the new `task_filter_tests` module and the untouched `event_filter_tests`/other modules.

- [ ] **Step 9: Commit**

```bash
git add src/outlook/client.rs src/outlook/mod.rs README.md
git commit -m "Make list_tasks query match task body, not just subject"
```

---

### Task 3: Add a live regression test for task body search, then final verification

Mirrors Task 1's email test, but for tasks — proves the Task 2 change works against the real, running Outlook, not just the fake/unit layer.

**Files:**
- Modify: `tests/live_outlook.rs` (add a new test after `list_tasks_filters_and_create_task_additions_round_trip`, currently ending at line 661)
- Modify: `TESTING.md:78-83` (extend the existing `list_tasks` paragraph)

**Interfaces:**
- Consumes: `WindowsOutlookClient::create_task`, `::list_tasks(TaskQuery)`, `::delete_task` — all existing, unchanged signatures. `TaskQuery` gained no new fields in Task 2 (only its `query` field's *meaning* changed), so this test uses the same `TaskQuery` shape as the existing live tests.
- Produces: nothing — this is the last task in the plan.

- [ ] **Step 1: Write the live test**

Add to `tests/live_outlook.rs`, directly after `list_tasks_filters_and_create_task_additions_round_trip` (after line 661):

```rust
#[test]
#[ignore]
fn list_tasks_query_matches_real_body_text() {
    let c = client();
    let token = "zztaskbodytoken8842";
    let created = c.create_task(
        "[outlook-mcp-rs body-search live] task probe".to_string(),
        Some(format!("this task's body contains {token} and the subject does not")),
        None, "normal".to_string(), None, None, None,
    ).expect("create_task should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let found = c.list_tasks(TaskQuery {
        include_completed: false,
        category: None,
        importance: None,
        query: Some(token.to_string()),
    }).expect("list_tasks query should succeed");

    c.delete_task(id.clone()).expect("cleanup: delete the task");

    assert!(
        found.iter().any(|t| t.id == id),
        "list_tasks query {token:?} should find a task whose ONLY occurrence \
         of that token is in the body"
    );
}
```

- [ ] **Step 2: Run it against the real, running Outlook**

Run: `cargo test --test live_outlook -- --ignored list_tasks_query_matches_real_body_text --nocapture`

Expected: `test result: ok. 1 passed`

- [ ] **Step 3: Re-run both new live tests together, to confirm no interference**

Run: `cargo test --test live_outlook -- --ignored list_emails_query_matches_real_body_text list_tasks_query_matches_real_body_text --nocapture`

Expected: `test result: ok. 2 passed`

- [ ] **Step 4: Document the new live test in TESTING.md**

In `TESTING.md`, change the `list_tasks` paragraph (lines 78-83) from:

```markdown
`list_tasks` filters (`category`, `importance`, `query`, `include_completed`),
`create_task`'s additions (`categories`, `start_date`, `reminder_time`), and
`update_task`/`delete_task` (which retired the standalone `complete_task`
tool — `mark_complete: true`/`false` on `update_task` now covers completing
*and* reopening a task) are covered by the live suite:
`cargo test --test live_outlook -- --ignored list_tasks_filters_and_create_task_additions_round_trip update_task_marks_complete_then_reopens delete_task_removes_it`.
```

to:

```markdown
`list_tasks` filters (`category`, `importance`, `query`, `include_completed`),
`create_task`'s additions (`categories`, `start_date`, `reminder_time`), and
`update_task`/`delete_task` (which retired the standalone `complete_task`
tool — `mark_complete: true`/`false` on `update_task` now covers completing
*and* reopening a task) are covered by the live suite:
`cargo test --test live_outlook -- --ignored list_tasks_filters_and_create_task_additions_round_trip update_task_marks_complete_then_reopens delete_task_removes_it`.
`list_tasks`'s `query` filter matching real task body text (not just
subject) is covered separately by:
`cargo test --test live_outlook -- --ignored list_tasks_query_matches_real_body_text`.
```

- [ ] **Step 5: Run the full unit suite one final time**

Run: `cargo test`

Expected: all tests pass (no regressions from the doc-only changes in this task).

- [ ] **Step 6: Commit**

```bash
git add tests/live_outlook.rs TESTING.md
git commit -m "Add live proof that list_tasks query matches real task body text"
```

---

## Self-Review Notes

- **Spec coverage:** "tasks should search body" → Task 2 (unit-level) + Task 3 (live proof). "mail should also do it" → mail already implements this in code; Task 1 supplies the missing proof and documents it, rather than re-implementing something that already exists.
- **Type consistency:** `task_matches`'s new signature (`body: &str, summary: &TaskSummary, q: &TaskQuery`) is used identically in its Step 1 tests (Task 2) and its Step 4 call site update (Task 2) — parameter order matches `note_matches(body: &str, summary: &NoteSummary, q: &NoteQuery)` exactly, so the pattern stays consistent across both matchers.
- **No placeholders:** every step above has the literal code/diff/command to run; no task defers "add tests" or "handle errors" to later.
