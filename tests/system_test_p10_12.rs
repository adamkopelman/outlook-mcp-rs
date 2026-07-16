//! System test for Plans 10-12 (check_availability, Tasks, Notes), executed
//! against the real, running Outlook mailbox per
//! `SYSTEM_TEST_PLAN_2026-07-16-P10-12.md`. NOT run by plain `cargo test` -
//! `#[ignore]`d. Run explicitly:
//!   cargo test --test system_test_p10_12 -- --ignored --nocapture
//!
//! Every step records PASS/FAIL into a running log instead of panicking, so
//! a single failure never skips cleanup or later steps. Cleanup runs
//! unconditionally at the end. The final line is a real `assert!` over the
//! collected results, so the process exit code reflects overall success -
//! but only after cleanup has already happened. Mirrors the house style of
//! `tests/system_test.rs` (Plans 1-9).

use outlook_mcp_rs::outlook::client::WindowsOutlookClient;
use outlook_mcp_rs::outlook::{
    CheckAvailabilityInput, NoteQuery, NoteUpdate, OutlookClient, TaskQuery, TaskUpdate,
};
use std::collections::HashSet;

const TAG: &str = "[outlook-mcp-rs systest P10-12]";
const SELF_ADDR: &str = "adamkopelman@outlook.com";

fn client() -> WindowsOutlookClient {
    WindowsOutlookClient::new()
}

struct Results {
    entries: Vec<(String, bool, String)>,
}
impl Results {
    fn new() -> Self {
        Self { entries: Vec::new() }
    }
    fn record(&mut self, id: &str, pass: bool, note: impl Into<String>) {
        let note = note.into();
        println!("[{id}] {} - {note}", if pass { "PASS" } else { "FAIL" });
        self.entries.push((id.to_string(), pass, note));
    }
    fn print_summary(&self) {
        println!("\n=== SUMMARY ===");
        for (id, pass, note) in &self.entries {
            println!("{:<10} {:<4} {}", id, if *pass { "PASS" } else { "FAIL" }, note);
        }
        let pass = self.entries.iter().filter(|e| e.1).count();
        println!("\n{pass}/{} passed", self.entries.len());
    }
    fn all_passed(&self) -> bool {
        self.entries.iter().all(|e| e.1)
    }
}

/// Tagged subject suffixes among `list_tasks` results - used the same way
/// `system_test.rs`'s `tagged_suffixes` narrows a filtered result set down
/// to just our own seed data, ignoring any pre-existing real tasks that
/// happen to share a category/importance value.
fn tagged_task_suffixes<'a>(items: &'a [outlook_mcp_rs::outlook::types::TaskSummary]) -> HashSet<&'a str> {
    items.iter().filter_map(|t| t.subject.strip_prefix(TAG).map(|s| s.trim())).collect()
}

/// Same idea for `list_notes`: a note's `subject` is derived from its body's
/// first line, so a tagged body always yields a tagged subject too.
fn tagged_note_suffixes<'a>(items: &'a [outlook_mcp_rs::outlook::types::NoteSummary]) -> HashSet<&'a str> {
    items.iter().filter_map(|n| n.subject.strip_prefix(TAG).map(|s| s.trim())).collect()
}

fn check_task_set(r: &mut Results, id: &str, actual: HashSet<&str>, expected: &[&str], note: &str) {
    let expected_set: HashSet<&str> = expected.iter().copied().collect();
    let pass = actual == expected_set;
    r.record(id, pass, format!("{note}: expected {expected:?}, got {actual:?}"));
}

#[test]
#[ignore]
fn system_test_plans_10_to_12() {
    let c = client();
    let mut r = Results::new();
    // Tracks anything created that still needs cleanup at the very end,
    // regardless of what passed/failed above - populated on create, removed
    // as soon as each test's own delete_task/delete_note succeeds.
    let mut cleanup_tasks: Vec<(String, String)> = Vec::new();
    let mut cleanup_notes: Vec<(String, String)> = Vec::new();

    // ================= C1: check_availability (own mailbox) =================
    println!("\n--- C1: check_availability ---");
    match c.check_availability(CheckAvailabilityInput {
        people: vec![SELF_ADDR.to_string()],
        start: "2099-08-01T09:00".to_string(),
        end: "2099-08-01T11:00".to_string(),
        interval_minutes: 30,
        treat_as_free: vec!["free".to_string()],
    }) {
        Ok(result) => {
            let one_person = result.people.len() == 1;
            let person = result.people.first();
            let resolved = person.map(|p| p.resolved).unwrap_or(false);
            let four_slots = person.map(|p| p.slots.len() == 4).unwrap_or(false);
            let valid_words = person
                .map(|p| p.slots.iter().all(|s| {
                    ["free", "tentative", "busy", "out_of_office", "working_elsewhere"]
                        .contains(&s.status.as_str())
                }))
                .unwrap_or(false);
            let common_free_nonempty = !result.common_free.is_empty();
            let ok = one_person && resolved && four_slots && valid_words && common_free_nonempty;
            r.record("C1", ok, format!(
                "one_person={one_person} resolved={resolved} four_slots={four_slots} \
                 valid_words={valid_words} common_free_nonempty={common_free_nonempty} \
                 (slots={:?}, common_free={:?})",
                person.map(|p| &p.slots), result.common_free
            ));
        }
        Err(e) => r.record("C1", false, format!("check_availability failed: {e}")),
    }

    // ================= T1: list_tasks filters =================
    println!("\n--- T1: list_tasks filters ---");
    {
        let t1a_subject = format!("{TAG} T1 quokkaTask");
        let t1b_subject = format!("{TAG} T1 plain");
        let mut t1a_id: Option<String> = None;
        let mut t1b_id: Option<String> = None;

        match c.create_task(t1a_subject.clone(), None, None, "high".to_string(),
            Some(vec!["Red Category".to_string()]), None, None) {
            Ok(v) => {
                if let Some(id) = v["id"].as_str() {
                    t1a_id = Some(id.to_string());
                    cleanup_tasks.push(("T1a".to_string(), id.to_string()));
                }
            }
            Err(e) => r.record("T1-setup-a", false, format!("create_task failed: {e}")),
        }
        match c.create_task(t1b_subject.clone(), None, None, "normal".to_string(), None, None, None) {
            Ok(v) => {
                if let Some(id) = v["id"].as_str() {
                    t1b_id = Some(id.to_string());
                    cleanup_tasks.push(("T1b".to_string(), id.to_string()));
                }
            }
            Err(e) => r.record("T1-setup-b", false, format!("create_task failed: {e}")),
        }

        if t1a_id.is_some() && t1b_id.is_some() {
            match c.list_tasks(TaskQuery { category: Some("Red Category".to_string()), ..Default::default() }) {
                Ok(list) => check_task_set(&mut r, "T1-category", tagged_task_suffixes(&list),
                    &["T1 quokkaTask"], "category:Red Category"),
                Err(e) => r.record("T1-category", false, format!("failed: {e}")),
            }
            match c.list_tasks(TaskQuery { importance: Some("high".to_string()), ..Default::default() }) {
                Ok(list) => check_task_set(&mut r, "T1-importance", tagged_task_suffixes(&list),
                    &["T1 quokkaTask"], "importance:high"),
                Err(e) => r.record("T1-importance", false, format!("failed: {e}")),
            }
            match c.list_tasks(TaskQuery { query: Some("quokkaTask".to_string()), ..Default::default() }) {
                Ok(list) => check_task_set(&mut r, "T1-query", tagged_task_suffixes(&list),
                    &["T1 quokkaTask"], "query:quokkaTask"),
                Err(e) => r.record("T1-query", false, format!("failed: {e}")),
            }
        } else {
            r.record("T1-category", false, "skipped: setup failed");
            r.record("T1-importance", false, "skipped: setup failed");
            r.record("T1-query", false, "skipped: setup failed");
        }

        for (label, id) in [("T1a", t1a_id), ("T1b", t1b_id)] {
            if let Some(id) = id {
                match c.delete_task(id) {
                    Ok(_) => { cleanup_tasks.retain(|(l, _)| l != label); }
                    Err(e) => println!("T1 cleanup FAILED for {label}: {e}"),
                }
            }
        }
    }

    // ================= T2: create_task additions =================
    println!("\n--- T2: create_task additions ---");
    {
        let subject = format!("{TAG} T2 additions");
        match c.create_task(subject.clone(), None, None, "normal".to_string(),
            Some(vec!["Blue Category".to_string()]), Some("2099-01-01".to_string()),
            Some("2099-01-01T09:00".to_string())) {
            Ok(v) => {
                if let Some(id) = v["id"].as_str() {
                    let id = id.to_string();
                    cleanup_tasks.push(("T2".to_string(), id.clone()));
                    let found = c.list_tasks(TaskQuery { query: Some("T2 additions".to_string()), ..Default::default() })
                        .map(|l| l.iter().find(|t| t.id == id).map(|t| t.categories.clone()))
                        .unwrap_or(None);
                    let has_blue = found.as_ref().map(|cats| cats.iter().any(|c| c == "Blue Category")).unwrap_or(false);
                    r.record("T2", has_blue, format!("categories after create: {found:?}"));
                    match c.delete_task(id) {
                        Ok(_) => { cleanup_tasks.retain(|(l, _)| l != "T2"); }
                        Err(e) => println!("T2 cleanup FAILED: {e}"),
                    }
                } else {
                    r.record("T2", false, "create_task succeeded but no id returned");
                }
            }
            Err(e) => r.record("T2", false, format!("create_task failed: {e}")),
        }
    }

    // ================= T3: update_task - complete/reopen/field edits =================
    println!("\n--- T3: update_task ---");
    {
        let subject = format!("{TAG} T3 base");
        match c.create_task(subject.clone(), None, None, "normal".to_string(), None, None, None) {
            Ok(v) => {
                if let Some(id) = v["id"].as_str() {
                    let id = id.to_string();
                    cleanup_tasks.push(("T3".to_string(), id.clone()));

                    let mark_ok = match c.update_task(TaskUpdate {
                        task_id: id.clone(), mark_complete: Some(true), ..Default::default()
                    }) {
                        Ok(v) => v["changed"].as_array().map(|a| a.iter().any(|x| x == "mark_complete")).unwrap_or(false),
                        Err(e) => { println!("T3 mark_complete FAILED: {e}"); false }
                    };
                    let is_complete = c.list_tasks(TaskQuery { include_completed: true, ..Default::default() })
                        .map(|l| l.iter().find(|t| t.id == id).map(|t| t.complete).unwrap_or(false))
                        .unwrap_or(false);
                    let reopen_ok = match c.update_task(TaskUpdate {
                        task_id: id.clone(), mark_complete: Some(false), ..Default::default()
                    }) {
                        Ok(v) => v["changed"].as_array().map(|a| a.iter().any(|x| x == "mark_complete")).unwrap_or(false),
                        Err(e) => { println!("T3 reopen FAILED: {e}"); false }
                    };
                    let is_reopened = c.list_tasks(TaskQuery { include_completed: true, ..Default::default() })
                        .map(|l| l.iter().find(|t| t.id == id).map(|t| !t.complete).unwrap_or(false))
                        .unwrap_or(false);

                    let renamed = format!("{TAG} T3 base (renamed)");
                    let edit_ok = match c.update_task(TaskUpdate {
                        task_id: id.clone(), subject: Some(renamed.clone()), importance: Some("high".to_string()),
                        add_categories: Some(vec!["Orange Category".to_string()]), ..Default::default()
                    }) {
                        Ok(v) => {
                            let changed: HashSet<String> = v["changed"].as_array()
                                .map(|a| a.iter().filter_map(|x| x.as_str().map(|s| s.to_string())).collect())
                                .unwrap_or_default();
                            changed.contains("subject") && changed.contains("importance") && changed.contains("add_categories")
                        }
                        Err(e) => { println!("T3 field edit FAILED: {e}"); false }
                    };

                    let ok = mark_ok && is_complete && reopen_ok && is_reopened && edit_ok;
                    r.record("T3", ok, format!(
                        "mark_ok={mark_ok} is_complete={is_complete} reopen_ok={reopen_ok} \
                         is_reopened={is_reopened} edit_ok={edit_ok}"
                    ));

                    match c.delete_task(id) {
                        Ok(_) => { cleanup_tasks.retain(|(l, _)| l != "T3"); }
                        Err(e) => println!("T3 cleanup FAILED: {e}"),
                    }
                } else {
                    r.record("T3", false, "create_task succeeded but no id returned");
                }
            }
            Err(e) => r.record("T3", false, format!("create_task failed: {e}")),
        }
    }

    // ================= T4: delete_task =================
    println!("\n--- T4: delete_task ---");
    {
        let subject = format!("{TAG} T4 delete probe");
        match c.create_task(subject.clone(), None, None, "normal".to_string(), None, None, None) {
            Ok(v) => {
                if let Some(id) = v["id"].as_str() {
                    let id = id.to_string();
                    cleanup_tasks.push(("T4".to_string(), id.clone()));
                    match c.delete_task(id.clone()) {
                        Ok(v) => {
                            r.record("T4", v["status"] == "deleted", format!("{v}"));
                            cleanup_tasks.retain(|(l, _)| l != "T4");
                        }
                        Err(e) => r.record("T4", false, format!("delete_task failed: {e}")),
                    }
                } else {
                    r.record("T4", false, "create_task succeeded but no id returned");
                }
            }
            Err(e) => r.record("T4", false, format!("create_task failed: {e}")),
        }
    }

    // ================= N1: list_notes filters =================
    println!("\n--- N1: list_notes filters ---");
    {
        let n1a_body = format!("{TAG} N1 category note - remember zephyrling");
        let n1b_body = format!("{TAG} N1 plain note");
        let mut n1a_id: Option<String> = None;
        let mut n1b_id: Option<String> = None;

        match c.create_note(n1a_body.clone(), Some(vec!["Green Category".to_string()]), None) {
            Ok(v) => {
                if let Some(id) = v["id"].as_str() {
                    n1a_id = Some(id.to_string());
                    cleanup_notes.push(("N1a".to_string(), id.to_string()));
                }
            }
            Err(e) => r.record("N1-setup-a", false, format!("create_note failed: {e}")),
        }
        match c.create_note(n1b_body.clone(), None, None) {
            Ok(v) => {
                if let Some(id) = v["id"].as_str() {
                    n1b_id = Some(id.to_string());
                    cleanup_notes.push(("N1b".to_string(), id.to_string()));
                }
            }
            Err(e) => r.record("N1-setup-b", false, format!("create_note failed: {e}")),
        }

        if n1a_id.is_some() && n1b_id.is_some() {
            match c.list_notes(NoteQuery { category: Some("Green Category".to_string()), ..Default::default() }) {
                Ok(list) => check_task_set(&mut r, "N1-category", tagged_note_suffixes(&list),
                    &["N1 category note - remember zephyrling"], "category:Green Category"),
                Err(e) => r.record("N1-category", false, format!("failed: {e}")),
            }
            match c.list_notes(NoteQuery { query: Some("zephyrling".to_string()), ..Default::default() }) {
                Ok(list) => check_task_set(&mut r, "N1-query", tagged_note_suffixes(&list),
                    &["N1 category note - remember zephyrling"], "query:zephyrling"),
                Err(e) => r.record("N1-query", false, format!("failed: {e}")),
            }
        } else {
            r.record("N1-category", false, "skipped: setup failed");
            r.record("N1-query", false, "skipped: setup failed");
        }

        for (label, id) in [("N1a", n1a_id), ("N1b", n1b_id)] {
            if let Some(id) = id {
                match c.delete_note(id) {
                    Ok(_) => { cleanup_notes.retain(|(l, _)| l != label); }
                    Err(e) => println!("N1 cleanup FAILED for {label}: {e}"),
                }
            }
        }
    }

    // ================= N2: create_note additions + get_note modified =================
    println!("\n--- N2: create_note additions + get_note modified ---");
    {
        let body = format!("{TAG} N2 additions probe");
        match c.create_note(body.clone(), Some(vec!["Yellow Category".to_string()]), Some("yellow".to_string())) {
            Ok(v) => {
                if let Some(id) = v["id"].as_str() {
                    let id = id.to_string();
                    cleanup_notes.push(("N2".to_string(), id.clone()));

                    // Baseline: capture categories + modified right after
                    // creation, BEFORE update_note. create_note's own Save()
                    // already populates LastModificationTime, so asserting
                    // modified.is_some() only after the update would prove
                    // nothing (see skill doc / live_outlook.rs precedent) -
                    // capture the baseline and require non-decreasing after,
                    // paired with an assertion on the actually-changed body.
                    let before = c.get_note(id.clone()).ok();
                    let has_yellow = before.as_ref()
                        .map(|d| d.summary.categories.iter().any(|cat| cat == "Yellow Category"))
                        .unwrap_or(false);
                    let before_modified = before.as_ref().and_then(|d| d.modified.clone());

                    let edited_body = format!("{body} (edited)");
                    let update_ok = c.update_note(NoteUpdate {
                        note_id: id.clone(), body: Some(edited_body.clone()), ..Default::default()
                    }).is_ok();

                    let after = c.get_note(id.clone()).ok();
                    let after_modified = after.as_ref().and_then(|d| d.modified.clone());
                    let body_changed = after.as_ref().map(|d| d.body.starts_with(&edited_body)).unwrap_or(false);
                    let modified_nondecreasing = match (&before_modified, &after_modified) {
                        (Some(b), Some(a)) => a >= b,
                        _ => false,
                    };

                    let ok = has_yellow && update_ok && body_changed && modified_nondecreasing
                        && before_modified.is_some() && after_modified.is_some();
                    r.record("N2", ok, format!(
                        "has_yellow={has_yellow} update_ok={update_ok} body_changed={body_changed} \
                         before_modified={before_modified:?} after_modified={after_modified:?} \
                         nondecreasing={modified_nondecreasing}"
                    ));

                    match c.delete_note(id) {
                        Ok(_) => { cleanup_notes.retain(|(l, _)| l != "N2"); }
                        Err(e) => println!("N2 cleanup FAILED: {e}"),
                    }
                } else {
                    r.record("N2", false, "create_note succeeded but no id returned");
                }
            }
            Err(e) => r.record("N2", false, format!("create_note failed: {e}")),
        }
    }

    // ================= N3: update_note - categories and color =================
    println!("\n--- N3: update_note ---");
    {
        let body = format!("{TAG} N3 base");
        match c.create_note(body.clone(), None, None) {
            Ok(v) => {
                if let Some(id) = v["id"].as_str() {
                    let id = id.to_string();
                    cleanup_notes.push(("N3".to_string(), id.clone()));

                    let add_ok = match c.update_note(NoteUpdate {
                        note_id: id.clone(), add_categories: Some(vec!["Pink Category".to_string()]),
                        color: Some("pink".to_string()), ..Default::default()
                    }) {
                        Ok(v) => {
                            let changed: HashSet<String> = v["changed"].as_array()
                                .map(|a| a.iter().filter_map(|x| x.as_str().map(|s| s.to_string())).collect())
                                .unwrap_or_default();
                            changed.contains("add_categories") && changed.contains("color")
                        }
                        Err(e) => { println!("N3 add FAILED: {e}"); false }
                    };
                    let has_pink = c.get_note(id.clone()).ok()
                        .map(|d| d.summary.categories.iter().any(|cat| cat == "Pink Category"))
                        .unwrap_or(false);

                    let remove_ok = match c.update_note(NoteUpdate {
                        note_id: id.clone(), remove_categories: Some(vec!["Pink Category".to_string()]),
                        ..Default::default()
                    }) {
                        Ok(v) => v["changed"].as_array().map(|a| a.iter().any(|x| x == "remove_categories")).unwrap_or(false),
                        Err(e) => { println!("N3 remove FAILED: {e}"); false }
                    };
                    let pink_gone = c.get_note(id.clone()).ok()
                        .map(|d| !d.summary.categories.iter().any(|cat| cat == "Pink Category"))
                        .unwrap_or(false);

                    let ok = add_ok && has_pink && remove_ok && pink_gone;
                    r.record("N3", ok, format!(
                        "add_ok={add_ok} has_pink={has_pink} remove_ok={remove_ok} pink_gone={pink_gone}"
                    ));

                    match c.delete_note(id) {
                        Ok(_) => { cleanup_notes.retain(|(l, _)| l != "N3"); }
                        Err(e) => println!("N3 cleanup FAILED: {e}"),
                    }
                } else {
                    r.record("N3", false, "create_note succeeded but no id returned");
                }
            }
            Err(e) => r.record("N3", false, format!("create_note failed: {e}")),
        }
    }

    // ================= N4: delete_note =================
    println!("\n--- N4: delete_note ---");
    {
        let body = format!("{TAG} N4 delete probe");
        match c.create_note(body.clone(), None, None) {
            Ok(v) => {
                if let Some(id) = v["id"].as_str() {
                    let id = id.to_string();
                    cleanup_notes.push(("N4".to_string(), id.clone()));
                    match c.delete_note(id.clone()) {
                        Ok(v) => {
                            r.record("N4", v["status"] == "deleted", format!("{v}"));
                            cleanup_notes.retain(|(l, _)| l != "N4");
                        }
                        Err(e) => r.record("N4", false, format!("delete_note failed: {e}")),
                    }
                } else {
                    r.record("N4", false, "create_note succeeded but no id returned");
                }
            }
            Err(e) => r.record("N4", false, format!("create_note failed: {e}")),
        }
    }

    // ================= Cleanup checklist =================
    println!("\n--- Cleanup ---");
    let mut leftovers: Vec<String> = Vec::new();
    for (label, id) in &cleanup_tasks {
        match c.delete_task(id.clone()) {
            Ok(_) => println!("cleaned up leftover task {label} ({id})"),
            Err(e) => {
                println!("FAILED to clean up task {label} ({id}): {e}");
                leftovers.push(format!("task {label} id={id}: {e}"));
            }
        }
    }
    for (label, id) in &cleanup_notes {
        match c.delete_note(id.clone()) {
            Ok(_) => println!("cleaned up leftover note {label} ({id})"),
            Err(e) => {
                println!("FAILED to clean up note {label} ({id}): {e}");
                leftovers.push(format!("note {label} id={id}: {e}"));
            }
        }
    }
    let task_sweep_clean = c.list_tasks(TaskQuery { include_completed: true, ..Default::default() })
        .map(|l| !l.iter().any(|t| t.subject.starts_with(TAG)))
        .unwrap_or(false);
    let note_sweep_clean = c.list_notes(NoteQuery::default())
        .map(|l| !l.iter().any(|n| n.subject.starts_with(TAG)))
        .unwrap_or(false);
    r.record("cleanup", leftovers.is_empty() && task_sweep_clean && note_sweep_clean,
        format!("leftovers={leftovers:?} task_sweep_clean={task_sweep_clean} note_sweep_clean={note_sweep_clean}"));

    r.print_summary();
    assert!(r.all_passed(), "one or more system test steps failed - see summary above");
}
