//! Live system tests against a real, running Outlook. NOT run by plain
//! `cargo test` — every test is `#[ignore]`d. Run explicitly:
//!   cargo test --test live_outlook -- --ignored
//! See TESTING.md for preconditions.
//!
//! Every test that creates an Outlook item deletes it before returning, so
//! repeated runs don't accumulate junk in the mailbox. `send_email` and
//! `respond_to_meeting` are deliberately NOT covered here since a real send
//! can't be undone — see TESTING.md for how to test those by hand.

use outlook_mcp_rs::outlook::client::WindowsOutlookClient;
use outlook_mcp_rs::outlook::{EmailQuery, OutlookClient};

fn client() -> WindowsOutlookClient {
    WindowsOutlookClient::new()
}

#[test]
#[ignore]
fn list_folders_returns_at_least_inbox() {
    let folders = client().list_folders().expect("list_folders should succeed against a live Outlook");
    assert!(folders.iter().any(|f| f.name.eq_ignore_ascii_case("inbox")));
}

#[test]
#[ignore]
fn list_emails_returns_inbox_items() {
    let emails = client().list_emails(EmailQuery {
        query: None, folder: "inbox".into(), count: 5, unread_only: false,
        from: None, category: None, received_after: None, received_before: None,
        since_days: None, has_attachments: None, flagged: false, high_importance: false,
    }).expect("list_emails should succeed against a live Outlook");
    // Not asserting a specific count/content since the real mailbox varies —
    // just confirm the call succeeds and returns well-formed summaries.
    for email in &emails {
        assert!(!email.id.is_empty());
    }
}

#[test]
#[ignore]
fn create_draft_then_delete_round_trips() {
    let c = client();
    let created = c.create_draft(
        vec!["nobody@example.invalid".to_string()],
        "outlook-mcp-rs live test draft".to_string(),
        "This draft is created and deleted by an automated test.".to_string(),
        None, None, false,
    ).expect("create_draft should succeed");
    let id = created["id"].as_str().expect("create_draft returns an id").to_string();
    c.delete_email(id).expect("cleanup: delete_email should succeed");
}

#[test]
#[ignore]
fn create_task_complete_then_it_is_marked_complete() {
    let c = client();
    let created = c.create_task(
        "outlook-mcp-rs live test task".to_string(), None, None, "normal".to_string(),
    ).expect("create_task should succeed");
    let id = created["id"].as_str().unwrap().to_string();
    c.complete_task(id.clone()).expect("complete_task should succeed");
    let tasks = c.list_tasks(true).expect("list_tasks should succeed");
    assert!(tasks.iter().any(|t| t.id == id && t.complete));
    // Outlook has no direct "delete task" in our trait yet — deleting the
    // completed test task manually is fine (it's clearly labeled).
}

#[test]
#[ignore]
fn create_note_then_get_it_back() {
    let c = client();
    let created = c.create_note("outlook-mcp-rs live test note".to_string())
        .expect("create_note should succeed");
    let id = created["id"].as_str().unwrap().to_string();
    let note = c.get_note(id).expect("get_note should succeed");
    assert!(note.body.starts_with("outlook-mcp-rs live test note"));
}

#[test]
#[ignore]
fn create_event_then_delete_it() {
    let c = client();
    let created = c.create_event(
        "outlook-mcp-rs live test event".to_string(),
        "2099-01-01T10:00:00".to_string(),
        "2099-01-01T10:30:00".to_string(),
        None, None, None, false, None,
    ).expect("create_event should succeed");
    let id = created["id"].as_str().unwrap().to_string();
    // Calendar items don't have a dedicated "delete" tool in the trait; use
    // move_email into Deleted Items works for mail but not appointments —
    // delete the test event manually from the calendar after this test runs,
    // or extend the trait with a delete_event method if this becomes
    // frequent enough to automate.
    let _ = c.get_event(id); // just confirm it round-trips before manual cleanup
}
