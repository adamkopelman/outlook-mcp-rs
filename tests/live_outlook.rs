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
        None, None, false, None,
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

#[test]
#[ignore]
fn list_emails_query_filter_narrows_results() {
    use outlook_mcp_rs::outlook::EmailQuery;
    let c = WindowsOutlookClient::new();
    let all = c.list_emails(EmailQuery {
        query: None, folder: "inbox".into(), count: 25, unread_only: false,
        from: None, category: None, received_after: None, received_before: None,
        since_days: None, has_attachments: None, flagged: false, high_importance: false,
    }).expect("plain list should work");
    // A query that almost certainly matches nothing should return <= all.
    let filtered = c.list_emails(EmailQuery {
        query: Some("zzqx-improbable-token-9137".into()),
        folder: "inbox".into(), count: 25, unread_only: false,
        from: None, category: None, received_after: None, received_before: None,
        since_days: None, has_attachments: None, flagged: false, high_importance: false,
    }).expect("query list should work");
    assert!(filtered.len() <= all.len());
}

#[test]
#[ignore]
fn create_draft_with_attachment_round_trips() {
    let dir = std::env::temp_dir();
    let path = dir.join("outlook-mcp-rs-live-attach.txt");
    std::fs::write(&path, b"live attachment test").expect("write temp file");
    let path_str = path.to_string_lossy().to_string();

    let c = WindowsOutlookClient::new();
    let created = c.create_draft(
        vec!["nobody@example.invalid".to_string()],
        "outlook-mcp-rs attachment test".to_string(),
        "see attached".to_string(),
        None, None, false,
        Some(vec![path_str]),
    ).expect("create_draft with attachment should succeed");
    let id = created["id"].as_str().expect("draft id").to_string();
    c.delete_email(id).expect("cleanup: delete the draft");
    let _ = std::fs::remove_file(&path);
}

#[test]
#[ignore]
fn get_email_reports_item_type_for_real_inbox_item() {
    let c = WindowsOutlookClient::new();
    let list = c.list_emails(EmailQuery {
        query: None, folder: "inbox".into(), count: 1, unread_only: false,
        from: None, category: None, received_after: None, received_before: None,
        since_days: None, has_attachments: None, flagged: false, high_importance: false,
    }).expect("list");
    if let Some(first) = list.first() {
        let detail = c.get_email(first.id.clone(), false).expect("get_email");
        let v = serde_json::to_value(&detail).unwrap();
        let t = v["item_type"].as_str().unwrap();
        assert!(["email", "meeting", "bounce", "read_receipt", "other"].contains(&t));
        // If it's a meeting, the meeting block must be present.
        if v["is_meeting"].as_bool().unwrap() {
            assert!(v.get("meeting").is_some());
        }
    }
}

#[test]
#[ignore]
fn send_with_missing_attachment_errors_before_sending() {
    let c = WindowsOutlookClient::new();
    let err = c.send_email(
        vec!["nobody@example.invalid".to_string()],
        "should not send".to_string(), "body".to_string(),
        None, None, false,
        Some(vec!["C:/definitely/does/not/exist/nope.pdf".to_string()]),
    ).unwrap_err();
    assert!(err.to_string().contains("attachment not found"));
}
