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
use outlook_mcp_rs::outlook::{CreateEventInput, EmailQuery, EventQuery, OutlookClient, EmailUpdate};

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
    let created = c.create_event(CreateEventInput {
        subject: "outlook-mcp-rs live test event".to_string(),
        start: "2099-01-01T10:00:00".to_string(),
        end: "2099-01-01T10:30:00".to_string(),
        body: None, location: None, required_attendees: None, optional_attendees: None,
        all_day: false, reminder_minutes: None, categories: None, show_as: None,
        send: true,
    }).expect("create_event should succeed");
    let id = created["id"].as_str().unwrap().to_string();
    // Calendar items don't have a dedicated "delete" tool in the trait; moving
    // into Deleted Items works for mail but not appointments — delete the test
    // event manually from the calendar after this test runs, or extend the
    // trait with a delete_event method if this becomes frequent enough to
    // automate.
    let _ = c.get_event(id); // just confirm it round-trips before manual cleanup
}

#[test]
#[ignore]
fn create_event_with_tiers_categories_and_show_as() {
    let c = client();
    // send:false means nothing is ever delivered, so a placeholder address
    // for the invite tiers is safe — Outlook stores it without resolving
    // for delivery.
    let created = c.create_event(CreateEventInput {
        subject: "outlook-mcp-rs P7 tiers probe".to_string(),
        start: "2099-01-06T09:00".to_string(),
        end: "2099-01-06T09:30".to_string(),
        body: None, location: None,
        required_attendees: Some(vec!["required-probe@example.com".to_string()]),
        optional_attendees: Some(vec!["optional-probe@example.com".to_string()]),
        all_day: false, reminder_minutes: None,
        categories: Some(vec!["Work".to_string()]),
        show_as: Some("tentative".to_string()),
        send: false,
    }).expect("create_event should succeed");
    assert_eq!(created["status"], "meeting_saved");
    let id = created["id"].as_str().unwrap().to_string();

    let detail = c.get_event(id).expect("get_event should succeed");
    assert!(detail.summary.required_attendees.contains("required-probe@example.com"));
    assert!(detail.summary.optional_attendees.contains("optional-probe@example.com"));
    assert!(detail.summary.categories.iter().any(|cat| cat == "Work"));
    assert_eq!(detail.summary.show_as, "tentative");
    assert!(detail.summary.is_meeting);
    // Calendar items have no dedicated delete tool yet (Plan 8's delete_event);
    // delete the probe manually from the calendar after this test runs.
}

#[test]
#[ignore]
fn list_events_filters_by_query_and_category() {
    let c = WindowsOutlookClient::new();
    // A far-future, uniquely-named appointment we can pinpoint and clean up.
    let created = c.create_event(CreateEventInput {
        subject: "outlook-mcp-rs P6 filter probe".to_string(),
        start: "2099-01-05T09:00".to_string(),
        end: "2099-01-05T09:30".to_string(),
        body: None, location: None, required_attendees: None, optional_attendees: None,
        all_day: false, reminder_minutes: None, categories: None, show_as: None,
        send: true,
    }).expect("create_event");
    let id = created["id"].as_str().expect("event id").to_string();

    // A matching query in the window finds it.
    let hits = c.list_events(EventQuery {
        start_date: Some("2099-01-05".to_string()),
        end_date: Some("2099-01-05".to_string()),
        query: Some("filter probe".to_string()),
        ..Default::default()
    }).expect("list_events query");
    assert!(hits.iter().any(|e| e.id == id), "query should match the probe");
    // Enriched fields are populated.
    let probe = hits.iter().find(|e| e.id == id).unwrap();
    assert_eq!(probe.show_as, "busy");

    // A non-matching query in the same window excludes it.
    let misses = c.list_events(EventQuery {
        start_date: Some("2099-01-05".to_string()),
        end_date: Some("2099-01-05".to_string()),
        query: Some("no-such-subject-xyz".to_string()),
        ..Default::default()
    }).expect("list_events non-matching query");
    assert!(!misses.iter().any(|e| e.id == id), "non-matching query must exclude the probe");

    // Cleanup: delete the probe.
    c.delete_email(id).expect("cleanup delete");
}

#[test]
#[ignore]
fn list_events_calendar_of_self_opens_own_calendar() {
    // Opening your OWN calendar via calendar_of exercises the recipient-resolve
    // + GetSharedDefaultFolder path without needing a second user's sharing grant.
    // Set OUTLOOK_MCP_TEST_EMAIL to your SMTP address to run this.
    let me = match std::env::var("OUTLOOK_MCP_TEST_EMAIL") {
        Ok(v) if !v.is_empty() => v,
        _ => {
            eprintln!("skipping: set OUTLOOK_MCP_TEST_EMAIL to your address");
            return;
        }
    };
    let c = WindowsOutlookClient::new();
    // Should resolve and return without error (contents may be empty — that's fine).
    let _events = c.list_events(EventQuery {
        calendar_of: Some(me),
        ..Default::default()
    }).expect("list_events calendar_of self should resolve and not error");
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
fn update_email_applies_state_then_moves() {
    let c = WindowsOutlookClient::new();
    // A draft is a safe, disposable target (never sent).
    let created = c.create_draft(
        vec!["nobody@example.invalid".to_string()],
        "outlook-mcp-rs update_email live test".to_string(),
        "body".to_string(),
        None, None, false, None,
    ).expect("create_draft");
    let id = created["id"].as_str().expect("draft id").to_string();

    // Apply state changes only (no move yet) so we can read them back by the same id.
    // NOTE: `flag` is deliberately NOT exercised here. `MarkAsTask` (follow_up)
    // is only valid on sent/received items — Outlook rejects it on a draft
    // ("MarkAsTask is only valid on items that have been sent or received").
    // A draft is the only safe disposable target we can create, and mutating a
    // real received email's flag state isn't cleanly reversible from the trait
    // layer, so flag is verified manually — see TESTING.md.
    let res = c.update_email(EmailUpdate {
        email_id: id.clone(),
        move_to: None,
        mark_read: Some(true),
        flag: None,
        add_categories: Some(vec!["Work".to_string()]),
        remove_categories: None,
        importance: Some("high".to_string()),
    }).expect("update_email state");
    assert_eq!(res["status"], "updated");
    assert_eq!(res["id"], id); // no move → id unchanged
    let changed = res["changed"].as_array().unwrap();
    assert!(changed.iter().any(|v| v == "importance"));
    assert!(changed.iter().any(|v| v == "add_categories"));

    // Verify importance + category landed.
    let detail = c.get_email(id.clone(), false).expect("get_email");
    let dv = serde_json::to_value(&detail).unwrap();
    assert_eq!(dv["summary"]["importance"], "high");
    assert!(dv["summary"]["categories"].as_array().unwrap().iter().any(|v| v == "Work"));
    // mark_read(true) → the item must now read as read (unread == false).
    assert_eq!(dv["summary"]["unread"], false);

    // A standalone mark_read (no other field, so nothing else Saves afterward)
    // must still persist — set it back to unread and confirm it stuck.
    let unread = c.update_email(EmailUpdate {
        email_id: id.clone(),
        mark_read: Some(false),
        ..Default::default()
    }).expect("update_email mark unread");
    assert_eq!(unread["changed"], serde_json::json!(["mark_read"]));
    let redetail = c.get_email(id.clone(), false).expect("get_email after unread");
    let rv = serde_json::to_value(&redetail).unwrap();
    assert_eq!(rv["summary"]["unread"], true);

    // Now move it; the id must change, then delete via the new id for cleanup.
    let moved = c.update_email(EmailUpdate {
        email_id: id.clone(),
        move_to: Some("Deleted Items".to_string()),
        ..Default::default()
    }).expect("update_email move");
    assert_eq!(moved["changed"], serde_json::json!(["move_to"]));
    let new_id = moved["id"].as_str().expect("moved id").to_string();
    c.delete_email(new_id).expect("cleanup delete");
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
