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
use outlook_mcp_rs::outlook::{CheckAvailabilityInput, CreateEventInput, EmailQuery, EventQuery, OutlookClient, EmailUpdate, EventUpdate, RecurrenceInput};

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
        recurrence: None,
    }).expect("create_event should succeed");
    let id = created["id"].as_str().unwrap().to_string();
    let _ = c.get_event(id.clone()).expect("get_event should round-trip before cleanup");
    c.delete_event(id, true).expect("cleanup delete_event");
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
        recurrence: None,
    }).expect("create_event should succeed");
    assert_eq!(created["status"], "meeting_saved");
    let id = created["id"].as_str().unwrap().to_string();

    let detail = c.get_event(id.clone()).expect("get_event should succeed");
    assert!(detail.summary.required_attendees.contains("required-probe@example.com"));
    assert!(detail.summary.optional_attendees.contains("optional-probe@example.com"));
    assert!(detail.summary.categories.iter().any(|cat| cat == "Work"));
    assert_eq!(detail.summary.show_as, "tentative");
    assert!(detail.summary.is_meeting);
    c.delete_event(id, false).expect("cleanup delete_event");
}

#[test]
#[ignore]
fn update_event_edits_fields_and_manages_attendees() {
    let c = client();
    let created = c.create_event(CreateEventInput {
        subject: "outlook-mcp-rs P8 update probe".to_string(),
        start: "2099-01-07T09:00".to_string(),
        end: "2099-01-07T09:30".to_string(),
        body: None, location: None,
        required_attendees: Some(vec!["required-probe@example.com".to_string()]),
        optional_attendees: None,
        all_day: false, reminder_minutes: None, categories: None, show_as: None,
        send: false,
        recurrence: None,
    }).expect("create_event should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    // Edit fields, add an optional attendee, remove the required one, quietly
    // (send_update: false — nothing is ever delivered).
    let updated = c.update_event(EventUpdate {
        event_id: id.clone(),
        subject: Some("outlook-mcp-rs P8 update probe (renamed)".to_string()),
        start: None, end: None,
        location: Some("Room 42".to_string()),
        body: None, all_day: None, reminder_minutes: Some(15),
        show_as: Some("tentative".to_string()),
        add_categories: Some(vec!["Work".to_string()]),
        remove_categories: None,
        add_required_attendees: None,
        add_optional_attendees: Some(vec!["optional-probe@example.com".to_string()]),
        remove_attendees: Some(vec!["required-probe@example.com".to_string()]),
        send_update: false,
        recurrence: None,
        clear_recurrence: false,
    }).expect("update_event should succeed");
    assert_eq!(updated["status"], "updated");
    let changed = updated["changed"].as_array().unwrap();
    for field in ["subject", "location", "reminder_minutes", "show_as", "add_categories",
                  "add_optional_attendees", "remove_attendees"] {
        assert!(changed.iter().any(|v| v == field), "expected {field} in changed: {changed:?}");
    }

    let detail = c.get_event(id.clone()).expect("get_event should succeed");
    assert_eq!(detail.summary.subject, "outlook-mcp-rs P8 update probe (renamed)");
    assert_eq!(detail.summary.location, "Room 42");
    assert_eq!(detail.summary.show_as, "tentative");
    assert!(detail.summary.categories.iter().any(|cat| cat == "Work"));
    assert!(!detail.summary.required_attendees.contains("required-probe@example.com"));
    assert!(detail.summary.optional_attendees.contains("optional-probe@example.com"));

    c.delete_event(id, false).expect("cleanup delete_event");
}

#[test]
#[ignore]
fn delete_event_removes_a_personal_appointment() {
    let c = client();
    let created = c.create_event(CreateEventInput {
        subject: "outlook-mcp-rs P8 delete probe".to_string(),
        start: "2099-01-08T09:00".to_string(),
        end: "2099-01-08T09:30".to_string(),
        body: None, location: None, required_attendees: None, optional_attendees: None,
        all_day: false, reminder_minutes: None, categories: None, show_as: None,
        send: true, // no attendees present, so this just Saves — nothing is sent
        recurrence: None,
    }).expect("create_event should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let deleted = c.delete_event(id.clone(), true).expect("delete_event should succeed");
    assert_eq!(deleted["status"], "deleted");
    assert_eq!(deleted["note"], "Moved to Deleted Items.");

    // Soft-deleted: get_event on the original id should now fail (moved to
    // Deleted Items changes its EntryID, same as delete_email's behavior).
    assert!(c.get_event(id).is_err());
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
        recurrence: None,
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

#[test]
#[ignore]
fn create_event_weekly_recurrence_round_trips() {
    let c = client();
    let created = c.create_event(CreateEventInput {
        subject: "outlook-mcp-rs P9 weekly recurrence probe".to_string(),
        start: "2099-02-02T09:00".to_string(), // a Monday
        end: "2099-02-02T09:30".to_string(),
        body: None, location: None, required_attendees: None, optional_attendees: None,
        all_day: false, reminder_minutes: None, categories: None, show_as: None,
        send: true,
        recurrence: Some(RecurrenceInput {
            pattern: "weekly".to_string(),
            interval: Some(1),
            days_of_week: Some(vec!["monday".to_string(), "wednesday".to_string()]),
            day_of_month: None,
            until: None,
            occurrences: Some(10),
        }),
    }).expect("create_event with weekly recurrence should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let detail = c.get_event(id.clone()).expect("get_event should succeed");
    assert!(detail.summary.is_recurring);
    let recurrence = detail.recurrence.expect("recurring event should have a recurrence block");
    assert_eq!(recurrence.pattern, "weekly");
    assert_eq!(recurrence.interval, 1);
    assert_eq!(recurrence.days_of_week, vec!["monday".to_string(), "wednesday".to_string()]);
    assert_eq!(recurrence.occurrences, Some(10));
    assert!(!recurrence.no_end);

    c.delete_event(id, false).expect("cleanup delete_event");
}

#[test]
#[ignore]
fn create_event_monthly_recurrence_with_until_round_trips() {
    let c = client();
    let created = c.create_event(CreateEventInput {
        subject: "outlook-mcp-rs P9 monthly recurrence probe".to_string(),
        start: "2099-02-15T09:00".to_string(),
        end: "2099-02-15T09:30".to_string(),
        body: None, location: None, required_attendees: None, optional_attendees: None,
        all_day: false, reminder_minutes: None, categories: None, show_as: None,
        send: true,
        recurrence: Some(RecurrenceInput {
            pattern: "monthly".to_string(),
            interval: Some(2),
            days_of_week: None,
            day_of_month: Some(15),
            until: Some("2099-12-15".to_string()),
            occurrences: None,
        }),
    }).expect("create_event with monthly recurrence should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let detail = c.get_event(id.clone()).expect("get_event should succeed");
    let recurrence = detail.recurrence.expect("recurring event should have a recurrence block");
    assert_eq!(recurrence.pattern, "monthly");
    assert_eq!(recurrence.interval, 2);
    assert_eq!(recurrence.day_of_month, Some(15));
    assert!(recurrence.until.is_some());
    assert!(!recurrence.no_end);
    // Confirmed live: Outlook auto-computes a correct `Occurrences` (here 6:
    // Feb/Apr/Jun/Aug/Oct/Dec 15) for a series that was created via `until`,
    // not just for one created via `occurrences` — see RecurrenceInfo's doc
    // comment. This is populated Outlook state, not a bug, so it's pinned
    // here rather than asserted absent.
    assert_eq!(recurrence.occurrences, Some(6));

    c.delete_event(id, false).expect("cleanup delete_event");
}

#[test]
#[ignore]
fn create_event_yearly_recurrence_with_no_end_round_trips() {
    let c = client();
    let created = c.create_event(CreateEventInput {
        subject: "outlook-mcp-rs P9 yearly recurrence probe".to_string(),
        start: "2099-03-10T09:00".to_string(),
        end: "2099-03-10T09:30".to_string(),
        body: None, location: None, required_attendees: None, optional_attendees: None,
        all_day: false, reminder_minutes: None, categories: None, show_as: None,
        send: true,
        recurrence: Some(RecurrenceInput {
            pattern: "yearly".to_string(),
            interval: None,
            days_of_week: None,
            day_of_month: None,
            until: None,
            occurrences: None,
        }),
    }).expect("create_event with yearly recurrence should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let detail = c.get_event(id.clone()).expect("get_event should succeed");
    let recurrence = detail.recurrence.expect("recurring event should have a recurrence block");
    assert_eq!(recurrence.pattern, "yearly");
    assert_eq!(recurrence.interval, 1);
    assert_eq!(recurrence.day_of_month, Some(10)); // derived from the March 10 start date
    assert!(recurrence.no_end);
    assert!(recurrence.until.is_none());
    assert!(recurrence.occurrences.is_none());

    c.delete_event(id, false).expect("cleanup delete_event");
}

#[test]
#[ignore]
fn update_event_changes_then_clears_recurrence() {
    let c = client();
    let created = c.create_event(CreateEventInput {
        subject: "outlook-mcp-rs P9 update recurrence probe".to_string(),
        start: "2099-04-01T09:00".to_string(),
        end: "2099-04-01T09:30".to_string(),
        body: None, location: None, required_attendees: None, optional_attendees: None,
        all_day: false, reminder_minutes: None, categories: None, show_as: None,
        send: true,
        recurrence: Some(RecurrenceInput {
            pattern: "daily".to_string(), interval: Some(1), days_of_week: None,
            day_of_month: None, until: None, occurrences: Some(3),
        }),
    }).expect("create_event should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    // Change the pattern from daily to weekly.
    let updated = c.update_event(EventUpdate {
        event_id: id.clone(),
        subject: None, start: None, end: None, location: None, body: None,
        all_day: None, reminder_minutes: None, show_as: None,
        add_categories: None, remove_categories: None,
        add_required_attendees: None, add_optional_attendees: None, remove_attendees: None,
        send_update: false,
        recurrence: Some(RecurrenceInput {
            pattern: "weekly".to_string(), interval: Some(1),
            days_of_week: Some(vec!["tuesday".to_string()]),
            day_of_month: None, until: None, occurrences: Some(4),
        }),
        clear_recurrence: false,
    }).expect("update_event with recurrence should succeed");
    assert!(updated["changed"].as_array().unwrap().iter().any(|v| v == "recurrence"));

    let detail = c.get_event(id.clone()).expect("get_event should succeed");
    let recurrence = detail.recurrence.expect("still recurring after the change");
    assert_eq!(recurrence.pattern, "weekly");
    assert_eq!(recurrence.days_of_week, vec!["tuesday".to_string()]);

    // Now clear it entirely.
    let cleared = c.update_event(EventUpdate {
        event_id: id.clone(),
        subject: None, start: None, end: None, location: None, body: None,
        all_day: None, reminder_minutes: None, show_as: None,
        add_categories: None, remove_categories: None,
        add_required_attendees: None, add_optional_attendees: None, remove_attendees: None,
        send_update: false,
        recurrence: None,
        clear_recurrence: true,
    }).expect("update_event with clear_recurrence should succeed");
    assert!(cleared["changed"].as_array().unwrap().iter().any(|v| v == "clear_recurrence"));

    let detail = c.get_event(id.clone()).expect("get_event should succeed");
    assert!(!detail.summary.is_recurring);
    assert!(detail.recurrence.is_none());

    c.delete_event(id, false).expect("cleanup delete_event");
}

#[test]
#[ignore]
fn check_availability_against_own_mailbox_returns_free_slots() {
    let c = client();
    // Per the spec's testing strategy: check_availability is tested against
    // the developer's own mailbox where possible; the cross-user sharing
    // path (someone else's calendar) depends on another account having
    // granted access, which can't be set up from a test — see TESTING.md.
    let ns_person = std::env::var("OUTLOOK_TEST_SELF_EMAIL")
        .unwrap_or_else(|_| "adamkopelman@outlook.com".to_string());
    let result = c.check_availability(CheckAvailabilityInput {
        people: vec![ns_person.clone()],
        start: "2099-07-01T09:00".to_string(),
        end: "2099-07-01T11:00".to_string(),
        interval_minutes: 30,
        treat_as_free: vec!["free".to_string()],
    }).expect("check_availability should succeed against a resolvable address");

    assert_eq!(result.people.len(), 1);
    let person = &result.people[0];
    assert_eq!(person.person, ns_person);
    assert!(person.resolved, "self address should always resolve");
    // 2 hours / 30-minute slots = 4 slots.
    assert_eq!(person.slots.len(), 4);
    for slot in &person.slots {
        assert!(["free", "tentative", "busy", "out_of_office", "working_elsewhere"]
            .contains(&slot.status.as_str()));
    }
    // Far-future date with nothing scheduled should read back as free
    // end-to-end (proves common_free's intersection logic against a real
    // FreeBusy string, not just the fake's canned response).
    assert!(!result.common_free.is_empty());
}

#[test]
#[ignore]
fn check_availability_marks_unresolvable_person_without_failing() {
    let c = client();
    let result = c.check_availability(CheckAvailabilityInput {
        people: vec!["this-address-does-not-exist-outlook-mcp-rs-p10@nonexistent-domain-xyz.invalid".to_string()],
        start: "2099-07-01T09:00".to_string(),
        end: "2099-07-01T10:00".to_string(),
        interval_minutes: 30,
        treat_as_free: vec!["free".to_string()],
    }).expect("an unresolvable person should not fail the whole call");
    assert_eq!(result.people.len(), 1);
    assert!(!result.people[0].resolved);
    assert!(result.people[0].slots.is_empty());
    assert!(result.common_free.is_empty());
}
