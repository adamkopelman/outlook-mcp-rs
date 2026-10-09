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
use outlook_mcp_rs::outlook::{CheckAvailabilityInput, CreateEventInput, DraftUpdate, EmailQuery, EventQuery, OutlookClient, EmailUpdate, EventUpdate, NoteQuery, NoteUpdate, RecurrenceInput, TaskQuery, TaskUpdate, InlineImage};

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
        query: None, folder: "inbox".into(), count: 5, offset: 0, unread_only: false,
        from: None, to: None, category: None, received_after: None, received_before: None,
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
fn list_emails_offset_pages_tile_without_overlap() {
    let c = client();
    let page = |count: i32, offset: i32| -> Vec<String> {
        c.list_emails(EmailQuery {
            query: None, folder: "inbox".into(), count, offset, unread_only: false,
            from: None, to: None, category: None, received_after: None, received_before: None,
            since_days: None, has_attachments: None, flagged: false, high_importance: false,
        }).expect("list_emails should succeed against a live Outlook")
            .into_iter().map(|e| e.id).collect()
    };
    let p1 = page(5, 0);
    let p2 = page(5, 5);
    let both = page(10, 0);
    // Pages are disjoint, and together they are exactly one count=10 call
    // (assumes no mail arrives in the inbox mid-test).
    assert!(p1.iter().all(|id| !p2.contains(id)), "page 1 and page 2 overlap");
    assert_eq!([p1, p2].concat(), both);
}

#[test]
#[ignore]
fn create_draft_then_delete_round_trips() {
    let c = client();
    let created = c.create_draft(
        vec!["nobody@example.invalid".to_string()],
        "outlook-mcp-rs live test draft".to_string(),
        "This draft is created and deleted by an automated test.".to_string(),
        None, None, false, None, None,
    ).expect("create_draft should succeed");
    let id = created["id"].as_str().expect("create_draft returns an id").to_string();
    c.delete_email(id, false).expect("cleanup: delete_email should succeed");
}

#[test]
#[ignore]
fn permanent_delete_of_draft_skips_deleted_items() {
    let c = client();
    let subject = "outlook-mcp-rs permanent delete probe zzqx-7731";
    let created = c.create_draft(
        vec!["nobody@example.invalid".to_string()],
        subject.to_string(),
        "This draft is created and permanently deleted by an automated test.".to_string(),
        None, None, false, None, None,
    ).expect("create_draft should succeed");
    let id = created["id"].as_str().expect("create_draft returns an id").to_string();

    let result = c.delete_email(id.clone(), true).expect("permanent delete_email should succeed");
    assert_eq!(result["permanent"], true);

    // The old id must no longer resolve...
    assert!(c.get_email(id, false, None).is_err(), "deleted draft's id should no longer resolve");
    // ...and nothing with that subject may be sitting in Deleted Items.
    let leftovers = c.list_emails(EmailQuery {
        query: Some(subject.to_string()), folder: "deleted".into(), count: 50, offset: 0,
        unread_only: false, from: None, to: None, category: None, received_after: None,
        received_before: None, since_days: None, has_attachments: None,
        flagged: false, high_importance: false,
    }).expect("list_emails on Deleted Items should succeed");
    assert!(
        !leftovers.iter().any(|e| e.subject == subject),
        "a permanently deleted draft must not remain in Deleted Items"
    );
}

#[test]
#[ignore]
fn create_task_update_task_marks_complete() {
    let c = client();
    let created = c.create_task(
        "outlook-mcp-rs live test task".to_string(), None, None, "normal".to_string(),
        None, None, None,
    ).expect("create_task should succeed");
    let id = created["id"].as_str().unwrap().to_string();
    c.update_task(TaskUpdate { task_id: id.clone(), mark_complete: Some(true), ..Default::default() })
        .expect("update_task should succeed");
    let tasks = c.list_tasks(TaskQuery { include_completed: true, ..Default::default() })
        .expect("list_tasks should succeed");
    assert!(tasks.iter().any(|t| t.id == id && t.complete));
    c.delete_task(id).expect("cleanup delete_task");
}

#[test]
#[ignore]
fn create_note_then_get_it_back() {
    let c = client();
    let created = c.create_note("outlook-mcp-rs live test note".to_string(), None, None)
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
    c.delete_email(id, false).expect("cleanup delete");
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
        query: None, folder: "inbox".into(), count: 25, offset: 0, unread_only: false,
        from: None, to: None, category: None, received_after: None, received_before: None,
        since_days: None, has_attachments: None, flagged: false, high_importance: false,
    }).expect("plain list should work");
    // A query that almost certainly matches nothing should return <= all.
    let filtered = c.list_emails(EmailQuery {
        query: Some("zzqx-improbable-token-9137".into()),
        folder: "inbox".into(), count: 25, offset: 0, unread_only: false,
        from: None, to: None, category: None, received_after: None, received_before: None,
        since_days: None, has_attachments: None, flagged: false, high_importance: false,
    }).expect("query list should work");
    assert!(filtered.len() <= all.len());
}

#[test]
#[ignore]
fn list_emails_query_matches_real_body_text() {
    let c = client();
    let token = "zzbodytoken8842";
    let created = c.create_draft(
        vec!["nobody@example.invalid".to_string()],
        "[outlook-mcp-rs body-search live] draft probe".to_string(),
        format!("this draft's body contains {token} and the subject does not"),
        None, None, false, None, None,
    ).expect("create_draft should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let found = c.list_emails(EmailQuery {
        query: Some(token.to_string()), folder: "drafts".into(), count: 25, offset: 0,
        unread_only: false, from: None, to: None, category: None, received_after: None,
        received_before: None, since_days: None, has_attachments: None,
        flagged: false, high_importance: false,
    }).expect("list_emails query should succeed");

    c.delete_email(id.clone(), false).expect("cleanup: delete the draft");

    assert!(
        found.iter().any(|e| e.id == id),
        "list_emails query {token:?} should find a draft whose ONLY occurrence \
         of that token is in the body, proving the existing @SQL textdescription \
         clause matches body content and not just subject/sender"
    );
}

#[test]
#[ignore]
fn list_emails_hebrew_query_finds_matching_subject() {
    // Issue #2: take a Hebrew word from a recent inbox subject (or pin one via
    // OUTLOOK_MCP_LIVE_HEBREW_QUERY) and check list_emails(query=...) finds
    // that email. Skips (passes) if there's no Hebrew subject to use.
    let c = client();
    let inbox = |query: Option<String>| EmailQuery {
        query, folder: "inbox".into(), count: 50, offset: 0, unread_only: false,
        from: None, to: None, category: None, received_after: None, received_before: None,
        since_days: None, has_attachments: None, flagged: false, high_importance: false,
    };
    let is_hebrew = |ch: char| ('\u{0590}'..='\u{05FF}').contains(&ch);
    let mut term = std::env::var("OUTLOOK_MCP_LIVE_HEBREW_QUERY").ok().filter(|s| !s.is_empty());
    let mut expected_id = None;
    if term.is_none() {
        let recent = c.list_emails(inbox(None)).expect("plain list should work");
        'outer: for email in &recent {
            for word in email.subject.split(|ch: char| !is_hebrew(ch)) {
                if word.chars().count() >= 3 {
                    term = Some(word.to_string());
                    expected_id = Some(email.id.clone());
                    break 'outer;
                }
            }
        }
    }
    let Some(term) = term else {
        eprintln!("skipping: no Hebrew subject in the 50 newest inbox emails");
        return;
    };
    let found = c.list_emails(inbox(Some(term.clone()))).expect("Hebrew query list should work");
    assert!(!found.is_empty(), "no results for Hebrew query {term:?}");
    if let Some(id) = expected_id {
        assert!(
            found.iter().any(|e| e.id == id),
            "Hebrew query {term:?} should find the inbox email whose subject it came from"
        );
    }
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
        Some(vec![path_str]), None,
    ).expect("create_draft with attachment should succeed");
    let id = created["id"].as_str().expect("draft id").to_string();
    c.delete_email(id, false).expect("cleanup: delete the draft");
    let _ = std::fs::remove_file(&path);
}

/// Read a draft's first attachment's MAPI Content-ID / hidden flag straight
/// from COM (the client API doesn't expose PropertyAccessor reads).
fn first_attachment_cid_and_hidden(item_id: &str) -> (String, bool) {
    use outlook_mcp_rs::constants as k;
    use outlook_mcp_rs::outlook::com::{
        call_method, create_com_object, get_property, parse_item_id, variant_from_i32,
        variant_from_str, variant_to_bool, variant_to_string, ComGuard,
    };
    use windows::Win32::System::Com::IDispatch;
    let disp = |v: windows::Win32::System::Variant::VARIANT| IDispatch::try_from(&v).expect("IDispatch");
    let _guard = ComGuard::new().expect("CoInitialize");
    let app = create_com_object("Outlook.Application").expect("Outlook.Application");
    let ns = disp(call_method(&app, "GetNamespace", &mut [variant_from_str("MAPI")]).unwrap());
    let (entry, store) = parse_item_id(item_id).unwrap();
    let item = disp(call_method(&ns, "GetItemFromID", &mut [variant_from_str(&entry), variant_from_str(&store)]).unwrap());
    let atts = disp(get_property(&item, "Attachments").unwrap());
    let att = disp(call_method(&atts, "Item", &mut [variant_from_i32(1)]).unwrap());
    let pa = disp(get_property(&att, "PropertyAccessor").unwrap());
    let cid = variant_to_string(&call_method(&pa, "GetProperty", &mut [variant_from_str(k::PR_ATTACH_CONTENT_ID)]).unwrap());
    let hidden = call_method(&pa, "GetProperty", &mut [variant_from_str(k::PR_ATTACHMENT_HIDDEN)])
        .ok()
        .and_then(|v| variant_to_bool(&v))
        .unwrap_or(false);
    (cid, hidden)
}

#[test]
#[ignore]
fn create_draft_with_inline_base64_image_sets_content_id() {
    // A 1x1 transparent PNG.
    const PNG_1X1_B64: &str = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";
    let c = WindowsOutlookClient::new();
    let created = c.create_draft(
        vec!["nobody@example.invalid".to_string()],
        "outlook-mcp-rs inline image test".to_string(),
        "<p>Pixel:</p><img src=\"cid:pixel\">".to_string(),
        None, None, true, None,
        Some(vec![InlineImage {
            content_id: "pixel".into(),
            data_base64: Some(format!("data:image/png;base64,{PNG_1X1_B64}")),
            ..Default::default()
        }]),
    ).expect("create_draft with an inline image should succeed");
    let id = created["id"].as_str().expect("draft id").to_string();
    // Read back before asserting so cleanup still runs on a mismatch.
    let read_back = std::panic::catch_unwind(|| first_attachment_cid_and_hidden(&id));
    c.delete_email(id, false).expect("cleanup: delete the draft");
    let (cid, hidden) = read_back.expect("reading the attachment's MAPI properties should succeed");
    assert_eq!(cid, "pixel");
    // Not asserted: whether PR_ATTACHMENT_HIDDEN survives Save varies by
    // Outlook version. Printed for manual verification.
    eprintln!("inline image attachment hidden flag after Save: {hidden}");
}

#[test]
#[ignore]
fn list_attachments_reports_metadata_for_a_draft_attachment() {
    let dir = std::env::temp_dir();
    let path = dir.join("outlook-mcp-rs-live-attach-meta.txt");
    std::fs::write(&path, b"live attachment metadata test").expect("write temp file");
    let path_str = path.to_string_lossy().to_string();

    let c = WindowsOutlookClient::new();
    let created = c.create_draft(
        vec!["nobody@example.invalid".to_string()],
        "outlook-mcp-rs attachment metadata test".to_string(),
        "see attached".to_string(),
        None, None, false,
        Some(vec![path_str]), None,
    ).expect("create_draft with attachment should succeed");
    let id = created["id"].as_str().expect("draft id").to_string();

    // Capture the result before cleanup so a failed assertion doesn't leak the draft.
    let listed = c.list_attachments(id.clone());
    c.delete_email(id, false).expect("cleanup: delete the draft");
    let _ = std::fs::remove_file(&path);

    let atts = listed.expect("list_attachments should succeed");
    assert_eq!(atts.len(), 1);
    let v = serde_json::to_value(&atts[0]).unwrap();
    assert_eq!(v["index"], 1);
    assert_eq!(v["filename"], "outlook-mcp-rs-live-attach-meta.txt");
    assert!(v["size"].as_i64().unwrap() > 0);
    assert_eq!(v["type"], "file");
    // A plain file attachment has no Content-ID; its MIME type comes from the
    // tag or the .txt extension.
    assert!(v["content_id"].is_null());
    assert_eq!(v["mime_type"], "text/plain");
    assert_eq!(v["hidden"], false);
}

#[test]
#[ignore]
fn inline_flag_consistent_on_a_real_inbox_email() {
    let c = WindowsOutlookClient::new();
    let list = c.list_emails(EmailQuery {
        query: None, folder: "inbox".into(), count: 25, offset: 0, unread_only: false,
        from: None, to: None, category: None, received_after: None, received_before: None,
        since_days: None, has_attachments: None, flagged: false, high_importance: false,
    }).expect("list");
    let Some(email) = list.iter().find(|e| e.has_attachments) else {
        eprintln!("skipping: none of the newest 25 inbox emails has attachments");
        return;
    };
    let atts = c.list_attachments(email.id.clone()).expect("list_attachments");
    for att in &atts {
        let v = serde_json::to_value(att).unwrap();
        assert!(v["is_inline"].is_boolean(), "is_inline missing: {v}");
        // Inline content is always cid:-addressable.
        if att.is_inline {
            assert!(att.content_id.is_some(), "inline attachment without a content_id: {v}");
        }
    }
}

#[test]
#[ignore]
fn get_inline_image_round_trips_a_real_content_id() {
    // Read-only: looks for an existing inline attachment; never sends mail.
    use base64::Engine as _;
    let c = WindowsOutlookClient::new();
    let emails = c.list_emails(EmailQuery {
        query: None, folder: "inbox".into(), count: 25, offset: 0, unread_only: false,
        from: None, to: None, category: None, received_after: None, received_before: None,
        since_days: None, has_attachments: None, flagged: false, high_importance: false,
    }).expect("list_emails");
    let found = emails.iter().find_map(|e| {
        // Some item types don't support attachments; just skip those.
        let atts = c.list_attachments(e.id.clone()).ok()?;
        atts.into_iter()
            .find(|a| a.content_id.is_some() && a.size <= 10 * 1024 * 1024)
            .map(|a| (e.id.clone(), a))
    });
    let Some((email_id, att)) = found else {
        eprintln!("skipping: no attachment with a Content-ID in the newest 25 inbox items");
        return;
    };
    let cid = att.content_id.clone().unwrap();
    let image = c.get_inline_image(email_id.clone(), format!("cid:{cid}"), None).expect("get_inline_image");
    assert_eq!(image.content_id, cid);
    let (header, payload) = image.data_uri.split_once(',').expect("data URI has a comma");
    assert_eq!(header, format!("data:{};base64", image.mime_type));
    let data = base64::engine::general_purpose::STANDARD.decode(payload).expect("valid base64");
    assert_eq!(data.len(), image.size);
    assert!(image.size > 0);
    assert!(image.context.is_none(), "context only when context_lines is given");

    // Same image with surrounding text requested: `context` is always present
    // (possibly "" if the HTML body never references this Content-ID).
    let with_ctx = c.get_inline_image(email_id, cid.clone(), Some(3)).expect("get_inline_image with context");
    assert_eq!(with_ctx.content_id, cid);
    let context = with_ctx.context.expect("context requested");
    assert!(context.lines().count() <= 3, "{context:?}");
}

#[test]
#[ignore]
fn get_email_reports_item_type_for_real_inbox_item() {
    let c = WindowsOutlookClient::new();
    let list = c.list_emails(EmailQuery {
        query: None, folder: "inbox".into(), count: 1, offset: 0, unread_only: false,
        from: None, to: None, category: None, received_after: None, received_before: None,
        since_days: None, has_attachments: None, flagged: false, high_importance: false,
    }).expect("list");
    if let Some(first) = list.first() {
        let detail = c.get_email(first.id.clone(), false, None).expect("get_email");
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
fn get_email_reports_truncation_and_honours_max_body_chars() {
    let c = client();
    // A >100k-char HTML body (well past the 100,000 default cut). A draft is
    // a safe, disposable target (never sent).
    let filler = "<p>outlook-mcp-rs truncation live test line.</p>\n".repeat(3_000);
    let html = format!("<html><body>{filler}</body></html>");
    assert!(html.chars().count() > 100_000);
    let created = c.create_draft(
        vec!["nobody@example.invalid".to_string()],
        "outlook-mcp-rs truncation live test".to_string(),
        html,
        None, None, true, None, None,
    ).expect("create_draft");
    let id = created["id"].as_str().expect("draft id").to_string();

    let result = || {
        // Default limit: HTML is cut and the flags/lengths say so. Outlook
        // may normalise the stored HTML, so compare against what it reports.
        let detail = c.get_email(id.clone(), true, None).expect("get_email default");
        let html_length = detail.html_length.expect("html_length with prefer_html");
        assert!(html_length > 100_000, "html_length {html_length}");
        assert_eq!(detail.html_truncated, Some(true));
        assert!(detail.html_body.as_deref().unwrap().contains("[... truncated at 100000 characters]"));
        assert_eq!(detail.body_truncated, detail.body_length > 100_000);

        // Re-fetch with a limit of the reported length: the complete HTML.
        let limit = html_length.max(detail.body_length) as u32;
        let full = c.get_email(id.clone(), true, Some(limit)).expect("get_email larger limit");
        assert_eq!(full.html_truncated, Some(false));
        assert!(!full.body_truncated);
        let full_html = full.html_body.expect("html_body");
        assert_eq!(full_html.chars().count(), html_length);
        assert!(full_html.contains("</body>"));

        // Without prefer_html the HTML fields are absent.
        let plain = c.get_email(id.clone(), false, None).expect("get_email plain");
        assert!(plain.html_truncated.is_none() && plain.html_length.is_none());
    };
    let outcome = std::panic::catch_unwind(std::panic::AssertUnwindSafe(result));
    c.delete_email(id, false).expect("cleanup delete");
    if let Err(panic) = outcome {
        std::panic::resume_unwind(panic);
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
        None, None, false, None, None,
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
    let detail = c.get_email(id.clone(), false, None).expect("get_email");
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
    let redetail = c.get_email(id.clone(), false, None).expect("get_email after unread");
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
    c.delete_email(new_id, false).expect("cleanup delete");
}

#[test]
#[ignore]
fn update_draft_edits_subject_body_and_recipients() {
    let c = WindowsOutlookClient::new();
    // A draft is a safe, disposable target. update_draft only saves; it never sends.
    let created = c.create_draft(
        vec!["nobody@example.invalid".to_string()],
        "outlook-mcp-rs update_draft live test".to_string(),
        "original body".to_string(),
        None, None, false, None, None,
    ).expect("create_draft");
    let id = created["id"].as_str().expect("draft id").to_string();

    let res = c.update_draft(DraftUpdate {
        draft_id: id.clone(),
        subject: Some("outlook-mcp-rs update_draft live test (edited)".to_string()),
        body: Some("edited body zzdraftedit5521".to_string()),
        to: Some(vec!["someone-else@example.invalid".to_string()]),
        ..Default::default()
    });
    // Read back before asserting so a failure still cleans up.
    let detail = c.get_email(id.clone(), false, None);
    c.delete_email(id.clone(), false).expect("cleanup: delete the draft");

    let res = res.expect("update_draft");
    assert_eq!(res["status"], "draft_updated");
    assert_eq!(res["changed"], serde_json::json!(["subject", "body", "to"]));
    let dv = serde_json::to_value(detail.expect("get_email")).unwrap();
    assert_eq!(dv["subject"], "outlook-mcp-rs update_draft live test (edited)");
    assert!(dv["body"].as_str().unwrap().contains("zzdraftedit5521"));
    let to = dv["to"].as_str().unwrap();
    assert!(to.contains("someone-else@example.invalid"), "to was {to:?}");
    assert!(!to.contains("nobody@example.invalid"), "to was {to:?}");
}

#[test]
#[ignore]
fn send_with_missing_attachment_errors_before_sending() {
    let c = WindowsOutlookClient::new();
    let err = c.send_email(
        vec!["nobody@example.invalid".to_string()],
        "should not send".to_string(), "body".to_string(),
        None, None, false,
        Some(vec!["C:/definitely/does/not/exist/nope.pdf".to_string()]), None,
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
    // Set OUTLOOK_MCP_TEST_EMAIL to your SMTP address to run this (same
    // convention as list_events_calendar_of_self_opens_own_calendar above).
    let ns_person = match std::env::var("OUTLOOK_MCP_TEST_EMAIL") {
        Ok(v) if !v.is_empty() => v,
        _ => {
            eprintln!("skipping: set OUTLOOK_MCP_TEST_EMAIL to your address");
            return;
        }
    };
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

#[test]
#[ignore]
fn list_tasks_filters_and_create_task_additions_round_trip() {
    let c = client();
    let created = c.create_task(
        "[outlook-mcp-rs P11 live] filtered task".to_string(),
        None,
        None,
        "high".to_string(),
        Some(vec!["Red Category".to_string()]),
        Some("2099-01-01".to_string()),
        Some("2099-01-01T09:00".to_string()),
    ).expect("create_task with additions should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let found = c.list_tasks(TaskQuery {
        include_completed: false,
        category: Some("Red Category".to_string()),
        importance: Some("high".to_string()),
        query: Some("filtered task".to_string()),
    }).expect("list_tasks should succeed");
    assert!(found.iter().any(|t| t.id == id), "filtered list_tasks should find the new task");

    c.delete_task(id).expect("cleanup delete_task");
}

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

#[test]
#[ignore]
fn update_task_marks_complete_then_reopens() {
    let c = client();
    let created = c.create_task(
        "[outlook-mcp-rs P11 live] update probe".to_string(),
        None, None, "normal".to_string(), None, None, None,
    ).expect("create_task should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let updated = c.update_task(TaskUpdate {
        task_id: id.clone(),
        mark_complete: Some(true),
        ..Default::default()
    }).expect("update_task mark_complete should succeed");
    assert!(updated["changed"].as_array().unwrap().iter().any(|v| v == "mark_complete"));

    let after_complete = c.list_tasks(TaskQuery { include_completed: true, ..Default::default() })
        .expect("list_tasks should succeed");
    let task = after_complete.iter().find(|t| t.id == id).expect("task should still exist");
    assert!(task.complete);

    let reopened = c.update_task(TaskUpdate {
        task_id: id.clone(),
        mark_complete: Some(false),
        ..Default::default()
    }).expect("update_task reopen should succeed");
    assert!(reopened["changed"].as_array().unwrap().iter().any(|v| v == "mark_complete"));

    c.delete_task(id).expect("cleanup delete_task");
}

#[test]
#[ignore]
fn delete_task_removes_it() {
    let c = client();
    let created = c.create_task(
        "[outlook-mcp-rs P11 live] delete probe".to_string(),
        None, None, "normal".to_string(), None, None, None,
    ).expect("create_task should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let deleted = c.delete_task(id).expect("delete_task should succeed");
    assert_eq!(deleted["status"], "deleted");
}

#[test]
#[ignore]
fn list_notes_filters_and_create_note_additions_round_trip() {
    let c = client();
    let created = c.create_note(
        "outlook-mcp-rs P12 live filtered note - remember to renew".to_string(),
        Some(vec!["Green Category".to_string()]),
        Some("green".to_string()),
    ).expect("create_note with additions should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let found = c.list_notes(NoteQuery {
        category: Some("Green Category".to_string()),
        query: Some("renew".to_string()),
    }).expect("list_notes should succeed");
    assert!(found.iter().any(|n| n.id == id), "filtered list_notes should find the new note");

    c.delete_note(id).expect("cleanup delete_note");
}

#[test]
#[ignore]
fn get_note_includes_modified_after_update() {
    let c = client();
    let created = c.create_note(
        "outlook-mcp-rs P12 live modified probe".to_string(), None, None,
    ).expect("create_note should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    // Outlook already sets LastModificationTime on the initial Save() inside
    // create_note, so `modified.is_some()` alone (checked only after the
    // update) would pass even if update_note were a no-op. Capture the
    // baseline first and assert it's still populated and non-decreasing
    // after the edit — a strict "must be later" check would risk flaking on
    // same-second saves if Outlook's timestamp resolution is coarse.
    let before = c.get_note(id.clone()).expect("get_note (baseline) should succeed");
    let before_modified = before.modified.expect("modified should be populated right after create_note's own Save()");

    c.update_note(NoteUpdate {
        note_id: id.clone(),
        body: Some("outlook-mcp-rs P12 live modified probe (edited)".to_string()),
        ..Default::default()
    }).expect("update_note should succeed");

    let note = c.get_note(id.clone()).expect("get_note (after update) should succeed");
    let after_modified = note.modified.expect("modified should still be populated after an edit");
    assert!(after_modified >= before_modified,
        "modified ({after_modified}) should not go backwards after update_note ({before_modified})");
    assert!(note.body.starts_with("outlook-mcp-rs P12 live modified probe (edited)"));

    c.delete_note(id).expect("cleanup delete_note");
}

#[test]
#[ignore]
fn update_note_manages_categories_and_color() {
    let c = client();
    let created = c.create_note(
        "outlook-mcp-rs P12 live category probe".to_string(), None, None,
    ).expect("create_note should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let updated = c.update_note(NoteUpdate {
        note_id: id.clone(),
        add_categories: Some(vec!["Blue Category".to_string()]),
        color: Some("blue".to_string()),
        ..Default::default()
    }).expect("update_note should succeed");
    assert!(updated["changed"].as_array().unwrap().iter().any(|v| v == "add_categories"));
    assert!(updated["changed"].as_array().unwrap().iter().any(|v| v == "color"));

    let note = c.get_note(id.clone()).expect("get_note should succeed");
    assert!(note.summary.categories.iter().any(|cat| cat == "Blue Category"));

    c.delete_note(id).expect("cleanup delete_note");
}

#[test]
#[ignore]
fn delete_note_removes_it() {
    let c = client();
    let created = c.create_note(
        "outlook-mcp-rs P12 live delete probe".to_string(), None, None,
    ).expect("create_note should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let deleted = c.delete_note(id).expect("delete_note should succeed");
    assert_eq!(deleted["status"], "deleted");
}

#[test]
#[ignore]
fn list_emails_to_filter_matches_draft_recipient() {
    let c = client();
    let created = c.create_draft(
        vec!["nobody@example.invalid".to_string()],
        "[outlook-mcp-rs to-filter live] draft probe".to_string(),
        "recipient filter probe; never sent".to_string(),
        None, None, false, None, None,
    ).expect("create_draft should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let query = |to: &str| EmailQuery {
        query: None, folder: "drafts".into(), count: 50, offset: 0, unread_only: false,
        from: None, to: Some(to.to_string()), category: None, received_after: None,
        received_before: None, since_days: None, has_attachments: None,
        flagged: false, high_importance: false,
    };
    let hit = c.list_emails(query("nobody@example.invalid"));
    let miss = c.list_emails(query("someone-else@example.invalid"));

    c.delete_email(id.clone(), false).expect("cleanup: delete the draft");

    let hit = hit.expect("list_emails to=nobody should succeed");
    let miss = miss.expect("list_emails to=someone-else should succeed");
    assert!(hit.iter().any(|e| e.id == id), "to filter should find the draft addressed to nobody@example.invalid");
    assert!(!miss.iter().any(|e| e.id == id), "to filter must not match a different recipient");
}

/// Issue #3: Hebrew written through COM comes back byte-identical, both via
/// `get_email` and via a Hebrew `list_emails` query (Restrict filter text).
#[test]
#[ignore]
fn hebrew_subject_and_body_round_trip_through_com() {
    let c = client();
    let subject = "[outlook-mcp-rs utf8 live] מייל שיקוף".to_string();
    let body = "סיכום עשייה — שורה ראשונה".to_string();
    let created = c.create_draft(
        vec!["nobody@example.invalid".to_string()],
        subject.clone(), body.clone(), None, None, false, None, None,
    ).expect("create_draft should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let detail = c.get_email(id.clone(), false, None);
    let found = c.list_emails(EmailQuery {
        query: Some("מייל שיקוף".to_string()), folder: "drafts".into(), count: 25, offset: 0,
        unread_only: false, from: None, to: None, category: None, received_after: None,
        received_before: None, since_days: None, has_attachments: None,
        flagged: false, high_importance: false,
    });

    c.delete_email(id.clone(), false).expect("cleanup: delete the draft");

    let detail = detail.expect("get_email should succeed");
    assert_eq!(detail.summary.subject.as_bytes(), subject.as_bytes());
    assert!(
        detail.body.contains(&body),
        "body should contain the Hebrew text verbatim, got {:?}", detail.body
    );
    let found = found.expect("list_emails with a Hebrew query should succeed");
    assert!(found.iter().any(|e| e.id == id && e.subject == subject));
}

// ---- Issue #1: date filters must not depend on the regional date format ----
//
// Run these on a day-first locale (en-GB, en-IL, de-DE…) as well as en-US:
// before the fix, `Restrict` read the US-format filter in the user's order, so
// 2026-09-11..2026-09-18 returned nothing and 2026-11-09..2026-12-09 returned
// the September events. Only personal appointments are created (nothing is
// sent), and every one is deleted before the assertions run.

const ISSUE_1_SUBJECT: &str = "outlook-mcp-rs issue-1 date window probe";

fn issue_1_appointment(c: &WindowsOutlookClient, start: &str, end: &str, recurrence: Option<RecurrenceInput>) -> String {
    let created = c.create_event(CreateEventInput {
        subject: ISSUE_1_SUBJECT.to_string(),
        start: start.to_string(),
        end: end.to_string(),
        body: None, location: None, required_attendees: None, optional_attendees: None,
        all_day: false, reminder_minutes: None, categories: None, show_as: None,
        send: false,
        recurrence,
    }).expect("create_event");
    created["id"].as_str().expect("event id").to_string()
}

/// Starts (ISO) of the probe events `list_events` returns for the range.
fn issue_1_probe_starts(c: &WindowsOutlookClient, start: &str, end: &str) -> Result<Vec<String>, String> {
    c.list_events(EventQuery {
        start_date: Some(start.to_string()),
        end_date: Some(end.to_string()),
        query: Some(ISSUE_1_SUBJECT.to_string()),
        ..Default::default()
    })
    .map(|events| events.into_iter().filter_map(|e| e.start).collect())
    .map_err(|e| e.0)
}

#[test]
#[ignore]
fn list_events_date_range_is_read_the_same_in_every_locale() {
    let c = client();
    let ids = [
        issue_1_appointment(&c, "2026-09-10T23:00", "2026-09-10T23:30", None), // day before
        issue_1_appointment(&c, "2026-09-11T10:00", "2026-09-11T10:30", None), // first day
        issue_1_appointment(&c, "2026-09-18T16:00", "2026-09-18T16:30", None), // last day
        issue_1_appointment(&c, "2026-09-19T08:00", "2026-09-19T08:30", None), // day after
    ];
    let week = issue_1_probe_starts(&c, "2026-09-11", "2026-09-18");
    // The day/month-swapped range of the issue's repro.
    let swapped = issue_1_probe_starts(&c, "2026-11-09", "2026-12-09");
    // A start with a time: exact to the minute, not just the day.
    let from_10_30 = issue_1_probe_starts(&c, "2026-09-11T10:30", "2026-09-18");
    for id in ids {
        c.delete_event(id, false).expect("cleanup delete_event");
    }
    let mut week = week.expect("list_events 2026-09-11..2026-09-18");
    week.sort();
    assert_eq!(week, vec!["2026-09-11T10:00:00", "2026-09-18T16:00:00"]);
    assert_eq!(swapped.expect("list_events 2026-11-09..2026-12-09"), Vec::<String>::new());
    assert_eq!(from_10_30.expect("list_events from 10:30"), vec!["2026-09-18T16:00:00"]);
}

#[test]
#[ignore]
fn list_events_returns_only_in_range_occurrences_of_an_earlier_series() {
    // A weekly series that starts weeks before the window: its occurrences
    // under IncludeRecurrences must be filtered to the window without tripping
    // the "Outlook misread the date filter" check.
    let c = client();
    let id = issue_1_appointment(&c, "2026-08-03T09:00", "2026-08-03T09:30", Some(RecurrenceInput {
        pattern: "weekly".to_string(),
        interval: Some(1),
        days_of_week: Some(vec!["monday".to_string()]),
        day_of_month: None,
        until: None,
        occurrences: Some(10),
    }));
    let week = issue_1_probe_starts(&c, "2026-09-11", "2026-09-18");
    c.delete_event(id, false).expect("cleanup delete_event");
    assert_eq!(week.expect("list_events over a recurring series"), vec!["2026-09-14T09:00:00"]);
}

#[test]
#[ignore]
fn list_emails_received_after_matches_an_unfiltered_scan() {
    let c = client();
    let query = |received_after: Option<String>, since_days: Option<i32>| {
        c.list_emails(EmailQuery {
            query: None, folder: "inbox".into(), count: 200, offset: 0, unread_only: false,
            from: None, to: None, category: None, received_after, received_before: None,
            since_days, has_attachments: None, flagged: false, high_importance: false,
        }).expect("list_emails should succeed against a live Outlook")
    };
    let cutoff = (chrono::Local::now().naive_local() - chrono::Duration::days(30))
        .format("%Y-%m-%dT%H:%M:%S")
        .to_string();
    let all = query(None, None);
    let expected: Vec<String> = all
        .iter()
        .filter(|e| e.received.as_deref().is_some_and(|r| r >= cutoff.as_str()))
        .map(|e| e.id.clone())
        .collect();
    if expected.len() == all.len() {
        eprintln!("skipping: more than 200 inbox emails in the last 30 days");
        return;
    }
    let filtered: Vec<String> = query(Some(cutoff.clone()), None).into_iter().map(|e| e.id).collect();
    assert_eq!(filtered, expected, "received_after {cutoff}");
    let since: Vec<String> = query(None, Some(30)).into_iter().map(|e| e.id).collect();
    assert_eq!(since, expected, "since_days 30");
}
