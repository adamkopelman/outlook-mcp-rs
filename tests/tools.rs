use std::sync::Arc;

use outlook_mcp_rs::outlook::fake::{FakeOutlookClient, EMAIL_ID};
use outlook_mcp_rs::server::{
    CompleteTaskParams, CreateDraftParams, CreateEventParams, CreateNoteParams, CreateTaskParams,
    DeleteEmailParams, DeleteEventParams, GetEmailParams, GetEventParams, GetNoteParams, ListAttachmentsParams,
    ListEmailsParams, ListEventsParams, ListTasksParams, OutlookMcpServer,
    RecurrenceParams, ReplyEmailParams, RespondToMeetingParams, SaveAttachmentsParams,
    SendEmailParams, UpdateEmailParams, UpdateEventParams,
};
use rmcp::handler::server::wrapper::Parameters;
use rmcp::model::CallToolResult;
use serde_json::{json, Value};

fn result_json(result: &CallToolResult) -> Value {
    let text = result.content[0]
        .as_text()
        .expect("expected text content")
        .text
        .clone();
    serde_json::from_str(&text).unwrap()
}

#[tokio::test]
async fn list_folders_records_call() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server.list_folders().await.unwrap();
    assert_eq!(fake.calls(), vec![("list_folders".to_string(), json!({}))]);
}

#[tokio::test]
async fn list_emails_passes_arguments() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEmailsParams = serde_json::from_value(json!({
        "folder": "sent", "count": 5, "unread_only": true
    }))
    .unwrap();
    server.list_emails(Parameters(params)).await.unwrap();
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "list_emails");
    assert_eq!(args["folder"], "sent");
    assert_eq!(args["count"], 5);
    assert_eq!(args["unread_only"], true);
}

#[tokio::test]
async fn list_emails_uses_defaults() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEmailsParams = serde_json::from_value(json!({})).unwrap();
    server.list_emails(Parameters(params)).await.unwrap();
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "list_emails");
    assert_eq!(args["folder"], "inbox");
    assert_eq!(args["count"], 10);
    assert_eq!(args["unread_only"], false);
}

#[tokio::test]
async fn list_emails_forwards_query_and_filters() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEmailsParams = serde_json::from_value(json!({
        "query": "invoice", "from": "ada@x.com", "category": "Work",
        "since_days": 30, "has_attachments": true, "flagged": true, "high_importance": true
    }))
    .unwrap();
    server.list_emails(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["query"], "invoice");
    assert_eq!(args["from"], "ada@x.com");
    assert_eq!(args["category"], "Work");
    assert_eq!(args["since_days"], 30);
    assert_eq!(args["has_attachments"], true);
    assert_eq!(args["flagged"], true);
    assert_eq!(args["high_importance"], true);
}

#[tokio::test]
async fn list_emails_returns_categories() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEmailsParams = serde_json::from_value(json!({})).unwrap();
    let result = server.list_emails(Parameters(params)).await.unwrap();
    let json = result_json(&result);
    assert_eq!(json[0]["categories"], serde_json::json!(["Work"]));
}

#[tokio::test]
async fn get_email_returns_body() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_email(Parameters(GetEmailParams {
            email_id: EMAIL_ID.to_string(),
            prefer_html: false,
        }))
        .await
        .unwrap();
    assert_eq!(result_json(&result)["body"], "Hi there");
}

#[tokio::test]
async fn get_email_includes_item_type() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_email(Parameters(GetEmailParams {
            email_id: EMAIL_ID.to_string(),
            prefer_html: false,
        }))
        .await
        .unwrap();
    assert_eq!(result_json(&result)["item_type"], "email");
    assert_eq!(result_json(&result)["is_meeting"], false);
}

#[tokio::test]
async fn send_email_passes_recipients_and_html_flag() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .send_email(Parameters(SendEmailParams {
            to: vec!["a@example.com".to_string(), "b@example.com".to_string()],
            subject: "Hi".to_string(),
            body: "Hello!".to_string(),
            cc: None,
            bcc: None,
            html: false,
            attachments: None,
        }))
        .await
        .unwrap();
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "send_email");
    assert_eq!(args["to"], json!(["a@example.com", "b@example.com"]));
    assert_eq!(args["html"], false);
}

#[tokio::test]
async fn create_draft_returns_draft_saved_status() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .create_draft(Parameters(CreateDraftParams {
            to: vec!["a@example.com".to_string()],
            subject: "Hi".to_string(),
            body: "Hello!".to_string(),
            cc: None,
            bcc: None,
            html: false,
            attachments: None,
        }))
        .await
        .unwrap();
    assert_eq!(result_json(&result)["status"], "draft_saved");
}

#[tokio::test]
async fn reply_email_passes_reply_all_and_send_flags() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .reply_email(Parameters(ReplyEmailParams {
            email_id: EMAIL_ID.to_string(),
            body: "Thanks!".to_string(),
            reply_all: true,
            html: false,
            send: false,
            attachments: None,
        }))
        .await
        .unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["reply_all"], true);
    assert_eq!(args["send"], false);
}

#[tokio::test]
async fn send_email_forwards_attachments() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: SendEmailParams = serde_json::from_value(json!({
        "to": ["a@x.com"], "subject": "Hi", "body": "yo",
        "attachments": ["C:/tmp/a.pdf", "C:/tmp/b.png"]
    })).unwrap();
    server.send_email(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["attachments"], serde_json::json!(["C:/tmp/a.pdf", "C:/tmp/b.png"]));
}

#[tokio::test]
async fn update_email_move_returns_new_id() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .update_email(Parameters(UpdateEmailParams {
            email_id: EMAIL_ID.to_string(),
            move_to: Some("Archive".to_string()),
            mark_read: None, flag: None, add_categories: None,
            remove_categories: None, importance: None,
        }))
        .await
        .unwrap();
    let v = result_json(&result);
    assert_eq!(v["id"], "new-entry|store-1");
    assert_eq!(v["status"], "updated");
    assert_eq!(v["changed"], serde_json::json!(["move_to"]));
}

#[tokio::test]
async fn update_email_state_only_keeps_same_id_and_lists_changes() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .update_email(Parameters(UpdateEmailParams {
            email_id: EMAIL_ID.to_string(),
            move_to: None,
            mark_read: Some(true),
            flag: Some("follow_up".to_string()),
            add_categories: Some(vec!["Work".to_string()]),
            remove_categories: None,
            importance: Some("high".to_string()),
        }))
        .await
        .unwrap();
    let v = result_json(&result);
    // No move → id unchanged.
    assert_eq!(v["id"], EMAIL_ID);
    assert_eq!(v["changed"], serde_json::json!(["mark_read", "flag", "add_categories", "importance"]));
    // The client saw the full update.
    let (name, args) = fake.calls().pop().unwrap();
    assert_eq!(name, "update_email");
    assert_eq!(args["flag"], "follow_up");
    assert_eq!(args["importance"], "high");
}

#[tokio::test]
async fn delete_email_records_call() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .delete_email(Parameters(DeleteEmailParams { email_id: EMAIL_ID.to_string() }))
        .await
        .unwrap();
    assert_eq!(
        fake.calls(),
        vec![("delete_email".to_string(), json!({"email_id": EMAIL_ID}))]
    );
}

#[tokio::test]
async fn client_error_propagates_as_tool_error() {
    let fake = Arc::new(FakeOutlookClient::new());
    fake.set_fail_with("Outlook exploded");
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEmailsParams = serde_json::from_value(json!({})).unwrap();
    let err = server.list_emails(Parameters(params)).await.unwrap_err();
    assert!(err.message.contains("Outlook exploded"));
}

#[tokio::test]
async fn list_events_passes_date_range() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .list_events(Parameters(ListEventsParams {
            start_date: Some("2026-06-10".to_string()),
            end_date: Some("2026-06-17".to_string()),
            query: None, category: None, show_as: None, my_response: None,
            attendees: None, attendee_role: None, meetings_only: false,
            all_day: None, calendar_of: None,
        }))
        .await
        .unwrap();
    let (name, args) = fake.calls().pop().unwrap();
    assert_eq!(name, "list_events");
    assert_eq!(args["start_date"], "2026-06-10");
    assert_eq!(args["end_date"], "2026-06-17");
}

#[tokio::test]
async fn list_events_forwards_all_filters() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .list_events(Parameters(ListEventsParams {
            start_date: None, end_date: None,
            query: Some("review".to_string()),
            category: Some("Work".to_string()),
            show_as: Some("busy".to_string()),
            my_response: Some("accepted".to_string()),
            attendees: Some(vec!["alice@example.com".to_string()]),
            attendee_role: Some("required".to_string()),
            meetings_only: true,
            all_day: Some(false),
            calendar_of: Some("bob@example.com".to_string()),
        }))
        .await
        .unwrap();
    let (name, args) = fake.calls().pop().unwrap();
    assert_eq!(name, "list_events");
    assert_eq!(args["query"], "review");
    assert_eq!(args["category"], "Work");
    assert_eq!(args["show_as"], "busy");
    assert_eq!(args["my_response"], "accepted");
    assert_eq!(args["attendees"], serde_json::json!(["alice@example.com"]));
    assert_eq!(args["attendee_role"], "required");
    assert_eq!(args["meetings_only"], true);
    assert_eq!(args["all_day"], false);
    assert_eq!(args["calendar_of"], "bob@example.com");
}

#[tokio::test]
async fn get_event_returns_subject_and_friendly_fields() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_event(Parameters(GetEventParams { event_id: EVENT_ID.to_string() }))
        .await
        .unwrap();
    let v = result_json(&result);
    assert_eq!(v["subject"], "Standup");
    // New enriched fields surface at the top level (EventDetail flattens the summary).
    assert_eq!(v["show_as"], "busy");
    assert_eq!(v["my_response"], "accepted");
    assert_eq!(v["required_attendees"], "");
    assert_eq!(v["optional_attendees"], "");
    // The old nested "response" key is gone (renamed to my_response in the summary).
    assert!(v.get("response").is_none());
}

#[tokio::test]
async fn create_event_passes_attendees() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .create_event(Parameters(CreateEventParams {
            subject: "Sync".to_string(),
            start: "2026-06-12T14:00".to_string(),
            end: "2026-06-12T15:00".to_string(),
            body: None,
            location: None,
            attendees: Some(vec!["a@example.com".to_string()]),
            required_attendees: None,
            optional_attendees: None,
            all_day: false,
            reminder_minutes: None,
            categories: None,
            show_as: None,
            send: true,
            recurrence: None,
        }))
        .await
        .unwrap();
    let (_, args) = &fake.calls()[0];
    // The legacy `attendees` alias merges into `required_attendees`.
    assert_eq!(args["required_attendees"], json!(["a@example.com"]));
}

#[tokio::test]
async fn create_event_status_reflects_attendees_and_send() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());

    let base = |required: Option<Vec<String>>, send: bool| CreateEventParams {
        subject: "Sync".to_string(),
        start: "2026-06-12T14:00".to_string(),
        end: "2026-06-12T15:00".to_string(),
        body: None, location: None, attendees: None,
        required_attendees: required, optional_attendees: None,
        all_day: false, reminder_minutes: None, categories: None, show_as: None,
        send,
        recurrence: None,
    };

    let r = server.create_event(Parameters(base(Some(vec!["a@example.com".to_string()]), true)))
        .await.unwrap();
    assert_eq!(result_json(&r)["status"], "meeting_sent");

    let r = server.create_event(Parameters(base(Some(vec!["a@example.com".to_string()]), false)))
        .await.unwrap();
    assert_eq!(result_json(&r)["status"], "meeting_saved");

    let r = server.create_event(Parameters(base(None, true))).await.unwrap();
    assert_eq!(result_json(&r)["status"], "saved");
}

#[tokio::test]
async fn create_event_forwards_recurrence() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .create_event(Parameters(CreateEventParams {
            subject: "Standup".to_string(),
            start: "2026-06-12T09:00".to_string(),
            end: "2026-06-12T09:15".to_string(),
            body: None, location: None, attendees: None,
            required_attendees: None, optional_attendees: None,
            all_day: false, reminder_minutes: None, categories: None, show_as: None,
            send: true,
            recurrence: Some(RecurrenceParams {
                pattern: "weekly".to_string(),
                interval: Some(1),
                days_of_week: Some(vec!["monday".to_string(), "wednesday".to_string()]),
                day_of_month: None,
                until: None,
                occurrences: Some(10),
            }),
        }))
        .await
        .unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["recurrence"]["pattern"], "weekly");
    assert_eq!(args["recurrence"]["days_of_week"], json!(["monday", "wednesday"]));
    assert_eq!(args["recurrence"]["occurrences"], 10);
}

#[tokio::test]
async fn get_event_recurrence_is_none_by_default() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_event(Parameters(GetEventParams { event_id: EVENT_ID.to_string() }))
        .await
        .unwrap();
    let v = result_json(&result);
    assert!(v["recurrence"].is_null());
}

#[tokio::test]
async fn respond_to_meeting_defaults_send_true() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: RespondToMeetingParams =
        serde_json::from_value(json!({"event_id": EVENT_ID, "response": "accept"})).unwrap();
    server.respond_to_meeting(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["response"], "accept");
    assert_eq!(args["send"], true);
}

#[tokio::test]
async fn update_event_lists_changed_fields() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .update_event(Parameters(UpdateEventParams {
            event_id: EVENT_ID.to_string(),
            subject: Some("Renamed sync".to_string()),
            start: None, end: None, location: None, body: None, all_day: None,
            reminder_minutes: None, show_as: Some("tentative".to_string()),
            add_categories: Some(vec!["Work".to_string()]),
            remove_categories: None,
            add_required_attendees: Some(vec!["a@example.com".to_string()]),
            add_optional_attendees: None, remove_attendees: None,
            send_update: true,
            recurrence: None, clear_recurrence: false,
        }))
        .await
        .unwrap();
    let v = result_json(&result);
    assert_eq!(v["status"], "updated");
    assert_eq!(v["id"], EVENT_ID);
    assert_eq!(
        v["changed"],
        json!(["subject", "show_as", "add_categories", "add_required_attendees"])
    );
    let (name, args) = fake.calls().pop().unwrap();
    assert_eq!(name, "update_event");
    assert_eq!(args["send_update"], true);
}

#[tokio::test]
async fn update_event_remove_attendees_is_tracked() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .update_event(Parameters(UpdateEventParams {
            event_id: EVENT_ID.to_string(),
            subject: None, start: None, end: None, location: None, body: None,
            all_day: None, reminder_minutes: None, show_as: None,
            add_categories: None, remove_categories: None,
            add_required_attendees: None, add_optional_attendees: None,
            remove_attendees: Some(vec!["a@example.com".to_string()]),
            send_update: false,
            recurrence: None, clear_recurrence: false,
        }))
        .await
        .unwrap();
    let v = result_json(&result);
    assert_eq!(v["changed"], json!(["remove_attendees"]));
}

#[tokio::test]
async fn update_event_forwards_recurrence() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .update_event(Parameters(UpdateEventParams {
            event_id: EVENT_ID.to_string(),
            subject: None, start: None, end: None, location: None, body: None,
            all_day: None, reminder_minutes: None, show_as: None,
            add_categories: None, remove_categories: None,
            add_required_attendees: None, add_optional_attendees: None, remove_attendees: None,
            send_update: false,
            recurrence: Some(RecurrenceParams {
                pattern: "daily".to_string(), interval: Some(2), days_of_week: None,
                day_of_month: None, until: Some("2099-06-01".to_string()), occurrences: None,
            }),
            clear_recurrence: false,
        }))
        .await
        .unwrap();
    let v = result_json(&result);
    assert_eq!(v["changed"], json!(["recurrence"]));
    let (_, args) = fake.calls().last().unwrap().clone();
    assert_eq!(args["recurrence"]["pattern"], "daily");
    assert_eq!(args["recurrence"]["until"], "2099-06-01");
}

#[tokio::test]
async fn update_event_forwards_clear_recurrence() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .update_event(Parameters(UpdateEventParams {
            event_id: EVENT_ID.to_string(),
            subject: None, start: None, end: None, location: None, body: None,
            all_day: None, reminder_minutes: None, show_as: None,
            add_categories: None, remove_categories: None,
            add_required_attendees: None, add_optional_attendees: None, remove_attendees: None,
            send_update: false,
            recurrence: None,
            clear_recurrence: true,
        }))
        .await
        .unwrap();
    let v = result_json(&result);
    assert_eq!(v["changed"], json!(["clear_recurrence"]));
}

#[tokio::test]
async fn delete_event_returns_deleted_status() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .delete_event(Parameters(DeleteEventParams {
            event_id: EVENT_ID.to_string(),
            send_cancellation: true,
        }))
        .await
        .unwrap();
    assert_eq!(result_json(&result)["status"], "deleted");
    let (name, args) = fake.calls().pop().unwrap();
    assert_eq!(name, "delete_event");
    assert_eq!(args["send_cancellation"], true);
}

#[tokio::test]
async fn list_attachments_returns_filename() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .list_attachments(Parameters(ListAttachmentsParams { email_id: EMAIL_ID.to_string() }))
        .await
        .unwrap();
    assert_eq!(result_json(&result)[0]["filename"], "report.pdf");
}

#[tokio::test]
async fn save_attachments_passes_dir_and_names() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .save_attachments(Parameters(SaveAttachmentsParams {
            email_id: EMAIL_ID.to_string(),
            save_dir: "/tmp/x".to_string(),
            attachment_names: Some(vec!["report.pdf".to_string()]),
        }))
        .await
        .unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["save_dir"], "/tmp/x");
    assert_eq!(args["attachment_names"], json!(["report.pdf"]));
}

// ---- Tasks ----

#[tokio::test]
async fn list_tasks_passes_include_completed() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .list_tasks(Parameters(ListTasksParams { include_completed: true }))
        .await
        .unwrap();
    assert_eq!(fake.calls(), vec![
        ("list_tasks".to_string(), json!({"include_completed": true})),
    ]);
}

#[tokio::test]
async fn create_task_passes_importance() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: CreateTaskParams = serde_json::from_value(json!({
        "subject": "Buy milk", "due_date": "2026-06-15", "importance": "high"
    })).unwrap();
    server.create_task(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["importance"], "high");
}

#[tokio::test]
async fn complete_task_records_call() {
    use outlook_mcp_rs::outlook::fake::TASK_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .complete_task(Parameters(CompleteTaskParams { task_id: TASK_ID.to_string() }))
        .await
        .unwrap();
    assert_eq!(fake.calls(), vec![
        ("complete_task".to_string(), json!({"task_id": TASK_ID})),
    ]);
}

// ---- Notes ----

#[tokio::test]
async fn list_notes_records_call() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server.list_notes().await.unwrap();
    assert_eq!(fake.calls(), vec![("list_notes".to_string(), json!({}))]);
}

#[tokio::test]
async fn get_note_returns_body() {
    use outlook_mcp_rs::outlook::fake::NOTE_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_note(Parameters(GetNoteParams { note_id: NOTE_ID.to_string() }))
        .await
        .unwrap();
    assert!(result_json(&result)["body"].as_str().unwrap().starts_with("Ideas"));
}

#[tokio::test]
async fn create_note_records_body() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .create_note(Parameters(CreateNoteParams { body: "Ideas\n- one".to_string() }))
        .await
        .unwrap();
    assert_eq!(fake.calls(), vec![
        ("create_note".to_string(), json!({"body": "Ideas\n- one"})),
    ]);
}
