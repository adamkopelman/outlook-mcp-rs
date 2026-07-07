use std::sync::Arc;

use outlook_mcp_rs::outlook::fake::{FakeOutlookClient, EMAIL_ID};
use outlook_mcp_rs::server::{
    CompleteTaskParams, CreateDraftParams, CreateEventParams, CreateNoteParams, CreateTaskParams,
    DeleteEmailParams, GetEmailParams, GetEventParams, GetNoteParams, ListAttachmentsParams,
    ListEmailsParams, ListEventsParams, ListTasksParams, MoveEmailParams, OutlookMcpServer,
    ReplyEmailParams, RespondToMeetingParams, SaveAttachmentsParams,
    SendEmailParams,
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
async fn move_email_returns_new_id() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .move_email(Parameters(MoveEmailParams {
            email_id: EMAIL_ID.to_string(),
            target_folder: "Archive".to_string(),
        }))
        .await
        .unwrap();
    assert_eq!(result_json(&result)["id"], "new-entry|store-1");
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
        }))
        .await
        .unwrap();
    assert_eq!(fake.calls(), vec![
        ("list_events".to_string(), json!({"start_date": "2026-06-10", "end_date": "2026-06-17"})),
    ]);
}

#[tokio::test]
async fn get_event_returns_subject() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_event(Parameters(GetEventParams { event_id: EVENT_ID.to_string() }))
        .await
        .unwrap();
    assert_eq!(result_json(&result)["subject"], "Standup");
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
            all_day: false,
            reminder_minutes: None,
        }))
        .await
        .unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["attendees"], json!(["a@example.com"]));
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
