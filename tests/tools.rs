use std::sync::Arc;

use outlook_mcp_rs::outlook::fake::{FakeOutlookClient, EMAIL_ID};
use outlook_mcp_rs::server::{
    CheckAvailabilityParams, CreateDraftParams, CreateEventParams, CreateNoteParams, CreateTaskParams,
    DeleteEmailParams, DeleteEventParams, DeleteNoteParams, DeleteTaskParams, GetEmailParams,
    GetEventParams, GetInlineImageParams, GetNoteParams, GetTaskParams, ListAttachmentsParams,
    EmptyDeletedItemsParams,
    ListEmailsParams, ListEventsParams, ListNotesParams, ListTasksParams, OutlookMcpServer,
    RecurrenceParams, ReplyEmailParams, RespondToMeetingParams, SaveAttachmentsParams,
    SendEmailParams, UpdateDraftParams, UpdateEmailParams, UpdateEventParams, UpdateNoteParams, UpdateTaskParams,
};
use rmcp::handler::server::wrapper::Parameters;
use rmcp::model::CallToolResult;
use serde_json::{json, Value};

/// Write `bytes` to a fresh file under the temp dir and return its path.
fn temp_file(name: &str, bytes: &[u8]) -> String {
    use std::sync::atomic::{AtomicU32, Ordering};
    static N: AtomicU32 = AtomicU32::new(0);
    let dir = std::env::temp_dir().join(format!(
        "outlook-mcp-rs-tools-{}-{}", std::process::id(), N.fetch_add(1, Ordering::Relaxed)
    ));
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join(name);
    std::fs::write(&path, bytes).unwrap();
    path.to_string_lossy().into_owned()
}

/// The error text of a failed tool call.
fn err_text(err: rmcp::ErrorData) -> String {
    err.message.to_string()
}

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
        "folder": "sent", "count": 5, "offset": 15, "unread_only": true
    }))
    .unwrap();
    server.list_emails(Parameters(params)).await.unwrap();
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "list_emails");
    assert_eq!(args["folder"], "sent");
    assert_eq!(args["count"], 5);
    assert_eq!(args["offset"], 15);
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
    assert_eq!(args["offset"], 0);
    assert_eq!(args["unread_only"], false);
    assert_eq!(args["to"], json!([]));
    assert_eq!(args["item_type"], json!([]));
    assert_eq!(args["importance"], json!([]));
    assert_eq!(args["flag"], json!([]));
    assert!(args["received_after"].is_null());
}

#[tokio::test]
async fn list_emails_forwards_recipient_filter() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEmailsParams = serde_json::from_value(json!({
        "folder": "sent", "to": "ada@x.com", "from": "me@x.com"
    }))
    .unwrap();
    server.list_emails(Parameters(params)).await.unwrap();
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "list_emails");
    assert_eq!(args["folder"], "sent");
    assert_eq!(args["to"], json!(["ada@x.com"]));
    assert_eq!(args["from"], json!(["me@x.com"]));
}

#[tokio::test]
async fn list_emails_forwards_query_and_filters() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEmailsParams = serde_json::from_value(json!({
        "query": "subject:\"weekly update\" q3*", "from": "ada@x.com", "category": "Work",
        "received_after": "-30d", "received_before": "today", "has_attachments": true,
        "flag": "follow_up", "importance": "HIGH", "item_type": ["email", "meeting"]
    }))
    .unwrap();
    server.list_emails(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["query"], "subject:\"weekly update\" q3*");
    assert_eq!(args["from"], json!(["ada@x.com"]));
    assert_eq!(args["category"], json!(["Work"]));
    assert_eq!(args["received_after"], "-30d");
    assert_eq!(args["received_before"], "today");
    assert_eq!(args["has_attachments"], true);
    assert_eq!(args["flag"], json!(["follow_up"]));
    assert_eq!(args["importance"], json!(["high"]));
    assert_eq!(args["item_type"], json!(["email", "meeting"]));
}

#[tokio::test]
async fn list_emails_string_filters_accept_lists() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEmailsParams = serde_json::from_value(json!({
        "from": ["Person A", " ", "Person B"], "to": ["ada@x.com", "bob@x.com"],
        "category": ["Work", "Red Category"], "importance": ["low", "high", "low"]
    }))
    .unwrap();
    server.list_emails(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["from"], json!(["Person A", "Person B"]));
    assert_eq!(args["to"], json!(["ada@x.com", "bob@x.com"]));
    assert_eq!(args["category"], json!(["Work", "Red Category"]));
    assert_eq!(args["importance"], json!(["low", "high"]));
}

#[tokio::test]
async fn list_emails_maps_deprecated_filters() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEmailsParams = serde_json::from_value(json!({
        "since_days": 14, "flagged": true, "high_importance": true
    }))
    .unwrap();
    server.list_emails(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["received_after"], "-14d");
    assert_eq!(args["flag"], json!(["follow_up", "complete"]));
    assert_eq!(args["importance"], json!(["high"]));
    assert!(args.get("since_days").is_none());
    assert!(args.get("flagged").is_none());
    assert!(args.get("high_importance").is_none());
}

#[tokio::test]
async fn list_emails_deprecated_filters_agree_with_new_ones() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    // since_days 0 was "no filter"; flagged + a flagged state narrows; the
    // same importance twice is not a contradiction.
    let params: ListEmailsParams = serde_json::from_value(json!({
        "since_days": 0, "received_after": "2026-06-01", "flagged": true, "flag": "complete",
        "high_importance": true, "importance": "high"
    }))
    .unwrap();
    server.list_emails(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["received_after"], "2026-06-01");
    assert_eq!(args["flag"], json!(["complete"]));
    assert_eq!(args["importance"], json!(["high"]));
}

#[tokio::test]
async fn list_emails_rejects_contradicting_deprecated_filters() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    for (args, names) in [
        (json!({"since_days": 7, "received_after": "-14d"}), ["received_after", "since_days"]),
        (json!({"flagged": true, "flag": "clear"}), ["flag", "flagged"]),
        (json!({"high_importance": true, "importance": "low"}), ["importance", "high_importance"]),
    ] {
        let params: ListEmailsParams = serde_json::from_value(args).unwrap();
        let err = server.list_emails(Parameters(params)).await.unwrap_err();
        let both = format!("pass either `{}` or `{}`, not both", names[0], names[1]);
        assert!(err.message.contains(&both), "{}", err.message);
    }
    assert!(fake.calls().is_empty());
}

#[tokio::test]
async fn list_emails_rejects_unknown_enum_values() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    for (args, needle) in [
        (json!({"importance": "urgent"}), "Invalid importance \"urgent\""),
        (json!({"flag": ["follow_up", "flagged"]}), "Invalid flag \"flagged\""),
        (json!({"item_type": "invite"}), "Invalid item_type \"invite\""),
    ] {
        let params: ListEmailsParams = serde_json::from_value(args).unwrap();
        let err = server.list_emails(Parameters(params)).await.unwrap_err();
        assert!(err.message.contains(needle), "{}", err.message);
    }
    assert!(fake.calls().is_empty());
}

#[tokio::test]
async fn list_emails_rejects_bad_dates_naming_the_parameter() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEmailsParams =
        serde_json::from_value(json!({"received_before": "next tuesday"})).unwrap();
    let err = server.list_emails(Parameters(params)).await.unwrap_err();
    assert!(err.message.contains("Invalid received_before \"next tuesday\""), "{}", err.message);
    assert!(err.message.contains("'-14d'"), "{}", err.message);
    let params: ListEmailsParams =
        serde_json::from_value(json!({"received_after": "today", "received_before": "yesterday"})).unwrap();
    let err = server.list_emails(Parameters(params)).await.unwrap_err();
    assert!(err.message.contains("later than received_before"), "{}", err.message);
    assert!(fake.calls().is_empty());
}

#[tokio::test]
async fn list_emails_forwards_hebrew_query_unchanged() {
    // Issue #2: the non-ASCII query must reach the client intact (the
    // DASL/fallback handling lives in the COM client).
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEmailsParams = serde_json::from_value(json!({
        "query": "מייל שיקוף Q3"
    }))
    .unwrap();
    server.list_emails(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["query"], "מייל שיקוף Q3");
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
        .get_email(Parameters(serde_json::from_value::<GetEmailParams>(json!({"email_id": EMAIL_ID})).unwrap()))
        .await
        .unwrap();
    assert_eq!(result_json(&result)["body"], "Hi there");
}

#[tokio::test]
async fn get_email_includes_item_type() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_email(Parameters(serde_json::from_value::<GetEmailParams>(json!({"email_id": EMAIL_ID})).unwrap()))
        .await
        .unwrap();
    assert_eq!(result_json(&result)["item_type"], "email");
    assert_eq!(result_json(&result)["is_meeting"], false);
}

// ---- Non-ASCII (Hebrew) round trip, issue #3 ----------------------------

const HE_SUBJECT: &str = "מייל שיקוף";
const HE_SENDER: &str = "עדה לאבלייס";
const HE_BODY: &str = "סיכום עשייה\nשורה שנייה";

/// The raw text of a tool result: the exact JSON string rmcp puts on the wire.
fn result_text(result: &CallToolResult) -> String {
    result.content[0].as_text().expect("expected text content").text.clone()
}

/// Hebrew must appear as literal UTF-8 in the JSON, never as `\u05xx` escapes.
fn assert_raw_utf8(text: &str, expected: &[&str]) {
    for s in expected {
        assert!(text.contains(s), "{s:?} not found verbatim in {text}");
    }
    assert!(!text.contains("\\u05"), "unexpected \\u escape in {text}");
}

#[tokio::test]
async fn list_emails_returns_hebrew_text_verbatim() {
    let fake = Arc::new(FakeOutlookClient::new());
    fake.set_email_text(HE_SUBJECT, HE_SENDER, HE_BODY);
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEmailsParams = serde_json::from_value(json!({})).unwrap();
    let result = server.list_emails(Parameters(params)).await.unwrap();
    let text = result_text(&result);
    assert_raw_utf8(&text, &[HE_SUBJECT, HE_SENDER]);
    let json = result_json(&result);
    assert_eq!(json[0]["subject"].as_str().unwrap().as_bytes(), HE_SUBJECT.as_bytes());
    assert_eq!(json[0]["sender"].as_str().unwrap().as_bytes(), HE_SENDER.as_bytes());
}

#[tokio::test]
async fn get_email_returns_hebrew_text_verbatim() {
    let fake = Arc::new(FakeOutlookClient::new());
    fake.set_email_text(HE_SUBJECT, HE_SENDER, HE_BODY);
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_email(Parameters(serde_json::from_value::<GetEmailParams>(json!({"email_id": EMAIL_ID})).unwrap()))
        .await
        .unwrap();
    // The newline in the body is escaped as `\n` in JSON; check the line.
    assert_raw_utf8(&result_text(&result), &[HE_SUBJECT, HE_SENDER, "סיכום עשייה"]);
    let json = result_json(&result);
    assert_eq!(json["subject"].as_str().unwrap().as_bytes(), HE_SUBJECT.as_bytes());
    assert_eq!(json["sender"].as_str().unwrap().as_bytes(), HE_SENDER.as_bytes());
    assert_eq!(json["body"].as_str().unwrap().as_bytes(), HE_BODY.as_bytes());
}

#[tokio::test]
async fn hebrew_arguments_reach_the_client_unchanged() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEmailsParams = serde_json::from_value(json!({
        "query": HE_SUBJECT, "from": HE_SENDER, "category": "סיכום עשייה"
    }))
    .unwrap();
    server.list_emails(Parameters(params)).await.unwrap();
    server
        .create_draft(Parameters(CreateDraftParams {
            to: vec!["a@example.com".to_string()],
            subject: HE_SUBJECT.to_string(),
            body: Some(HE_BODY.to_string()),
            ..Default::default()
        }))
        .await
        .unwrap();
    let calls = fake.calls();
    assert_eq!(calls[0].1["query"], HE_SUBJECT);
    assert_eq!(calls[0].1["from"], json!([HE_SENDER]));
    assert_eq!(calls[0].1["category"], json!(["סיכום עשייה"]));
    assert_eq!(calls[1].1["subject"], HE_SUBJECT);
    assert_eq!(calls[1].1["body"], HE_BODY);
}

/// End to end over the same newline-delimited JSON-RPC codec that rmcp's
/// stdio transport uses (an in-memory duplex pipe stands in for
/// stdin/stdout): Hebrew in the request reaches the client, and Hebrew in the
/// result leaves the server as raw UTF-8 bytes.
#[tokio::test]
async fn hebrew_round_trips_as_raw_utf8_over_the_stdio_codec() {
    use rmcp::ServiceExt;
    use tokio::io::{AsyncBufReadExt, AsyncWriteExt, BufReader};

    let fake = Arc::new(FakeOutlookClient::new());
    fake.set_email_text(HE_SUBJECT, HE_SENDER, HE_BODY);
    let server = OutlookMcpServer::new(fake.clone());
    let (client_io, server_io) = tokio::io::duplex(64 * 1024);
    tokio::spawn(async move {
        if let Ok(running) = server.serve(server_io).await {
            let _ = running.waiting().await;
        }
    });
    let (read_half, mut write_half) = tokio::io::split(client_io);
    let mut reader = BufReader::new(read_half);

    // Reads raw bytes up to the next response carrying `id`.
    async fn read_response<R: tokio::io::AsyncBufRead + Unpin>(reader: &mut R, id: u64) -> Vec<u8> {
        loop {
            let mut line = Vec::new();
            assert!(reader.read_until(b'\n', &mut line).await.unwrap() > 0, "server closed");
            let msg: Value = serde_json::from_slice(&line).unwrap();
            if msg["id"] == id {
                return line;
            }
        }
    }

    let requests = [
        json!({"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {
            "protocolVersion": "2025-06-18", "capabilities": {},
            "clientInfo": {"name": "utf8-test", "version": "0"}}}),
        json!({"jsonrpc": "2.0", "method": "notifications/initialized"}),
        json!({"jsonrpc": "2.0", "id": 2, "method": "tools/call", "params": {
            "name": "list_emails", "arguments": {"query": HE_SUBJECT}}}),
    ];
    for (i, req) in requests.iter().enumerate() {
        // serde_json writes the Hebrew as raw UTF-8, like a real client would.
        let mut bytes = serde_json::to_vec(req).unwrap();
        bytes.push(b'\n');
        write_half.write_all(&bytes).await.unwrap();
        if i == 0 {
            read_response(&mut reader, 1).await;
        }
    }
    let line = read_response(&mut reader, 2).await;

    // Strict UTF-8 decode of the wire bytes; the tool result is JSON text
    // nested in a JSON string, so its Hebrew is still raw, not `\u` escaped.
    let wire = String::from_utf8(line).expect("server output must be valid UTF-8");
    assert_raw_utf8(&wire, &[HE_SUBJECT, HE_SENDER]);
    let msg: Value = serde_json::from_str(&wire).unwrap();
    let emails: Value =
        serde_json::from_str(msg["result"]["content"][0]["text"].as_str().unwrap()).unwrap();
    assert_eq!(emails[0]["subject"], HE_SUBJECT);
    assert_eq!(emails[0]["sender"], HE_SENDER);
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "list_emails");
    assert_eq!(args["query"], HE_SUBJECT);
}

#[tokio::test]
async fn get_email_max_body_chars_defaults_to_none() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: GetEmailParams = serde_json::from_value(json!({"email_id": EMAIL_ID})).unwrap();
    assert_eq!(params.max_body_chars, None);
    server.get_email(Parameters(params)).await.unwrap();
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "get_email");
    assert_eq!(args["max_body_chars"], Value::Null);
}

#[tokio::test]
async fn get_email_forwards_max_body_chars() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: GetEmailParams = serde_json::from_value(json!({
        "email_id": EMAIL_ID, "prefer_html": true, "max_body_chars": 2_000_000
    }))
    .unwrap();
    server.get_email(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    // The deprecated prefer_html maps onto html_body.
    assert_eq!(args["html_body"], true);
    assert_eq!(args["max_body_chars"], 2_000_000);
}

#[tokio::test]
async fn get_email_reports_truncation_fields() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    // Plain: body flags present, html ones omitted.
    let params: GetEmailParams = serde_json::from_value(json!({"email_id": EMAIL_ID})).unwrap();
    let json = result_json(&server.get_email(Parameters(params)).await.unwrap());
    assert_eq!(json["body_truncated"], false);
    assert_eq!(json["body_length"], 8);
    assert!(json.get("html_truncated").is_none());
    assert!(json.get("html_length").is_none());
    // prefer_html: html flags present too.
    let params: GetEmailParams =
        serde_json::from_value(json!({"email_id": EMAIL_ID, "prefer_html": true})).unwrap();
    let json = result_json(&server.get_email(Parameters(params)).await.unwrap());
    assert_eq!(json["html_truncated"], false);
    assert_eq!(json["html_length"], 15);
}

#[tokio::test]
async fn get_event_and_get_note_report_body_truncated() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: GetEventParams = serde_json::from_value(json!({"event_id": "x"})).unwrap();
    let json = result_json(&server.get_event(Parameters(params)).await.unwrap());
    assert_eq!(json["body_truncated"], false);
    let params: GetNoteParams = serde_json::from_value(json!({"note_id": "x"})).unwrap();
    let json = result_json(&server.get_note(Parameters(params)).await.unwrap());
    assert_eq!(json["body_truncated"], false);
}

#[tokio::test]
async fn send_email_passes_recipients_and_html_flag() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .send_email(Parameters(SendEmailParams {
            to: vec!["a@example.com".to_string(), "b@example.com".to_string()],
            subject: "Hi".to_string(),
            body: Some("Hello!".to_string()),
            html: Some(false),
            ..Default::default()
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
            body: Some("Hello!".to_string()),
            html: Some(false),
            ..Default::default()
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
            body: Some("Thanks!".to_string()),
            reply_all: true,
            send: false,
            ..Default::default()
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
async fn update_draft_forwards_fields_and_lists_changes_in_apply_order() {
    let path = std::env::temp_dir().join("outlook-mcp-rs-tools-update-draft.txt");
    std::fs::write(&path, b"x").unwrap();
    let path_str = path.to_string_lossy().to_string();
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: UpdateDraftParams = serde_json::from_value(json!({
        "draft_id": EMAIL_ID, "attachments": [path_str], "bcc": [],
        "to": ["a@x.com", "b@x.com"], "html_body": "<p>hi</p>", "subject": "New"
    })).unwrap();
    let result = server.update_draft(Parameters(params)).await;
    let _ = std::fs::remove_file(&path);
    let v = result_json(&result.unwrap());
    assert_eq!(v["status"], "draft_updated");
    assert_eq!(v["id"], EMAIL_ID);
    assert_eq!(v["changed"], json!(["subject", "html_body", "to", "bcc", "attachments"]));
    let (name, args) = fake.calls().pop().unwrap();
    assert_eq!(name, "update_draft");
    // `draft_id` (deprecated alias) still deserializes into `email_id`.
    assert_eq!(args["email_id"], EMAIL_ID);
    assert_eq!(args["subject"], "New");
    assert_eq!(args["html_body"], "<p>hi</p>");
    assert_eq!(args["body"], Value::Null);
    assert_eq!(args["to"], json!(["a@x.com", "b@x.com"]));
    assert_eq!(args["cc"], Value::Null);
    assert_eq!(args["bcc"], json!([]));
    assert_eq!(args["attachments"], json!([path_str]));
}

#[tokio::test]
async fn update_draft_rejects_body_and_html_body_together() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: UpdateDraftParams = serde_json::from_value(json!({
        "draft_id": EMAIL_ID, "body": "plain", "html_body": "<p>html</p>"
    })).unwrap();
    assert!(server.update_draft(Parameters(params)).await.is_err());
    assert!(fake.calls().is_empty());
}

#[tokio::test]
async fn update_draft_rejects_an_empty_update() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: UpdateDraftParams = serde_json::from_value(json!({"draft_id": EMAIL_ID})).unwrap();
    assert!(server.update_draft(Parameters(params)).await.is_err());
    assert!(fake.calls().is_empty());
}

#[tokio::test]
async fn send_email_forwards_inline_images() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    // Paths are validated (in the fake too), so use a real file.
    let logo = temp_file("logo.png", b"\x89PNG\r\n\x1a\n");
    let params: SendEmailParams = serde_json::from_value(json!({
        "to": ["a@x.com"], "subject": "Hi", "html": true,
        "body": "<img src=\"cid:logo\">",
        "inline_images": [
            {"content_id": "logo", "path": logo},
            {"content_id": "chart", "data_base64": "aGk=", "filename": "chart.png", "mime_type": "image/png"}
        ]
    })).unwrap();
    server.send_email(Parameters(params)).await.unwrap();
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "send_email");
    assert_eq!(args["html"], true);
    let imgs = &args["inline_images"];
    assert_eq!(imgs[0]["content_id"], "logo");
    assert_eq!(imgs[0]["path"], logo);
    assert_eq!(imgs[0]["data_base64"], Value::Null);
    assert_eq!(imgs[1]["content_id"], "chart");
    assert_eq!(imgs[1]["data_base64"], "aGk=");
    assert_eq!(imgs[1]["filename"], "chart.png");
    assert_eq!(imgs[1]["mime_type"], "image/png");
}

#[tokio::test]
async fn create_draft_forwards_inline_images_and_defaults_to_none() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: CreateDraftParams = serde_json::from_value(json!({
        "to": ["a@x.com"], "subject": "Hi", "body": "<img src=\"cid:logo\">", "html": true,
        "inline_images": [{"content_id": "logo", "data_base64": "aGk="}]
    })).unwrap();
    let result = server.create_draft(Parameters(params)).await.unwrap();
    assert_eq!(result_json(&result)["status"], "draft_saved");
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "create_draft");
    assert_eq!(args["inline_images"][0]["content_id"], "logo");
    assert_eq!(args["inline_images"][0]["data_base64"], "aGk=");

    // Omitted -> None (recorded as null).
    let params: CreateDraftParams = serde_json::from_value(json!({
        "to": ["a@x.com"], "subject": "Hi", "body": "plain"
    })).unwrap();
    server.create_draft(Parameters(params)).await.unwrap();
    assert_eq!(fake.calls()[1].1["inline_images"], Value::Null);
}

#[test]
fn inline_images_appear_in_send_and_draft_schemas() {
    for schema in [schemars::schema_for!(SendEmailParams), schemars::schema_for!(CreateDraftParams)] {
        let v = serde_json::to_value(&schema).unwrap();
        let prop = &v["properties"]["inline_images"];
        assert!(prop.is_object(), "inline_images missing from schema: {v}");
        assert!(prop["description"].as_str().unwrap_or("").contains("cid:CONTENT_ID"), "{prop}");
        let def = &v["$defs"]["InlineImage"];
        for field in ["content_id", "path", "data_base64", "filename", "mime_type"] {
            assert!(def["properties"][field]["description"].is_string(), "InlineImage.{field} lacks a description: {def}");
        }
        assert_eq!(def["required"], json!(["content_id"]));
    }
}

// ---- #28 / #29: body inputs, files, metadata, data: URI images ----

/// Issue #28: a tool call with a large HTML body (76 KB in the report;
/// ~1.5 MB here, with an embedded base64 image) goes through the same
/// newline-delimited JSON-RPC codec rmcp's stdio transport uses, which has
/// no size limit. The data: URI arrives and is turned into a cid: image.
#[tokio::test]
async fn large_html_body_round_trips_over_the_stdio_codec() {
    use rmcp::model::CallToolRequestParams;
    use rmcp::ServiceExt;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let (client_io, server_io) = tokio::io::duplex(64 * 1024);
    tokio::spawn(async move {
        if let Ok(running) = server.serve(server_io).await {
            let _ = running.waiting().await;
        }
    });
    let client = ().serve(client_io).await.expect("handshake");
    let html = format!(
        "<p>{}</p><img src=\"data:image/png;base64,{PNG_B64}\">",
        "lorem ipsum ".repeat(130_000)
    );
    let args = json!({"email_id": EMAIL_ID, "html_body": html});
    let result = client
        .call_tool(CallToolRequestParams::new("update_draft").with_arguments(args.as_object().unwrap().clone()))
        .await
        .expect("a ~1.5 MB update_draft call should succeed");
    assert_ne!(result.is_error, Some(true), "{result:?}");
    let (name, recorded) = fake.calls().pop().unwrap();
    assert_eq!(name, "update_draft");
    let body = recorded["html_body"].as_str().unwrap();
    assert!(body.len() > 1_500_000 && body.contains("cid:img-") && !body.contains("base64"));
    client.cancel().await.ok();
}

/// A 1x1 transparent PNG.
const PNG_B64: &str = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

async fn send(fake: &Arc<FakeOutlookClient>, args: Value) -> Result<CallToolResult, rmcp::ErrorData> {
    let server = OutlookMcpServer::new(fake.clone());
    let params: SendEmailParams = serde_json::from_value(args).unwrap();
    server.send_email(Parameters(params)).await
}

#[tokio::test]
async fn send_email_html_body_is_html() {
    let fake = Arc::new(FakeOutlookClient::new());
    send(&fake, json!({"to": ["a@x.com"], "subject": "Hi", "html_body": "<p>hi</p>"})).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["html"], true);
    assert_eq!(args["body"], "<p>hi</p>");
    // Plain body stays text; the new fields default to null.
    send(&fake, json!({"to": ["a@x.com"], "subject": "Hi", "body": "<p>hi</p>"})).await.unwrap();
    let (_, args) = &fake.calls()[1];
    assert_eq!(args["html"], false);
    assert_eq!(args["categories"], Value::Null);
    assert_eq!(args["importance"], Value::Null);
}

#[tokio::test]
async fn send_email_deprecated_html_flag_still_works_and_contradictions_fail() {
    let fake = Arc::new(FakeOutlookClient::new());
    send(&fake, json!({"to": ["a@x.com"], "subject": "Hi", "body": "<b>x</b>", "html": true})).await.unwrap();
    assert_eq!(fake.calls()[0].1["html"], true);
    let e = err_text(send(&fake, json!({
        "to": ["a@x.com"], "subject": "Hi", "html_body": "<b>x</b>", "html": false
    })).await.unwrap_err());
    assert!(e.contains("`html: false`") && e.contains("`html_body`"), "{e}");
    assert_eq!(fake.calls().len(), 1);
}

#[tokio::test]
async fn send_email_needs_exactly_one_body_source() {
    let fake = Arc::new(FakeOutlookClient::new());
    let e = err_text(send(&fake, json!({"to": ["a@x.com"], "subject": "Hi"})).await.unwrap_err());
    assert!(e.contains("needs a body"), "{e}");
    let e = err_text(send(&fake, json!({
        "to": ["a@x.com"], "subject": "Hi", "body": "a", "html_body": "<p>b</p>"
    })).await.unwrap_err());
    assert_eq!(e, "pass either `body` or `html_body`, not both");
    let e = err_text(send(&fake, json!({
        "to": ["a@x.com"], "subject": "Hi", "html_body": "<p>b</p>", "html_body_file": "x.html"
    })).await.unwrap_err());
    assert_eq!(e, "pass either `html_body` or `html_body_file`, not both");
    assert!(fake.calls().is_empty());
}

#[tokio::test]
async fn compose_tools_read_body_files() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let html = "<p>שלום</p>".repeat(20_000); // ~360 KB, well past the 76 KB from #28
    let html_path = temp_file("big.html", html.as_bytes());
    let txt_path = temp_file("body.txt", b"plain from file");
    let params: CreateDraftParams = serde_json::from_value(json!({
        "to": ["a@x.com"], "subject": "Big", "html_body_file": html_path
    })).unwrap();
    server.create_draft(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["html"], true);
    assert_eq!(args["body"].as_str().unwrap(), html);
    send(&fake, json!({"to": ["a@x.com"], "subject": "T", "body_file": txt_path})).await.unwrap();
    let (_, args) = &fake.calls()[1];
    assert_eq!(args["html"], false);
    assert_eq!(args["body"], "plain from file");
    // Deprecated html=true applies to body_file too.
    send(&fake, json!({"to": ["a@x.com"], "subject": "T", "body_file": txt_path, "html": true})).await.unwrap();
    assert_eq!(fake.calls()[2].1["html"], true);
}

#[tokio::test]
async fn missing_body_file_fails_before_the_client_is_called() {
    let fake = Arc::new(FakeOutlookClient::new());
    let e = err_text(send(&fake, json!({
        "to": ["a@x.com"], "subject": "Hi", "html_body_file": "/definitely/not/here.html"
    })).await.unwrap_err());
    assert!(e.starts_with("html_body_file: could not read"), "{e}");
    assert!(fake.calls().is_empty());
}

#[tokio::test]
async fn data_uri_images_in_html_body_become_inline_attachments() {
    let fake = Arc::new(FakeOutlookClient::new());
    send(&fake, json!({
        "to": ["a@x.com"], "subject": "Hi",
        "html_body": format!("<p>Look:</p><img src=\"data:image/png;base64,{PNG_B64}\"><img src='data:image/png;base64,{PNG_B64}'>"),
    })).await.unwrap();
    let (_, args) = &fake.calls()[0];
    let ids = args["inline_content_ids"].as_array().unwrap();
    assert_eq!(ids.len(), 1, "the same image twice is attached once: {args}");
    let cid = ids[0].as_str().unwrap();
    assert_eq!(
        args["body"],
        format!("<p>Look:</p><img src=\"cid:{cid}\"><img src='cid:{cid}'>")
    );
    assert!(!args["body"].as_str().unwrap().contains("base64"));
}

#[tokio::test]
async fn invalid_data_uri_fails_before_anything_is_created() {
    let fake = Arc::new(FakeOutlookClient::new());
    let e = err_text(send(&fake, json!({
        "to": ["a@x.com"], "subject": "Hi",
        "html_body": "<img src=\"data:image/png;base64,@@@not base64@@@\">",
    })).await.unwrap_err());
    assert!(e.contains("not valid base64"), "{e}");
    assert!(fake.calls().is_empty());
}

#[tokio::test]
async fn inline_images_need_an_html_body() {
    let fake = Arc::new(FakeOutlookClient::new());
    let e = err_text(send(&fake, json!({
        "to": ["a@x.com"], "subject": "Hi", "body": "plain",
        "inline_images": [{"content_id": "a", "data_base64": PNG_B64}],
    })).await.unwrap_err());
    assert!(e.contains("requires an HTML body"), "{e}");
    assert!(fake.calls().is_empty());
}

#[tokio::test]
async fn send_and_draft_forward_categories_and_importance() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    send(&fake, json!({
        "to": ["a@x.com"], "subject": "Hi", "body": "x",
        "categories": ["Red Category"], "importance": "high",
    })).await.unwrap();
    let params: CreateDraftParams = serde_json::from_value(json!({
        "to": ["a@x.com"], "subject": "Hi", "body": "x",
        "categories": ["Blue Category"], "importance": "low",
    })).unwrap();
    server.create_draft(Parameters(params)).await.unwrap();
    let calls = fake.calls();
    assert_eq!(calls[0].1["categories"], json!(["Red Category"]));
    assert_eq!(calls[0].1["importance"], "high");
    assert_eq!(calls[1].1["categories"], json!(["Blue Category"]));
    assert_eq!(calls[1].1["importance"], "low");
    let e = err_text(send(&fake, json!({
        "to": ["a@x.com"], "subject": "Hi", "body": "x", "importance": "urgent",
    })).await.unwrap_err());
    assert!(e.contains("invalid importance"), "{e}");
    assert_eq!(fake.calls().len(), 2);
}

#[tokio::test]
async fn reply_email_takes_html_body_inline_images_and_metadata() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ReplyEmailParams = serde_json::from_value(json!({
        "email_id": EMAIL_ID, "send": false,
        "html_body": format!("<p>See</p><img src=\"cid:logo\"><img src=\"data:image/png;base64,{PNG_B64}\">"),
        "inline_images": [{"content_id": "logo", "data_base64": PNG_B64}],
        "categories": ["Work"], "importance": "high",
    })).unwrap();
    server.reply_email(Parameters(params)).await.unwrap();
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "reply_email");
    assert_eq!(args["html"], true);
    let ids = args["inline_content_ids"].as_array().unwrap();
    assert_eq!(ids.len(), 2);
    assert_eq!(ids[0], "logo");
    assert!(args["body"].as_str().unwrap().contains(&format!("cid:{}", ids[1].as_str().unwrap())));
    assert_eq!(args["categories"], json!(["Work"]));
    assert_eq!(args["importance"], "high");
    // body_file works for replies too.
    let path = temp_file("reply.txt", b"thanks from a file");
    let params: ReplyEmailParams = serde_json::from_value(json!({"email_id": EMAIL_ID, "body_file": path})).unwrap();
    server.reply_email(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[1];
    assert_eq!(args["body"], "thanks from a file");
    assert_eq!(args["send"], true);
}

async fn update_draft(fake: &Arc<FakeOutlookClient>, args: Value) -> Result<CallToolResult, rmcp::ErrorData> {
    let server = OutlookMcpServer::new(fake.clone());
    let params: UpdateDraftParams = serde_json::from_value(args).unwrap();
    server.update_draft(Parameters(params)).await
}

#[tokio::test]
async fn update_draft_takes_email_id_and_new_fields() {
    let fake = Arc::new(FakeOutlookClient::new());
    let html_path = temp_file("draft.html", format!("<p>x</p><img src=\"data:image/png;base64,{PNG_B64}\">").as_bytes());
    let v = result_json(&update_draft(&fake, json!({
        "email_id": EMAIL_ID, "html_body_file": html_path,
        "inline_images": [{"content_id": "logo", "data_base64": PNG_B64}],
        "add_categories": ["Red"], "remove_categories": ["Blue"], "importance": "high",
    })).await.unwrap());
    assert_eq!(v["status"], "draft_updated");
    assert_eq!(v["changed"], json!(["html_body", "inline_images", "add_categories", "remove_categories", "importance"]));
    let (_, args) = fake.calls().pop().unwrap();
    assert_eq!(args["email_id"], EMAIL_ID);
    assert_eq!(args["send"], false);
    let ids = args["inline_content_ids"].as_array().unwrap();
    assert_eq!(ids.len(), 2);
    assert_eq!(args["html_body"], format!("<p>x</p><img src=\"cid:{}\">", ids[1].as_str().unwrap()));
    assert_eq!(args["add_categories"], json!(["Red"]));
    assert_eq!(args["remove_categories"], json!(["Blue"]));
}

#[tokio::test]
async fn update_draft_send_alone_sends_it() {
    let fake = Arc::new(FakeOutlookClient::new());
    let v = result_json(&update_draft(&fake, json!({"draft_id": EMAIL_ID, "send": true})).await.unwrap());
    assert_eq!(v["status"], "sent");
    assert_eq!(v["changed"], json!([]));
    let (_, args) = fake.calls().pop().unwrap();
    assert_eq!(args["send"], true);
    assert_eq!(args["email_id"], EMAIL_ID);
}

#[tokio::test]
async fn update_draft_rejects_bad_inputs_without_touching_the_draft() {
    let fake = Arc::new(FakeOutlookClient::new());
    let e = err_text(update_draft(&fake, json!({
        "email_id": EMAIL_ID, "body": "plain",
        "inline_images": [{"content_id": "a", "data_base64": PNG_B64}],
    })).await.unwrap_err());
    assert!(e.contains("html_body"), "{e}");
    let e = err_text(update_draft(&fake, json!({"email_id": EMAIL_ID, "body": "a", "body_file": "b.txt"})).await.unwrap_err());
    assert_eq!(e, "pass either `body` or `body_file`, not both");
    let e = err_text(update_draft(&fake, json!({"email_id": EMAIL_ID, "importance": "urgent"})).await.unwrap_err());
    assert!(e.contains("invalid importance"), "{e}");
    let e = err_text(update_draft(&fake, json!({"email_id": EMAIL_ID, "html_body": "<img src=\"data:image/png;base64,!!\">"})).await.unwrap_err());
    assert!(e.contains("not valid base64"), "{e}");
    assert!(fake.calls().is_empty());
}

#[test]
fn write_tool_schemas_reflect_the_new_body_convention() {
    // Compared sorted: the order schemars lists required fields in is not part of the contract.
    let required = |schema: schemars::Schema| {
        let mut names: Vec<String> = serde_json::from_value(serde_json::to_value(&schema).unwrap()["required"].clone()).unwrap();
        names.sort();
        json!(names)
    };
    // A body is no longer a required field (any one of four sources works).
    assert_eq!(required(schemars::schema_for!(SendEmailParams)), json!(["subject", "to"]));
    assert_eq!(required(schemars::schema_for!(CreateDraftParams)), json!(["subject", "to"]));
    assert_eq!(required(schemars::schema_for!(ReplyEmailParams)), json!(["email_id"]));
    assert_eq!(required(schemars::schema_for!(UpdateDraftParams)), json!(["email_id"]));
    for schema in [
        schemars::schema_for!(SendEmailParams), schemars::schema_for!(CreateDraftParams),
        schemars::schema_for!(ReplyEmailParams), schemars::schema_for!(UpdateDraftParams),
    ] {
        let v = serde_json::to_value(&schema).unwrap();
        for field in ["body", "html_body", "body_file", "html_body_file", "inline_images", "importance"] {
            assert!(v["properties"][field].is_object(), "{field} missing: {v}");
        }
    }
    let v = serde_json::to_value(schemars::schema_for!(UpdateDraftParams)).unwrap();
    for field in ["send", "add_categories", "remove_categories"] {
        assert!(v["properties"][field].is_object(), "{field} missing: {v}");
    }
    for schema in [
        schemars::schema_for!(CreateEventParams), schemars::schema_for!(UpdateEventParams),
        schemars::schema_for!(CreateTaskParams), schemars::schema_for!(UpdateTaskParams),
        schemars::schema_for!(CreateNoteParams), schemars::schema_for!(UpdateNoteParams),
    ] {
        let v = serde_json::to_value(&schema).unwrap();
        assert!(v["properties"]["body_file"].is_object(), "body_file missing: {v}");
    }
}

#[tokio::test]
async fn event_task_and_note_tools_read_body_file() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let path = temp_file("desc.txt", "תיאור from a file".as_bytes());
    let p: CreateEventParams = serde_json::from_value(json!({
        "subject": "S", "start": "2026-06-12T14:00", "end": "2026-06-12T15:00", "body_file": path
    })).unwrap();
    server.create_event(Parameters(p)).await.unwrap();
    let p: UpdateEventParams = serde_json::from_value(json!({"event_id": "e", "body_file": path})).unwrap();
    server.update_event(Parameters(p)).await.unwrap();
    let p: CreateTaskParams = serde_json::from_value(json!({"subject": "S", "body_file": path})).unwrap();
    server.create_task(Parameters(p)).await.unwrap();
    let p: UpdateTaskParams = serde_json::from_value(json!({"task_id": "t", "body_file": path})).unwrap();
    server.update_task(Parameters(p)).await.unwrap();
    let p: CreateNoteParams = serde_json::from_value(json!({"body_file": path})).unwrap();
    server.create_note(Parameters(p)).await.unwrap();
    let p: UpdateNoteParams = serde_json::from_value(json!({"note_id": "n", "body_file": path})).unwrap();
    server.update_note(Parameters(p)).await.unwrap();
    let calls = fake.calls();
    assert_eq!(calls.len(), 6);
    for (name, args) in &calls {
        assert_eq!(args["body"], "תיאור from a file", "{name}");
    }
}

#[tokio::test]
async fn body_and_body_file_are_exclusive_and_create_note_needs_one() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let p: CreateTaskParams = serde_json::from_value(json!({"subject": "S", "body": "a", "body_file": "b"})).unwrap();
    let e = err_text(server.create_task(Parameters(p)).await.unwrap_err());
    assert_eq!(e, "pass either `body` or `body_file`, not both");
    let p: UpdateNoteParams = serde_json::from_value(json!({"note_id": "n", "body": "a", "body_file": "b"})).unwrap();
    assert!(server.update_note(Parameters(p)).await.is_err());
    let p: CreateEventParams = serde_json::from_value(json!({
        "subject": "S", "start": "2026-06-12T14:00", "end": "2026-06-12T15:00", "body": "a", "body_file": "b"
    })).unwrap();
    assert!(server.create_event(Parameters(p)).await.is_err());
    let p: CreateNoteParams = serde_json::from_value(json!({})).unwrap();
    let e = err_text(server.create_note(Parameters(p)).await.unwrap_err());
    assert!(e.contains("create_note needs a body"), "{e}");
    assert!(fake.calls().is_empty());
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
async fn delete_email_defaults_to_soft_delete() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    // permanent omitted → serde default false.
    let params: DeleteEmailParams =
        serde_json::from_value(json!({"email_id": EMAIL_ID})).unwrap();
    let result = server.delete_email(Parameters(params)).await.unwrap();
    assert_eq!(result_json(&result)["permanent"], false);
    assert_eq!(
        fake.calls(),
        vec![("delete_email".to_string(), json!({"email_id": EMAIL_ID, "permanent": false}))]
    );
}

#[tokio::test]
async fn delete_email_forwards_permanent() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: DeleteEmailParams =
        serde_json::from_value(json!({"email_id": EMAIL_ID, "permanent": true})).unwrap();
    let result = server.delete_email(Parameters(params)).await.unwrap();
    assert_eq!(result_json(&result)["permanent"], true);
    assert_eq!(
        fake.calls(),
        vec![("delete_email".to_string(), json!({"email_id": EMAIL_ID, "permanent": true}))]
    );
}

#[tokio::test]
async fn empty_deleted_items_refuses_without_confirm() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    // confirm omitted → serde default false → the refusal surfaces as an error.
    let params: EmptyDeletedItemsParams = serde_json::from_value(json!({})).unwrap();
    let err = server.empty_deleted_items(Parameters(params)).await.unwrap_err();
    assert!(err.message.contains("confirm=true"));
    assert_eq!(
        fake.calls(),
        vec![("empty_deleted_items".to_string(), json!({"confirm": false}))]
    );
}

#[tokio::test]
async fn empty_deleted_items_with_confirm_returns_counts() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: EmptyDeletedItemsParams =
        serde_json::from_value(json!({"confirm": true})).unwrap();
    let result = server.empty_deleted_items(Parameters(params)).await.unwrap();
    let v = result_json(&result);
    assert_eq!(v["status"], "emptied");
    assert!(v["items_deleted"].is_number());
    assert!(v["folders_deleted"].is_number());
    assert!(v["failed"].is_number());
    assert_eq!(
        fake.calls(),
        vec![("empty_deleted_items".to_string(), json!({"confirm": true}))]
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
    let params: ListEventsParams = serde_json::from_value(json!({
        "start_after": "2026-06-10", "start_before": "2026-06-17"
    }))
    .unwrap();
    server.list_events(Parameters(params)).await.unwrap();
    let (name, args) = fake.calls().pop().unwrap();
    assert_eq!(name, "list_events");
    assert_eq!(args["start_after"], "2026-06-10");
    assert_eq!(args["start_before"], "2026-06-17");
}

#[tokio::test]
async fn list_events_accepts_deprecated_start_date_end_date() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEventsParams = serde_json::from_value(json!({
        "start_date": "start_of_week", "end_date": "end_of_week"
    }))
    .unwrap();
    server.list_events(Parameters(params)).await.unwrap();
    let (_, args) = fake.calls().pop().unwrap();
    assert_eq!(args["start_after"], "start_of_week");
    assert_eq!(args["start_before"], "end_of_week");
    // Old and new name for the same bound together is a duplicate field.
    assert!(serde_json::from_value::<ListEventsParams>(json!({
        "start_date": "today", "start_after": "today"
    }))
    .is_err());
}

#[tokio::test]
async fn list_events_rejects_bad_dates() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEventsParams = serde_json::from_value(json!({"start_after": "2026-13-01"})).unwrap();
    let err = server.list_events(Parameters(params)).await.unwrap_err();
    assert!(err.message.contains("Invalid start_after \"2026-13-01\""), "{}", err.message);
}

#[tokio::test]
async fn list_events_pages_with_count_and_offset() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEventsParams = serde_json::from_value(json!({})).unwrap();
    server.list_events(Parameters(params)).await.unwrap();
    let params: ListEventsParams = serde_json::from_value(json!({"count": 20, "offset": 40})).unwrap();
    server.list_events(Parameters(params)).await.unwrap();
    let calls = fake.calls();
    assert_eq!(calls[0].1["count"], 250);
    assert_eq!(calls[0].1["offset"], 0);
    assert_eq!(calls[1].1["count"], 20);
    assert_eq!(calls[1].1["offset"], 40);
}

#[tokio::test]
async fn list_events_forwards_all_filters() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEventsParams = serde_json::from_value(json!({
        "query": "review", "category": "Work", "show_as": "Busy", "my_response": "accepted",
        "attendees": ["alice@example.com"], "attendee_role": "required", "meetings_only": true,
        "all_day": false, "calendar_of": "bob@example.com"
    }))
    .unwrap();
    server.list_events(Parameters(params)).await.unwrap();
    let (name, args) = fake.calls().pop().unwrap();
    assert_eq!(name, "list_events");
    assert_eq!(args["query"], "review");
    assert_eq!(args["category"], json!(["Work"]));
    assert_eq!(args["show_as"], json!(["busy"]));
    assert_eq!(args["my_response"], json!(["accepted"]));
    assert_eq!(args["attendees"], serde_json::json!(["alice@example.com"]));
    assert_eq!(args["attendee_role"], "required");
    assert_eq!(args["meetings_only"], true);
    assert_eq!(args["all_day"], false);
    assert_eq!(args["calendar_of"], "bob@example.com");
}

#[tokio::test]
async fn list_events_filters_accept_lists_and_single_values() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEventsParams = serde_json::from_value(json!({
        "category": ["Work", "Home"], "show_as": ["busy", "out_of_office"],
        "my_response": ["accepted", "tentative"], "attendees": "alice@example.com"
    }))
    .unwrap();
    server.list_events(Parameters(params)).await.unwrap();
    let (_, args) = fake.calls().pop().unwrap();
    assert_eq!(args["category"], json!(["Work", "Home"]));
    assert_eq!(args["show_as"], json!(["busy", "out_of_office"]));
    assert_eq!(args["my_response"], json!(["accepted", "tentative"]));
    assert_eq!(args["attendees"], json!(["alice@example.com"]));
    let params: ListEventsParams = serde_json::from_value(json!({"show_as": "away"})).unwrap();
    let err = server.list_events(Parameters(params)).await.unwrap_err();
    assert!(err.message.contains("Invalid show_as \"away\""), "{}", err.message);
}

#[tokio::test]
async fn get_event_returns_subject_and_friendly_fields() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_event(Parameters(serde_json::from_value::<GetEventParams>(json!({"event_id": EVENT_ID})).unwrap()))
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
            body: None, body_file: None,
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
        body: None, body_file: None, location: None, attendees: None,
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
            body: None, body_file: None, location: None, attendees: None,
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
async fn check_availability_forwards_params() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .check_availability(Parameters(CheckAvailabilityParams {
            people: vec!["alice@example.com".to_string(), "bob@example.com".to_string()],
            start: "2099-01-01T09:00".to_string(),
            end: "2099-01-01T17:00".to_string(),
            interval_minutes: 30,
            treat_as_free: vec!["free".to_string()],
        }))
        .await
        .unwrap();
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "check_availability");
    assert_eq!(args["people"], json!(["alice@example.com", "bob@example.com"]));
    assert_eq!(args["start"], "2099-01-01T09:00");
    assert_eq!(args["end"], "2099-01-01T17:00");
    assert_eq!(args["interval_minutes"], 30);
    assert_eq!(args["treat_as_free"], json!(["free"]));
}

#[tokio::test]
async fn check_availability_returns_people_and_common_free() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .check_availability(Parameters(CheckAvailabilityParams {
            people: vec!["alice@example.com".to_string()],
            start: "2099-01-01T09:00".to_string(),
            end: "2099-01-01T09:30".to_string(),
            interval_minutes: 30,
            treat_as_free: vec!["free".to_string()],
        }))
        .await
        .unwrap();
    let v = result_json(&result);
    assert_eq!(v["people"][0]["person"], "alice@example.com");
    assert_eq!(v["people"][0]["resolved"], true);
    assert_eq!(v["common_free"][0]["start"], "2099-01-01T09:00");
}

#[tokio::test]
async fn get_event_recurrence_is_none_by_default() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_event(Parameters(serde_json::from_value::<GetEventParams>(json!({"event_id": EVENT_ID})).unwrap()))
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
            start: None, end: None, location: None, body: None, body_file: None, all_day: None,
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
            subject: None, start: None, end: None, location: None, body: None, body_file: None,
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
            subject: None, start: None, end: None, location: None, body: None, body_file: None,
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
            subject: None, start: None, end: None, location: None, body: None, body_file: None,
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
async fn update_event_rejects_recurrence_and_clear_recurrence_together() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let err = server
        .update_event(Parameters(UpdateEventParams {
            event_id: EVENT_ID.to_string(),
            subject: None, start: None, end: None, location: None, body: None, body_file: None,
            all_day: None, reminder_minutes: None, show_as: None,
            add_categories: None, remove_categories: None,
            add_required_attendees: None, add_optional_attendees: None, remove_attendees: None,
            send_update: false,
            recurrence: Some(RecurrenceParams {
                pattern: "daily".to_string(), interval: None, days_of_week: None,
                day_of_month: None, until: None, occurrences: None,
            }),
            clear_recurrence: true,
        }))
        .await
        .unwrap_err();
    assert!(err.message.contains("cannot set recurrence and clear_recurrence"));
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
        .list_attachments(Parameters(serde_json::from_value::<ListAttachmentsParams>(json!({"email_id": EMAIL_ID})).unwrap()))
        .await
        .unwrap();
    let v = result_json(&result);
    assert_eq!(v[0]["filename"], "report.pdf");
    assert_eq!(v[0]["index"], 1);
    assert_eq!(v[0]["size"], 1234);
    assert_eq!(v[0]["type"], "file");
    assert!(v[0]["content_id"].is_null());
    assert_eq!(v[0]["mime_type"], "application/pdf");
    assert_eq!(v[0]["hidden"], false);
    assert_eq!(v[0]["is_inline"], false);
    assert_eq!(v[1]["content_id"], "logo@example");
    assert_eq!(v[1]["mime_type"], "image/png");
    assert_eq!(v[1]["hidden"], true);
    assert_eq!(v[1]["is_inline"], true);
}

#[tokio::test]
async fn save_attachments_passes_dir_and_names() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .save_attachments(Parameters(SaveAttachmentsParams {
            email_id: EMAIL_ID.to_string(),
            save_dir: "/tmp/x".to_string(),
            attachment_names: Some(vec!["report.pdf".to_string()]),
            inline: None,
        }))
        .await
        .unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["save_dir"], "/tmp/x");
    assert_eq!(args["attachment_names"], json!(["report.pdf"]));
    assert_eq!(args["inline"], Value::Null);
    let v = result_json(&result);
    assert_eq!(v[0]["filename"], "report.pdf");
    assert_eq!(v[0]["status"], "saved");
    assert_eq!(v[0]["saved_to"], "/tmp/x");
    for key in ["index", "size", "type", "content_id", "mime_type", "hidden", "is_inline"] {
        assert!(v[0].get(key).is_some(), "missing {key}");
    }
}

async fn save_attachments_with_inline(inline: Option<bool>) -> (Value, Value) {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .save_attachments(Parameters(SaveAttachmentsParams {
            email_id: EMAIL_ID.to_string(),
            save_dir: "/tmp/x".to_string(),
            attachment_names: None,
            inline,
        }))
        .await
        .unwrap();
    let (_, args) = &fake.calls()[0];
    (args["inline"].clone(), result_json(&result))
}

#[tokio::test]
async fn save_attachments_inline_false_skips_inline_images() {
    let (arg, v) = save_attachments_with_inline(Some(false)).await;
    assert_eq!(arg, false);
    let names: Vec<_> = v.as_array().unwrap().iter().map(|a| a["filename"].clone()).collect();
    assert_eq!(names, vec![json!("report.pdf")]);
}

#[tokio::test]
async fn save_attachments_inline_true_saves_only_inline_images() {
    let (arg, v) = save_attachments_with_inline(Some(true)).await;
    assert_eq!(arg, true);
    let names: Vec<_> = v.as_array().unwrap().iter().map(|a| a["filename"].clone()).collect();
    assert_eq!(names, vec![json!("logo.png")]);
    assert_eq!(v[0]["is_inline"], true);
}

#[tokio::test]
async fn save_attachments_without_inline_saves_all() {
    let (arg, v) = save_attachments_with_inline(None).await;
    assert_eq!(arg, Value::Null);
    assert_eq!(v.as_array().unwrap().len(), 2);
}

#[tokio::test]
async fn get_inline_image_forwards_args_and_returns_data_uri() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_inline_image(Parameters(serde_json::from_value::<GetInlineImageParams>(json!({"email_id": EMAIL_ID, "content_id": "cid:logo@example"})).unwrap()))
        .await
        .unwrap();
    assert_eq!(
        fake.calls(),
        vec![(
            "get_inline_image".to_string(),
            json!({"email_id": EMAIL_ID, "content_ids": ["cid:logo@example"], "context_lines": null,
                "output_dir": null})
        )]
    );
    let v = result_json(&result);
    assert_eq!(v["content_id"], "logo@example");
    assert_eq!(v["filename"], "logo.png");
    assert_eq!(v["mime_type"], "image/png");
    assert_eq!(v["size"], 4);
    assert_eq!(v["data_uri"], "data:image/png;base64,iVBORw==");
    // Not requested: no `context` key at all.
    assert!(v.get("context").is_none(), "{v}");
}

#[tokio::test]
async fn get_inline_image_forwards_context_lines_and_returns_context() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_inline_image(Parameters(serde_json::from_value::<GetInlineImageParams>(json!({"email_id": EMAIL_ID, "content_id": "logo@example", "context_lines": 3})).unwrap()))
        .await
        .unwrap();
    assert_eq!(
        fake.calls(),
        vec![(
            "get_inline_image".to_string(),
            json!({"email_id": EMAIL_ID, "content_ids": ["logo@example"], "context_lines": 3,
                "output_dir": null})
        )]
    );
    let v = result_json(&result);
    assert_eq!(v["context"], "Here is our new logo:");
    assert_eq!(v["data_uri"], "data:image/png;base64,iVBORw==");
}

#[test]
fn get_inline_image_params_context_lines_defaults_to_none() {
    let p: GetInlineImageParams =
        serde_json::from_value(json!({"email_id": EMAIL_ID, "content_id": "a@b"})).unwrap();
    assert_eq!(p.context_lines, None);
    let p: GetInlineImageParams =
        serde_json::from_value(json!({"email_id": EMAIL_ID, "content_id": "a@b", "context_lines": 5})).unwrap();
    assert_eq!(p.context_lines, Some(5));
    let negative = serde_json::from_value::<GetInlineImageParams>(
        json!({"email_id": EMAIL_ID, "content_id": "a@b", "context_lines": -1}));
    assert!(negative.is_err());
}

#[tokio::test]
async fn get_inline_image_requires_exactly_one_of_content_id_and_content_ids() {
    let missing = serde_json::from_value::<GetInlineImageParams>(json!({"content_id": "a@b"}));
    assert!(missing.is_err());
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    for (args, needle) in [
        (json!({"email_id": EMAIL_ID}), "content_id"),
        (json!({"email_id": EMAIL_ID, "content_id": "a@b", "content_ids": ["c@d"]}),
            "pass either `content_id` or `content_ids`, not both"),
        (json!({"email_id": EMAIL_ID, "content_ids": []}), "content_ids must not be empty"),
    ] {
        let params: GetInlineImageParams = serde_json::from_value(args).unwrap();
        let err = server.get_inline_image(Parameters(params)).await.unwrap_err();
        assert!(err.message.contains(needle), "{}", err.message);
    }
    assert!(fake.calls().is_empty());
}

#[tokio::test]
async fn get_inline_image_surfaces_client_errors() {
    let fake = Arc::new(FakeOutlookClient::new());
    fake.set_fail_with("Content-ID 'x@y' not found.");
    let server = OutlookMcpServer::new(fake.clone());
    let err = server
        .get_inline_image(Parameters(serde_json::from_value::<GetInlineImageParams>(json!({"email_id": EMAIL_ID, "content_id": "x@y"})).unwrap()))
        .await
        .unwrap_err();
    assert!(err.message.contains("Content-ID 'x@y' not found."));
}

// ---- Tasks ----

#[tokio::test]
async fn list_tasks_passes_include_completed() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListTasksParams = serde_json::from_value(json!({"include_completed": true})).unwrap();
    server.list_tasks(Parameters(params)).await.unwrap();
    assert_eq!(fake.calls(), vec![
        ("list_tasks".to_string(), json!({
            "include_completed": true, "category": [], "importance": [], "query": null,
            "due_after": null, "due_before": null, "count": 500, "offset": 0,
        })),
    ]);
}

#[tokio::test]
async fn list_tasks_forwards_filters() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListTasksParams = serde_json::from_value(json!({
        "include_completed": true, "category": "Red Category", "importance": ["High", "normal"],
        "query": "body:milk", "due_after": "today", "due_before": "end_of_week",
        "count": 25, "offset": 25
    }))
    .unwrap();
    server.list_tasks(Parameters(params)).await.unwrap();
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "list_tasks");
    assert_eq!(args["include_completed"], true);
    assert_eq!(args["category"], json!(["Red Category"]));
    assert_eq!(args["importance"], json!(["high", "normal"]));
    assert_eq!(args["query"], "body:milk");
    assert_eq!(args["due_after"], "today");
    assert_eq!(args["due_before"], "end_of_week");
    assert_eq!(args["count"], 25);
    assert_eq!(args["offset"], 25);
}

#[tokio::test]
async fn list_tasks_rejects_bad_importance_and_dates() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListTasksParams = serde_json::from_value(json!({"importance": "urgent"})).unwrap();
    let err = server.list_tasks(Parameters(params)).await.unwrap_err();
    assert!(err.message.contains("Invalid importance \"urgent\""), "{}", err.message);
    let params: ListTasksParams = serde_json::from_value(json!({"due_before": "soon"})).unwrap();
    let err = server.list_tasks(Parameters(params)).await.unwrap_err();
    assert!(err.message.contains("Invalid due_before \"soon\""), "{}", err.message);
    assert!(fake.calls().is_empty());
}

#[tokio::test]
async fn list_tasks_defaults_all_filters_to_none() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListTasksParams = serde_json::from_value(json!({})).unwrap();
    server.list_tasks(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["include_completed"], false);
    assert_eq!(args["category"], json!([]));
    assert_eq!(args["importance"], json!([]));
    assert!(args["query"].is_null());
    assert_eq!(args["count"], 500);
    assert_eq!(args["offset"], 0);
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
async fn create_task_forwards_categories_start_date_and_reminder() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: CreateTaskParams = serde_json::from_value(json!({
        "subject": "Ship it",
        "categories": ["Blue Category"],
        "start_date": "2099-01-01",
        "reminder_time": "2099-01-01T09:00"
    })).unwrap();
    server.create_task(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["categories"], json!(["Blue Category"]));
    assert_eq!(args["start_date"], "2099-01-01");
    assert_eq!(args["reminder_time"], "2099-01-01T09:00");
}

#[tokio::test]
async fn update_task_marks_complete() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    use outlook_mcp_rs::outlook::fake::TASK_ID;
    server
        .update_task(Parameters(UpdateTaskParams {
            task_id: TASK_ID.to_string(), mark_complete: Some(true),
            subject: None, body: None, body_file: None, due_date: None, start_date: None,
            importance: None, add_categories: None, remove_categories: None,
            percent_complete: None, reminder_time: None,
        }))
        .await
        .unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["mark_complete"], true);
}

#[tokio::test]
async fn update_task_reopens_with_mark_complete_false() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    use outlook_mcp_rs::outlook::fake::TASK_ID;
    let result = server
        .update_task(Parameters(UpdateTaskParams {
            task_id: TASK_ID.to_string(), mark_complete: Some(false),
            subject: None, body: None, body_file: None, due_date: None, start_date: None,
            importance: None, add_categories: None, remove_categories: None,
            percent_complete: None, reminder_time: None,
        }))
        .await
        .unwrap();
    let json = result_json(&result);
    assert!(json["changed"].as_array().unwrap().iter().any(|v| v == "mark_complete"));
}

#[tokio::test]
async fn update_task_forwards_field_edits() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    use outlook_mcp_rs::outlook::fake::TASK_ID;
    server
        .update_task(Parameters(UpdateTaskParams {
            task_id: TASK_ID.to_string(), mark_complete: None,
            subject: Some("Renamed".to_string()), body: None, body_file: None,
            due_date: None, start_date: None, importance: Some("high".to_string()),
            add_categories: Some(vec!["Red Category".to_string()]), remove_categories: None,
            percent_complete: Some(50), reminder_time: None,
        }))
        .await
        .unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["subject"], "Renamed");
    assert_eq!(args["importance"], "high");
    assert_eq!(args["add_categories"], json!(["Red Category"]));
    assert_eq!(args["percent_complete"], 50);
}

#[tokio::test]
async fn delete_task_records_call() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    use outlook_mcp_rs::outlook::fake::TASK_ID;
    let result = server
        .delete_task(Parameters(DeleteTaskParams { task_id: TASK_ID.to_string() }))
        .await
        .unwrap();
    let json = result_json(&result);
    assert_eq!(json["status"], "deleted");
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "delete_task");
    assert_eq!(args["task_id"], TASK_ID);
}

// ---- Notes ----

#[tokio::test]
async fn list_notes_records_call() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListNotesParams = serde_json::from_value(json!({})).unwrap();
    server.list_notes(Parameters(params)).await.unwrap();
    assert_eq!(fake.calls(), vec![("list_notes".to_string(), json!({
        "category": [], "query": null, "created_after": null, "created_before": null,
        "count": 500, "offset": 0,
    }))]);
}

#[tokio::test]
async fn list_notes_forwards_filters() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListNotesParams = serde_json::from_value(json!({
        "category": ["Green Category", "Blue Category"], "query": "\"renew passport\"",
        "created_after": "2026-01-01", "created_before": "2026-06-30", "count": 5, "offset": 10
    }))
    .unwrap();
    server.list_notes(Parameters(params)).await.unwrap();
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "list_notes");
    assert_eq!(args["category"], json!(["Green Category", "Blue Category"]));
    assert_eq!(args["query"], "\"renew passport\"");
    assert_eq!(args["created_after"], "2026-01-01");
    assert_eq!(args["created_before"], "2026-06-30");
    assert_eq!(args["count"], 5);
    assert_eq!(args["offset"], 10);
}

#[tokio::test]
async fn list_notes_rejects_bad_dates() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListNotesParams = serde_json::from_value(json!({"created_after": "-3x"})).unwrap();
    let err = server.list_notes(Parameters(params)).await.unwrap_err();
    assert!(err.message.contains("Invalid created_after \"-3x\""), "{}", err.message);
}

#[tokio::test]
async fn list_notes_defaults_filters_to_none() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListNotesParams = serde_json::from_value(json!({})).unwrap();
    server.list_notes(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["category"], json!([]));
    assert!(args["query"].is_null());
}

#[tokio::test]
async fn get_note_returns_body() {
    use outlook_mcp_rs::outlook::fake::NOTE_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_note(Parameters(serde_json::from_value::<GetNoteParams>(json!({"note_id": NOTE_ID})).unwrap()))
        .await
        .unwrap();
    assert!(result_json(&result)["body"].as_str().unwrap().starts_with("Ideas"));
}

#[tokio::test]
async fn create_note_records_body() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .create_note(Parameters(CreateNoteParams {
            body: Some("Ideas\n- one".to_string()), ..Default::default()
        }))
        .await
        .unwrap();
    assert_eq!(fake.calls(), vec![
        ("create_note".to_string(), json!({"body": "Ideas\n- one", "categories": null, "color": null})),
    ]);
}

#[tokio::test]
async fn create_note_forwards_categories_and_color() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: CreateNoteParams = serde_json::from_value(json!({
        "body": "Remember to renew the domain",
        "categories": ["Yellow Category"],
        "color": "yellow"
    })).unwrap();
    server.create_note(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["categories"], json!(["Yellow Category"]));
    assert_eq!(args["color"], "yellow");
}

#[tokio::test]
async fn get_note_includes_modified() {
    use outlook_mcp_rs::outlook::fake::NOTE_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_note(Parameters(serde_json::from_value::<GetNoteParams>(json!({"note_id": NOTE_ID})).unwrap()))
        .await
        .unwrap();
    let v = result_json(&result);
    // The fake may return null for a note that was never "modified" —
    // assert the key exists in the JSON shape, not a specific non-null value.
    assert!(v.as_object().unwrap().contains_key("modified"));
}

#[tokio::test]
async fn update_note_forwards_body_and_color() {
    use outlook_mcp_rs::outlook::fake::NOTE_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .update_note(Parameters(UpdateNoteParams {
            note_id: NOTE_ID.to_string(),
            body: Some("Updated body".to_string()), body_file: None,
            add_categories: None, remove_categories: None,
            color: Some("pink".to_string()),
        }))
        .await
        .unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["body"], "Updated body");
    assert_eq!(args["color"], "pink");
}

#[tokio::test]
async fn update_note_manages_categories() {
    use outlook_mcp_rs::outlook::fake::NOTE_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .update_note(Parameters(UpdateNoteParams {
            note_id: NOTE_ID.to_string(),
            body: None, body_file: None,
            add_categories: Some(vec!["Blue Category".to_string()]),
            remove_categories: None,
            color: None,
        }))
        .await
        .unwrap();
    let v = result_json(&result);
    assert!(v["changed"].as_array().unwrap().iter().any(|c| c == "add_categories"));
}

#[tokio::test]
async fn delete_note_records_call() {
    use outlook_mcp_rs::outlook::fake::NOTE_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .delete_note(Parameters(DeleteNoteParams { note_id: NOTE_ID.to_string() }))
        .await
        .unwrap();
    let json = result_json(&result);
    assert_eq!(json["status"], "deleted");
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "delete_note");
    assert_eq!(args["note_id"], NOTE_ID);
}

// ---- Read tools: include / body_format / max_body_chars / output_dir,
// ---- batch ids (#34, #35) and resolve_inline_images (#29) ----

use outlook_mcp_rs::outlook::fake::{FAKE_IMAGE_BYTES, FAKE_IMAGE_DATA_URI, MISSING_ID, NOTE_ID, TASK_ID};

/// A fresh, empty directory under the system temp dir for output_dir tests.
fn out_dir(name: &str) -> std::path::PathBuf {
    let dir = std::env::temp_dir()
        .join(format!("outlook-mcp-rs-tools-{}-{name}", std::process::id()));
    let _ = std::fs::remove_dir_all(&dir);
    dir
}

async fn get_email_json(server: &OutlookMcpServer, args: Value) -> Result<Value, String> {
    let params: GetEmailParams = serde_json::from_value(args).unwrap();
    server
        .get_email(Parameters(params))
        .await
        .map(|r| result_json(&r))
        .map_err(|e| e.message.to_string())
}

#[tokio::test]
async fn get_email_default_output_is_unchanged() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let v = get_email_json(&server, json!({"email_id": EMAIL_ID})).await.unwrap();
    assert_eq!(v["body"], "Hi there");
    assert_eq!(v["body_truncated"], false);
    assert_eq!(v["body_length"], 8);
    assert_eq!(v["attachments"], json!([]));
    for absent in ["html_body", "html_truncated", "html_length", "body_file", "html_body_file",
        "inline_images_resolved", "inline_images_unresolved"] {
        assert!(v.get(absent).is_none(), "{absent} in {v}");
    }
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["email_ids"], json!([EMAIL_ID]));
    assert_eq!(args["body"], true);
    assert_eq!(args["html_body"], false);
    assert_eq!(args["attachments"], true);
    assert_eq!(args["meeting"], true);
}

#[tokio::test]
async fn get_email_body_format_html_matches_deprecated_prefer_html() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let html = get_email_json(&server, json!({"email_id": EMAIL_ID, "body_format": "html"})).await.unwrap();
    let legacy = get_email_json(&server, json!({"email_id": EMAIL_ID, "prefer_html": true})).await.unwrap();
    assert_eq!(html, legacy);
    assert_eq!(html["html_body"], "<p>Hi there</p>");
    assert_eq!(html["body"], "Hi there");
    let text = get_email_json(&server, json!({"email_id": EMAIL_ID, "body_format": "text"})).await.unwrap();
    assert!(text.get("html_body").is_none());
}

#[tokio::test]
async fn get_email_rejects_contradicting_prefer_html_and_body_format() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let err = get_email_json(&server, json!({
        "email_id": EMAIL_ID, "prefer_html": true, "body_format": "text"
    })).await.unwrap_err();
    assert!(err.contains("prefer_html") && err.contains("body_format"), "{err}");
    let err = get_email_json(&server, json!({"email_id": EMAIL_ID, "body_format": "rtf"}))
        .await.unwrap_err();
    assert!(err.contains("body_format"), "{err}");
    assert!(fake.calls().is_empty());
}

#[tokio::test]
async fn get_email_include_selects_fields() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    // Metadata only.
    let v = get_email_json(&server, json!({"email_id": EMAIL_ID, "include": []})).await.unwrap();
    for absent in ["body", "body_truncated", "body_length", "html_body", "attachments", "meeting"] {
        assert!(v.get(absent).is_none(), "{absent} in {v}");
    }
    assert_eq!(v["subject"], "Hello");
    assert_eq!(v["item_type"], "email");
    assert_eq!(v["is_meeting"], false);
    // HTML only.
    let v = get_email_json(&server, json!({"email_id": EMAIL_ID, "include": ["html_body"]})).await.unwrap();
    assert!(v.get("body").is_none());
    assert_eq!(v["html_body"], "<p>Hi there</p>");
    assert_eq!(v["html_truncated"], false);
    assert_eq!(v["html_length"], 15);
    // Unknown field: error naming the valid ones.
    let err = get_email_json(&server, json!({"email_id": EMAIL_ID, "include": ["headers"]}))
        .await.unwrap_err();
    assert!(err.contains("headers") && err.contains("html_body"), "{err}");
}

#[tokio::test]
async fn get_email_max_body_chars_truncates_and_reports_length() {
    let fake = Arc::new(FakeOutlookClient::new());
    fake.set_email_text("Long", "Ada", "x".repeat(2500));
    let server = OutlookMcpServer::new(fake.clone());
    let v = get_email_json(&server, json!({"email_id": EMAIL_ID, "max_body_chars": 1000})).await.unwrap();
    assert_eq!(v["body_truncated"], true);
    assert_eq!(v["body_length"], 2500);
    assert!(v["body"].as_str().unwrap().ends_with("[... truncated at 1000 characters]"));
}

#[tokio::test]
async fn get_email_output_dir_writes_full_bodies_to_files() {
    let fake = Arc::new(FakeOutlookClient::new());
    fake.set_email_text("Long", "Ada", "שלום ".repeat(500));
    let server = OutlookMcpServer::new(fake.clone());
    let dir = out_dir("email");
    let v = get_email_json(&server, json!({
        "email_id": EMAIL_ID, "body_format": "html", "max_body_chars": 1000,
        "output_dir": dir.to_str().unwrap(),
    })).await.unwrap();
    assert!(v.get("body").is_none() && v.get("html_body").is_none(), "{v}");
    let body_file = std::path::PathBuf::from(v["body_file"].as_str().unwrap());
    let html_file = std::path::PathBuf::from(v["html_body_file"].as_str().unwrap());
    assert!(body_file.is_absolute() && body_file.starts_with(&dir));
    // Files hold the full text: max_body_chars only limits inline bodies.
    assert_eq!(std::fs::read_to_string(&body_file).unwrap(), "שלום ".repeat(500));
    assert_eq!(std::fs::read_to_string(&html_file).unwrap(), "<p>Hi there</p>");
    assert_eq!(v["body_truncated"], false);
    assert_eq!(v["body_length"], 2500);
    assert_eq!(v["html_length"], 15);
    let _ = std::fs::remove_dir_all(&dir);
}

#[tokio::test]
async fn get_email_batch_returns_a_list_with_per_item_errors() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let v = get_email_json(&server, json!({"email_ids": [EMAIL_ID, MISSING_ID, "entry-9|store-1"]}))
        .await.unwrap();
    let items = v.as_array().unwrap();
    assert_eq!(items.len(), 3);
    assert_eq!(items[0]["id"], EMAIL_ID);
    assert_eq!(items[0]["body"], "Hi there");
    assert_eq!(items[1], json!({"id": MISSING_ID, "error": format!("Item not found: {MISSING_ID}")}));
    assert_eq!(items[2]["id"], "entry-9|store-1");
    // One call to the client for the whole batch.
    assert_eq!(fake.calls().len(), 1);
    // A one-element list is still a list.
    let v = get_email_json(&server, json!({"email_id": [EMAIL_ID]})).await.unwrap();
    assert!(v.is_array());
    // A single failing id (not a list) fails the call, as before batching.
    let err = get_email_json(&server, json!({"email_id": MISSING_ID})).await.unwrap_err();
    assert!(err.contains("Item not found"), "{err}");
    // An empty list is an error.
    let err = get_email_json(&server, json!({"email_id": []})).await.unwrap_err();
    assert!(err.contains("non-empty"), "{err}");
}

#[tokio::test]
async fn get_email_resolve_inline_images_inlines_known_cids_only() {
    let fake = Arc::new(FakeOutlookClient::new());
    fake.set_email_html(r#"<p>Logo:</p><img src="cid:LOGO@example"><img src="cid:other@x">"#);
    let server = OutlookMcpServer::new(fake.clone());
    let v = get_email_json(&server, json!({
        "email_id": EMAIL_ID, "include": [], "resolve_inline_images": true
    })).await.unwrap();
    assert_eq!(
        v["html_body"],
        format!(r#"<p>Logo:</p><img src="{FAKE_IMAGE_DATA_URI}"><img src="cid:other@x">"#)
    );
    assert_eq!(v["inline_images_resolved"], 1);
    assert_eq!(v["inline_images_unresolved"], json!(["other@x"]));
    assert!(v.get("body").is_none());
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["resolve_inline_images"], true);
    assert_eq!(args["html_body"], true);
}

#[tokio::test]
async fn get_email_resolved_html_goes_to_the_file_with_output_dir() {
    let fake = Arc::new(FakeOutlookClient::new());
    fake.set_email_html(r#"<img src="cid:logo@example">"#);
    let server = OutlookMcpServer::new(fake.clone());
    let dir = out_dir("resolve");
    let v = get_email_json(&server, json!({
        "email_id": EMAIL_ID, "resolve_inline_images": true, "output_dir": dir.to_str().unwrap(),
    })).await.unwrap();
    let html = std::fs::read_to_string(v["html_body_file"].as_str().unwrap()).unwrap();
    assert_eq!(html, format!(r#"<img src="{FAKE_IMAGE_DATA_URI}">"#));
    assert!(v.get("html_body").is_none());
    let _ = std::fs::remove_dir_all(&dir);
}

#[tokio::test]
async fn get_event_and_get_note_report_body_length_and_honour_include() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: GetNoteParams = serde_json::from_value(json!({"note_id": NOTE_ID})).unwrap();
    let v = result_json(&server.get_note(Parameters(params)).await.unwrap());
    assert_eq!(v["body_length"], 11);
    let params: GetNoteParams =
        serde_json::from_value(json!({"note_id": NOTE_ID, "include": []})).unwrap();
    let v = result_json(&server.get_note(Parameters(params)).await.unwrap());
    assert!(v.get("body").is_none() && v.get("body_length").is_none());
    assert_eq!(v["subject"], "Ideas");
    let params: GetEventParams =
        serde_json::from_value(json!({"event_id": "x", "max_body_chars": 2000})).unwrap();
    let v = result_json(&server.get_event(Parameters(params)).await.unwrap());
    assert_eq!(v["body_length"], 0);
    let (name, args) = fake.calls().last().unwrap().clone();
    assert_eq!(name, "get_event");
    assert_eq!(args["max_body_chars"], 2000);
}

#[tokio::test]
async fn non_email_get_tools_reject_html_body_format() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: GetEventParams =
        serde_json::from_value(json!({"event_id": "x", "body_format": "html"})).unwrap();
    let err = server.get_event(Parameters(params)).await.unwrap_err();
    assert!(err.message.contains("plain-text"), "{}", err.message);
    let params: GetTaskParams =
        serde_json::from_value(json!({"task_id": TASK_ID, "include": ["html_body"]})).unwrap();
    let err = server.get_task(Parameters(params)).await.unwrap_err();
    assert!(err.message.contains("html_body"), "{}", err.message);
    assert!(fake.calls().is_empty());
}

#[tokio::test]
async fn get_note_batch_and_output_dir() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let dir = out_dir("note");
    let params: GetNoteParams = serde_json::from_value(json!({
        "note_ids": [NOTE_ID, MISSING_ID], "output_dir": dir.to_str().unwrap()
    })).unwrap();
    let v = result_json(&server.get_note(Parameters(params)).await.unwrap());
    assert_eq!(v[0]["id"], NOTE_ID);
    assert_eq!(std::fs::read_to_string(v[0]["body_file"].as_str().unwrap()).unwrap(), "Ideas\n- one");
    assert!(v[1]["error"].as_str().unwrap().contains("not found"));
    let _ = std::fs::remove_dir_all(&dir);
}

#[tokio::test]
async fn get_task_returns_details_and_body() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: GetTaskParams = serde_json::from_value(json!({"task_id": TASK_ID})).unwrap();
    let v = result_json(&server.get_task(Parameters(params)).await.unwrap());
    assert_eq!(v["id"], TASK_ID);
    assert_eq!(v["subject"], "Buy milk");
    assert_eq!(v["status"], "not_started");
    assert_eq!(v["body"], "2 litres, semi-skimmed");
    assert_eq!(v["body_truncated"], false);
    assert_eq!(v["body_length"], 22);
    assert_eq!(v["percent_complete"], 0);
    assert_eq!(v["reminder_set"], false);
    for key in ["start_date", "date_completed", "reminder_time", "created", "modified"] {
        assert!(v.as_object().unwrap().contains_key(key), "missing {key}");
    }
    assert_eq!(fake.calls()[0].0, "get_task");
    assert_eq!(fake.calls()[0].1["task_ids"], json!([TASK_ID]));
}

#[tokio::test]
async fn get_task_batch_reports_per_item_errors() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: GetTaskParams =
        serde_json::from_value(json!({"task_id": [MISSING_ID, TASK_ID], "include": []})).unwrap();
    let v = result_json(&server.get_task(Parameters(params)).await.unwrap());
    assert_eq!(v[0]["id"], MISSING_ID);
    assert!(v[0]["error"].is_string());
    assert_eq!(v[1]["id"], TASK_ID);
    assert!(v[1].get("body").is_none());
}

#[tokio::test]
async fn list_attachments_batch_wraps_each_email() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListAttachmentsParams =
        serde_json::from_value(json!({"email_ids": [EMAIL_ID, MISSING_ID]})).unwrap();
    let v = result_json(&server.list_attachments(Parameters(params)).await.unwrap());
    assert_eq!(v[0]["id"], EMAIL_ID);
    assert_eq!(v[0]["attachments"][0]["filename"], "report.pdf");
    assert_eq!(v[1]["id"], MISSING_ID);
    assert!(v[1]["error"].is_string());
    assert_eq!(fake.calls(), vec![(
        "list_attachments".to_string(), json!({"email_ids": [EMAIL_ID, MISSING_ID]}),
    )]);
}

#[tokio::test]
async fn get_inline_image_content_ids_batch_with_per_item_errors() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: GetInlineImageParams = serde_json::from_value(json!({
        "email_id": EMAIL_ID, "content_ids": ["<LOGO@example>", "nope@x"]
    })).unwrap();
    let v = result_json(&server.get_inline_image(Parameters(params)).await.unwrap());
    assert_eq!(v[0]["content_id"], "logo@example");
    assert_eq!(v[0]["data_uri"], FAKE_IMAGE_DATA_URI);
    assert_eq!(v[1]["id"], "nope@x");
    assert!(v[1]["error"].as_str().unwrap().contains("not found"));
}

#[tokio::test]
async fn get_inline_image_output_dir_writes_image_bytes() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let dir = out_dir("image");
    let params: GetInlineImageParams = serde_json::from_value(json!({
        "email_id": EMAIL_ID, "content_id": "logo@example", "output_dir": dir.to_str().unwrap()
    })).unwrap();
    let v = result_json(&server.get_inline_image(Parameters(params)).await.unwrap());
    assert!(v.get("data_uri").is_none(), "{v}");
    let file = v["data_file"].as_str().unwrap();
    assert!(file.ends_with(".png"), "{file}");
    assert_eq!(std::fs::read(file).unwrap(), FAKE_IMAGE_BYTES);
    assert_eq!(v["size"], 4);
    let _ = std::fs::remove_dir_all(&dir);
}
