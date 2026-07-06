use std::sync::Arc;

use outlook_mcp_rs::outlook::fake::{FakeOutlookClient, EMAIL_ID};
use outlook_mcp_rs::server::{
    DeleteEmailParams, GetEmailParams, ListEmailsParams, MoveEmailParams, OutlookMcpServer,
    ReplyEmailParams, SearchEmailsParams, SendEmailParams, CreateDraftParams,
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
    server
        .list_emails(Parameters(ListEmailsParams {
            folder: "sent".to_string(),
            count: 5,
            unread_only: true,
        }))
        .await
        .unwrap();
    assert_eq!(
        fake.calls(),
        vec![(
            "list_emails".to_string(),
            json!({"folder": "sent", "count": 5, "unread_only": true})
        )]
    );
}

#[tokio::test]
async fn list_emails_uses_defaults() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEmailsParams = serde_json::from_value(json!({})).unwrap();
    server.list_emails(Parameters(params)).await.unwrap();
    assert_eq!(
        fake.calls(),
        vec![(
            "list_emails".to_string(),
            json!({"folder": "inbox", "count": 10, "unread_only": false})
        )]
    );
}

#[tokio::test]
async fn search_emails_passes_query_and_since_days() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: SearchEmailsParams =
        serde_json::from_value(json!({"query": "invoice", "since_days": 30})).unwrap();
    server.search_emails(Parameters(params)).await.unwrap();
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "search_emails");
    assert_eq!(args["query"], "invoice");
    assert_eq!(args["since_days"], 30);
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
        }))
        .await
        .unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["reply_all"], true);
    assert_eq!(args["send"], false);
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
    let err = server
        .list_emails(Parameters(ListEmailsParams {
            folder: "inbox".to_string(),
            count: 10,
            unread_only: false,
        }))
        .await
        .unwrap_err();
    assert!(err.message.contains("Outlook exploded"));
}
