use std::sync::Arc;

use rmcp::{
    ErrorData as McpError, ServerHandler,
    handler::server::wrapper::Parameters,
    model::{CallToolResult, ContentBlock, ServerCapabilities, ServerInfo},
    tool, tool_handler, tool_router,
};
use serde::Deserialize;

use crate::error::ToolError;
use crate::outlook::OutlookClient;

/// Runs a blocking `OutlookClient` call on a dedicated blocking thread so the
/// tokio scheduler never migrates it mid-call (COM apartment-threading
/// requires the same OS thread for the lifetime of a call).
async fn run_blocking<T, F>(f: F) -> Result<T, ToolError>
where
    T: Send + 'static,
    F: FnOnce() -> Result<T, ToolError> + Send + 'static,
{
    tokio::task::spawn_blocking(f)
        .await
        .map_err(|e| ToolError::new(format!("internal task error: {e}")))?
}

fn json_content<T: serde::Serialize>(value: &T) -> Result<ContentBlock, McpError> {
    ContentBlock::json(value)
}

#[derive(Clone)]
pub struct OutlookMcpServer {
    client: Arc<dyn OutlookClient>,
}

impl OutlookMcpServer {
    pub fn new(client: Arc<dyn OutlookClient>) -> Self {
        Self { client }
    }
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ListEmailsParams {
    #[serde(default = "default_folder")]
    pub folder: String,
    #[serde(default = "default_count")]
    pub count: i32,
    #[serde(default)]
    pub unread_only: bool,
}
fn default_folder() -> String { "inbox".to_string() }
fn default_count() -> i32 { 10 }

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct SearchEmailsParams {
    pub query: String,
    #[serde(default = "default_folder")]
    pub folder: String,
    #[serde(default = "default_count")]
    pub count: i32,
    #[serde(default)]
    pub since_days: Option<i32>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct GetEmailParams {
    pub email_id: String,
    #[serde(default)]
    pub prefer_html: bool,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct SendEmailParams {
    pub to: Vec<String>,
    pub subject: String,
    pub body: String,
    #[serde(default)]
    pub cc: Option<Vec<String>>,
    #[serde(default)]
    pub bcc: Option<Vec<String>>,
    #[serde(default)]
    pub html: bool,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CreateDraftParams {
    pub to: Vec<String>,
    pub subject: String,
    pub body: String,
    #[serde(default)]
    pub cc: Option<Vec<String>>,
    #[serde(default)]
    pub bcc: Option<Vec<String>>,
    #[serde(default)]
    pub html: bool,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ReplyEmailParams {
    pub email_id: String,
    pub body: String,
    #[serde(default)]
    pub reply_all: bool,
    #[serde(default)]
    pub html: bool,
    #[serde(default = "default_true")]
    pub send: bool,
}
fn default_true() -> bool { true }

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct MoveEmailParams {
    pub email_id: String,
    pub target_folder: String,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct DeleteEmailParams {
    pub email_id: String,
}

// ---- Calendar ----

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ListEventsParams {
    #[serde(default)]
    pub start_date: Option<String>,
    #[serde(default)]
    pub end_date: Option<String>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct GetEventParams {
    pub event_id: String,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CreateEventParams {
    pub subject: String,
    pub start: String,
    pub end: String,
    #[serde(default)]
    pub body: Option<String>,
    #[serde(default)]
    pub location: Option<String>,
    #[serde(default)]
    pub attendees: Option<Vec<String>>,
    #[serde(default)]
    pub all_day: bool,
    #[serde(default)]
    pub reminder_minutes: Option<i32>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct RespondToMeetingParams {
    pub event_id: String,
    pub response: String,
    #[serde(default)]
    pub comment: Option<String>,
    #[serde(default = "default_true")]
    pub send: bool,
}

// ---- Attachments ----

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ListAttachmentsParams {
    pub email_id: String,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct SaveAttachmentsParams {
    pub email_id: String,
    pub save_dir: String,
    #[serde(default)]
    pub attachment_names: Option<Vec<String>>,
}

// ---- Tasks ----

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ListTasksParams {
    #[serde(default)]
    pub include_completed: bool,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CreateTaskParams {
    pub subject: String,
    #[serde(default)]
    pub body: Option<String>,
    #[serde(default)]
    pub due_date: Option<String>,
    #[serde(default = "default_importance")]
    pub importance: String,
}
fn default_importance() -> String { "normal".to_string() }

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CompleteTaskParams {
    pub task_id: String,
}

#[tool_router]
impl OutlookMcpServer {
    #[tool(description = "List Outlook mail folders (name, path, item counts).")]
    pub async fn list_folders(&self) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.list_folders()).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "List recent emails in a folder (default: inbox).")]
    pub async fn list_emails(
        &self,
        Parameters(ListEmailsParams { folder, count, unread_only }): Parameters<ListEmailsParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.list_emails(folder, count, unread_only)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Search emails by subject/sender/body text in a folder.")]
    pub async fn search_emails(
        &self,
        Parameters(SearchEmailsParams { query, folder, count, since_days }): Parameters<SearchEmailsParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.search_emails(query, folder, count, since_days)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Get the full body and attachment list of one email by id.")]
    pub async fn get_email(
        &self,
        Parameters(GetEmailParams { email_id, prefer_html }): Parameters<GetEmailParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.get_email(email_id, prefer_html)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Send a new email immediately.")]
    pub async fn send_email(
        &self,
        Parameters(SendEmailParams { to, subject, body, cc, bcc, html }): Parameters<SendEmailParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.send_email(to, subject, body, cc, bcc, html)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Create (but don't send) a draft email.")]
    pub async fn create_draft(
        &self,
        Parameters(CreateDraftParams { to, subject, body, cc, bcc, html }): Parameters<CreateDraftParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.create_draft(to, subject, body, cc, bcc, html)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Reply to an email, optionally to all recipients, optionally as a draft.")]
    pub async fn reply_email(
        &self,
        Parameters(ReplyEmailParams { email_id, body, reply_all, html, send }): Parameters<ReplyEmailParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.reply_email(email_id, body, reply_all, html, send)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Move an email to another folder.")]
    pub async fn move_email(
        &self,
        Parameters(MoveEmailParams { email_id, target_folder }): Parameters<MoveEmailParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.move_email(email_id, target_folder)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Delete an email (moves it to Deleted Items).")]
    pub async fn delete_email(
        &self,
        Parameters(DeleteEmailParams { email_id }): Parameters<DeleteEmailParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.delete_email(email_id)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    // ---- Calendar ----

    #[tool(description = "List calendar events in a date range (default: next 7 days).")]
    pub async fn list_events(
        &self,
        Parameters(ListEventsParams { start_date, end_date }): Parameters<ListEventsParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.list_events(start_date, end_date)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Get the full details of one calendar event by id.")]
    pub async fn get_event(
        &self,
        Parameters(GetEventParams { event_id }): Parameters<GetEventParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.get_event(event_id)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Create a calendar event, optionally sending meeting invites to attendees.")]
    pub async fn create_event(
        &self,
        Parameters(CreateEventParams {
            subject, start, end, body, location, attendees, all_day, reminder_minutes,
        }): Parameters<CreateEventParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || {
            client.create_event(subject, start, end, body, location, attendees, all_day, reminder_minutes)
        }).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Respond to a meeting invite: accept, decline, or tentative.")]
    pub async fn respond_to_meeting(
        &self,
        Parameters(RespondToMeetingParams { event_id, response, comment, send }): Parameters<RespondToMeetingParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.respond_to_meeting(event_id, response, comment, send)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    // ---- Attachments ----

    #[tool(description = "List an email's attachments (filename and size).")]
    pub async fn list_attachments(
        &self,
        Parameters(ListAttachmentsParams { email_id }): Parameters<ListAttachmentsParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.list_attachments(email_id)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Save an email's attachments to a local directory.")]
    pub async fn save_attachments(
        &self,
        Parameters(SaveAttachmentsParams { email_id, save_dir, attachment_names }): Parameters<SaveAttachmentsParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.save_attachments(email_id, save_dir, attachment_names)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    // ---- Tasks ----

    #[tool(description = "List Outlook tasks (default: not-yet-completed only).")]
    pub async fn list_tasks(
        &self,
        Parameters(ListTasksParams { include_completed }): Parameters<ListTasksParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.list_tasks(include_completed)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Create a new task.")]
    pub async fn create_task(
        &self,
        Parameters(CreateTaskParams { subject, body, due_date, importance }): Parameters<CreateTaskParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.create_task(subject, body, due_date, importance)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Mark a task complete.")]
    pub async fn complete_task(
        &self,
        Parameters(CompleteTaskParams { task_id }): Parameters<CompleteTaskParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.complete_task(task_id)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
}

#[tool_handler]
impl ServerHandler for OutlookMcpServer {
    fn get_info(&self) -> ServerInfo {
        ServerInfo::new(ServerCapabilities::builder().enable_tools().build())
            .with_instructions(
                "Controls Microsoft Outlook desktop (email, calendar, tasks, notes) via COM.",
            )
    }
}
