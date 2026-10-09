use std::sync::Arc;

use rmcp::{
    ErrorData as McpError, ServerHandler,
    handler::server::wrapper::Parameters,
    model::{CallToolResult, ContentBlock, ServerCapabilities, ServerInfo},
    tool, tool_handler, tool_router,
};
use serde::Deserialize;

use crate::error::ToolError;
use crate::outlook::{CheckAvailabilityInput, CreateEventInput, DraftUpdate, EmailQuery, EmailUpdate, EventQuery, EventUpdate, NoteQuery, NoteUpdate, OutlookClient, RecurrenceInput, TaskQuery, TaskUpdate, InlineImage};
use crate::outlook::{require_mail_body, resolve_mail_body, resolve_text_input, BodyInputs, MailBody, NewEmail, ReplyInput};

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
    #[serde(default)]
    pub query: Option<String>,
    #[serde(default = "default_folder")]
    pub folder: String,
    #[serde(default = "default_count")]
    pub count: i32,
    /// Matches to skip before this page starts (default 0). To page, call
    /// again with offset += count until fewer than count results come back.
    #[serde(default)]
    pub offset: i32,
    #[serde(default)]
    pub unread_only: bool,
    #[serde(default)]
    pub from: Option<String>,
    #[serde(default)]
    pub to: Option<String>,
    #[serde(default)]
    pub category: Option<String>,
    #[serde(default)]
    pub received_after: Option<String>,
    #[serde(default)]
    pub received_before: Option<String>,
    #[serde(default)]
    pub since_days: Option<i32>,
    #[serde(default)]
    pub has_attachments: Option<bool>,
    #[serde(default)]
    pub flagged: bool,
    #[serde(default)]
    pub high_importance: bool,
}
fn default_folder() -> String { "inbox".to_string() }
fn default_count() -> i32 { 10 }

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct GetEmailParams {
    pub email_id: String,
    #[serde(default)]
    pub prefer_html: bool,
    /// Cut body/html_body at this many characters. Default 100,000;
    /// clamped to 1,000..=5,000,000.
    #[serde(default)]
    pub max_body_chars: Option<u32>,
}

// Field docs below are the public schema descriptions, so the shared body
// fields repeat them on every mail tool.

#[derive(Debug, Default, Deserialize, schemars::JsonSchema)]
pub struct SendEmailParams {
    pub to: Vec<String>,
    pub subject: String,
    /// Plain-text body. Give exactly one of body, html_body, body_file, html_body_file.
    #[serde(default)]
    pub body: Option<String>,
    /// HTML body. `data:` image URIs in it (e.g. <img src="data:image/png;base64,...">)
    /// are sent as inline Content-ID attachments.
    #[serde(default)]
    pub html_body: Option<String>,
    /// Local path of a UTF-8 text file to use as the plain-text body (for large bodies).
    #[serde(default)]
    pub body_file: Option<String>,
    /// Local path of a UTF-8 HTML file to use as the HTML body (for large bodies).
    #[serde(default)]
    pub html_body_file: Option<String>,
    #[serde(default)]
    pub cc: Option<Vec<String>>,
    #[serde(default)]
    pub bcc: Option<Vec<String>>,
    /// Deprecated: use html_body / html_body_file instead. true = body/body_file is HTML.
    #[serde(default)]
    pub html: Option<bool>,
    #[serde(default)]
    pub attachments: Option<Vec<String>>,
    /// Images to embed as hidden Content-ID attachments (requires an HTML body).
    /// Reference each in the HTML body as <img src="cid:CONTENT_ID">.
    #[serde(default)]
    pub inline_images: Option<Vec<InlineImage>>,
    /// Category names to assign.
    #[serde(default)]
    pub categories: Option<Vec<String>>,
    /// "low" | "normal" | "high".
    #[serde(default)]
    pub importance: Option<String>,
}

#[derive(Debug, Default, Deserialize, schemars::JsonSchema)]
pub struct CreateDraftParams {
    pub to: Vec<String>,
    pub subject: String,
    /// Plain-text body. Give exactly one of body, html_body, body_file, html_body_file.
    #[serde(default)]
    pub body: Option<String>,
    /// HTML body. `data:` image URIs in it (e.g. <img src="data:image/png;base64,...">)
    /// are saved as inline Content-ID attachments.
    #[serde(default)]
    pub html_body: Option<String>,
    /// Local path of a UTF-8 text file to use as the plain-text body (for large bodies).
    #[serde(default)]
    pub body_file: Option<String>,
    /// Local path of a UTF-8 HTML file to use as the HTML body (for large bodies).
    #[serde(default)]
    pub html_body_file: Option<String>,
    #[serde(default)]
    pub cc: Option<Vec<String>>,
    #[serde(default)]
    pub bcc: Option<Vec<String>>,
    /// Deprecated: use html_body / html_body_file instead. true = body/body_file is HTML.
    #[serde(default)]
    pub html: Option<bool>,
    #[serde(default)]
    pub attachments: Option<Vec<String>>,
    /// Images to embed as hidden Content-ID attachments (requires an HTML body).
    /// Reference each in the HTML body as <img src="cid:CONTENT_ID">.
    #[serde(default)]
    pub inline_images: Option<Vec<InlineImage>>,
    /// Category names to assign.
    #[serde(default)]
    pub categories: Option<Vec<String>>,
    /// "low" | "normal" | "high".
    #[serde(default)]
    pub importance: Option<String>,
}

#[derive(Debug, Default, Deserialize, schemars::JsonSchema)]
pub struct ReplyEmailParams {
    pub email_id: String,
    /// Plain-text reply, put above the quoted original. Give exactly one of
    /// body, html_body, body_file, html_body_file.
    #[serde(default)]
    pub body: Option<String>,
    /// HTML reply, put above the quoted original. `data:` image URIs in it
    /// become inline Content-ID attachments.
    #[serde(default)]
    pub html_body: Option<String>,
    /// Local path of a UTF-8 text file to use as the plain-text reply.
    #[serde(default)]
    pub body_file: Option<String>,
    /// Local path of a UTF-8 HTML file to use as the HTML reply.
    #[serde(default)]
    pub html_body_file: Option<String>,
    #[serde(default)]
    pub reply_all: bool,
    /// Deprecated: use html_body / html_body_file instead. true = body/body_file is HTML.
    #[serde(default)]
    pub html: Option<bool>,
    #[serde(default = "default_true")]
    pub send: bool,
    #[serde(default)]
    pub attachments: Option<Vec<String>>,
    /// Images to embed as hidden Content-ID attachments (requires an HTML reply).
    /// Reference each in the HTML as <img src="cid:CONTENT_ID">.
    #[serde(default)]
    pub inline_images: Option<Vec<InlineImage>>,
    /// Category names to assign to the reply.
    #[serde(default)]
    pub categories: Option<Vec<String>>,
    /// "low" | "normal" | "high".
    #[serde(default)]
    pub importance: Option<String>,
}
fn default_true() -> bool { true }

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct UpdateEmailParams {
    pub email_id: String,
    /// Destination folder name (e.g. "Archive"). Applied last; changes the id.
    #[serde(default)]
    pub move_to: Option<String>,
    /// true = mark read, false = mark unread.
    #[serde(default)]
    pub mark_read: Option<bool>,
    /// "follow_up" | "complete" | "clear".
    #[serde(default)]
    pub flag: Option<String>,
    /// Category names to add (existing categories are preserved).
    #[serde(default)]
    pub add_categories: Option<Vec<String>>,
    /// Category names to remove.
    #[serde(default)]
    pub remove_categories: Option<Vec<String>>,
    /// "low" | "normal" | "high".
    #[serde(default)]
    pub importance: Option<String>,
}

#[derive(Debug, Default, Deserialize, schemars::JsonSchema)]
pub struct UpdateDraftParams {
    /// Id of an unsent draft (e.g. from create_draft or reply_email with send=false).
    /// `draft_id` is accepted as a deprecated alias.
    #[serde(alias = "draft_id")]
    pub email_id: String,
    /// New subject (replaces the current one).
    #[serde(default)]
    pub subject: Option<String>,
    /// New plain-text body (replaces the current body). At most one of body,
    /// html_body, body_file, html_body_file.
    #[serde(default)]
    pub body: Option<String>,
    /// New HTML body (replaces the current body). `data:` image URIs in it
    /// become inline Content-ID attachments.
    #[serde(default)]
    pub html_body: Option<String>,
    /// Local path of a UTF-8 text file to use as the new plain-text body.
    #[serde(default)]
    pub body_file: Option<String>,
    /// Local path of a UTF-8 HTML file to use as the new HTML body.
    #[serde(default)]
    pub html_body_file: Option<String>,
    /// Replaces the whole To line; an empty list clears it.
    #[serde(default)]
    pub to: Option<Vec<String>>,
    /// Replaces the whole CC line; an empty list clears it.
    #[serde(default)]
    pub cc: Option<Vec<String>>,
    /// Replaces the whole BCC line; an empty list clears it.
    #[serde(default)]
    pub bcc: Option<Vec<String>>,
    /// Local file paths appended to the existing attachments.
    #[serde(default)]
    pub attachments: Option<Vec<String>>,
    /// Images to embed as hidden Content-ID attachments (requires html_body or
    /// html_body_file). Reference each in the HTML as <img src="cid:CONTENT_ID">.
    /// An existing attachment with the same Content-ID is replaced.
    #[serde(default)]
    pub inline_images: Option<Vec<InlineImage>>,
    /// Category names to add (existing categories are preserved).
    #[serde(default)]
    pub add_categories: Option<Vec<String>>,
    /// Category names to remove.
    #[serde(default)]
    pub remove_categories: Option<Vec<String>>,
    /// "low" | "normal" | "high".
    #[serde(default)]
    pub importance: Option<String>,
    /// true = send the draft after applying the changes (default false: only save).
    #[serde(default)]
    pub send: bool,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct DeleteEmailParams {
    pub email_id: String,
    /// true = hard delete (like shift+delete); IRREVERSIBLE. Default false
    /// moves the email to Deleted Items.
    #[serde(default)]
    pub permanent: bool,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct EmptyDeletedItemsParams {
    /// Must be true, or the call is refused. Only pass true when the user
    /// has explicitly asked to empty Deleted Items.
    #[serde(default)]
    pub confirm: bool,
}

// ---- Calendar ----

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ListEventsParams {
    #[serde(default)]
    pub start_date: Option<String>,
    #[serde(default)]
    pub end_date: Option<String>,
    /// Text match on subject + location.
    #[serde(default)]
    pub query: Option<String>,
    /// Filter to a color category.
    #[serde(default)]
    pub category: Option<String>,
    /// "free" | "tentative" | "busy" | "out_of_office" | "working_elsewhere".
    #[serde(default)]
    pub show_as: Option<String>,
    /// This mailbox's response: "organizer" | "accepted" | "declined" | "tentative" | "not_responded".
    #[serde(default)]
    pub my_response: Option<String>,
    /// Names/emails; match events where ANY listed person participates.
    #[serde(default)]
    pub attendees: Option<Vec<String>>,
    /// "required" | "optional" | "any" (default "any").
    #[serde(default)]
    pub attendee_role: Option<String>,
    /// Only events that have other attendees (meetings).
    #[serde(default)]
    pub meetings_only: bool,
    /// Only all-day (true) or only non-all-day (false) events.
    #[serde(default)]
    pub all_day: Option<bool>,
    /// Email/name of another person whose shared calendar to view (default: your own).
    #[serde(default)]
    pub calendar_of: Option<String>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct GetEventParams {
    pub event_id: String,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct RecurrenceParams {
    /// "daily" | "weekly" | "monthly" | "yearly".
    pub pattern: String,
    /// Repeat every N days/weeks/months/years (default 1).
    #[serde(default)]
    pub interval: Option<i32>,
    /// Required for "weekly": full weekday names, e.g. ["monday", "wednesday"].
    #[serde(default)]
    pub days_of_week: Option<Vec<String>>,
    /// Required for "monthly": day of the month (1-31). Not used for "yearly"
    /// (the event's own start date supplies the month/day).
    #[serde(default)]
    pub day_of_month: Option<i32>,
    /// End date (ISO). At most one of `until`/`occurrences`; neither = no end date.
    #[serde(default)]
    pub until: Option<String>,
    /// Number of occurrences. At most one of `until`/`occurrences`.
    #[serde(default)]
    pub occurrences: Option<i32>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CreateEventParams {
    pub subject: String,
    pub start: String,
    pub end: String,
    /// Plain-text description. Not with body_file.
    #[serde(default)]
    pub body: Option<String>,
    /// Local path of a UTF-8 text file to use as the body. Not with body.
    #[serde(default)]
    pub body_file: Option<String>,
    #[serde(default)]
    pub location: Option<String>,
    /// Legacy alias for `required_attendees`; merged in if both are given.
    #[serde(default)]
    pub attendees: Option<Vec<String>>,
    #[serde(default)]
    pub required_attendees: Option<Vec<String>>,
    #[serde(default)]
    pub optional_attendees: Option<Vec<String>>,
    #[serde(default)]
    pub all_day: bool,
    #[serde(default)]
    pub reminder_minutes: Option<i32>,
    #[serde(default)]
    pub categories: Option<Vec<String>>,
    /// "free" | "tentative" | "busy" | "out_of_office" | "working_elsewhere".
    #[serde(default)]
    pub show_as: Option<String>,
    /// If false, a meeting with attendees is saved (not sent) for later review.
    #[serde(default = "default_true")]
    pub send: bool,
    /// Repeat this event daily/weekly/monthly/yearly. Omit for a one-off event.
    #[serde(default)]
    pub recurrence: Option<RecurrenceParams>,
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

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct UpdateEventParams {
    pub event_id: String,
    #[serde(default)]
    pub subject: Option<String>,
    #[serde(default)]
    pub start: Option<String>,
    #[serde(default)]
    pub end: Option<String>,
    #[serde(default)]
    pub location: Option<String>,
    /// New plain-text description. Not with body_file.
    #[serde(default)]
    pub body: Option<String>,
    /// Local path of a UTF-8 text file to use as the new body. Not with body.
    #[serde(default)]
    pub body_file: Option<String>,
    #[serde(default)]
    pub all_day: Option<bool>,
    #[serde(default)]
    pub reminder_minutes: Option<i32>,
    /// "free" | "tentative" | "busy" | "out_of_office" | "working_elsewhere".
    #[serde(default)]
    pub show_as: Option<String>,
    /// Category names to add (existing categories are preserved).
    #[serde(default)]
    pub add_categories: Option<Vec<String>>,
    /// Category names to remove.
    #[serde(default)]
    pub remove_categories: Option<Vec<String>>,
    /// Adding either attendee list converts a personal appointment into a meeting.
    #[serde(default)]
    pub add_required_attendees: Option<Vec<String>>,
    #[serde(default)]
    pub add_optional_attendees: Option<Vec<String>>,
    /// Names/emails to remove from either attendee tier.
    #[serde(default)]
    pub remove_attendees: Option<Vec<String>>,
    /// If the event is a meeting, notify attendees of these changes (default true).
    /// false = apply quietly to your own copy only. Ignored for non-meetings.
    #[serde(default = "default_true")]
    pub send_update: bool,
    /// Set or replace the event's recurrence pattern (whole series, not one
    /// occurrence). Mutually exclusive with `clear_recurrence`.
    #[serde(default)]
    pub recurrence: Option<RecurrenceParams>,
    /// Remove the event's recurrence pattern, converting it back to a single
    /// occurrence. Mutually exclusive with `recurrence`.
    #[serde(default)]
    pub clear_recurrence: bool,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct DeleteEventParams {
    pub event_id: String,
    /// If you organize the meeting, notify attendees of the cancellation (default true).
    #[serde(default = "default_true")]
    pub send_cancellation: bool,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CheckAvailabilityParams {
    pub people: Vec<String>,
    pub start: String,
    pub end: String,
    #[serde(default = "default_interval_minutes")]
    pub interval_minutes: i32,
    #[serde(default = "default_treat_as_free")]
    pub treat_as_free: Vec<String>,
}
fn default_interval_minutes() -> i32 { 30 }
fn default_treat_as_free() -> Vec<String> { vec!["free".to_string()] }

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

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct GetInlineImageParams {
    pub email_id: String,
    /// The attachment's Content-ID (`content_id` from list_attachments, or the
    /// `cid:...` reference from the HTML body). A `cid:` prefix and `<>` are accepted.
    pub content_id: String,
    /// Also return `context`: up to this many lines (max 50) of plain text
    /// immediately before the image's first `cid:` reference in the HTML body.
    #[serde(default)]
    pub context_lines: Option<u32>,
}

// ---- Tasks ----

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ListTasksParams {
    #[serde(default)]
    pub include_completed: bool,
    /// Filter to a color category.
    #[serde(default)]
    pub category: Option<String>,
    /// "low" | "normal" | "high".
    #[serde(default)]
    pub importance: Option<String>,
    /// Text match on the task's subject.
    #[serde(default)]
    pub query: Option<String>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CreateTaskParams {
    pub subject: String,
    /// Plain-text body. Not with body_file.
    #[serde(default)]
    pub body: Option<String>,
    /// Local path of a UTF-8 text file to use as the body. Not with body.
    #[serde(default)]
    pub body_file: Option<String>,
    #[serde(default)]
    pub due_date: Option<String>,
    #[serde(default = "default_importance")]
    pub importance: String,
    /// Category names to assign on creation.
    #[serde(default)]
    pub categories: Option<Vec<String>>,
    #[serde(default)]
    pub start_date: Option<String>,
    /// ISO datetime — an absolute reminder time (unlike appointment
    /// reminders, which are minutes-before-start).
    #[serde(default)]
    pub reminder_time: Option<String>,
}
fn default_importance() -> String { "normal".to_string() }

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct UpdateTaskParams {
    pub task_id: String,
    /// true = mark complete (100%), false = reopen.
    #[serde(default)]
    pub mark_complete: Option<bool>,
    #[serde(default)]
    pub subject: Option<String>,
    /// New plain-text body. Not with body_file.
    #[serde(default)]
    pub body: Option<String>,
    /// Local path of a UTF-8 text file to use as the new body. Not with body.
    #[serde(default)]
    pub body_file: Option<String>,
    #[serde(default)]
    pub due_date: Option<String>,
    #[serde(default)]
    pub start_date: Option<String>,
    #[serde(default)]
    pub importance: Option<String>,
    #[serde(default)]
    pub add_categories: Option<Vec<String>>,
    #[serde(default)]
    pub remove_categories: Option<Vec<String>>,
    /// 0-100.
    #[serde(default)]
    pub percent_complete: Option<i32>,
    #[serde(default)]
    pub reminder_time: Option<String>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct DeleteTaskParams {
    pub task_id: String,
}

// ---- Notes ----

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ListNotesParams {
    /// Filter to a color category.
    #[serde(default)]
    pub category: Option<String>,
    /// Text match on the note's body.
    #[serde(default)]
    pub query: Option<String>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct GetNoteParams {
    pub note_id: String,
}

#[derive(Debug, Default, Deserialize, schemars::JsonSchema)]
pub struct CreateNoteParams {
    /// The note's text. Give exactly one of body or body_file.
    #[serde(default)]
    pub body: Option<String>,
    /// Local path of a UTF-8 text file to use as the note's text.
    #[serde(default)]
    pub body_file: Option<String>,
    /// Category names to assign on creation.
    #[serde(default)]
    pub categories: Option<Vec<String>>,
    /// "blue" | "green" | "pink" | "yellow" | "white".
    #[serde(default)]
    pub color: Option<String>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct UpdateNoteParams {
    pub note_id: String,
    /// New text. Not with body_file.
    #[serde(default)]
    pub body: Option<String>,
    /// Local path of a UTF-8 text file to use as the new text. Not with body.
    #[serde(default)]
    pub body_file: Option<String>,
    #[serde(default)]
    pub add_categories: Option<Vec<String>>,
    #[serde(default)]
    pub remove_categories: Option<Vec<String>>,
    /// "blue" | "green" | "pink" | "yellow" | "white".
    #[serde(default)]
    pub color: Option<String>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct DeleteNoteParams {
    pub note_id: String,
}

#[tool_router]
impl OutlookMcpServer {
    #[tool(description = "List Outlook mail folders (name, path, item counts).")]
    pub async fn list_folders(&self) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.list_folders()).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Find emails in a folder (newest first) with optional text query and filters (sender, recipient, category, date range, attachments, flagged, importance). `from` matches the sender name or address; `to` matches any To/CC recipient name or address (case-insensitive substring). count is capped at 200. To page through more, call again with offset += count until fewer than count results come back.")]
    pub async fn list_emails(
        &self,
        Parameters(p): Parameters<ListEmailsParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let q = EmailQuery {
            query: p.query, folder: p.folder, count: p.count, offset: p.offset,
            unread_only: p.unread_only, from: p.from, category: p.category,
            received_after: p.received_after, to: p.to,
            received_before: p.received_before, since_days: p.since_days,
            has_attachments: p.has_attachments, flagged: p.flagged,
            high_importance: p.high_importance,
        };
        let result = run_blocking(move || client.list_emails(q)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Get the full body and attachment list of one email by id. Set prefer_html to also get html_body. body (and html_body) are cut at max_body_chars characters (default 100,000; allowed 1,000-5,000,000). The result always reports body_truncated and body_length (the full original length in characters), plus html_truncated and html_length when prefer_html is set. If a *_truncated flag is true, the cut may fall mid-tag or mid-image-data: call again with max_body_chars of at least the reported length (max 5,000,000) to get the complete text.")]
    pub async fn get_email(
        &self,
        Parameters(GetEmailParams { email_id, prefer_html, max_body_chars }): Parameters<GetEmailParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.get_email(email_id, prefer_html, max_body_chars)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Send a new email immediately. Give the body as exactly one of `body` (plain text), `html_body`, `body_file` or `html_body_file` (local UTF-8 file paths; use these for large bodies). `html` is a deprecated flag: html=true makes body/body_file HTML. `attachments` is a list of local file paths. Images can be embedded two ways, both sent as hidden Content-ID attachments: `data:` image URIs inside the HTML (e.g. <img src=\"data:image/png;base64,...\">) are converted automatically, and `inline_images` takes images (a local path or base64 data) you reference as <img src=\"cid:CONTENT_ID\">. `categories` and `importance` (low/normal/high) are optional. Everything is validated before anything is sent.")]
    pub async fn send_email(
        &self,
        Parameters(p): Parameters<SendEmailParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || {
            let body = require_mail_body(BodyInputs {
                body: p.body, html_body: p.html_body, body_file: p.body_file,
                html_body_file: p.html_body_file, html: p.html,
            }, "send_email")?;
            client.send_email(NewEmail {
                to: p.to, subject: p.subject, body, cc: p.cc, bcc: p.bcc,
                attachments: p.attachments, inline_images: p.inline_images,
                categories: p.categories, importance: p.importance,
            })
        }).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Create (but don't send) a draft email. Give the body as exactly one of `body` (plain text), `html_body`, `body_file` or `html_body_file` (local UTF-8 file paths; use these for large bodies). `html` is a deprecated flag: html=true makes body/body_file HTML. `attachments` is a list of local file paths. Images can be embedded two ways, both saved as hidden Content-ID attachments: `data:` image URIs inside the HTML (e.g. <img src=\"data:image/png;base64,...\">) are converted automatically, and `inline_images` takes images (a local path or base64 data) you reference as <img src=\"cid:CONTENT_ID\">. `categories` and `importance` (low/normal/high) are optional. Everything is validated before the draft is created. Edit or send it later with update_draft.")]
    pub async fn create_draft(
        &self,
        Parameters(p): Parameters<CreateDraftParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || {
            let body = require_mail_body(BodyInputs {
                body: p.body, html_body: p.html_body, body_file: p.body_file,
                html_body_file: p.html_body_file, html: p.html,
            }, "create_draft")?;
            client.create_draft(NewEmail {
                to: p.to, subject: p.subject, body, cc: p.cc, bcc: p.bcc,
                attachments: p.attachments, inline_images: p.inline_images,
                categories: p.categories, importance: p.importance,
            })
        }).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Reply to an email, optionally to all recipients (reply_all), and send it (default) or save it as a draft (send=false). Give the reply text as exactly one of `body` (plain text), `html_body`, `body_file` or `html_body_file` (local UTF-8 file paths); it goes above the quoted original. `html` is a deprecated flag: html=true makes body/body_file HTML. `data:` image URIs in an HTML reply become hidden Content-ID attachments, and `inline_images` adds images referenced as <img src=\"cid:CONTENT_ID\">. Optional `attachments` (local file paths), `categories` and `importance` (low/normal/high). Everything is validated before the reply is created.")]
    pub async fn reply_email(
        &self,
        Parameters(p): Parameters<ReplyEmailParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || {
            let body = require_mail_body(BodyInputs {
                body: p.body, html_body: p.html_body, body_file: p.body_file,
                html_body_file: p.html_body_file, html: p.html,
            }, "reply_email")?;
            client.reply_email(ReplyInput {
                email_id: p.email_id, body, reply_all: p.reply_all, send: p.send,
                attachments: p.attachments, inline_images: p.inline_images,
                categories: p.categories, importance: p.importance,
            })
        }).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Update an existing email: move to a folder, mark read/unread, flag (follow_up/complete/clear), add/remove categories, or set importance. Combine any of these in one call.")]
    pub async fn update_email(
        &self,
        Parameters(p): Parameters<UpdateEmailParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let u = EmailUpdate {
            email_id: p.email_id, move_to: p.move_to, mark_read: p.mark_read,
            flag: p.flag, add_categories: p.add_categories,
            remove_categories: p.remove_categories, importance: p.importance,
        };
        let result = run_blocking(move || client.update_email(u)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Edit an existing unsent draft (e.g. one from create_draft or reply_email with send=false) by `email_id` (`draft_id` is a deprecated alias), save it, and optionally send it (send=true; default false). Only the fields you pass change: subject; the body, as at most one of `body` (plain text), `html_body`, `body_file` or `html_body_file` (local UTF-8 file paths; use these for large bodies), replaces the current body; to/cc/bcc replace that recipient line entirely (an empty list clears it); attachments are local file paths appended to the existing ones; add_categories/remove_categories; importance (low/normal/high). `data:` image URIs in a new HTML body become hidden Content-ID attachments, and `inline_images` (needs an HTML body) adds images referenced as <img src=\"cid:CONTENT_ID\">. send=true alone sends the draft as is. Sent or received emails are rejected. Returns the draft id and the list of changed fields (status \"draft_updated\"), or status \"sent\" when sent.")]
    pub async fn update_draft(
        &self,
        Parameters(p): Parameters<UpdateDraftParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || {
            let body = resolve_mail_body(BodyInputs {
                body: p.body, html_body: p.html_body, body_file: p.body_file,
                html_body_file: p.html_body_file, html: None,
            })?;
            let (body, html_body) = match body {
                Some(MailBody::Text(t)) => (Some(t), None),
                Some(MailBody::Html(h)) => (None, Some(h)),
                None => (None, None),
            };
            client.update_draft(DraftUpdate {
                email_id: p.email_id, subject: p.subject, body, html_body,
                to: p.to, cc: p.cc, bcc: p.bcc, attachments: p.attachments,
                inline_images: p.inline_images, add_categories: p.add_categories,
                remove_categories: p.remove_categories, importance: p.importance,
                send: p.send,
            })
        }).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Delete an email. By default (permanent=false) it moves to Deleted Items and is recoverable. With permanent=true it is hard-deleted like Outlook's shift+delete (moved to Deleted Items, then deleted from there): IRREVERSIBLE, it cannot be recovered from Deleted Items. On Exchange/Microsoft 365 with retention, it may still be held in Recoverable Items per server policy.")]
    pub async fn delete_email(
        &self,
        Parameters(DeleteEmailParams { email_id, permanent }): Parameters<DeleteEmailParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.delete_email(email_id, permanent)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Permanently delete EVERYTHING in the default mailbox's Deleted Items folder (items and subfolders). IRREVERSIBLE: refuses unless confirm=true; only pass confirm=true when the user has explicitly asked to empty Deleted Items. Returns counts of items and folders deleted and failures. On Exchange/Microsoft 365 with retention, items may still be held in Recoverable Items per server policy.")]
    pub async fn empty_deleted_items(
        &self,
        Parameters(EmptyDeletedItemsParams { confirm }): Parameters<EmptyDeletedItemsParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.empty_deleted_items(confirm)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    // ---- Calendar ----

    #[tool(description = "List/search calendar events. Filter by date range, text (subject/location), category, show_as, your response, attendees (+role), meetings-only, all-day; or view another person's shared calendar via calendar_of.")]
    pub async fn list_events(
        &self,
        Parameters(p): Parameters<ListEventsParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let q = EventQuery {
            start_date: p.start_date, end_date: p.end_date, query: p.query,
            category: p.category, show_as: p.show_as, my_response: p.my_response,
            attendees: p.attendees, attendee_role: p.attendee_role,
            meetings_only: p.meetings_only, all_day: p.all_day,
            calendar_of: p.calendar_of,
        };
        let result = run_blocking(move || client.list_events(q)).await?;
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

    #[tool(description = "Create a calendar event. required_attendees/optional_attendees invite two tiers (attendees is a legacy alias merged into required_attendees); any attendee makes it a meeting. categories and show_as (busy status) can be set on creation. body (plain text) or body_file (a local UTF-8 text file path, for large text), not both. recurrence repeats the event (daily/weekly/monthly/yearly, with an interval and an until date or occurrence count). send (default true) controls whether a meeting is actually sent to attendees or just saved for review.")]
    pub async fn create_event(
        &self,
        Parameters(CreateEventParams {
            subject, start, end, body, body_file, location, attendees, required_attendees,
            optional_attendees, all_day, reminder_minutes, categories, show_as, send,
            recurrence,
        }): Parameters<CreateEventParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let body = run_blocking(move || resolve_text_input(body, body_file, "body", "body_file")).await?;
        // `attendees` is a legacy alias for the required tier; merge it in.
        let mut required = required_attendees.unwrap_or_default();
        required.extend(attendees.unwrap_or_default());
        let required_attendees = (!required.is_empty()).then_some(required);
        let recurrence = recurrence.map(|r| RecurrenceInput {
            pattern: r.pattern, interval: r.interval, days_of_week: r.days_of_week,
            day_of_month: r.day_of_month, until: r.until, occurrences: r.occurrences,
        });
        let input = CreateEventInput {
            subject, start, end, body, location, required_attendees, optional_attendees,
            all_day, reminder_minutes, categories, show_as, send, recurrence,
        };
        let result = run_blocking(move || client.create_event(input)).await?;
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

    #[tool(description = "Update an existing calendar event: subject, start/end, location, body (or body_file: a local UTF-8 text file path), show_as, add/remove categories, add/remove attendees, reminder, all_day, recurrence (set/replace) or clear_recurrence (remove it). Adding an attendee converts a personal appointment into a meeting. Recurrence edits apply to the whole series. send_update (default true) notifies attendees if the event is a meeting.")]
    pub async fn update_event(
        &self,
        Parameters(p): Parameters<UpdateEventParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let recurrence = p.recurrence.map(|r| RecurrenceInput {
            pattern: r.pattern, interval: r.interval, days_of_week: r.days_of_week,
            day_of_month: r.day_of_month, until: r.until, occurrences: r.occurrences,
        });
        let (body, body_file) = (p.body, p.body_file);
        let body = run_blocking(move || resolve_text_input(body, body_file, "body", "body_file")).await?;
        let u = EventUpdate {
            event_id: p.event_id, subject: p.subject, start: p.start, end: p.end,
            location: p.location, body, all_day: p.all_day,
            reminder_minutes: p.reminder_minutes, show_as: p.show_as,
            add_categories: p.add_categories, remove_categories: p.remove_categories,
            add_required_attendees: p.add_required_attendees,
            add_optional_attendees: p.add_optional_attendees,
            remove_attendees: p.remove_attendees, send_update: p.send_update,
            recurrence, clear_recurrence: p.clear_recurrence,
        };
        let result = run_blocking(move || client.update_event(u)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Delete/cancel a calendar event (moves it to Deleted Items). If you organize the meeting, send_cancellation (default true) notifies attendees; if false, it's canceled quietly.")]
    pub async fn delete_event(
        &self,
        Parameters(DeleteEventParams { event_id, send_cancellation }): Parameters<DeleteEventParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.delete_event(event_id, send_cancellation)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Check free/busy availability for one or more people over a time window. Returns each person's raw status per time slot (never event details) plus common_free: the windows where everyone is available. treat_as_free (default [\"free\"]) controls which statuses count as available when computing common_free; a person whose address can't be resolved, or who resolves but has no loadable free/busy data, is marked resolved:false with empty slots and doesn't fail the call.")]
    pub async fn check_availability(
        &self,
        Parameters(CheckAvailabilityParams { people, start, end, interval_minutes, treat_as_free }):
            Parameters<CheckAvailabilityParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let input = CheckAvailabilityInput { people, start, end, interval_minutes, treat_as_free };
        let result = run_blocking(move || client.check_availability(input)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    // ---- Attachments ----

    #[tool(description = "List an email's attachments. Each entry has `index` (1-based position), `filename`, `size` (bytes), `type` (\"file\", \"link\", \"item\", \"ole\" or \"unknown\"), `content_id` (the Content-ID an HTML body references as `cid:...`, without `<>`; null if none), `mime_type` (from the attachment, else guessed from the extension; may be null), `hidden`, and `is_inline` (true for inline content an HTML body shows via `cid:` rather than a standalone attachment: it has a `content_id` and is either referenced as `cid:<content_id>` in the HTML body or hidden).")]
    pub async fn list_attachments(
        &self,
        Parameters(ListAttachmentsParams { email_id }): Parameters<ListAttachmentsParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.list_attachments(email_id)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Save an email's attachments to a local directory. Pass attachment_names to save only specific files (case-insensitive). Each result carries the same metadata as list_attachments, including `is_inline` (`index` is the original position even when filtering) plus `saved_to` and `status` (\"saved\"), or `status` \"failed\" with `error`.")]
    pub async fn save_attachments(
        &self,
        Parameters(SaveAttachmentsParams { email_id, save_dir, attachment_names }): Parameters<SaveAttachmentsParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.save_attachments(email_id, save_dir, attachment_names)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Fetch an email's attachment by Content-ID (e.g. an inline image an HTML body references as `cid:...`; see `content_id` from list_attachments) as a base64 data URI. Returns `content_id`, `filename`, `mime_type` (application/octet-stream if unknown), `size` (bytes) and `data_uri`. A `cid:` prefix and surrounding `<>` are accepted; matching is case-insensitive. Limited to 10 MB; use save_attachments for larger files. Optional `context_lines` (max 50) also returns `context`: the last that-many non-empty lines of plain text immediately before the image's first `cid:` reference in the HTML body (tags stripped, entities decoded), useful for knowing what the image shows. `context` is \"\" when the HTML body never references the image, and is omitted when `context_lines` is not given.")]
    pub async fn get_inline_image(
        &self,
        Parameters(GetInlineImageParams { email_id, content_id, context_lines }): Parameters<GetInlineImageParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.get_inline_image(email_id, content_id, context_lines)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    // ---- Tasks ----

    #[tool(description = "List Outlook tasks (default: not-yet-completed only). Filter by category, importance, or a text query matching the subject.")]
    pub async fn list_tasks(
        &self,
        Parameters(p): Parameters<ListTasksParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let q = TaskQuery {
            include_completed: p.include_completed, category: p.category,
            importance: p.importance, query: p.query,
        };
        let result = run_blocking(move || client.list_tasks(q)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Create a new task. body (plain text) or body_file (a local UTF-8 text file path, for large text), not both. reminder_time (ISO datetime) is an absolute reminder time, unlike appointment reminders which are minutes-before-start.")]
    pub async fn create_task(
        &self,
        Parameters(CreateTaskParams {
            subject, body, body_file, due_date, importance, categories, start_date, reminder_time,
        }): Parameters<CreateTaskParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || {
            let body = resolve_text_input(body, body_file, "body", "body_file")?;
            client.create_task(subject, body, due_date, importance, categories, start_date, reminder_time)
        }).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Update an existing task: mark_complete (true=complete, false=reopen), subject, body (or body_file: a local UTF-8 text file path), due_date, start_date, importance, add/remove categories, percent_complete, reminder_time. Combine any of these in one call. Note: mark_complete is applied last and both complete/reopen set percent_complete themselves, so mark_complete:false always resets percent_complete to 0 even if you also pass an explicit percent_complete in the same call — set it in a separate call afterward if you need it to stick.")]
    pub async fn update_task(
        &self,
        Parameters(p): Parameters<UpdateTaskParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || {
            let body = resolve_text_input(p.body, p.body_file, "body", "body_file")?;
            client.update_task(TaskUpdate {
                task_id: p.task_id, mark_complete: p.mark_complete, subject: p.subject,
                body, due_date: p.due_date, start_date: p.start_date,
                importance: p.importance, add_categories: p.add_categories,
                remove_categories: p.remove_categories, percent_complete: p.percent_complete,
                reminder_time: p.reminder_time,
            })
        }).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Delete a task (moves it to Deleted Items).")]
    pub async fn delete_task(
        &self,
        Parameters(DeleteTaskParams { task_id }): Parameters<DeleteTaskParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.delete_task(task_id)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    // ---- Notes ----

    #[tool(description = "List Outlook notes. Filter by category or a text query matching the note's body.")]
    pub async fn list_notes(
        &self,
        Parameters(ListNotesParams { category, query }): Parameters<ListNotesParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let q = NoteQuery { category, query };
        let result = run_blocking(move || client.list_notes(q)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Get the full body of one note by id.")]
    pub async fn get_note(
        &self,
        Parameters(GetNoteParams { note_id }): Parameters<GetNoteParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.get_note(note_id)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Create a new note. Give its text as exactly one of `body` or `body_file` (a local UTF-8 text file path, for large text). Optional categories and color.")]
    pub async fn create_note(
        &self,
        Parameters(CreateNoteParams { body, body_file, categories, color }): Parameters<CreateNoteParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || {
            let body = resolve_text_input(body, body_file, "body", "body_file")?.ok_or_else(|| {
                ToolError::new("create_note needs a body: pass `body` or `body_file`")
            })?;
            client.create_note(body, categories, color)
        }).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Update an existing note: body (or body_file: a local UTF-8 text file path), add/remove categories, color. Combine any of these in one call.")]
    pub async fn update_note(
        &self,
        Parameters(p): Parameters<UpdateNoteParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || {
            let body = resolve_text_input(p.body, p.body_file, "body", "body_file")?;
            client.update_note(NoteUpdate {
                note_id: p.note_id, body, add_categories: p.add_categories,
                remove_categories: p.remove_categories, color: p.color,
            })
        }).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Delete a note (moves it to Deleted Items).")]
    pub async fn delete_note(
        &self,
        Parameters(DeleteNoteParams { note_id }): Parameters<DeleteNoteParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.delete_note(note_id)).await?;
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
