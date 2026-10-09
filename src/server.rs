use std::sync::Arc;

use rmcp::{
    ErrorData as McpError, ServerHandler,
    handler::server::wrapper::Parameters,
    model::{CallToolResult, ContentBlock, ServerCapabilities, ServerInfo},
    tool, tool_handler, tool_router,
};
use serde::Deserialize;

use crate::error::ToolError;
use crate::outlook::filters::{self, normalize_choices};
use crate::outlook::read::{read_options, ReadRequest, ReadTool};
use crate::outlook::{DEFAULT_EMAIL_COUNT, MAX_EVENT_COUNT, MAX_NOTE_COUNT, MAX_TASK_COUNT};
use crate::params::OneOrMany;
use serde_json::{json, Value};
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

/// A batch-read id parameter as `(ids, is_list)`; an empty list is an error.
fn batch_ids(ids: OneOrMany<String>, param: &str) -> Result<(Vec<String>, bool), ToolError> {
    let many = ids.is_many();
    let ids = ids.into_vec();
    if ids.is_empty() {
        return Err(ToolError::new(format!("{param} must be an id or a non-empty list of ids.")));
    }
    Ok((ids, many))
}

/// The shared batch-read result shape. One id given (not a list): exactly
/// that item's object, and its error fails the call (as before batching).
/// A list given: a JSON array in input order, where a failed item is
/// `{"id": <id>, "error": <message>}` and a successful one is
/// `ok_item(id, value)`.
fn batch_content<T: serde::Serialize>(
    ids: &[String],
    many: bool,
    results: Vec<Result<T, ToolError>>,
    ok_item: impl Fn(&str, Value) -> Value,
) -> Result<CallToolResult, McpError> {
    if !many {
        let one = results
            .into_iter()
            .next()
            .unwrap_or_else(|| Err(ToolError::new("no result returned for the requested id")))?;
        return Ok(CallToolResult::success(vec![json_content(&one)?]));
    }
    let items = ids
        .iter()
        .zip(results)
        .map(|(id, result)| match result {
            Ok(item) => serde_json::to_value(item)
                .map(|v| ok_item(id, v))
                .map_err(|e| McpError::internal_error(e.to_string(), None)),
            Err(e) => Ok(json!({"id": id, "error": e.0})),
        })
        .collect::<Result<Vec<_>, _>>()?;
    Ok(CallToolResult::success(vec![json_content(&items)?]))
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
    /// Text search. Terms are ANDed, case-insensitive, any language;
    /// "quoted phrase" keeps words together; * is a wildcard; field:term
    /// limits a term to subject, from, to or body. Unscoped terms search
    /// subject, sender and body.
    #[serde(default)]
    pub query: Option<String>,
    #[serde(default = "default_folder")]
    pub folder: String,
    /// Page size: default 10, max 200.
    #[serde(default = "default_count")]
    pub count: i32,
    /// Matches to skip before this page starts (default 0). To page, call
    /// again with offset += count until fewer than count results come back.
    #[serde(default)]
    pub offset: i32,
    #[serde(default)]
    pub unread_only: bool,
    /// Sender name or address (caseless substring); a list matches any of them.
    #[serde(default)]
    pub from: Option<OneOrMany<String>>,
    /// To/CC recipient name or address (caseless substring); a list matches any of them.
    #[serde(default)]
    pub to: Option<OneOrMany<String>>,
    /// Category name; a list matches any of them.
    #[serde(default)]
    pub category: Option<OneOrMany<String>>,
    /// Received at or after this date, e.g. '2026-06-01', 'start_of_week', '-14d'.
    #[serde(default)]
    pub received_after: Option<String>,
    /// Received at or before this date; a bare date like '2026-06-30'
    /// includes that whole day, 'today' means before today.
    #[serde(default)]
    pub received_before: Option<String>,
    /// Deprecated: use received_after: "-<N>d" instead.
    #[serde(default)]
    pub since_days: Option<i32>,
    #[serde(default)]
    pub has_attachments: Option<bool>,
    /// Item type as get_email reports it: "email" | "meeting" | "bounce" |
    /// "read_receipt" | "other"; a list matches any of them.
    #[serde(default)]
    pub item_type: Option<OneOrMany<String>>,
    /// "low" | "normal" | "high"; a list matches any of them.
    #[serde(default)]
    pub importance: Option<OneOrMany<String>>,
    /// Flag state, same values as update_email's flag: "follow_up" |
    /// "complete" | "clear" (not flagged); a list matches any of them.
    #[serde(default)]
    pub flag: Option<OneOrMany<String>>,
    /// Deprecated: use flag: ["follow_up", "complete"] instead.
    #[serde(default)]
    pub flagged: bool,
    /// Deprecated: use importance: "high" instead.
    #[serde(default)]
    pub high_importance: bool,
}
fn default_folder() -> String { "inbox".to_string() }
fn default_count() -> i32 { DEFAULT_EMAIL_COUNT }

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct GetEmailParams {
    /// One email id, or a list of ids to read in one call (alias `email_ids`).
    /// A list returns a list in the same order, where an id that fails
    /// becomes {"id": ..., "error": ...}.
    #[serde(alias = "email_ids")]
    pub email_id: OneOrMany<String>,
    /// Optional fields to return: any of "body", "html_body", "attachments",
    /// "meeting". Omitted = ["body", "attachments", "meeting"]; [] = metadata
    /// only (headers, item_type, is_meeting).
    #[serde(default)]
    pub include: Option<Vec<String>>,
    /// "text" (default) or "html". "html" also returns html_body (the same
    /// as adding "html_body" to include).
    #[serde(default)]
    pub body_format: Option<String>,
    /// Deprecated: use body_format instead. true = body_format "html".
    #[serde(default)]
    pub prefer_html: Option<bool>,
    /// Cut body/html_body at this many characters. Default 100,000;
    /// clamped to 1,000..=5,000,000. Does not apply to files (output_dir).
    #[serde(default)]
    pub max_body_chars: Option<u32>,
    /// Local directory (created if missing). body / html_body are written
    /// there in full and returned as body_file / html_body_file (absolute
    /// paths) instead of inline text.
    #[serde(default)]
    pub output_dir: Option<String>,
    /// Replace each cid: reference in html_body with that inline image's
    /// data: URI so the HTML is self-contained (implies html_body).
    #[serde(default)]
    pub resolve_inline_images: bool,
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
    /// Events starting at or after this date (default today 00:00), e.g.
    /// '2026-06-10', 'start_of_week', '-1d'. `start_date` is a deprecated alias.
    #[serde(default, alias = "start_date")]
    pub start_after: Option<String>,
    /// Events starting at or before this date (default 7 days after
    /// start_after); a bare date includes that whole day. `end_date` is a
    /// deprecated alias.
    #[serde(default, alias = "end_date")]
    pub start_before: Option<String>,
    /// Text search with the shared list query syntax; field:term scopes are
    /// subject, location, organizer and attendees. Unscoped terms search
    /// subject and location.
    #[serde(default)]
    pub query: Option<String>,
    /// Category name; a list matches any of them.
    #[serde(default)]
    pub category: Option<OneOrMany<String>>,
    /// "free" | "tentative" | "busy" | "out_of_office" | "working_elsewhere"; a list matches any.
    #[serde(default)]
    pub show_as: Option<OneOrMany<String>>,
    /// This mailbox's response: "organizer" | "accepted" | "declined" |
    /// "tentative" | "not_responded" | "none"; a list matches any.
    #[serde(default)]
    pub my_response: Option<OneOrMany<String>>,
    /// Names/emails; match events where ANY listed person participates.
    #[serde(default)]
    pub attendees: Option<OneOrMany<String>>,
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
    /// Page size: default and max 250.
    #[serde(default = "default_event_count")]
    pub count: i32,
    /// Matches to skip before this page starts (default 0).
    #[serde(default)]
    pub offset: i32,
}
fn default_event_count() -> i32 { MAX_EVENT_COUNT }

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct GetEventParams {
    /// One event id, or a list of ids (alias `event_ids`); a list returns a
    /// list in the same order, where a failing id becomes {"id", "error"}.
    #[serde(alias = "event_ids")]
    pub event_id: OneOrMany<String>,
    /// Optional fields to return: "body". Omitted = ["body"]; [] = no body.
    #[serde(default)]
    pub include: Option<Vec<String>>,
    /// "text" (default). Event bodies are plain text only.
    #[serde(default)]
    pub body_format: Option<String>,
    /// Cut the body at this many characters. Default 100,000; clamped to
    /// 1,000..=5,000,000. Does not apply to files (output_dir).
    #[serde(default)]
    pub max_body_chars: Option<u32>,
    /// Local directory (created if missing). The body is written there in
    /// full and returned as body_file (absolute path) instead of body.
    #[serde(default)]
    pub output_dir: Option<String>,
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
    /// End date (any date form, e.g. '2026-12-31' or 'end_of_year'). At most one of
    /// `until`/`occurrences`; neither = no end date.
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
    /// Independent of `permanent`: the cancellation is sent before the delete.
    #[serde(default = "default_true")]
    pub send_cancellation: bool,
    /// true = hard delete (like shift+delete); IRREVERSIBLE. Default false
    /// moves the event to Deleted Items.
    #[serde(default)]
    pub permanent: bool,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CheckAvailabilityParams {
    pub people: Vec<String>,
    /// Window start (any date form, e.g. '2026-06-10T09:00' or 'tomorrow+9h').
    pub start: String,
    /// Window end (any date form).
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
    /// One email id, or a list of ids (alias `email_ids`). A list returns a
    /// list in the same order of {"id", "attachments": [...]}, or
    /// {"id", "error"} for an id that fails.
    #[serde(alias = "email_ids")]
    pub email_id: OneOrMany<String>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct SaveAttachmentsParams {
    pub email_id: String,
    pub save_dir: String,
    #[serde(default)]
    pub attachment_names: Option<Vec<String>>,
    /// false = save only regular attachments (skip inline images an HTML
    /// body shows via `cid:`); true = only inline images. Omit to save all.
    #[serde(default)]
    pub inline: Option<bool>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct GetInlineImageParams {
    pub email_id: String,
    /// The attachment's Content-ID (`content_id` from list_attachments, or the
    /// `cid:...` reference from the HTML body). A `cid:` prefix and `<>` are
    /// accepted. Pass this or `content_ids`, not both.
    #[serde(default)]
    pub content_id: Option<String>,
    /// Several Content-IDs of the same email; returns a list in the same
    /// order, where a failing one becomes {"id": <content id>, "error"}.
    #[serde(default)]
    pub content_ids: Option<Vec<String>>,
    /// Also return `context`: up to this many lines (max 50) of plain text
    /// immediately before the image's first `cid:` reference in the HTML body.
    #[serde(default)]
    pub context_lines: Option<u32>,
    /// Local directory (created if missing). Each image's decoded bytes are
    /// written there and returned as data_file (absolute path) instead of
    /// data_uri.
    #[serde(default)]
    pub output_dir: Option<String>,
}

// ---- Tasks ----

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ListTasksParams {
    #[serde(default)]
    pub include_completed: bool,
    /// Category name; a list matches any of them.
    #[serde(default)]
    pub category: Option<OneOrMany<String>>,
    /// "low" | "normal" | "high"; a list matches any of them.
    #[serde(default)]
    pub importance: Option<OneOrMany<String>>,
    /// Text search with the shared list query syntax; field:term scopes are
    /// subject and body. Unscoped terms search both.
    #[serde(default)]
    pub query: Option<String>,
    /// Due at or after this date, e.g. 'today'. Tasks without a due date
    /// never match a due range.
    #[serde(default)]
    pub due_after: Option<String>,
    /// Due at or before this date, e.g. 'end_of_week'; a bare date includes that whole day.
    #[serde(default)]
    pub due_before: Option<String>,
    /// Page size: default and max 500.
    #[serde(default = "default_task_count")]
    pub count: i32,
    /// Matches to skip before this page starts (default 0).
    #[serde(default)]
    pub offset: i32,
}
fn default_task_count() -> i32 { MAX_TASK_COUNT }

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
pub struct GetTaskParams {
    /// One task id, or a list of ids (alias `task_ids`); a list returns a
    /// list in the same order, where a failing id becomes {"id", "error"}.
    #[serde(alias = "task_ids")]
    pub task_id: OneOrMany<String>,
    /// Optional fields to return: "body". Omitted = ["body"]; [] = no body.
    #[serde(default)]
    pub include: Option<Vec<String>>,
    /// "text" (default). Task bodies are plain text only.
    #[serde(default)]
    pub body_format: Option<String>,
    /// Cut the body at this many characters. Default 100,000; clamped to
    /// 1,000..=5,000,000. Does not apply to files (output_dir).
    #[serde(default)]
    pub max_body_chars: Option<u32>,
    /// Local directory (created if missing). The body is written there in
    /// full and returned as body_file (absolute path) instead of body.
    #[serde(default)]
    pub output_dir: Option<String>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct DeleteTaskParams {
    pub task_id: String,
    /// true = hard delete (like shift+delete); IRREVERSIBLE. Default false
    /// moves the task to Deleted Items.
    #[serde(default)]
    pub permanent: bool,
}

// ---- Notes ----

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ListNotesParams {
    /// Category name; a list matches any of them.
    #[serde(default)]
    pub category: Option<OneOrMany<String>>,
    /// Text search with the shared list query syntax; field:term scopes are
    /// subject (the first line) and body. Unscoped terms search the body.
    #[serde(default)]
    pub query: Option<String>,
    /// Created at or after this date, e.g. '-30d'.
    #[serde(default)]
    pub created_after: Option<String>,
    /// Created at or before this date; a bare date includes that whole day.
    #[serde(default)]
    pub created_before: Option<String>,
    /// Page size: default and max 500.
    #[serde(default = "default_note_count")]
    pub count: i32,
    /// Matches to skip before this page starts (default 0).
    #[serde(default)]
    pub offset: i32,
}
fn default_note_count() -> i32 { MAX_NOTE_COUNT }

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct GetNoteParams {
    /// One note id, or a list of ids (alias `note_ids`); a list returns a
    /// list in the same order, where a failing id becomes {"id", "error"}.
    #[serde(alias = "note_ids")]
    pub note_id: OneOrMany<String>,
    /// Optional fields to return: "body". Omitted = ["body"]; [] = no body.
    #[serde(default)]
    pub include: Option<Vec<String>>,
    /// "text" (default). Note bodies are plain text only.
    #[serde(default)]
    pub body_format: Option<String>,
    /// Cut the body at this many characters. Default 100,000; clamped to
    /// 1,000..=5,000,000. Does not apply to files (output_dir).
    #[serde(default)]
    pub max_body_chars: Option<u32>,
    /// Local directory (created if missing). The body is written there in
    /// full and returned as body_file (absolute path) instead of body.
    #[serde(default)]
    pub output_dir: Option<String>,
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
    /// true = hard delete (like shift+delete); IRREVERSIBLE. Default false
    /// moves the note to Deleted Items.
    #[serde(default)]
    pub permanent: bool,
}

/// An optional one-or-many tool parameter as a plain list (empty = not given).
fn list(v: Option<OneOrMany<String>>) -> Vec<String> {
    v.map(OneOrMany::into_vec).unwrap_or_default()
}

/// A free-text one-or-many filter as a trimmed list without empty values.
fn list_values(v: Option<OneOrMany<String>>) -> Vec<String> {
    filters::clean_values(list(v))
}

/// Maps `list_emails`' parameters onto an [`EmailQuery`]: list-valued
/// filters become lists, enum values are validated, and the deprecated
/// `since_days` / `flagged` / `high_importance` are folded into
/// `received_after` / `flag` / `importance`, refusing contradictions.
pub fn email_query(p: ListEmailsParams) -> Result<EmailQuery, ToolError> {
    let mut received_after = p.received_after.filter(|s| !s.trim().is_empty());
    if let Some(days) = p.since_days.filter(|d| *d != 0) {
        if received_after.is_some() {
            return Err(ToolError::new(
                "pass either `received_after` or `since_days`, not both (since_days is deprecated: use received_after: \"-14d\")",
            ));
        }
        received_after = Some(format!("{:+}d", -i64::from(days)));
    }
    let mut importance = normalize_choices(list(p.importance), "importance", filters::IMPORTANCES)?;
    if p.high_importance {
        if !importance.is_empty() && importance != ["high"] {
            return Err(ToolError::new(
                "pass either `importance` or `high_importance`, not both (high_importance is deprecated: use importance: \"high\")",
            ));
        }
        importance = vec!["high".to_string()];
    }
    let mut flag = normalize_choices(list(p.flag), "flag", filters::FLAGS)?;
    if p.flagged {
        if flag.iter().any(|f| f == "clear") {
            return Err(ToolError::new(
                "pass either `flag` or `flagged`, not both (flagged is deprecated: use flag: [\"follow_up\", \"complete\"])",
            ));
        }
        if flag.is_empty() {
            flag = vec!["follow_up".to_string(), "complete".to_string()];
        }
    }
    Ok(EmailQuery {
        query: p.query, folder: p.folder, count: p.count, offset: p.offset,
        unread_only: p.unread_only, from: list_values(p.from), to: list_values(p.to),
        category: list_values(p.category), received_after,
        received_before: p.received_before, has_attachments: p.has_attachments,
        item_type: normalize_choices(list(p.item_type), "item_type", filters::ITEM_TYPES)?,
        importance, flag,
    })
}

#[tool_router]
impl OutlookMcpServer {
    #[tool(description = "List Outlook mail folders (name, path, item counts).")]
    pub async fn list_folders(&self) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.list_folders()).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Find emails in a folder (newest first). Filters (all optional, ANDed): query; from / to / category (one value or a list, matching any); item_type (as get_email reports it, e.g. [\"email\"] to skip meeting invites and bounces); importance; flag (follow_up/complete/clear, as in update_email); unread_only; has_attachments; received_after / received_before. `from` matches the sender name or address; `to` matches any To/CC recipient name or address (case-insensitive substring). query syntax (shared by all list_* tools): space-separated terms that must all match, case-insensitive, any language; \"quoted phrase\" keeps words together; * is a wildcard (status*report); field:term limits a term to subject, from, to or body. Dates accept ISO ('2026-06-10'), keywords ('today', 'start_of_week') or offsets ('-14d'). Paging: count (default 10, max 200) and offset; to page, call again with offset += count until fewer than count results come back. Deprecated: since_days (use received_after: \"-14d\"), flagged (use flag), high_importance (use importance).")]
    pub async fn list_emails(
        &self,
        Parameters(p): Parameters<ListEmailsParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let q = email_query(p)?;
        let result = run_blocking(move || client.list_emails(q)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Get one email by id, or several (email_id as a list: the result is a list in the same order, and an id that fails becomes {\"id\", \"error\"} instead of failing the call). Returns the headers, item_type, is_meeting, plus the optional fields picked by `include` (any of \"body\", \"html_body\", \"attachments\" (file names), \"meeting\"; default [\"body\", \"attachments\", \"meeting\"]; [] = metadata only). body_format \"html\" (the deprecated prefer_html=true) also returns html_body. Each returned body reports *_truncated and *_length (full original length in characters): body_truncated/body_length, html_truncated/html_length. Inline bodies are cut at max_body_chars characters (default 100,000; allowed 1,000-5,000,000); if a *_truncated flag is true the cut may fall mid-tag or mid-image-data: call again with max_body_chars of at least the reported length, or use output_dir. output_dir writes each body in full to a file in that directory (created if missing) and returns body_file / html_body_file (absolute paths) instead of the text; re-reading the same email overwrites the same files. resolve_inline_images=true (implies html_body) replaces every cid: reference in html_body with the referenced attachment's base64 data: URI so the HTML is self-contained; Content-ID matching is case-insensitive, references with no matching attachment, or whose image is over 10 MB or unreadable, are left as cid: and listed in inline_images_unresolved (inline_images_resolved counts the replaced ones). Resolved HTML is usually large: combine it with output_dir.")]
    pub async fn get_email(
        &self,
        Parameters(p): Parameters<GetEmailParams>,
    ) -> Result<CallToolResult, McpError> {
        let (ids, many) = batch_ids(p.email_id, "email_id")?;
        let opts = read_options(ReadTool::Email, ReadRequest {
            include: p.include, body_format: p.body_format, prefer_html: p.prefer_html,
            resolve_inline_images: p.resolve_inline_images, max_body_chars: p.max_body_chars,
            output_dir: p.output_dir,
        })?;
        let client = self.client.clone();
        let batch = ids.clone();
        let results = run_blocking(move || client.get_email(batch, &opts)).await?;
        batch_content(&ids, many, results, |_, v| v)
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

    #[tool(description = "List/search calendar events (sorted by start). start_after / start_before bound the event start (default: today 00:00 to 7 days later; start_date / end_date are deprecated aliases). Filters (all optional, ANDed): query (shared list query syntax; field scopes subject, location, organizer, attendees; unscoped terms search subject and location); category, show_as, my_response, attendees (one value or a list, matching any) with attendee_role; meetings_only; all_day. calendar_of views another person's shared calendar. Paging: count (default and max 250) and offset.")]
    pub async fn list_events(
        &self,
        Parameters(p): Parameters<ListEventsParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let q = EventQuery {
            start_after: p.start_after, start_before: p.start_before, query: p.query,
            category: list_values(p.category),
            show_as: normalize_choices(list(p.show_as), "show_as", filters::SHOW_AS)?,
            my_response: normalize_choices(list(p.my_response), "my_response", filters::MY_RESPONSES)?,
            attendees: list_values(p.attendees), attendee_role: p.attendee_role,
            meetings_only: p.meetings_only, all_day: p.all_day,
            calendar_of: p.calendar_of, count: p.count, offset: p.offset,
        };
        let result = run_blocking(move || client.list_events(q)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Get the full details of one calendar event by id, or several (event_id as a list: the result is a list in the same order, and an id that fails becomes {\"id\", \"error\"}). `include` picks the optional fields: \"body\" (default [\"body\"]; [] = details without the body). The body reports body_truncated and body_length (full length in characters) and is cut at max_body_chars (default 100,000; allowed 1,000-5,000,000). output_dir writes the full body to a file there and returns body_file (absolute path) instead. body_format accepts only \"text\".")]
    pub async fn get_event(
        &self,
        Parameters(p): Parameters<GetEventParams>,
    ) -> Result<CallToolResult, McpError> {
        let (ids, many) = batch_ids(p.event_id, "event_id")?;
        let opts = read_options(ReadTool::Event, ReadRequest {
            include: p.include, body_format: p.body_format, max_body_chars: p.max_body_chars,
            output_dir: p.output_dir, ..ReadRequest::default()
        })?;
        let client = self.client.clone();
        let batch = ids.clone();
        let results = run_blocking(move || client.get_event(batch, &opts)).await?;
        batch_content(&ids, many, results, |_, v| v)
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

    #[tool(description = "Delete/cancel a calendar event. If you organize the meeting, send_cancellation (default true) notifies attendees; if false, it's canceled quietly; this is independent of permanent (the cancellation is sent first). For a recurring event, the id (including an occurrence's id from list_events) names the whole series, so the whole series is deleted. By default (permanent=false) it moves to Deleted Items and is recoverable. With permanent=true it is hard-deleted like Outlook's shift+delete (moved to Deleted Items, then deleted from there): IRREVERSIBLE, it cannot be recovered from Deleted Items. On Exchange/Microsoft 365 with retention, it may still be held in Recoverable Items per server policy.")]
    pub async fn delete_event(
        &self,
        Parameters(DeleteEventParams { event_id, send_cancellation, permanent }):
            Parameters<DeleteEventParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result =
            run_blocking(move || client.delete_event(event_id, send_cancellation, permanent)).await?;
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

    #[tool(description = "List an email's attachments. Each entry has `index` (1-based position), `filename`, `size` (bytes), `type` (\"file\", \"link\", \"item\", \"ole\" or \"unknown\"), `content_id` (the Content-ID an HTML body references as `cid:...`, without `<>`; null if none), `mime_type` (from the attachment, else guessed from the extension; may be null), `hidden`, and `is_inline` (true for inline content an HTML body shows via `cid:` rather than a standalone attachment: it has a `content_id` and is either referenced as `cid:<content_id>` in the HTML body or hidden). email_id may be a list: the result is then a list in the same order of {\"id\", \"attachments\": [...]}, or {\"id\", \"error\"} for an id that fails.")]
    pub async fn list_attachments(
        &self,
        Parameters(ListAttachmentsParams { email_id }): Parameters<ListAttachmentsParams>,
    ) -> Result<CallToolResult, McpError> {
        let (ids, many) = batch_ids(email_id, "email_id")?;
        let client = self.client.clone();
        let batch = ids.clone();
        let results = run_blocking(move || client.list_attachments(batch)).await?;
        batch_content(&ids, many, results, |id, v| json!({"id": id, "attachments": v}))
    }

    #[tool(description = "Save an email's attachments to a local directory. Pass attachment_names to save only specific files (case-insensitive). Pass `inline: false` to save only regular attachments and skip inline images (`is_inline`), e.g. when re-sending an email whose HTML already carries its images (get_email with resolve_inline_images) so they aren't attached twice; `inline: true` saves only the inline images. Each result carries the same metadata as list_attachments, including `is_inline` (`index` is the original position even when filtering) plus `saved_to` and `status` (\"saved\"), or `status` \"failed\" with `error`.")]
    pub async fn save_attachments(
        &self,
        Parameters(SaveAttachmentsParams { email_id, save_dir, attachment_names, inline }): Parameters<SaveAttachmentsParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.save_attachments(email_id, save_dir, attachment_names, inline)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Fetch an email's attachment by Content-ID (e.g. an inline image an HTML body references as `cid:...`; see `content_id` from list_attachments) as a base64 data URI. Returns `content_id`, `filename`, `mime_type` (application/octet-stream if unknown), `size` (bytes) and `data_uri`. A `cid:` prefix and surrounding `<>` are accepted; matching is case-insensitive. Limited to 10 MB; use save_attachments for larger files. Pass `content_ids` (a list) instead of `content_id` to fetch several images of the same email in one call: the result is a list in the same order, where a Content-ID that fails becomes {\"id\": <content id>, \"error\"}. output_dir writes each image's decoded bytes to a file there (created if missing) and returns `data_file` (absolute path) instead of `data_uri`. Optional `context_lines` (max 50) also returns `context`: the last that-many non-empty lines of plain text immediately before the image's first `cid:` reference in the HTML body (tags stripped, entities decoded), useful for knowing what the image shows. `context` is \"\" when the HTML body never references the image, and is omitted when `context_lines` is not given. To get an email's HTML with every image already inlined, use get_email with resolve_inline_images.")]
    pub async fn get_inline_image(
        &self,
        Parameters(GetInlineImageParams { email_id, content_id, content_ids, context_lines, output_dir }):
            Parameters<GetInlineImageParams>,
    ) -> Result<CallToolResult, McpError> {
        let (ids, many) = match (content_id, content_ids) {
            (Some(_), Some(_)) => {
                return Err(ToolError::new("pass either `content_id` or `content_ids`, not both").into())
            }
            (Some(one), None) => (vec![one], false),
            (None, Some(list)) if !list.is_empty() => (list, true),
            (None, Some(_)) => return Err(ToolError::new("content_ids must not be empty.").into()),
            (None, None) => {
                return Err(ToolError::new(
                    "pass `content_id` (one Content-ID) or `content_ids` (a list).",
                ).into())
            }
        };
        let client = self.client.clone();
        let batch = ids.clone();
        let results = run_blocking(move || {
            client.get_inline_image(email_id, batch, context_lines, output_dir)
        }).await?;
        batch_content(&ids, many, results, |_, v| v)
    }

    // ---- Tasks ----

    #[tool(description = "List Outlook tasks (default: not-yet-completed only). Filters (all optional, ANDed): category and importance (one value or a list, matching any); due_after / due_before (any date form, e.g. 'end_of_week'); query (shared list query syntax; field scopes subject and body; unscoped terms search both). Paging: count (default and max 500) and offset.")]
    pub async fn list_tasks(
        &self,
        Parameters(p): Parameters<ListTasksParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let q = TaskQuery {
            include_completed: p.include_completed, category: list_values(p.category),
            importance: normalize_choices(list(p.importance), "importance", filters::IMPORTANCES)?,
            query: p.query, due_after: p.due_after, due_before: p.due_before,
            count: p.count, offset: p.offset,
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

    #[tool(description = "Get the full details of one task by id, or several (task_id as a list: the result is a list in the same order, and an id that fails becomes {\"id\", \"error\"}). Returns the list_tasks fields (subject, due_date, complete, status, importance, categories) plus start_date, date_completed, percent_complete, reminder_set, reminder_time, created, modified (unset dates are null) and the optional body. `include` picks the optional fields: \"body\" (default [\"body\"]; [] = details without the body). The body reports body_truncated and body_length (full length in characters) and is cut at max_body_chars (default 100,000; allowed 1,000-5,000,000). output_dir writes the full body to a file there and returns body_file (absolute path) instead. body_format accepts only \"text\".")]
    pub async fn get_task(
        &self,
        Parameters(p): Parameters<GetTaskParams>,
    ) -> Result<CallToolResult, McpError> {
        let (ids, many) = batch_ids(p.task_id, "task_id")?;
        let opts = read_options(ReadTool::Task, ReadRequest {
            include: p.include, body_format: p.body_format, max_body_chars: p.max_body_chars,
            output_dir: p.output_dir, ..ReadRequest::default()
        })?;
        let client = self.client.clone();
        let batch = ids.clone();
        let results = run_blocking(move || client.get_task(batch, &opts)).await?;
        batch_content(&ids, many, results, |_, v| v)
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

    #[tool(description = "Delete a task. By default (permanent=false) it moves to Deleted Items and is recoverable. With permanent=true it is hard-deleted like Outlook's shift+delete (moved to Deleted Items, then deleted from there): IRREVERSIBLE, it cannot be recovered from Deleted Items. On Exchange/Microsoft 365 with retention, it may still be held in Recoverable Items per server policy.")]
    pub async fn delete_task(
        &self,
        Parameters(DeleteTaskParams { task_id, permanent }): Parameters<DeleteTaskParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.delete_task(task_id, permanent)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    // ---- Notes ----

    #[tool(description = "List Outlook notes. Filters (all optional, ANDed): category (one value or a list, matching any); created_after / created_before (any date form); query (shared list query syntax; field scopes subject, i.e. the first line, and body; unscoped terms search the body). Paging: count (default and max 500) and offset.")]
    pub async fn list_notes(
        &self,
        Parameters(p): Parameters<ListNotesParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let q = NoteQuery {
            category: list_values(p.category), query: p.query,
            created_after: p.created_after, created_before: p.created_before,
            count: p.count, offset: p.offset,
        };
        let result = run_blocking(move || client.list_notes(q)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Get the full body of one note by id, or several (note_id as a list: the result is a list in the same order, and an id that fails becomes {\"id\", \"error\"}). `include` picks the optional fields: \"body\" (default [\"body\"]; [] = subject, dates and categories only). The body reports body_truncated and body_length (full length in characters) and is cut at max_body_chars (default 100,000; allowed 1,000-5,000,000). output_dir writes the full body to a file there and returns body_file (absolute path) instead. body_format accepts only \"text\".")]
    pub async fn get_note(
        &self,
        Parameters(p): Parameters<GetNoteParams>,
    ) -> Result<CallToolResult, McpError> {
        let (ids, many) = batch_ids(p.note_id, "note_id")?;
        let opts = read_options(ReadTool::Note, ReadRequest {
            include: p.include, body_format: p.body_format, max_body_chars: p.max_body_chars,
            output_dir: p.output_dir, ..ReadRequest::default()
        })?;
        let client = self.client.clone();
        let batch = ids.clone();
        let results = run_blocking(move || client.get_note(batch, &opts)).await?;
        batch_content(&ids, many, results, |_, v| v)
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

    #[tool(description = "Delete a note. By default (permanent=false) it moves to Deleted Items and is recoverable. With permanent=true it is hard-deleted like Outlook's shift+delete (moved to Deleted Items, then deleted from there): IRREVERSIBLE, it cannot be recovered from Deleted Items. On Exchange/Microsoft 365 with retention, it may still be held in Recoverable Items per server policy.")]
    pub async fn delete_note(
        &self,
        Parameters(DeleteNoteParams { note_id, permanent }): Parameters<DeleteNoteParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.delete_note(note_id, permanent)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
}

#[tool_handler]
impl ServerHandler for OutlookMcpServer {
    fn get_info(&self) -> ServerInfo {
        ServerInfo::new(ServerCapabilities::builder().enable_tools().build())
            .with_instructions(concat!(
                "Controls Microsoft Outlook desktop (email, calendar, tasks, notes) via COM. ",
                "Every date parameter on every tool accepts the same grammar, in local time: ",
                "an ISO date or datetime ('2026-06-10', '2026-06-10T14:30'); a keyword (now, ",
                "today, yesterday, tomorrow, start_of_week, end_of_week, start_of_month, ",
                "end_of_month, start_of_year, end_of_year; weeks start on the first day of the week in ",
                "the Windows user's regional settings (e.g. Monday, Sunday or Saturday); today/yesterday/",
                "tomorrow/start_of_* are midnight, end_of_* is 23:59:59 on the last day) optionally ",
                "followed by signed offsets ('today-1d', 'start_of_week-1w'); or offsets from now ",
                "('-14d', '+3h', '-2w'; units m=minutes, h=hours, d=days, w=weeks, mo=months, y=years).",
            ))
    }
}
