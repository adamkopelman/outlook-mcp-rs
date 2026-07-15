pub mod client;
pub mod com;
pub mod fake;
pub mod types;

use crate::error::ToolError;
use serde_json::Value;
use types::*;

/// All filters for `list_emails`. All optional except `folder`/`count`
/// (which the server fills with defaults). Supplying several ANDs them.
#[derive(Debug, Clone)]
pub struct EmailQuery {
    pub query: Option<String>,
    pub folder: String,
    pub count: i32,
    pub unread_only: bool,
    pub from: Option<String>,
    pub category: Option<String>,
    pub received_after: Option<String>,
    pub received_before: Option<String>,
    pub since_days: Option<i32>,
    pub has_attachments: Option<bool>,
    pub flagged: bool,
    pub high_importance: bool,
}

/// All changes `update_email` can apply to one existing email. Every field
/// except `email_id` is optional; supplying several applies all of them.
/// State changes are applied first and `move_to` last (Move changes the
/// EntryID, so it must come after everything that addresses the item by id).
#[derive(Debug, Clone, Default)]
pub struct EmailUpdate {
    pub email_id: String,
    pub move_to: Option<String>,
    pub mark_read: Option<bool>,
    pub flag: Option<String>,               // "follow_up" | "complete" | "clear"
    pub add_categories: Option<Vec<String>>,
    pub remove_categories: Option<Vec<String>>,
    pub importance: Option<String>,         // "low" | "normal" | "high"
}

/// All filters for `list_events`. Every field is optional; supplying several
/// ANDs them. `start_date`/`end_date` bound the (recurrence-expanded) scan;
/// the rest filter the streamed events client-side. `calendar_of` (an
/// email/name) opens another person's shared calendar instead of your own.
#[derive(Debug, Clone, Default)]
pub struct EventQuery {
    pub start_date: Option<String>,
    pub end_date: Option<String>,
    pub query: Option<String>,                 // text match on subject + location
    pub category: Option<String>,
    pub show_as: Option<String>,               // "free"|"tentative"|"busy"|"out_of_office"|"working_elsewhere"
    pub my_response: Option<String>,           // "organizer"|"accepted"|"declined"|"tentative"|"not_responded"
    pub attendees: Option<Vec<String>>,        // match events where ANY listed person participates
    pub attendee_role: Option<String>,         // "required"|"optional"|"any" (default "any")
    pub meetings_only: bool,
    pub all_day: Option<bool>,
    pub calendar_of: Option<String>,
}

/// All inputs for `create_event`. `required_attendees`/`optional_attendees`
/// are the two invite tiers Outlook shows a meeting organizer; any attendee
/// in either tier makes the item a meeting. `send` (default true in the
/// tool layer) controls whether a meeting is actually sent to attendees or
/// merely saved for later review — see `create_event_status` below for the
/// resulting status string.
#[derive(Debug, Clone)]
pub struct CreateEventInput {
    pub subject: String,
    pub start: String,
    pub end: String,
    pub body: Option<String>,
    pub location: Option<String>,
    pub required_attendees: Option<Vec<String>>,
    pub optional_attendees: Option<Vec<String>>,
    pub all_day: bool,
    pub reminder_minutes: Option<i32>,
    pub categories: Option<Vec<String>>,
    pub show_as: Option<String>,
    pub send: bool,
}

pub trait OutlookClient: Send + Sync {
    fn list_folders(&self) -> Result<Vec<FolderInfo>, ToolError>;
    fn list_emails(&self, q: EmailQuery) -> Result<Vec<EmailSummary>, ToolError>;
    fn get_email(&self, email_id: String, prefer_html: bool)
        -> Result<EmailDetail, ToolError>;
    fn send_email(&self, to: Vec<String>, subject: String, body: String,
        cc: Option<Vec<String>>, bcc: Option<Vec<String>>, html: bool,
        attachments: Option<Vec<String>>) -> Result<Value, ToolError>;
    fn create_draft(&self, to: Vec<String>, subject: String, body: String,
        cc: Option<Vec<String>>, bcc: Option<Vec<String>>, html: bool,
        attachments: Option<Vec<String>>) -> Result<Value, ToolError>;
    fn reply_email(&self, email_id: String, body: String, reply_all: bool,
        html: bool, send: bool, attachments: Option<Vec<String>>)
        -> Result<Value, ToolError>;
    fn update_email(&self, u: EmailUpdate) -> Result<Value, ToolError>;
    fn delete_email(&self, email_id: String) -> Result<Value, ToolError>;

    fn list_events(&self, q: EventQuery) -> Result<Vec<EventSummary>, ToolError>;
    fn get_event(&self, event_id: String) -> Result<EventDetail, ToolError>;
    fn create_event(&self, input: CreateEventInput) -> Result<Value, ToolError>;
    fn respond_to_meeting(&self, event_id: String, response: String,
        comment: Option<String>, send: bool) -> Result<Value, ToolError>;

    fn list_attachments(&self, email_id: String)
        -> Result<Vec<AttachmentInfo>, ToolError>;
    fn save_attachments(&self, email_id: String, save_dir: String,
        attachment_names: Option<Vec<String>>) -> Result<Vec<Value>, ToolError>;

    fn list_tasks(&self, include_completed: bool)
        -> Result<Vec<TaskSummary>, ToolError>;
    fn create_task(&self, subject: String, body: Option<String>,
        due_date: Option<String>, importance: String) -> Result<Value, ToolError>;
    fn complete_task(&self, task_id: String) -> Result<Value, ToolError>;

    fn list_notes(&self) -> Result<Vec<NoteSummary>, ToolError>;
    fn get_note(&self, note_id: String) -> Result<NoteDetail, ToolError>;
    fn create_note(&self, body: String) -> Result<Value, ToolError>;
}
