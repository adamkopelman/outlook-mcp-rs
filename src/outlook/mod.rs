pub mod com;
pub mod fake;
pub mod types;

use crate::error::ToolError;
use serde_json::Value;
use types::*;

pub trait OutlookClient: Send + Sync {
    fn list_folders(&self) -> Result<Vec<FolderInfo>, ToolError>;
    fn list_emails(&self, folder: String, count: i32, unread_only: bool)
        -> Result<Vec<EmailSummary>, ToolError>;
    fn search_emails(&self, query: String, folder: String, count: i32,
        since_days: Option<i32>) -> Result<Vec<EmailSummary>, ToolError>;
    fn get_email(&self, email_id: String, prefer_html: bool)
        -> Result<EmailDetail, ToolError>;
    fn send_email(&self, to: Vec<String>, subject: String, body: String,
        cc: Option<Vec<String>>, bcc: Option<Vec<String>>, html: bool)
        -> Result<Value, ToolError>;
    fn create_draft(&self, to: Vec<String>, subject: String, body: String,
        cc: Option<Vec<String>>, bcc: Option<Vec<String>>, html: bool)
        -> Result<Value, ToolError>;
    fn reply_email(&self, email_id: String, body: String, reply_all: bool,
        html: bool, send: bool) -> Result<Value, ToolError>;
    fn move_email(&self, email_id: String, target_folder: String)
        -> Result<Value, ToolError>;
    fn delete_email(&self, email_id: String) -> Result<Value, ToolError>;

    fn list_events(&self, start_date: Option<String>, end_date: Option<String>)
        -> Result<Vec<EventSummary>, ToolError>;
    fn get_event(&self, event_id: String) -> Result<EventDetail, ToolError>;
    fn create_event(&self, subject: String, start: String, end: String,
        body: Option<String>, location: Option<String>,
        attendees: Option<Vec<String>>, all_day: bool,
        reminder_minutes: Option<i32>) -> Result<Value, ToolError>;
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
