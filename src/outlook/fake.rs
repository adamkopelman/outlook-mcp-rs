use std::sync::Mutex;

use serde_json::{json, Value};

use crate::error::ToolError;
use super::types::*;
use super::{CreateEventInput, EmailQuery, EmailUpdate, EventQuery, OutlookClient};

pub const EMAIL_ID: &str = "entry-1|store-1";
pub const EVENT_ID: &str = "entry-2|store-1";
pub const TASK_ID: &str = "entry-3|store-1";
pub const NOTE_ID: &str = "entry-4|store-1";

/// In-memory stand-in for COM Outlook; records every call. Mirrors
/// `tests/conftest.py::FakeOutlookClient` in the Python project.
pub struct FakeOutlookClient {
    calls: Mutex<Vec<(String, Value)>>,
    fail_with: Mutex<Option<String>>,
}

impl FakeOutlookClient {
    pub fn new() -> Self {
        Self { calls: Mutex::new(Vec::new()), fail_with: Mutex::new(None) }
    }

    pub fn calls(&self) -> Vec<(String, Value)> {
        self.calls.lock().unwrap().clone()
    }

    pub fn set_fail_with(&self, msg: impl Into<String>) {
        *self.fail_with.lock().unwrap() = Some(msg.into());
    }

    pub fn clear_fail_with(&self) {
        *self.fail_with.lock().unwrap() = None;
    }

    fn record(&self, name: &str, args: Value) -> Result<(), ToolError> {
        if let Some(msg) = self.fail_with.lock().unwrap().clone() {
            return Err(ToolError::new(msg));
        }
        self.calls.lock().unwrap().push((name.to_string(), args));
        Ok(())
    }
}

impl OutlookClient for FakeOutlookClient {
    fn list_folders(&self) -> Result<Vec<FolderInfo>, ToolError> {
        self.record("list_folders", json!({}))?;
        Ok(vec![FolderInfo {
            name: "Inbox".into(), path: "Inbox".into(), items: 2, unread: 1,
        }])
    }

    fn list_emails(&self, q: EmailQuery) -> Result<Vec<EmailSummary>, ToolError> {
        self.record("list_emails", json!({
            "query": q.query, "folder": q.folder, "count": q.count,
            "unread_only": q.unread_only, "from": q.from, "category": q.category,
            "received_after": q.received_after, "received_before": q.received_before,
            "since_days": q.since_days, "has_attachments": q.has_attachments,
            "flagged": q.flagged, "high_importance": q.high_importance,
        }))?;
        Ok(vec![EmailSummary {
            id: EMAIL_ID.into(), subject: "Hello".into(), sender: "Ada".into(),
            sender_email: "".into(), to: "".into(), received: None,
            unread: true, has_attachments: false,
            categories: vec!["Work".to_string()],
        }])
    }

    fn get_email(&self, email_id: String, prefer_html: bool)
        -> Result<EmailDetail, ToolError> {
        self.record("get_email", json!({"email_id": email_id, "prefer_html": prefer_html}))?;
        Ok(EmailDetail {
            summary: EmailSummary {
                id: email_id, subject: "Hello".into(), sender: "".into(),
                sender_email: "".into(), to: "".into(), received: None,
                unread: false, has_attachments: false, categories: vec![],
            },
            cc: "".into(), bcc: "".into(), body: "Hi there".into(),
            html_body: None, attachments: vec![],
            item_type: "email".to_string(),
            is_meeting: false,
            meeting: None,
        })
    }

    fn send_email(&self, to: Vec<String>, subject: String, body: String,
        cc: Option<Vec<String>>, bcc: Option<Vec<String>>, html: bool,
        attachments: Option<Vec<String>>) -> Result<Value, ToolError> {
        self.record("send_email",
            json!({"to": to, "subject": subject, "body": body, "cc": cc, "bcc": bcc, "html": html, "attachments": attachments}))?;
        Ok(json!({"status": "sent", "to": to.join("; "), "subject": subject}))
    }

    fn create_draft(&self, to: Vec<String>, subject: String, body: String,
        cc: Option<Vec<String>>, bcc: Option<Vec<String>>, html: bool,
        attachments: Option<Vec<String>>) -> Result<Value, ToolError> {
        self.record("create_draft",
            json!({"to": to, "subject": subject, "body": body, "cc": cc, "bcc": bcc, "html": html, "attachments": attachments}))?;
        Ok(json!({"status": "draft_saved", "id": EMAIL_ID, "subject": subject}))
    }

    fn reply_email(&self, email_id: String, body: String, reply_all: bool,
        html: bool, send: bool, attachments: Option<Vec<String>>) -> Result<Value, ToolError> {
        self.record("reply_email",
            json!({"email_id": email_id, "body": body, "reply_all": reply_all, "html": html, "send": send, "attachments": attachments}))?;
        Ok(json!({"status": if send { "sent" } else { "draft_saved" }}))
    }

    fn update_email(&self, u: EmailUpdate) -> Result<Value, ToolError> {
        self.record("update_email", json!({
            "email_id": u.email_id, "move_to": u.move_to, "mark_read": u.mark_read,
            "flag": u.flag, "add_categories": u.add_categories,
            "remove_categories": u.remove_categories, "importance": u.importance,
        }))?;
        // Mirror the real client's `changed` ordering: state changes first, move last.
        let mut changed: Vec<&str> = Vec::new();
        if u.mark_read.is_some() { changed.push("mark_read"); }
        if u.flag.is_some() { changed.push("flag"); }
        if u.add_categories.is_some() { changed.push("add_categories"); }
        if u.remove_categories.is_some() { changed.push("remove_categories"); }
        if u.importance.is_some() { changed.push("importance"); }
        // Move changes the EntryID; simulate a new id only when we moved.
        let id = if u.move_to.is_some() {
            changed.push("move_to");
            "new-entry|store-1".to_string()
        } else {
            u.email_id.clone()
        };
        Ok(json!({"status": "updated", "id": id, "changed": changed}))
    }

    fn delete_email(&self, email_id: String) -> Result<Value, ToolError> {
        self.record("delete_email", json!({"email_id": email_id}))?;
        Ok(json!({"status": "deleted"}))
    }

    fn list_events(&self, q: EventQuery) -> Result<Vec<EventSummary>, ToolError> {
        self.record("list_events", json!({
            "start_date": q.start_date, "end_date": q.end_date, "query": q.query,
            "category": q.category, "show_as": q.show_as, "my_response": q.my_response,
            "attendees": q.attendees, "attendee_role": q.attendee_role,
            "meetings_only": q.meetings_only, "all_day": q.all_day,
            "calendar_of": q.calendar_of,
        }))?;
        Ok(vec![EventSummary {
            id: EVENT_ID.into(), subject: "Standup".into(), start: None, end: None,
            location: "".into(), organizer: "".into(), all_day: false,
            is_recurring: false, is_meeting: false, categories: vec![],
            show_as: "busy".into(), my_response: "accepted".into(),
            required_attendees: "".into(), optional_attendees: "".into(),
        }])
    }

    fn get_event(&self, event_id: String) -> Result<EventDetail, ToolError> {
        self.record("get_event", json!({"event_id": event_id}))?;
        Ok(EventDetail {
            summary: EventSummary {
                id: event_id, subject: "Standup".into(), start: None, end: None,
                location: "".into(), organizer: "".into(), all_day: false,
                is_recurring: false, is_meeting: false, categories: vec![],
                show_as: "busy".into(), my_response: "accepted".into(),
                required_attendees: "".into(), optional_attendees: "".into(),
            },
            body: "".into(),
        })
    }

    fn create_event(&self, input: CreateEventInput) -> Result<Value, ToolError> {
        self.record("create_event", json!({
            "subject": input.subject, "start": input.start, "end": input.end,
            "body": input.body, "location": input.location,
            "required_attendees": input.required_attendees,
            "optional_attendees": input.optional_attendees,
            "all_day": input.all_day, "reminder_minutes": input.reminder_minutes,
            "categories": input.categories, "show_as": input.show_as,
            "send": input.send,
        }))?;
        let has_attendees = input.required_attendees.as_ref().is_some_and(|v| !v.is_empty())
            || input.optional_attendees.as_ref().is_some_and(|v| !v.is_empty());
        let status = super::create_event_status(has_attendees, input.send);
        Ok(json!({"status": status, "id": EVENT_ID, "subject": input.subject}))
    }

    fn respond_to_meeting(&self, event_id: String, response: String,
        comment: Option<String>, send: bool) -> Result<Value, ToolError> {
        self.record("respond_to_meeting",
            json!({"event_id": event_id, "response": response, "comment": comment, "send": send}))?;
        Ok(json!({"status": format!("{response}_sent")}))
    }

    fn list_attachments(&self, email_id: String)
        -> Result<Vec<AttachmentInfo>, ToolError> {
        self.record("list_attachments", json!({"email_id": email_id}))?;
        Ok(vec![AttachmentInfo { index: 1, filename: "report.pdf".into(), size: 1234 }])
    }

    fn save_attachments(&self, email_id: String, save_dir: String,
        attachment_names: Option<Vec<String>>) -> Result<Vec<Value>, ToolError> {
        self.record("save_attachments",
            json!({"email_id": email_id, "save_dir": save_dir, "attachment_names": attachment_names}))?;
        Ok(vec![json!({"filename": "report.pdf", "saved_to": save_dir, "status": "saved"})])
    }

    fn list_tasks(&self, include_completed: bool) -> Result<Vec<TaskSummary>, ToolError> {
        self.record("list_tasks", json!({"include_completed": include_completed}))?;
        Ok(vec![TaskSummary {
            id: TASK_ID.into(), subject: "Buy milk".into(), due_date: None,
            complete: false, status: "not_started".to_string(), importance: "normal".to_string(), categories: vec![],
        }])
    }

    fn create_task(&self, subject: String, body: Option<String>,
        due_date: Option<String>, importance: String) -> Result<Value, ToolError> {
        self.record("create_task",
            json!({"subject": subject, "body": body, "due_date": due_date, "importance": importance}))?;
        Ok(json!({"status": "created", "id": TASK_ID, "subject": subject}))
    }

    fn complete_task(&self, task_id: String) -> Result<Value, ToolError> {
        self.record("complete_task", json!({"task_id": task_id}))?;
        Ok(json!({"status": "completed"}))
    }

    fn list_notes(&self) -> Result<Vec<NoteSummary>, ToolError> {
        self.record("list_notes", json!({}))?;
        Ok(vec![NoteSummary { id: NOTE_ID.into(), subject: "Ideas".into(), created: None, categories: vec![] }])
    }

    fn get_note(&self, note_id: String) -> Result<NoteDetail, ToolError> {
        self.record("get_note", json!({"note_id": note_id}))?;
        Ok(NoteDetail {
            summary: NoteSummary { id: note_id, subject: "Ideas".into(), created: None, categories: vec![] },
            body: "Ideas\n- one".into(),
        })
    }

    fn create_note(&self, body: String) -> Result<Value, ToolError> {
        self.record("create_note", json!({"body": body}))?;
        Ok(json!({"status": "created", "id": NOTE_ID}))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn basic_query() -> EmailQuery {
        EmailQuery {
            query: None, folder: "inbox".into(), count: 10, unread_only: false,
            from: None, category: None, received_after: None, received_before: None,
            since_days: None, has_attachments: None, flagged: false, high_importance: false,
        }
    }

    #[test]
    fn records_calls_in_order() {
        let fake = FakeOutlookClient::new();
        fake.list_folders().unwrap();
        fake.list_emails(basic_query()).unwrap();
        assert_eq!(fake.calls(), vec![
            ("list_folders".to_string(), json!({})),
            ("list_emails".to_string(), json!({
                "query": null, "folder": "inbox", "count": 10, "unread_only": false,
                "from": null, "category": null, "received_after": null,
                "received_before": null, "since_days": null, "has_attachments": null,
                "flagged": false, "high_importance": false,
            })),
        ]);
    }

    #[test]
    fn fail_with_makes_every_call_error_before_recording() {
        let fake = FakeOutlookClient::new();
        fake.set_fail_with("Outlook exploded");
        let err = fake.list_emails(basic_query()).unwrap_err();
        assert_eq!(err.to_string(), "Outlook exploded");
        assert!(fake.calls().is_empty());
    }
}
