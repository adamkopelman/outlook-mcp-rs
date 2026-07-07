use serde::Serialize;

#[derive(Debug, Clone, Serialize)]
pub struct FolderInfo {
    pub name: String,
    pub path: String,
    pub items: i32,
    pub unread: i32,
}

#[derive(Debug, Clone, Serialize)]
pub struct EmailSummary {
    pub id: String,
    pub subject: String,
    pub sender: String,
    pub sender_email: String,
    pub to: String,
    pub received: Option<String>,
    pub unread: bool,
    pub has_attachments: bool,
    pub categories: Vec<String>,
}

#[derive(Debug, Clone, Serialize)]
pub struct EmailDetail {
    #[serde(flatten)]
    pub summary: EmailSummary,
    pub cc: String,
    pub bcc: String,
    pub body: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub html_body: Option<String>,
    pub attachments: Vec<String>,
}

#[derive(Debug, Clone, Serialize)]
pub struct EventSummary {
    pub id: String,
    pub subject: String,
    pub start: Option<String>,
    pub end: Option<String>,
    pub location: String,
    pub organizer: String,
    pub all_day: bool,
    pub is_recurring: bool,
    pub is_meeting: bool,
    pub categories: Vec<String>,
}

#[derive(Debug, Clone, Serialize)]
pub struct EventDetail {
    #[serde(flatten)]
    pub summary: EventSummary,
    pub body: String,
    pub required_attendees: String,
    pub optional_attendees: String,
    pub response_status: Option<i32>,
}

#[derive(Debug, Clone, Serialize)]
pub struct TaskSummary {
    pub id: String,
    pub subject: String,
    pub due_date: Option<String>,
    pub complete: bool,
    pub status: i32,
    pub importance: i32,
    pub categories: Vec<String>,
}

#[derive(Debug, Clone, Serialize)]
pub struct NoteSummary {
    pub id: String,
    pub subject: String,
    pub created: Option<String>,
    pub categories: Vec<String>,
}

#[derive(Debug, Clone, Serialize)]
pub struct NoteDetail {
    #[serde(flatten)]
    pub summary: NoteSummary,
    pub body: String,
}

#[derive(Debug, Clone, Serialize)]
pub struct AttachmentInfo {
    pub index: i32,
    pub filename: String,
    pub size: i32,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn email_detail_flattens_summary_fields_at_top_level() {
        let detail = EmailDetail {
            summary: EmailSummary {
                id: "e1|s1".into(), subject: "Hi".into(), sender: "Ada".into(),
                sender_email: "ada@example.com".into(), to: "bob@example.com".into(),
                received: Some("2026-06-10T12:00:00".into()), unread: true,
                has_attachments: false, categories: vec![],
            },
            cc: "".into(), bcc: "".into(), body: "Hello".into(),
            html_body: None, attachments: vec![],
        };
        let value = serde_json::to_value(&detail).unwrap();
        // Flattened: "id" and "subject" appear at the top level, not nested
        // under a "summary" key, and html_body is omitted when None.
        assert_eq!(value["id"], "e1|s1");
        assert_eq!(value["subject"], "Hi");
        assert_eq!(value["body"], "Hello");
        assert!(value.get("html_body").is_none());
        assert!(value.get("summary").is_none());
    }
}
