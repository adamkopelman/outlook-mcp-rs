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
    pub item_type: String,
    pub is_meeting: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub meeting: Option<MeetingInfo>,
}

#[derive(Debug, Clone, Serialize)]
pub struct MeetingInfo {
    pub meeting_type: String,
    pub start: Option<String>,
    pub end: Option<String>,
    pub location: String,
    pub organizer: String,
    pub required_attendees: String,
    pub optional_attendees: String,
    pub is_recurring: bool,
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
    /// Busy status as a friendly word: "free"/"tentative"/"busy"/"out_of_office"/"working_elsewhere".
    pub show_as: String,
    /// This mailbox's response as a friendly word: "organizer"/"accepted"/"declined"/"tentative"/"not_responded"/"none".
    pub my_response: String,
    pub required_attendees: String,
    pub optional_attendees: String,
}

#[derive(Debug, Clone, Serialize)]
pub struct EventDetail {
    #[serde(flatten)]
    pub summary: EventSummary,
    pub body: String,
    pub recurrence: Option<RecurrenceInfo>,
}

/// The recurrence pattern of a recurring event, read back via
/// `AppointmentItem.GetRecurrencePattern()`. `None` on `EventDetail` when the
/// event isn't recurring.
///
/// On the *write* side, `RecurrenceInput`'s `until`/`occurrences` are
/// mutually exclusive (`validate_recurrence` rejects both being set), and a
/// series with neither set is unbounded (`no_end`). On this *read-back*
/// side, `no_end: true` still means the other two are `None`, but when the
/// series has a finite end (`no_end: false`), `until` and `occurrences` are
/// populated *together* — confirmed live: Outlook's `RecurrencePattern`
/// keeps `Occurrences` and `PatternEndDate` mutually consistent regardless
/// of which one the series was originally created with (e.g. a series
/// created with only `until` still reports a correct, auto-computed
/// `Occurrences`, and vice versa). There is no COM-level signal for which
/// field the caller originally specified, so this struct does not attempt
/// to suppress either one.
#[derive(Debug, Clone, Serialize)]
pub struct RecurrenceInfo {
    /// "daily" | "weekly" | "monthly" | "yearly".
    pub pattern: String,
    pub interval: i32,
    /// Populated only for "weekly"; e.g. ["monday", "wednesday"].
    pub days_of_week: Vec<String>,
    /// Populated only for "monthly"/"yearly".
    pub day_of_month: Option<i32>,
    /// ISO end date, if the series has a finite end (`no_end: false`).
    pub until: Option<String>,
    /// Outlook's auto-computed occurrence count, if the series has a finite
    /// end (`no_end: false`) — populated alongside `until`, not only when
    /// the series was created via `occurrences`; see struct doc comment.
    pub occurrences: Option<i32>,
    /// True if the series never ends.
    pub no_end: bool,
}

#[derive(Debug, Clone, Serialize, PartialEq)]
pub struct AvailabilitySlot {
    pub start: String,
    pub end: String,
    pub status: String,
}

#[derive(Debug, Clone, Serialize)]
pub struct PersonAvailability {
    pub person: String,
    /// `false` covers two distinct COM outcomes that both degrade this one
    /// person rather than failing the whole `check_availability` call:
    /// the address itself couldn't be resolved (`Recipient.Resolve()`
    /// returned `false`), or it resolved fine but no free/busy data could
    /// be loaded for it (`Recipient.FreeBusy()` errored — e.g. a
    /// syntactically valid but nonexistent/unpublished address). Callers
    /// cannot distinguish the two from this field alone; `slots` is empty
    /// either way.
    pub resolved: bool,
    pub slots: Vec<AvailabilitySlot>,
}

#[derive(Debug, Clone, Serialize, PartialEq)]
pub struct FreeWindow {
    pub start: String,
    pub end: String,
}

#[derive(Debug, Clone, Serialize)]
pub struct AvailabilityResult {
    pub people: Vec<PersonAvailability>,
    pub common_free: Vec<FreeWindow>,
}

#[derive(Debug, Clone, Serialize)]
pub struct TaskSummary {
    pub id: String,
    pub subject: String,
    pub due_date: Option<String>,
    pub complete: bool,
    pub status: String,
    pub importance: String,
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
            item_type: "email".into(), is_meeting: false, meeting: None,
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
