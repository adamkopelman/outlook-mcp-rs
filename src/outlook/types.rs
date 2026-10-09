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

/// `get_email`'s result. The optional fields follow the request's
/// `include` / `body_format` / `output_dir` (see `read::read_options`):
/// a body field not requested is omitted along with its `*_truncated` /
/// `*_length`, and a body written to `output_dir` comes back as
/// `<field>_file` instead of `<field>`.
#[derive(Debug, Clone, Serialize)]
pub struct EmailDetail {
    #[serde(flatten)]
    pub summary: EmailSummary,
    pub cc: String,
    pub bcc: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub body: Option<String>,
    /// Absolute path of the file holding the full plain-text body (`output_dir`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub body_file: Option<String>,
    /// True when `body` was cut at the caller's `max_body_chars`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub body_truncated: Option<bool>,
    /// Full original length of the body, in characters.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub body_length: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub html_body: Option<String>,
    /// Absolute path of the file holding the full HTML body (`output_dir`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub html_body_file: Option<String>,
    /// Whether `html_body` was cut.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub html_truncated: Option<bool>,
    /// Full HTML length in characters (after `resolve_inline_images`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub html_length: Option<usize>,
    /// With `resolve_inline_images`: how many distinct Content-IDs were
    /// replaced by `data:` URIs.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub inline_images_resolved: Option<usize>,
    /// With `resolve_inline_images`: referenced Content-IDs left as `cid:`
    /// (no such attachment, over 10 MB, or unreadable).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub inline_images_unresolved: Option<Vec<String>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub attachments: Option<Vec<String>>,
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

/// `get_event`'s result; the body fields behave as on [`EmailDetail`].
#[derive(Debug, Clone, Serialize)]
pub struct EventDetail {
    #[serde(flatten)]
    pub summary: EventSummary,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub body: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub body_file: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub body_truncated: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub body_length: Option<usize>,
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

/// `get_task`'s result: the `list_tasks` summary plus the body (fields as
/// on [`EmailDetail`]) and the remaining scheduling details. Dates are ISO;
/// Outlook's "none" date comes back as null.
#[derive(Debug, Clone, Serialize)]
pub struct TaskDetail {
    #[serde(flatten)]
    pub summary: TaskSummary,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub body: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub body_file: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub body_truncated: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub body_length: Option<usize>,
    pub start_date: Option<String>,
    pub date_completed: Option<String>,
    /// 0-100.
    pub percent_complete: i32,
    pub reminder_set: bool,
    pub reminder_time: Option<String>,
    pub created: Option<String>,
    pub modified: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
pub struct NoteSummary {
    pub id: String,
    pub subject: String,
    pub created: Option<String>,
    pub categories: Vec<String>,
}

/// `get_note`'s result; the body fields behave as on [`EmailDetail`].
#[derive(Debug, Clone, Serialize)]
pub struct NoteDetail {
    #[serde(flatten)]
    pub summary: NoteSummary,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub body: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub body_file: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub body_truncated: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub body_length: Option<usize>,
    pub modified: Option<String>,
}

/// Metadata for one attachment, shared by `list_attachments` and (flattened,
/// plus `saved_to`/`status`/`error`) `save_attachments`.
#[derive(Debug, Clone, Serialize)]
pub struct AttachmentInfo {
    /// COM's 1-based position in the item's `Attachments` collection.
    pub index: i32,
    pub filename: String,
    pub size: i32,
    /// `"file"`, `"link"`, `"item"`, `"ole"` or `"unknown"` (see
    /// `constants::attachment_type_name`).
    #[serde(rename = "type")]
    pub att_type: String,
    /// Content-ID without `<>`, as an HTML body references it (`cid:...`).
    pub content_id: Option<String>,
    pub mime_type: Option<String>,
    /// MAPI `PR_ATTACHMENT_HIDDEN`; false when the property is absent.
    pub hidden: bool,
    /// Inline (`cid:`-referenced) content rather than a standalone attachment:
    /// has a Content-ID and is either hidden or referenced by the HTML body
    /// (see `com::is_inline`).
    pub is_inline: bool,
}

/// An attachment fetched by Content-ID (`get_inline_image`), inlined as a
/// base64 `data:` URI.
#[derive(Debug, Clone, Serialize)]
pub struct InlineImageData {
    /// The attachment's Content-ID without `<>` (as `list_attachments` shows it).
    pub content_id: String,
    pub filename: String,
    /// From the attachment's metadata, else `application/octet-stream`.
    pub mime_type: String,
    /// Actual byte length of the decoded payload.
    pub size: usize,
    /// `data:<mime_type>;base64,<payload>`; omitted with `output_dir`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data_uri: Option<String>,
    /// With `output_dir`: absolute path of the file holding the decoded
    /// image bytes (instead of `data_uri`).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data_file: Option<String>,
    /// Only when `context_lines` was requested: the plain-text lines just
    /// before the first `cid:` reference to this image in the HTML body
    /// (`""` if the body never references it).
    #[serde(skip_serializing_if = "Option::is_none")]
    pub context: Option<String>,
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
            cc: "".into(), bcc: "".into(), body: Some("Hello".into()), body_file: None,
            body_truncated: Some(false), body_length: Some(5),
            html_body: None, html_body_file: None, html_truncated: None, html_length: None,
            inline_images_resolved: None, inline_images_unresolved: None,
            attachments: Some(vec![]),
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
        // Truncation info is always present for the plain body, and the
        // HTML counterparts are omitted without prefer_html.
        assert_eq!(value["body_truncated"], false);
        assert_eq!(value["body_length"], 5);
        assert!(value.get("html_truncated").is_none());
        assert!(value.get("html_length").is_none());
        assert!(value.get("body_file").is_none());
        assert!(value.get("inline_images_resolved").is_none());
        assert_eq!(value["attachments"], serde_json::json!([]));
    }
}
