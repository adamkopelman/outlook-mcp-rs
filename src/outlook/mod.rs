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
    pub recurrence: Option<RecurrenceInput>,
}

/// All changes `update_event` can apply to one existing calendar event. Every
/// field except `event_id` is optional; supplying several applies all of
/// them. There is no `move_to` — events don't change folders in this API —
/// so, unlike `EmailUpdate`, field application order is cosmetic, not a
/// correctness constraint. Adding either attendee tier converts a personal
/// appointment into a meeting. `send_update` (no default here — the tool
/// layer defaults it to `true`) controls whether a meeting's edits are
/// delivered to attendees or applied quietly to your own copy only; a
/// personal (non-meeting) appointment always just saves, regardless.
#[derive(Debug, Clone)]
pub struct EventUpdate {
    pub event_id: String,
    pub subject: Option<String>,
    pub start: Option<String>,
    pub end: Option<String>,
    pub location: Option<String>,
    pub body: Option<String>,
    pub all_day: Option<bool>,
    pub reminder_minutes: Option<i32>,
    pub show_as: Option<String>,
    pub add_categories: Option<Vec<String>>,
    pub remove_categories: Option<Vec<String>>,
    pub add_required_attendees: Option<Vec<String>>,
    pub add_optional_attendees: Option<Vec<String>>,
    pub remove_attendees: Option<Vec<String>>,
    pub send_update: bool,
    pub recurrence: Option<RecurrenceInput>,
    pub clear_recurrence: bool,
}

/// One recurrence pattern for `create_event`/`update_event`. `pattern`
/// selects which of the other fields matter: `"weekly"` requires
/// `days_of_week`; `"monthly"` requires `day_of_month`; `"yearly"` derives
/// its month/day from the event's own start date (no field needed);
/// `"daily"` needs nothing extra. At most one of `until`/`occurrences` may
/// be set; if neither is set the series has no end date.
#[derive(Debug, Clone)]
pub struct RecurrenceInput {
    pub pattern: String,
    pub interval: Option<i32>,
    pub days_of_week: Option<Vec<String>>,
    pub day_of_month: Option<i32>,
    pub until: Option<String>,
    pub occurrences: Option<i32>,
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
    fn update_event(&self, u: EventUpdate) -> Result<Value, ToolError>;
    fn delete_event(&self, event_id: String, send_cancellation: bool) -> Result<Value, ToolError>;

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

/// The status string `create_event` returns: `"meeting_sent"` (attendees +
/// send), `"meeting_saved"` (attendees + no send), or `"saved"` (no
/// attendees, regardless of `send` — there's nothing to send or withhold).
pub fn create_event_status(has_attendees: bool, send: bool) -> &'static str {
    match (has_attendees, send) {
        (true, true) => "meeting_sent",
        (true, false) => "meeting_saved",
        (false, _) => "saved",
    }
}

/// Resolves `r.pattern` to its `OlRecurrenceType` id and checks the fields
/// each pattern requires. Called first by both `create_event`'s and
/// `update_event`'s real-COM recurrence-writing code, before any COM call is
/// made, so a bad `recurrence` object fails fast with a clear message.
pub fn validate_recurrence(r: &RecurrenceInput) -> Result<i32, ToolError> {
    let recurrence_type = crate::friendly::recurrence_pattern_to_id(&r.pattern).ok_or_else(|| {
        ToolError::new(format!(
            "invalid recurrence.pattern {:?}: expected \"daily\", \"weekly\", \"monthly\", or \"yearly\"",
            r.pattern
        ))
    })?;
    if r.pattern.eq_ignore_ascii_case("weekly")
        && !r.days_of_week.as_ref().is_some_and(|d| !d.is_empty())
    {
        return Err(ToolError::new(
            "recurrence.days_of_week is required for a \"weekly\" pattern",
        ));
    }
    if r.pattern.eq_ignore_ascii_case("monthly") && r.day_of_month.is_none() {
        return Err(ToolError::new(
            "recurrence.day_of_month is required for a \"monthly\" pattern",
        ));
    }
    if r.occurrences.is_some() && r.until.is_some() {
        return Err(ToolError::new(
            "recurrence: specify at most one of \"until\" or \"occurrences\", not both",
        ));
    }
    Ok(recurrence_type)
}

#[cfg(test)]
mod tests {
    use super::{create_event_status, validate_recurrence, RecurrenceInput};

    #[test]
    fn create_event_status_covers_all_three_outcomes() {
        assert_eq!(create_event_status(true, true), "meeting_sent");
        assert_eq!(create_event_status(true, false), "meeting_saved");
        assert_eq!(create_event_status(false, true), "saved");
        assert_eq!(create_event_status(false, false), "saved");
    }

    fn recurrence(pattern: &str) -> RecurrenceInput {
        RecurrenceInput {
            pattern: pattern.to_string(), interval: None, days_of_week: None,
            day_of_month: None, until: None, occurrences: None,
        }
    }

    #[test]
    fn validate_recurrence_accepts_daily_with_no_extra_fields() {
        assert_eq!(validate_recurrence(&recurrence("daily")).unwrap(), 0);
    }

    #[test]
    fn validate_recurrence_rejects_unknown_pattern() {
        assert!(validate_recurrence(&recurrence("biweekly")).is_err());
    }

    #[test]
    fn validate_recurrence_requires_days_of_week_for_weekly() {
        assert!(validate_recurrence(&recurrence("weekly")).is_err());
        let mut r = recurrence("weekly");
        r.days_of_week = Some(vec!["monday".to_string()]);
        assert_eq!(validate_recurrence(&r).unwrap(), 1);
    }

    #[test]
    fn validate_recurrence_requires_day_of_month_for_monthly() {
        assert!(validate_recurrence(&recurrence("monthly")).is_err());
        let mut r = recurrence("monthly");
        r.day_of_month = Some(15);
        assert_eq!(validate_recurrence(&r).unwrap(), 2);
    }

    #[test]
    fn validate_recurrence_rejects_both_until_and_occurrences() {
        let mut r = recurrence("daily");
        r.until = Some("2099-01-01".to_string());
        r.occurrences = Some(5);
        assert!(validate_recurrence(&r).is_err());
    }
}
