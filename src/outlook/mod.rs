pub mod client;
pub mod com;
pub mod dates;
pub mod fake;
pub mod filters;
pub mod read;
pub mod text_query;
pub mod types;

use crate::error::ToolError;
use serde_json::Value;
use types::*;

pub use read::{BatchResult, ReadOptions, ReadRequest, ReadTool};

/// Paging caps shared by the `list_*` tools (issue #34). Every list tool
/// takes `count` (clamped to `1..=` its cap) and `offset` (matches to skip,
/// after every filter). `list_emails` defaults to 10; the others default to
/// their cap, which is what they returned before paging existed.
pub const DEFAULT_EMAIL_COUNT: i32 = 10;
pub const MAX_EMAIL_COUNT: i32 = 200;
/// `list_events` cap: a recurring series without an end date expands
/// without bound under `IncludeRecurrences`.
pub const MAX_EVENT_COUNT: i32 = 250;
pub const MAX_TASK_COUNT: i32 = 500;
pub const MAX_NOTE_COUNT: i32 = 500;

/// All filters for `list_emails`. Every filter is optional; supplying
/// several ANDs them, and the values of one list-valued filter are ORed
/// (an empty list means "no filter"). Dates use the shared grammar in
/// [`dates`]; `query` uses the shared syntax in [`text_query`].
#[derive(Debug, Clone)]
pub struct EmailQuery {
    pub query: Option<String>,
    pub folder: String,
    pub count: i32,
    /// Matches to skip before the page starts (after every filter). Negative
    /// values are treated as 0.
    pub offset: i32,
    pub unread_only: bool,
    /// Sender filter: caseless substring of the sender's name or address.
    pub from: Vec<String>,
    /// Recipient filter: caseless substring of any To/CC recipient's
    /// display name or address.
    pub to: Vec<String>,
    pub category: Vec<String>,
    /// `ReceivedTime >=` this date.
    pub received_after: Option<String>,
    /// `ReceivedTime <=` this date (a bare ISO date includes that whole day).
    pub received_before: Option<String>,
    pub has_attachments: Option<bool>,
    /// `get_email`'s `item_type`: "email" | "meeting" | "bounce" | "read_receipt" | "other".
    pub item_type: Vec<String>,
    /// "low" | "normal" | "high".
    pub importance: Vec<String>,
    /// `update_email`'s `flag` values: "follow_up" | "complete" | "clear" (no flag).
    pub flag: Vec<String>,
}

impl Default for EmailQuery {
    fn default() -> Self {
        Self {
            query: None, folder: "inbox".to_string(), count: DEFAULT_EMAIL_COUNT, offset: 0,
            unread_only: false, from: Vec::new(), to: Vec::new(), category: Vec::new(),
            received_after: None, received_before: None, has_attachments: None,
            item_type: Vec::new(), importance: Vec::new(), flag: Vec::new(),
        }
    }
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

/// All changes `update_draft` can apply to one existing, unsent draft.
/// Every field except `email_id` is optional; supplying several applies all
/// of them in field order and then saves once. `subject`, `body` and
/// `html_body` replace the current value (`body` and `html_body` are
/// mutually exclusive; `data:` image URIs in `html_body` become inline
/// attachments). `to`/`cc`/`bcc` replace that whole recipient line, and
/// `Some(vec![])` clears it. `attachments` are local file paths appended to
/// the existing attachments; `inline_images` (which need `html_body`) are
/// added as hidden Content-ID attachments. `send: true` sends the draft
/// after the changes are saved (it may be the only "change").
#[derive(Debug, Clone, Default)]
pub struct DraftUpdate {
    pub email_id: String,
    pub subject: Option<String>,
    pub body: Option<String>,
    pub html_body: Option<String>,
    pub to: Option<Vec<String>>,
    pub cc: Option<Vec<String>>,
    pub bcc: Option<Vec<String>>,
    pub attachments: Option<Vec<String>>,
    pub inline_images: Option<Vec<InlineImage>>,
    pub add_categories: Option<Vec<String>>,
    pub remove_categories: Option<Vec<String>>,
    pub importance: Option<String>, // "low" | "normal" | "high"
    pub send: bool,
}

/// A new email for `send_email` / `create_draft`. `data:` image URIs in an
/// HTML body become inline attachments; `inline_images` need an HTML body.
#[derive(Debug, Clone, Default)]
pub struct NewEmail {
    pub to: Vec<String>,
    pub subject: String,
    pub body: MailBody,
    pub cc: Option<Vec<String>>,
    pub bcc: Option<Vec<String>>,
    pub attachments: Option<Vec<String>>,
    pub inline_images: Option<Vec<InlineImage>>,
    pub categories: Option<Vec<String>>,
    pub importance: Option<String>, // "low" | "normal" | "high"
}

/// A reply for `reply_email`. `body` is put above the quoted original (an
/// HTML body is prepended to the original's HTML). `send: false` saves the
/// reply as a draft instead of sending it.
#[derive(Debug, Clone, Default)]
pub struct ReplyInput {
    pub email_id: String,
    pub body: MailBody,
    pub reply_all: bool,
    pub send: bool,
    pub attachments: Option<Vec<String>>,
    pub inline_images: Option<Vec<InlineImage>>,
    pub categories: Option<Vec<String>>,
    pub importance: Option<String>, // "low" | "normal" | "high"
}

/// One image to embed in an HTML body as a hidden Content-ID attachment
/// (`send_email` / `create_draft` `inline_images`). The body references it
/// as `<img src="cid:CONTENT_ID">`. Also the tool-layer argument type, so
/// the field docs below are the public schema descriptions.
#[derive(Debug, Clone, Default, serde::Serialize, serde::Deserialize, schemars::JsonSchema)]
pub struct InlineImage {
    /// Content-ID the HTML body references as <img src="cid:CONTENT_ID">, e.g. "logo".
    pub content_id: String,
    /// Local image file path. Give exactly one of `path` or `data_base64`.
    #[serde(default)]
    pub path: Option<String>,
    /// Base64-encoded image bytes (a "data:image/png;base64," prefix is
    /// accepted). Give exactly one of `path` or `data_base64`.
    #[serde(default)]
    pub data_base64: Option<String>,
    /// Attachment file name; defaults to the path's file name, or
    /// CONTENT_ID plus an extension.
    #[serde(default)]
    pub filename: Option<String>,
    /// MIME type like "image/png"; guessed from the file name or the data
    /// when omitted.
    #[serde(default)]
    pub mime_type: Option<String>,
}

/// Where a validated inline image's bytes come from.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum InlineImageSource {
    /// An existing local file, attached as-is.
    Path(String),
    /// Decoded base64 bytes; the COM layer writes them to a temp file first,
    /// since `Attachments.Add` only takes a path.
    Data(Vec<u8>),
}

/// An [`InlineImage`] after [`validate_inline_images`]: normalized
/// content id, resolved bytes source, safe filename and MIME type.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ValidatedInlineImage {
    pub content_id: String,
    pub source: InlineImageSource,
    pub filename: String,
    pub mime_type: String,
}

/// All changes `update_task` can apply to one existing task. Every field
/// except `task_id` is optional; supplying several applies all of them.
/// `mark_complete: Some(true)` replaces the retired standalone
/// `complete_task` tool (`= update_task` with `mark_complete: true`);
/// `Some(false)` reopens a completed task, filling the "can't reopen" gap
/// the old `complete_task` had no way to close. `mark_complete` is applied
/// *last*, after every other field write (`MarkComplete()`/reopen both set
/// `PercentComplete` themselves) — so combining `percent_complete` with
/// `mark_complete: Some(false)` in one call silently resets
/// `percent_complete` to 0 regardless of the value supplied; set it in a
/// separate call afterward if a specific non-zero value should stick.
#[derive(Debug, Clone, Default)]
pub struct TaskUpdate {
    pub task_id: String,
    pub mark_complete: Option<bool>,
    pub subject: Option<String>,
    pub body: Option<String>,
    pub due_date: Option<String>,
    pub start_date: Option<String>,
    pub importance: Option<String>,
    pub add_categories: Option<Vec<String>>,
    pub remove_categories: Option<Vec<String>>,
    pub percent_complete: Option<i32>,
    pub reminder_time: Option<String>,
}

/// All changes `update_note` can apply to one existing note. Every field
/// except `note_id` is optional; supplying several applies all of them.
#[derive(Debug, Clone, Default)]
pub struct NoteUpdate {
    pub note_id: String,
    pub body: Option<String>,
    pub add_categories: Option<Vec<String>>,
    pub remove_categories: Option<Vec<String>>,
    pub color: Option<String>,
}

/// All filters for `list_events`. Every field is optional; supplying several
/// ANDs them, and the values of one list-valued filter are ORed.
/// `start_after`/`start_before` bound the (recurrence-expanded) scan on the
/// event's start (default: today 00:00 to 7 days later); the rest filter the
/// streamed events client-side (see [`filters::event_matches`]).
/// `calendar_of` (an email/name) opens another person's shared calendar
/// instead of your own.
#[derive(Debug, Clone)]
pub struct EventQuery {
    pub start_after: Option<String>,
    /// A bare ISO date includes that whole day.
    pub start_before: Option<String>,
    pub query: Option<String>,
    pub category: Vec<String>,
    pub show_as: Vec<String>,                  // "free"|"tentative"|"busy"|"out_of_office"|"working_elsewhere"
    pub my_response: Vec<String>,              // "organizer"|"accepted"|"declined"|"tentative"|"not_responded"|"none"
    pub attendees: Vec<String>,                // match events where ANY listed person participates
    pub attendee_role: Option<String>,         // "required"|"optional"|"any" (default "any")
    pub meetings_only: bool,
    pub all_day: Option<bool>,
    pub calendar_of: Option<String>,
    pub count: i32,
    pub offset: i32,
}

impl Default for EventQuery {
    fn default() -> Self {
        Self {
            start_after: None, start_before: None, query: None, category: Vec::new(),
            show_as: Vec::new(), my_response: Vec::new(), attendees: Vec::new(),
            attendee_role: None, meetings_only: false, all_day: None, calendar_of: None,
            count: MAX_EVENT_COUNT, offset: 0,
        }
    }
}

/// All filters for `list_tasks`. Every field is optional except
/// `include_completed`; supplying several ANDs them, and the values of one
/// list-valued filter are ORed. `include_completed` drives a server-side
/// `Restrict`; the rest filter the streamed tasks client-side (see
/// [`filters::task_matches`]). `query` searches the subject and the real
/// task body, read per item.
#[derive(Debug, Clone)]
pub struct TaskQuery {
    pub include_completed: bool,
    pub category: Vec<String>,
    pub importance: Vec<String>,
    pub query: Option<String>,
    /// `DueDate >=` this date; tasks without a due date never match.
    pub due_after: Option<String>,
    /// `DueDate <=` this date (a bare ISO date includes that whole day).
    pub due_before: Option<String>,
    pub count: i32,
    pub offset: i32,
}

impl Default for TaskQuery {
    fn default() -> Self {
        Self {
            include_completed: false, category: Vec::new(), importance: Vec::new(), query: None,
            due_after: None, due_before: None, count: MAX_TASK_COUNT, offset: 0,
        }
    }
}

/// All filters for `list_notes`. Every field is optional; supplying several
/// ANDs them. A note's only content is its body (its `subject` is the
/// body's first line), so `query` searches the real body text (see
/// [`filters::note_matches`]).
#[derive(Debug, Clone)]
pub struct NoteQuery {
    pub category: Vec<String>,
    pub query: Option<String>,
    /// `CreationTime >=` this date.
    pub created_after: Option<String>,
    /// `CreationTime <=` this date (a bare ISO date includes that whole day).
    pub created_before: Option<String>,
    pub count: i32,
    pub offset: i32,
}

impl Default for NoteQuery {
    fn default() -> Self {
        Self {
            category: Vec::new(), query: None, created_after: None, created_before: None,
            count: MAX_NOTE_COUNT, offset: 0,
        }
    }
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
    /// Batch read (one result per id, in order). `opts` picks the optional
    /// fields, the body cut (`read::clamp_body_limit`), `output_dir` and
    /// `resolve_inline_images`; see `read::read_options`.
    fn get_email(&self, email_ids: Vec<String>, opts: &ReadOptions) -> BatchResult<EmailDetail>;
    fn send_email(&self, m: NewEmail) -> Result<Value, ToolError>;
    fn create_draft(&self, m: NewEmail) -> Result<Value, ToolError>;
    fn reply_email(&self, r: ReplyInput) -> Result<Value, ToolError>;
    fn update_email(&self, u: EmailUpdate) -> Result<Value, ToolError>;
    fn update_draft(&self, u: DraftUpdate) -> Result<Value, ToolError>;
    /// `permanent = false` moves the email to Deleted Items; `true`
    /// hard-deletes it like Outlook's shift+delete.
    fn delete_email(&self, email_id: String, permanent: bool) -> Result<Value, ToolError>;
    /// Permanently delete everything in the default store's Deleted Items.
    /// Refuses (via [`require_empty_confirm`]) unless `confirm` is true.
    fn empty_deleted_items(&self, confirm: bool) -> Result<Value, ToolError>;

    fn list_events(&self, q: EventQuery) -> Result<Vec<EventSummary>, ToolError>;
    fn get_event(&self, event_ids: Vec<String>, opts: &ReadOptions) -> BatchResult<EventDetail>;
    fn create_event(&self, input: CreateEventInput) -> Result<Value, ToolError>;
    fn respond_to_meeting(&self, event_id: String, response: String,
        comment: Option<String>, send: bool) -> Result<Value, ToolError>;
    fn update_event(&self, u: EventUpdate) -> Result<Value, ToolError>;
    fn delete_event(&self, event_id: String, send_cancellation: bool) -> Result<Value, ToolError>;
    fn check_availability(&self, input: CheckAvailabilityInput) -> Result<AvailabilityResult, ToolError>;

    fn list_attachments(&self, email_ids: Vec<String>) -> BatchResult<Vec<AttachmentInfo>>;
    /// `inline`: `Some(b)` saves only attachments whose `is_inline == b`
    /// (applied together with `attachment_names`); `None` saves all.
    fn save_attachments(&self, email_id: String, save_dir: String,
        attachment_names: Option<Vec<String>>, inline: Option<bool>) -> Result<Vec<Value>, ToolError>;
    /// One result per requested Content-ID, in order. With `output_dir`,
    /// each image's bytes are written there (`data_file`) instead of being
    /// returned as `data_uri`.
    fn get_inline_image(&self, email_id: String, content_ids: Vec<String>,
        context_lines: Option<u32>, output_dir: Option<String>) -> BatchResult<InlineImageData>;

    fn list_tasks(&self, q: TaskQuery) -> Result<Vec<TaskSummary>, ToolError>;
    fn create_task(&self, subject: String, body: Option<String>,
        due_date: Option<String>, importance: String, categories: Option<Vec<String>>,
        start_date: Option<String>, reminder_time: Option<String>) -> Result<Value, ToolError>;
    fn get_task(&self, task_ids: Vec<String>, opts: &ReadOptions) -> BatchResult<TaskDetail>;
    fn update_task(&self, u: TaskUpdate) -> Result<Value, ToolError>;
    fn delete_task(&self, task_id: String) -> Result<Value, ToolError>;

    fn list_notes(&self, q: NoteQuery) -> Result<Vec<NoteSummary>, ToolError>;
    fn get_note(&self, note_ids: Vec<String>, opts: &ReadOptions) -> BatchResult<NoteDetail>;
    fn create_note(&self, body: String, categories: Option<Vec<String>>, color: Option<String>) -> Result<Value, ToolError>;
    fn update_note(&self, u: NoteUpdate) -> Result<Value, ToolError>;
    fn delete_note(&self, note_id: String) -> Result<Value, ToolError>;
}

/// Guard for `empty_deleted_items`: it is irreversible, so it refuses unless
/// the caller explicitly passed `confirm = true`. Shared by the real and fake
/// clients so the refusal is the same everywhere.
pub fn require_empty_confirm(confirm: bool) -> Result<(), ToolError> {
    if confirm {
        return Ok(());
    }
    Err(ToolError::new(
        "empty_deleted_items permanently deletes EVERYTHING in Deleted Items \
         (items and subfolders) and cannot be undone. Call again with \
         confirm=true to proceed.",
    ))
}

/// Whether a permanent `delete_email` must first move the item into Deleted
/// Items. The object model has no hard-delete call, but `Delete()` on an item
/// already in Deleted Items is permanent; so move there unless the item's
/// parent folder already is that folder (EntryIDs compared exactly).
pub fn permanent_delete_needs_move(parent_entry_id: &str, deleted_items_entry_id: &str) -> bool {
    parent_entry_id.is_empty() || parent_entry_id != deleted_items_entry_id
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

/// The `RecurrenceType`-dependent `Interval` value Outlook's COM object model
/// expects. Every pattern except `"yearly"` passes the user's `interval`
/// straight through ("every N days/weeks/months"). `"yearly"` is the one
/// exception: Outlook's `RecurrencePattern.Interval` is documented as being
/// in **months** for `olRecursYearly`, and must be a multiple of 12 — "every
/// 1 year" is `Interval = 12`, "every 2 years" is `Interval = 24`. Called by
/// `apply_recurrence` (`client.rs`) right after `validate_recurrence`.
pub fn com_recurrence_interval(r: &RecurrenceInput) -> i32 {
    let interval = r.interval.unwrap_or(1);
    if r.pattern.eq_ignore_ascii_case("yearly") {
        interval * 12
    } else {
        interval
    }
}

/// Rejects an `EventUpdate` that sets both `recurrence` and `clear_recurrence`
/// — Outlook has no single COM call that means "replace recurrence and clear
/// it," so exactly one (or neither) must be set. Called by both
/// `update_event` implementors before either field is applied.
pub fn validate_recurrence_update(u: &EventUpdate) -> Result<(), ToolError> {
    if u.recurrence.is_some() && u.clear_recurrence {
        return Err(ToolError::new(
            "cannot set recurrence and clear_recurrence in the same update_event call",
        ));
    }
    Ok(())
}

/// Rejects a `DraftUpdate` that can't be applied, before the draft is
/// touched: `body` and `html_body` together, an update that changes nothing
/// and doesn't send (an empty list counts as nothing), `inline_images`
/// without `html_body`, an invalid `importance`, or an attachment path that
/// isn't an existing file. Called by both `update_draft` implementors.
pub fn validate_draft_update(u: &DraftUpdate) -> Result<(), ToolError> {
    if u.body.is_some() && u.html_body.is_some() {
        return Err(ToolError::new("pass either `body` or `html_body`, not both"));
    }
    if draft_update_changes(u).is_empty() && !u.send {
        return Err(ToolError::new(
            "update_draft needs at least one of: subject, body, html_body, to, cc, bcc, \
             attachments, inline_images, add_categories, remove_categories, importance, send",
        ));
    }
    if u.inline_images.as_ref().is_some_and(|i| !i.is_empty()) && u.html_body.is_none() {
        return Err(inline_images_need_html());
    }
    parse_importance(u.importance.as_deref())?;
    for p in u.attachments.as_deref().unwrap_or(&[]) {
        if !std::path::Path::new(p).is_file() {
            return Err(ToolError::new(format!("attachment not found: {p}")));
        }
    }
    Ok(())
}

/// The `changed` list `update_draft` returns: the supplied fields, in the
/// order they are applied (`send` is not a change; it's reported by the
/// status). Empty lists of attachments/inline images/categories count as
/// not supplied.
pub fn draft_update_changes(u: &DraftUpdate) -> Vec<&'static str> {
    let non_empty = |v: &Option<Vec<String>>| v.as_ref().is_some_and(|a| !a.is_empty());
    let mut changed = Vec::new();
    if u.subject.is_some() { changed.push("subject"); }
    if u.body.is_some() { changed.push("body"); }
    if u.html_body.is_some() { changed.push("html_body"); }
    if u.to.is_some() { changed.push("to"); }
    if u.cc.is_some() { changed.push("cc"); }
    if u.bcc.is_some() { changed.push("bcc"); }
    if non_empty(&u.attachments) { changed.push("attachments"); }
    if u.inline_images.as_ref().is_some_and(|i| !i.is_empty()) { changed.push("inline_images"); }
    if non_empty(&u.add_categories) { changed.push("add_categories"); }
    if non_empty(&u.remove_categories) { changed.push("remove_categories"); }
    if u.importance.is_some() { changed.push("importance"); }
    changed
}

/// Inverse of [`com_recurrence_interval`]: converts a COM `Interval` value
/// read back from `RecurrencePattern` into the user-facing "every N
/// years/months/..." value. Only `olRecursYearly` (`OL_RECURS_YEARLY`) needs
/// unwinding, since it's the only pattern `com_recurrence_interval` scales.
/// Called by `recurrence_info` (`client.rs`) right after reading `Interval`.
pub fn friendly_recurrence_interval(recurrence_type: i32, com_interval: i32) -> i32 {
    if recurrence_type == crate::constants::OL_RECURS_YEARLY {
        com_interval / 12
    } else {
        com_interval
    }
}

/// All inputs for `check_availability`. `treat_as_free` decides which raw
/// statuses count as "free" when computing `common_free` — it never changes
/// what a person's own `slots` report (those always show the true status).
#[derive(Debug, Clone)]
pub struct CheckAvailabilityInput {
    pub people: Vec<String>,
    pub start: String,
    pub end: String,
    pub interval_minutes: i32,
    pub treat_as_free: Vec<String>,
}

/// Parses Outlook's raw `Recipient.FreeBusy` status-code string (one ASCII
/// digit per `interval_minutes`-sized slot: `'0'` free, `'1'` tentative,
/// `'2'` busy, `'3'` out-of-office, `'4'` working-elsewhere — the exact
/// `OlBusyStatus` numbering `friendly::busy_status_word` already maps) into
/// timestamped slots starting at `start`. `FreeBusy` returns a string
/// covering a much longer range than the caller's `[start, end)` window (it
/// has no `end` parameter), so callers must compute `max_slots` themselves
/// — `(end - start) / interval_minutes`, rounded up — and this function
/// truncates to it. Any digit outside 0-4 (Outlook shouldn't produce one,
/// but the string could be malformed) falls back to `"busy"`, the same
/// catch-all `busy_status_word` uses.
pub fn parse_freebusy_slots(
    raw: &str,
    start: &chrono::NaiveDateTime,
    interval_minutes: i32,
    max_slots: usize,
) -> Vec<AvailabilitySlot> {
    raw.chars()
        .take(max_slots)
        .enumerate()
        .map(|(i, ch)| {
            let code = ch.to_digit(10).map(|d| d as i32).unwrap_or(crate::constants::OL_BUSY);
            let slot_start = *start + chrono::Duration::minutes(i as i64 * interval_minutes as i64);
            let slot_end = slot_start + chrono::Duration::minutes(interval_minutes as i64);
            AvailabilitySlot {
                start: slot_start.format("%Y-%m-%dT%H:%M:%S").to_string(),
                end: slot_end.format("%Y-%m-%dT%H:%M:%S").to_string(),
                status: crate::friendly::busy_status_word(code).to_string(),
            }
        })
        .collect()
}

/// The windows where every **resolved** person's status is in
/// `treat_as_free` (case-insensitive). Unresolved people are skipped
/// entirely — they neither block nor contribute to a common-free window.
/// Assumes all resolved people's `slots` share the same slot boundaries
/// (true whenever they were built from the same `start`/`interval_minutes`,
/// which `check_availability` always uses); intersects only over the
/// shortest `slots` length present, so a person whose raw string was
/// unexpectedly short doesn't panic the lookup.
pub fn common_free(people: &[PersonAvailability], treat_as_free: &[String]) -> Vec<FreeWindow> {
    let resolved: Vec<&PersonAvailability> = people.iter().filter(|p| p.resolved).collect();
    if resolved.is_empty() {
        return Vec::new();
    }
    let treat_lower: Vec<String> = treat_as_free.iter().map(|s| s.to_lowercase()).collect();
    let min_len = resolved.iter().map(|p| p.slots.len()).min().unwrap_or(0);

    let mut windows = Vec::new();
    let mut run_start: Option<usize> = None;
    for i in 0..min_len {
        let all_free = resolved
            .iter()
            .all(|p| treat_lower.contains(&p.slots[i].status.to_lowercase()));
        if all_free {
            run_start.get_or_insert(i);
        } else if let Some(s) = run_start.take() {
            windows.push(FreeWindow {
                start: resolved[0].slots[s].start.clone(),
                end: resolved[0].slots[i - 1].end.clone(),
            });
        }
    }
    if let Some(s) = run_start {
        windows.push(FreeWindow {
            start: resolved[0].slots[s].start.clone(),
            end: resolved[0].slots[min_len - 1].end.clone(),
        });
    }
    windows
}

/// Upper bound on `get_inline_image`'s `context_lines`; larger requests are
/// clamped to it.
pub const MAX_CONTEXT_LINES: u32 = 50;

/// Characters that can continue a Content-ID (RFC 5322 `atext` plus `.`/`@`,
/// minus the HTML-significant `'` and `&`). A `cid:` reference only matches
/// when the next character is NOT one of these, so `cid:logo` does not match
/// `cid:logo.png` or `cid:logo2`.
fn is_cid_char(c: char) -> bool {
    c.is_alphanumeric() || "._@-+$%!#*/=?^`{|}~".contains(c)
}

/// Byte offset of the first `cid:<cid>` reference in `html` (case-insensitive,
/// whole-id), if any. `cid` is already normalized (no `cid:` prefix or `<>`).
/// Case folding is ASCII-only so byte offsets stay valid in `html`.
fn find_cid_reference(html: &str, cid: &str) -> Option<usize> {
    if cid.is_empty() {
        return None;
    }
    let haystack = html.to_ascii_lowercase();
    let needle = format!("cid:{}", cid.to_ascii_lowercase());
    haystack.match_indices(&needle).map(|(pos, _)| pos).find(|&pos| {
        // Left boundary: not part of a longer word like `xcid:`.
        let left_ok = haystack[..pos].chars().next_back().is_none_or(|c| !c.is_alphanumeric());
        let right_ok = haystack[pos + needle.len()..].chars().next().is_none_or(|c| !is_cid_char(c));
        left_ok && right_ok
    })
}

/// Decode the common HTML entities (`&nbsp;`, `&amp;`, `&lt;`, `&gt;`,
/// `&quot;`, `&apos;`/`&#39;`, and numeric `&#NNN;`/`&#xHH;`). Anything else
/// (unknown names, invalid code points, a missing `;`) is left as is.
fn decode_entities(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut rest = s;
    while let Some(amp) = rest.find('&') {
        out.push_str(&rest[..amp]);
        rest = &rest[amp..];
        // Entities are short; only look a few bytes ahead for the `;` (ASCII,
        // so its byte position is always a char boundary).
        let decoded = rest.bytes().skip(1).take(11).position(|b| b == b';').and_then(|semi| {
            let name = &rest[1..semi + 1];
            let ch = match name {
                "nbsp" => '\u{a0}',
                "amp" => '&',
                "lt" => '<',
                "gt" => '>',
                "quot" => '"',
                "apos" => '\'',
                _ => {
                    let num = name.strip_prefix('#')?;
                    let code = match num.strip_prefix(['x', 'X']) {
                        Some(hex) => u32::from_str_radix(hex, 16).ok()?,
                        None => num.parse::<u32>().ok()?,
                    };
                    char::from_u32(code)?
                }
            };
            Some((ch, semi + 2))
        });
        match decoded {
            Some((ch, len)) => {
                out.push(ch);
                rest = &rest[len..];
            }
            None => {
                out.push('&');
                rest = &rest[1..];
            }
        }
    }
    out.push_str(rest);
    out
}

/// A rough HTML-to-text conversion for `text_before_cid`: drops comments and
/// `<script>`/`<style>` blocks, turns `<br>` and the closing tags of block
/// elements (`p`, `div`, `li`, `tr`, `h1`-`h6`) into newlines, separates
/// table cells with a space, strips every other tag and decodes entities.
/// Source whitespace (including newlines) is just whitespace, as in a
/// browser; only tags produce line breaks.
fn html_to_text(html: &str) -> String {
    let mut out = String::new();
    let mut text = String::new();
    // Flush the pending text run: decode it and turn its whitespace into
    // spaces (line breaks come only from tags).
    let flush = |text: &mut String, out: &mut String| {
        let decoded = decode_entities(text);
        out.extend(decoded.chars().map(|c| if c.is_whitespace() { ' ' } else { c }));
        text.clear();
    };
    let mut rest = html;
    while let Some(lt) = rest.find('<') {
        text.push_str(&rest[..lt]);
        let tag = &rest[lt..];
        if let Some(comment) = tag.strip_prefix("<!--") {
            rest = comment.find("-->").map_or("", |end| &comment[end + 3..]);
            continue;
        }
        let after = tag[1..].chars().next();
        if !after.is_some_and(|c| c.is_ascii_alphabetic() || c == '/' || c == '!' || c == '?') {
            // A bare `<` (e.g. "a < b") is text.
            text.push('<');
            rest = &tag[1..];
            continue;
        }
        let Some(gt) = tag.find('>') else {
            // An unterminated tag runs to the end.
            rest = "";
            break;
        };
        let inner = &tag[1..gt];
        let closing = inner.starts_with('/');
        let name: String = inner
            .trim_start_matches('/')
            .chars()
            .take_while(|c| c.is_ascii_alphanumeric() || *c == ':')
            .collect::<String>()
            .to_ascii_lowercase();
        rest = &tag[gt + 1..];
        match name.as_str() {
            "script" | "style" if !closing => {
                // Skip to the matching close tag (or the end if there is none).
                let lower = rest.to_ascii_lowercase();
                rest = match lower.find(&format!("</{name}")) {
                    Some(end) => rest[end..].find('>').map_or("", |g| &rest[end + g + 1..]),
                    None => "",
                };
            }
            "br" => {
                flush(&mut text, &mut out);
                out.push('\n');
            }
            "p" | "div" | "li" | "tr" | "h1" | "h2" | "h3" | "h4" | "h5" | "h6" if closing => {
                flush(&mut text, &mut out);
                out.push('\n');
            }
            // Keep table cells apart on their row's line.
            "td" | "th" if closing => text.push(' '),
            _ => {}
        }
    }
    text.push_str(rest);
    flush(&mut text, &mut out);
    out
}

/// Up to `lines` lines (clamped to [`MAX_CONTEXT_LINES`]) of plain text
/// immediately preceding the first `cid:<cid>` reference in `html`, joined
/// with `\n`; the tag holding the reference (normally its `<img>`) is
/// excluded. Whitespace runs inside a line collapse to one space and empty
/// lines are dropped. `None` when `html` never references `cid` (matched
/// case-insensitively and as a whole id; `cid` is already normalized).
pub fn text_before_cid(html: &str, cid: &str, lines: u32) -> Option<String> {
    let pos = find_cid_reference(html, cid)?;
    let mut before = &html[..pos];
    // Inside a tag (the reference is an attribute value)? Cut at its `<`.
    if let Some(lt) = before.rfind('<').filter(|&lt| !before[lt..].contains('>')) {
        before = &before[..lt];
    }
    let text = html_to_text(before);
    let all: Vec<String> = text
        .lines()
        .map(|line| line.split_whitespace().collect::<Vec<_>>().join(" "))
        .filter(|line| !line.is_empty())
        .collect();
    let n = lines.min(MAX_CONTEXT_LINES) as usize;
    Some(all[all.len().saturating_sub(n)..].join("\n"))
}

/// One page of a lazy stream of filter matches: drops the first `offset`
/// (negative treated as 0), then keeps up to `count`, and stops pulling from
/// `matches` once the page is full. An error from any pulled item, including
/// a skipped one, is returned as-is.
pub fn take_page<T, E>(
    matches: impl Iterator<Item = Result<T, E>>,
    offset: i32,
    count: i32,
) -> Result<Vec<T>, E> {
    let mut matches = matches;
    for m in matches.by_ref().take(offset.max(0) as usize) {
        m?;
    }
    matches.take(count.max(0) as usize).collect()
}

// ---- inline images (send_email / create_draft) --------------------------

/// `"<cid:logo>"` / `"cid:logo"` / `"logo"` -> `"logo"`.
pub fn normalize_content_id(raw: &str) -> String {
    let mut cid = raw.trim();
    if cid.len() >= 2 && cid.starts_with('<') && cid.ends_with('>') {
        cid = cid[1..cid.len() - 1].trim();
    }
    if cid.get(..4).is_some_and(|p| p.eq_ignore_ascii_case("cid:")) {
        cid = cid[4..].trim();
    }
    cid.to_string()
}

/// Strictly decode inline-image base64, accepting an optional
/// `data:<mime>;base64,` prefix and embedded whitespace/newlines.
pub fn decode_inline_base64(raw: &str, cid: &str) -> Result<Vec<u8>, ToolError> {
    use base64::Engine as _;
    let mut text = raw.trim();
    if text.get(..5).is_some_and(|p| p.eq_ignore_ascii_case("data:")) {
        match text.split_once(',') {
            Some((header, rest)) if header.to_ascii_lowercase().contains(";base64") => text = rest,
            _ => {
                return Err(ToolError::new(format!(
                    "inline image {cid:?}: data URI must be base64-encoded"
                )))
            }
        }
    }
    let compact: String = text.chars().filter(|c| !c.is_whitespace()).collect();
    let data = base64::engine::general_purpose::STANDARD
        .decode(compact.as_bytes())
        .map_err(|_| ToolError::new(format!("inline image {cid:?}: data_base64 is not valid base64")))?;
    if data.is_empty() {
        return Err(ToolError::new(format!("inline image {cid:?}: data_base64 is empty")));
    }
    Ok(data)
}

/// Magic-number sniffing for base64 images given without filename/mime_type.
pub fn sniff_image_mime(data: &[u8]) -> Option<&'static str> {
    if data.starts_with(b"\x89PNG\r\n\x1a\n") {
        Some("image/png")
    } else if data.starts_with(b"\xff\xd8\xff") {
        Some("image/jpeg")
    } else if data.starts_with(b"GIF87a") || data.starts_with(b"GIF89a") {
        Some("image/gif")
    } else if data.len() >= 12 && &data[..4] == b"RIFF" && &data[8..12] == b"WEBP" {
        Some("image/webp")
    } else {
        None
    }
}

/// MIME type from a file name's extension (common image types only).
pub fn mime_from_filename(filename: &str) -> Option<&'static str> {
    let ext = std::path::Path::new(filename).extension()?.to_str()?.to_ascii_lowercase();
    Some(match ext.as_str() {
        "png" => "image/png",
        "jpg" | "jpeg" | "jpe" => "image/jpeg",
        "gif" => "image/gif",
        "webp" => "image/webp",
        "bmp" => "image/bmp",
        "svg" => "image/svg+xml",
        "tif" | "tiff" => "image/tiff",
        "ico" => "image/x-icon",
        _ => return None,
    })
}

/// File extension (with the dot) for a MIME type, or `""` if unknown.
pub fn extension_for_mime(mime: &str) -> &'static str {
    match mime.to_ascii_lowercase().as_str() {
        "image/png" => ".png",
        "image/jpeg" | "image/jpg" => ".jpg",
        "image/gif" => ".gif",
        "image/webp" => ".webp",
        "image/bmp" => ".bmp",
        "image/svg+xml" => ".svg",
        "image/tiff" => ".tif",
        "image/x-icon" | "image/vnd.microsoft.icon" => ".ico",
        _ => "",
    }
}

/// Validate every inline image FIRST (so a bad one fails before any COM
/// item is created, sent or saved) and normalize each to a
/// [`ValidatedInlineImage`]. Requires `html`; each entry needs a non-empty
/// content id (unique, case-insensitively) and exactly one of
/// `path` (must exist) / `data_base64` (must decode). The MIME type comes
/// from `mime_type`, else the filename extension, else the data's magic
/// bytes, else `application/octet-stream`.
pub fn validate_inline_images(images: &[InlineImage], html: bool)
    -> Result<Vec<ValidatedInlineImage>, ToolError> {
    if !html {
        return Err(inline_images_need_html());
    }
    let non_empty = |v: &Option<String>| v.as_deref().map(str::trim).filter(|s| !s.is_empty()).map(str::to_string);
    let mut out = Vec::with_capacity(images.len());
    let mut seen = std::collections::HashSet::new();
    for img in images {
        let cid = normalize_content_id(&img.content_id);
        if cid.is_empty() {
            return Err(ToolError::new("inline image is missing a content_id"));
        }
        if !seen.insert(cid.to_lowercase()) {
            return Err(ToolError::new(format!("duplicate inline image content_id: {cid:?}")));
        }
        let mut filename = non_empty(&img.filename);
        let source = match (non_empty(&img.path), img.data_base64.as_deref().filter(|s| !s.trim().is_empty())) {
            (Some(path), None) => {
                if !std::path::Path::new(&path).is_file() {
                    return Err(ToolError::new(format!("inline image not found: {path}")));
                }
                if filename.is_none() {
                    filename = std::path::Path::new(&path)
                        .file_name()
                        .map(|n| n.to_string_lossy().into_owned());
                }
                InlineImageSource::Path(path)
            }
            (None, Some(data)) => InlineImageSource::Data(decode_inline_base64(data, &cid)?),
            _ => {
                return Err(ToolError::new(format!(
                    "inline image {cid:?}: give exactly one of 'path' or 'data_base64'"
                )))
            }
        };
        let mime_type = non_empty(&img.mime_type)
            .or_else(|| filename.as_deref().and_then(mime_from_filename).map(str::to_string))
            .or_else(|| match &source {
                InlineImageSource::Data(d) => sniff_image_mime(d).map(str::to_string),
                InlineImageSource::Path(_) => None,
            })
            .unwrap_or_else(|| "application/octet-stream".to_string());
        let filename = filename.unwrap_or_else(|| format!("{cid}{}", extension_for_mime(&mime_type)));
        out.push(ValidatedInlineImage {
            content_id: cid,
            source,
            filename: com::safe_filename(&filename),
            mime_type,
        });
    }
    Ok(out)
}

// ---- data: URI images in an HTML body (#29, write half) ------------------

/// One image taken out of an HTML body's `data:` URI by
/// [`extract_data_uri_images`]; attached as a hidden Content-ID attachment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataUriImage {
    /// Generated Content-ID (`img-` + 16 hex digits of a hash of the MIME
    /// type and bytes), so the same image always gets the same id.
    pub content_id: String,
    /// The URI's declared media type, lowercased (always `image/...`).
    pub mime_type: String,
    /// The decoded image bytes.
    pub data: Vec<u8>,
}

/// 64-bit FNV-1a, used to derive a stable Content-ID from an image's bytes.
fn fnv1a64(parts: &[&[u8]]) -> u64 {
    let mut hash: u64 = 0xcbf2_9ce4_8422_2325;
    for part in parts {
        for &b in *part {
            hash ^= u64::from(b);
            hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
        }
    }
    hash
}

/// RFC 3986 percent-decoding to bytes. An invalid escape (like `%zz` or a
/// trailing `%`) is kept literally, as browsers do.
fn percent_decode(s: &str) -> Vec<u8> {
    let bytes = s.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 2 < bytes.len() {
            let hex = |b: u8| (b as char).to_digit(16);
            if let (Some(hi), Some(lo)) = (hex(bytes[i + 1]), hex(bytes[i + 2])) {
                out.push((hi * 16 + lo) as u8);
                i += 3;
                continue;
            }
        }
        out.push(bytes[i]);
        i += 1;
    }
    out
}

/// Where a `data:` URI found in the HTML ends: at the closing quote of a
/// quoted value, or at the first delimiter of an unquoted one.
#[derive(Clone, Copy)]
enum UriEnd {
    Quote(char),
    Unquoted,
}

/// Find every `data:image/...` URI used as an HTML attribute value (quoted
/// with `"` or `'`, or unquoted, e.g. `<img src="data:image/png;base64,...">`)
/// or as a CSS `url(...)` argument, decode it, and replace it with
/// `cid:<content id>`. Returns the rewritten HTML and the images, one per
/// distinct image (the same MIME type and bytes used twice share one
/// Content-ID and one entry).
///
/// - Both `;base64` and percent-encoded (`data:image/svg+xml,%3Csvg...`)
///   payloads are accepted. HTML entities in the value are decoded first,
///   then percent-escapes, then base64 (whitespace such as line wrapping is
///   ignored; `=` padding is optional).
/// - Non-image `data:` URIs (`data:text/plain,...`, `data:application/pdf;...`,
///   a missing media type) are left untouched, as is the text `data:` outside
///   an attribute value or `url(...)`.
/// - Errors, so nothing is created from a half-converted body: an image
///   `data:` URI with no `,`, invalid base64, an empty payload, or an
///   unterminated quoted value.
pub fn extract_data_uri_images(html: &str) -> Result<(String, Vec<DataUriImage>), ToolError> {
    use base64::Engine as _;
    let lenient = base64::engine::GeneralPurpose::new(
        &base64::alphabet::STANDARD,
        base64::engine::GeneralPurposeConfig::new()
            .with_decode_padding_mode(base64::engine::DecodePaddingMode::Indifferent),
    );
    // ASCII lowercasing keeps byte offsets identical to `html`.
    let lower = html.to_ascii_lowercase();
    let mut out = String::with_capacity(html.len());
    let mut images: Vec<DataUriImage> = Vec::new();
    let mut copied = 0; // `html[..copied]` is already in `out`
    let mut search = 0;
    while let Some(rel) = lower[search..].find("data:") {
        let start = search + rel;
        search = start + 5;
        // The URI must start a value: after `="`, `='`, `=`, `("`, `('` or
        // `url(`, optionally with whitespace in between.
        let before = html[..start].trim_end_matches(|c: char| c.is_ascii_whitespace());
        let end_kind = match before.chars().next_back() {
            Some(q @ ('"' | '\'')) => {
                let pre = before[..before.len() - 1].trim_end_matches(|c: char| c.is_ascii_whitespace());
                let opens_value = pre.ends_with('=')
                    || (pre.ends_with('(') && lower[..pre.len() - 1].ends_with("url"));
                if !opens_value {
                    continue;
                }
                UriEnd::Quote(q)
            }
            Some('=') => UriEnd::Unquoted,
            Some('(') if lower[..before.len() - 1].ends_with("url") => UriEnd::Unquoted,
            _ => continue,
        };
        // Only image URIs are converted; check before looking for the end so
        // other data: URIs are never an error.
        let header_lower = lower[start + 5..]
            .split([',', '"', '\'', '>', ')'])
            .next()
            .unwrap_or("");
        let mime_type = header_lower.split(';').next().unwrap_or("").trim().to_string();
        if !mime_type.starts_with("image/") {
            continue;
        }
        let end = match end_kind {
            UriEnd::Quote(q) => match html[start..].find(q) {
                Some(len) => start + html[start..start + len].trim_end().len(),
                None => {
                    return Err(ToolError::new(format!(
                        "html_body: unterminated {mime_type} data: URI (no closing {q})"
                    )))
                }
            },
            UriEnd::Unquoted => html[start..]
                .find(|c: char| c.is_ascii_whitespace() || "\"'<>)`".contains(c))
                .map_or(html.len(), |len| start + len),
        };
        let uri = decode_entities(&html[start + 5..end]);
        let Some((header, payload)) = uri.split_once(',') else {
            return Err(ToolError::new(format!(
                "html_body: malformed {mime_type} data: URI (missing ',' before the data)"
            )));
        };
        let is_base64 = header.split(';').skip(1).any(|p| p.trim().eq_ignore_ascii_case("base64"));
        let mut data = percent_decode(payload);
        if is_base64 {
            data.retain(|b| !b.is_ascii_whitespace());
            data = lenient.decode(&data).map_err(|_| {
                ToolError::new(format!(
                    "html_body: a {mime_type} data: URI is not valid base64 (image #{} in the HTML)",
                    images.len() + 1
                ))
            })?;
        }
        if data.is_empty() {
            return Err(ToolError::new(format!("html_body: a {mime_type} data: URI has no data")));
        }
        let content_id = format!("img-{:016x}", fnv1a64(&[mime_type.as_bytes(), &[0], &data]));
        if !images.iter().any(|i| i.content_id == content_id) {
            images.push(DataUriImage { content_id: content_id.clone(), mime_type, data });
        }
        out.push_str(&html[copied..start]);
        out.push_str("cid:");
        out.push_str(&content_id);
        copied = end;
        search = end;
    }
    out.push_str(&html[copied..]);
    Ok((out, images))
}

/// Everything a mail-writing tool attaches as inline images: the explicit
/// `inline_images` (validated by [`validate_inline_images`]) plus the images
/// pulled out of the HTML's `data:` URIs. Returns the rewritten HTML (data
/// URIs replaced by `cid:` references) and the full list, explicit images
/// first. A generated Content-ID that equals an explicit one is an error.
pub fn prepare_html_images(html: &str, explicit: &[InlineImage])
    -> Result<(String, Vec<ValidatedInlineImage>), ToolError> {
    let mut images = validate_inline_images(explicit, true)?;
    let (html, found) = extract_data_uri_images(html)?;
    for img in found {
        if images.iter().any(|i| i.content_id.eq_ignore_ascii_case(&img.content_id)) {
            return Err(ToolError::new(format!(
                "duplicate inline image content_id: {:?}", img.content_id
            )));
        }
        images.push(ValidatedInlineImage {
            filename: format!("{}{}", img.content_id, extension_for_mime(&img.mime_type)),
            content_id: img.content_id,
            source: InlineImageSource::Data(img.data),
            mime_type: img.mime_type,
        });
    }
    Ok((html, images))
}

/// The final body and inline images for `send_email` / `create_draft` /
/// `reply_email`: an HTML body goes through [`prepare_html_images`]; a
/// plain-text body takes no inline images.
pub fn prepare_mail_body(body: &MailBody, inline_images: &[InlineImage])
    -> Result<(MailBody, Vec<ValidatedInlineImage>), ToolError> {
    match body {
        MailBody::Text(text) => {
            if !inline_images.is_empty() {
                return Err(inline_images_need_html());
            }
            Ok((MailBody::Text(text.clone()), Vec::new()))
        }
        MailBody::Html(html) => {
            let (html, images) = prepare_html_images(html, inline_images)?;
            Ok((MailBody::Html(html), images))
        }
    }
}

fn inline_images_need_html() -> ToolError {
    ToolError::new(
        "inline_images requires an HTML body: pass html_body (or html_body_file) and reference \
         each image as <img src=\"cid:CONTENT_ID\">.",
    )
}

// ---- body / *_file inputs (#28) ------------------------------------------

/// A mail body: plain text or HTML.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MailBody {
    Text(String),
    Html(String),
}

impl Default for MailBody {
    fn default() -> Self {
        MailBody::Text(String::new())
    }
}

impl MailBody {
    pub fn is_html(&self) -> bool {
        matches!(self, MailBody::Html(_))
    }

    pub fn as_str(&self) -> &str {
        match self {
            MailBody::Text(s) | MailBody::Html(s) => s,
        }
    }
}

/// Read a `*_file` input: the whole file as UTF-8 (a leading BOM is dropped).
/// `param` names the parameter in the error.
pub fn read_text_file(path: &str, param: &str) -> Result<String, ToolError> {
    let bytes = std::fs::read(path)
        .map_err(|e| ToolError::new(format!("{param}: could not read {path:?}: {e}")))?;
    let text = String::from_utf8(bytes)
        .map_err(|_| ToolError::new(format!("{param}: {path:?} is not valid UTF-8 text")))?;
    Ok(text.strip_prefix('\u{feff}').map(str::to_string).unwrap_or(text))
}

/// Resolve a text parameter that has a `*_file` sibling (e.g. `body` /
/// `body_file`): at most one may be given; the file is read as UTF-8.
pub fn resolve_text_input(value: Option<String>, file: Option<String>, name: &str, file_name: &str)
    -> Result<Option<String>, ToolError> {
    match (value, file) {
        (Some(_), Some(_)) => Err(ToolError::new(format!(
            "pass either `{name}` or `{file_name}`, not both"
        ))),
        (Some(v), None) => Ok(Some(v)),
        (None, Some(path)) => read_text_file(&path, file_name).map(Some),
        (None, None) => Ok(None),
    }
}

/// The body inputs every mail-writing tool accepts. `html` is the deprecated
/// flag the compose tools used to take (`html: true` + `body` = HTML).
#[derive(Debug, Clone, Default)]
pub struct BodyInputs {
    pub body: Option<String>,
    pub html_body: Option<String>,
    pub body_file: Option<String>,
    pub html_body_file: Option<String>,
    pub html: Option<bool>,
}

/// Turn [`BodyInputs`] into one [`MailBody`] (or `None` when no body source
/// was given). At most one of `body`, `html_body`, `body_file` and
/// `html_body_file` may be given. The deprecated `html: true` turns `body` /
/// `body_file` into HTML; `html: false` together with `html_body` /
/// `html_body_file` is a contradiction and an error. Files are read here, so
/// call this before anything is created.
pub fn resolve_mail_body(b: BodyInputs) -> Result<Option<MailBody>, ToolError> {
    let given: Vec<&str> = [
        ("body", b.body.is_some()),
        ("html_body", b.html_body.is_some()),
        ("body_file", b.body_file.is_some()),
        ("html_body_file", b.html_body_file.is_some()),
    ]
    .into_iter()
    .filter_map(|(name, set)| set.then_some(name))
    .collect();
    if let [first, second, ..] = given[..] {
        return Err(ToolError::new(format!("pass either `{first}` or `{second}`, not both")));
    }
    let html_source = b.html_body.is_some() || b.html_body_file.is_some();
    if html_source && b.html == Some(false) {
        let name = if b.html_body.is_some() { "html_body" } else { "html_body_file" };
        return Err(ToolError::new(format!(
            "`html: false` contradicts `{name}`; drop the deprecated `html` flag"
        )));
    }
    let as_html = html_source || b.html == Some(true);
    let text = if let Some(t) = b.body.or(b.html_body) {
        t
    } else if let Some(path) = &b.body_file {
        read_text_file(path, "body_file")?
    } else if let Some(path) = &b.html_body_file {
        read_text_file(path, "html_body_file")?
    } else {
        return Ok(None);
    };
    Ok(Some(if as_html { MailBody::Html(text) } else { MailBody::Text(text) }))
}

/// [`resolve_mail_body`] for a tool where a body is required.
pub fn require_mail_body(b: BodyInputs, tool: &str) -> Result<MailBody, ToolError> {
    resolve_mail_body(b)?.ok_or_else(|| {
        ToolError::new(format!(
            "{tool} needs a body: pass one of `body`, `html_body`, `body_file` or `html_body_file`"
        ))
    })
}

/// Validate an optional importance word, returning its `OlImportance` id.
pub fn parse_importance(importance: Option<&str>) -> Result<Option<i32>, ToolError> {
    importance
        .map(|imp| {
            crate::constants::importance_name_to_id(imp).ok_or_else(|| {
                ToolError::new(format!(
                    "invalid importance {imp:?}: expected \"low\", \"normal\", or \"high\""
                ))
            })
        })
        .transpose()
}

/// `current` with `add` appended (skipping ones already present) and `remove`
/// dropped, all case-insensitively.
pub fn merge_categories(mut current: Vec<String>, add: &[String], remove: &[String]) -> Vec<String> {
    for a in add {
        if !current.iter().any(|c| c.eq_ignore_ascii_case(a)) {
            current.push(a.clone());
        }
    }
    current.retain(|c| !remove.iter().any(|r| r.eq_ignore_ascii_case(c)));
    current
}

#[cfg(test)]
mod tests {
    use super::{
        com_recurrence_interval, common_free, create_event_status, friendly_recurrence_interval,
        parse_freebusy_slots, take_page, validate_recurrence, validate_recurrence_update,
        EventUpdate, RecurrenceInput, permanent_delete_needs_move, require_empty_confirm,
        draft_update_changes, validate_draft_update, DraftUpdate, InlineImage, merge_categories,
        parse_importance,
    };
    use std::cell::Cell;

    fn ok_items(n: i32) -> impl Iterator<Item = Result<i32, String>> {
        (1..=n).map(Ok)
    }

    #[test]
    fn take_page_first_page_is_the_first_count_items() {
        assert_eq!(take_page(ok_items(10), 0, 3), Ok(vec![1, 2, 3]));
    }

    #[test]
    fn take_page_skips_offset_then_takes_count() {
        assert_eq!(take_page(ok_items(10), 3, 3), Ok(vec![4, 5, 6]));
    }

    #[test]
    fn take_page_consecutive_pages_tile_the_stream() {
        let p1 = take_page(ok_items(10), 0, 5).unwrap();
        let p2 = take_page(ok_items(10), 5, 5).unwrap();
        let both = take_page(ok_items(10), 0, 10).unwrap();
        assert_eq!([p1, p2].concat(), both);
    }

    #[test]
    fn take_page_last_page_is_short_and_past_the_end_is_empty() {
        assert_eq!(take_page(ok_items(7), 5, 5), Ok(vec![6, 7]));
        assert_eq!(take_page(ok_items(7), 7, 5), Ok(vec![]));
        assert_eq!(take_page(ok_items(7), 100, 5), Ok(vec![]));
    }

    #[test]
    fn take_page_negative_offset_is_treated_as_zero() {
        assert_eq!(take_page(ok_items(10), -4, 2), Ok(vec![1, 2]));
    }

    #[test]
    fn take_page_stops_pulling_once_the_page_is_full() {
        let pulled = Cell::new(0);
        let it = (1..=100).map(|i| {
            pulled.set(pulled.get() + 1);
            Ok::<i32, String>(i)
        });
        assert_eq!(take_page(it, 2, 3), Ok(vec![3, 4, 5]));
        assert_eq!(pulled.get(), 5);
    }

    #[test]
    fn take_page_propagates_errors_from_skipped_and_kept_items() {
        let with_err = |bad: i32| (1..=10).map(move |i| if i == bad { Err(format!("bad {i}")) } else { Ok(i) });
        assert_eq!(take_page(with_err(2), 3, 3), Err("bad 2".to_string()));
        assert_eq!(take_page(with_err(5), 3, 3), Err("bad 5".to_string()));
        // An error beyond the page is never pulled.
        assert_eq!(take_page(with_err(9), 3, 3), Ok(vec![4, 5, 6]));
    }
    use crate::outlook::types::{AvailabilitySlot, FreeWindow, PersonAvailability};

    #[test]
    fn require_empty_confirm_refuses_without_confirm() {
        let err = require_empty_confirm(false).unwrap_err();
        assert!(err.to_string().contains("confirm=true"));
        assert!(err.to_string().contains("cannot be undone"));
        assert!(require_empty_confirm(true).is_ok());
    }

    #[test]
    fn permanent_delete_skips_move_only_when_already_in_deleted_items() {
        assert!(!permanent_delete_needs_move("DELETED-ID", "DELETED-ID"));
        assert!(permanent_delete_needs_move("DRAFTS-ID", "DELETED-ID"));
        // An unreadable (empty) parent id never counts as "already there".
        assert!(permanent_delete_needs_move("", ""));
    }

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

    #[test]
    fn com_recurrence_interval_multiplies_yearly_by_12() {
        assert_eq!(com_recurrence_interval(&recurrence("yearly")), 12); // default interval 1 -> 12
        let mut r = recurrence("yearly");
        r.interval = Some(2);
        assert_eq!(com_recurrence_interval(&r), 24); // every 2 years -> 24 months

        assert_eq!(com_recurrence_interval(&recurrence("daily")), 1);
        let mut r = recurrence("weekly");
        r.interval = Some(3);
        assert_eq!(com_recurrence_interval(&r), 3); // unchanged for non-yearly
    }

    #[test]
    fn friendly_recurrence_interval_divides_yearly_by_12() {
        assert_eq!(
            friendly_recurrence_interval(crate::constants::OL_RECURS_YEARLY, 12),
            1
        );
        assert_eq!(
            friendly_recurrence_interval(crate::constants::OL_RECURS_YEARLY, 24),
            2
        );
        assert_eq!(
            friendly_recurrence_interval(crate::constants::OL_RECURS_DAILY, 1),
            1
        );
        assert_eq!(
            friendly_recurrence_interval(crate::constants::OL_RECURS_WEEKLY, 3),
            3
        );
    }

    fn event_update() -> EventUpdate {
        EventUpdate {
            event_id: "event-1".to_string(),
            subject: None, start: None, end: None, location: None, body: None,
            all_day: None, reminder_minutes: None, show_as: None,
            add_categories: None, remove_categories: None,
            add_required_attendees: None, add_optional_attendees: None, remove_attendees: None,
            send_update: false,
            recurrence: None, clear_recurrence: false,
        }
    }

    #[test]
    fn validate_recurrence_update_rejects_recurrence_and_clear_recurrence_together() {
        let mut u = event_update();
        u.recurrence = Some(recurrence("daily"));
        u.clear_recurrence = true;
        assert!(validate_recurrence_update(&u).is_err());
    }

    #[test]
    fn validate_recurrence_update_accepts_recurrence_only() {
        let mut u = event_update();
        u.recurrence = Some(recurrence("daily"));
        assert!(validate_recurrence_update(&u).is_ok());
    }

    #[test]
    fn validate_recurrence_update_accepts_clear_recurrence_only() {
        let mut u = event_update();
        u.clear_recurrence = true;
        assert!(validate_recurrence_update(&u).is_ok());
    }

    #[test]
    fn validate_recurrence_update_accepts_neither() {
        assert!(validate_recurrence_update(&event_update()).is_ok());
    }

    fn dt(s: &str) -> chrono::NaiveDateTime {
        chrono::NaiveDateTime::parse_from_str(s, "%Y-%m-%dT%H:%M:%S").unwrap()
    }

    #[test]
    fn parse_freebusy_slots_maps_codes_to_words_and_times() {
        // "02143" = free, busy, tentative, working_elsewhere, out_of_office
        let slots = parse_freebusy_slots("02143", &dt("2099-01-01T09:00:00"), 30, 5);
        assert_eq!(slots.len(), 5);
        assert_eq!(slots[0], AvailabilitySlot {
            start: "2099-01-01T09:00:00".to_string(),
            end: "2099-01-01T09:30:00".to_string(),
            status: "free".to_string(),
        });
        assert_eq!(slots[1].status, "busy");
        assert_eq!(slots[1].start, "2099-01-01T09:30:00");
        assert_eq!(slots[1].end, "2099-01-01T10:00:00");
        assert_eq!(slots[2].status, "tentative");
        assert_eq!(slots[3].status, "working_elsewhere");
        assert_eq!(slots[4].status, "out_of_office");
    }

    #[test]
    fn parse_freebusy_slots_truncates_to_max_slots() {
        // Outlook's raw FreeBusy string commonly covers a much longer range
        // than the caller's requested [start, end) window.
        let slots = parse_freebusy_slots("000000000000", &dt("2099-01-01T09:00:00"), 30, 3);
        assert_eq!(slots.len(), 3);
    }

    #[test]
    fn parse_freebusy_slots_treats_unrecognized_digit_as_busy() {
        let slots = parse_freebusy_slots("9", &dt("2099-01-01T09:00:00"), 30, 1);
        assert_eq!(slots[0].status, "busy");
    }

    fn avail(person: &str, resolved: bool, statuses: &[&str]) -> PersonAvailability {
        let mut slots = Vec::new();
        for (i, s) in statuses.iter().enumerate() {
            let start = dt("2099-01-01T09:00:00") + chrono::Duration::minutes(i as i64 * 30);
            slots.push(AvailabilitySlot {
                start: start.format("%Y-%m-%dT%H:%M:%S").to_string(),
                end: (start + chrono::Duration::minutes(30)).format("%Y-%m-%dT%H:%M:%S").to_string(),
                status: s.to_string(),
            });
        }
        PersonAvailability { person: person.to_string(), resolved, slots }
    }

    #[test]
    fn common_free_intersects_only_where_everyone_is_free() {
        let people = vec![
            avail("alice", true, &["free", "free", "busy"]),
            avail("bob", true, &["free", "busy", "busy"]),
        ];
        let windows = common_free(&people, &["free".to_string()]);
        assert_eq!(windows, vec![FreeWindow {
            start: "2099-01-01T09:00:00".to_string(),
            end: "2099-01-01T09:30:00".to_string(),
        }]);
    }

    #[test]
    fn common_free_merges_contiguous_free_slots_into_one_window() {
        let people = vec![avail("alice", true, &["free", "free", "busy", "free"])];
        let windows = common_free(&people, &["free".to_string()]);
        assert_eq!(windows, vec![
            FreeWindow { start: "2099-01-01T09:00:00".to_string(), end: "2099-01-01T10:00:00".to_string() },
            FreeWindow { start: "2099-01-01T10:30:00".to_string(), end: "2099-01-01T11:00:00".to_string() },
        ]);
    }

    #[test]
    fn common_free_respects_custom_treat_as_free() {
        let people = vec![avail("alice", true, &["tentative"])];
        assert_eq!(common_free(&people, &["free".to_string()]), vec![]);
        assert_eq!(
            common_free(&people, &["free".to_string(), "tentative".to_string()]).len(),
            1
        );
    }

    #[test]
    fn common_free_ignores_unresolved_people() {
        let people = vec![
            avail("alice", true, &["free"]),
            avail("bob", false, &[]),
        ];
        assert_eq!(common_free(&people, &["free".to_string()]).len(), 1);
    }

    #[test]
    fn common_free_empty_when_no_one_resolved() {
        let people = vec![avail("alice", false, &[])];
        assert_eq!(common_free(&people, &["free".to_string()]), vec![]);
    }

    fn draft(id: &str) -> DraftUpdate {
        DraftUpdate { email_id: id.to_string(), ..Default::default() }
    }

    #[test]
    fn validate_draft_update_rejects_body_and_html_body_together() {
        let u = DraftUpdate {
            body: Some("plain".into()), html_body: Some("<p>html</p>".into()), ..draft("d")
        };
        assert!(validate_draft_update(&u).unwrap_err().to_string().contains("not both"));
    }

    #[test]
    fn validate_draft_update_rejects_an_empty_update() {
        assert!(validate_draft_update(&draft("d")).is_err());
        // An empty attachments list changes nothing either.
        let u = DraftUpdate { attachments: Some(vec![]), ..draft("d") };
        assert!(validate_draft_update(&u).is_err());
    }

    #[test]
    fn validate_draft_update_accepts_a_single_field() {
        assert!(validate_draft_update(&DraftUpdate { subject: Some("s".into()), ..draft("d") }).is_ok());
        assert!(validate_draft_update(&DraftUpdate { html_body: Some("<b>x</b>".into()), ..draft("d") }).is_ok());
        // `[]` clears a recipient line, so it counts as a change.
        assert!(validate_draft_update(&DraftUpdate { cc: Some(vec![]), ..draft("d") }).is_ok());
    }

    #[test]
    fn validate_draft_update_rejects_a_missing_attachment() {
        let u = DraftUpdate {
            subject: Some("s".into()),
            attachments: Some(vec!["/definitely/not/here/outlook-mcp-rs.txt".into()]),
            ..draft("d")
        };
        assert!(validate_draft_update(&u).unwrap_err().to_string().contains("attachment not found"));
    }

    #[test]
    fn validate_draft_update_accepts_an_existing_attachment() {
        let path = std::env::temp_dir().join("outlook-mcp-rs-validate-draft-update.txt");
        std::fs::write(&path, b"x").unwrap();
        let u = DraftUpdate {
            attachments: Some(vec![path.to_string_lossy().to_string()]), ..draft("d")
        };
        let res = validate_draft_update(&u);
        let _ = std::fs::remove_file(&path);
        assert!(res.is_ok());
    }

    #[test]
    fn draft_update_changes_lists_fields_in_apply_order() {
        let u = DraftUpdate {
            bcc: Some(vec![]), subject: Some("s".into()), to: Some(vec!["a@x.com".into()]),
            body: Some("b".into()), attachments: Some(vec!["a.txt".into()]), ..draft("d")
        };
        assert_eq!(draft_update_changes(&u), vec!["subject", "body", "to", "bcc", "attachments"]);
        assert!(draft_update_changes(&DraftUpdate { attachments: Some(vec![]), ..draft("d") }).is_empty());
        let u = DraftUpdate {
            importance: Some("high".into()), remove_categories: Some(vec!["Old".into()]),
            add_categories: Some(vec!["New".into()]), html_body: Some("<p>x</p>".into()),
            inline_images: Some(vec![InlineImage { content_id: "a".into(), ..Default::default() }]),
            ..draft("d")
        };
        assert_eq!(
            draft_update_changes(&u),
            vec!["html_body", "inline_images", "add_categories", "remove_categories", "importance"]
        );
    }

    #[test]
    fn validate_draft_update_send_alone_is_enough() {
        assert!(validate_draft_update(&DraftUpdate { send: true, ..draft("d") }).is_ok());
        assert!(draft_update_changes(&DraftUpdate { send: true, ..draft("d") }).is_empty());
    }

    #[test]
    fn validate_draft_update_checks_importance_and_inline_images() {
        let u = DraftUpdate { importance: Some("urgent".into()), ..draft("d") };
        assert!(validate_draft_update(&u).unwrap_err().to_string().contains("invalid importance"));
        let img = InlineImage { content_id: "a".into(), data_base64: Some("aGk=".into()), ..Default::default() };
        let u = DraftUpdate { inline_images: Some(vec![img.clone()]), body: Some("plain".into()), ..draft("d") };
        assert!(validate_draft_update(&u).unwrap_err().to_string().contains("html_body"));
        let u = DraftUpdate { inline_images: Some(vec![img]), html_body: Some("<p/>".into()), ..draft("d") };
        assert!(validate_draft_update(&u).is_ok());
    }

    #[test]
    fn merge_categories_adds_and_removes_caselessly() {
        let cur = vec!["Red".to_string(), "Blue".to_string()];
        let got = merge_categories(cur, &["red".into(), "Green".into()], &["BLUE".into()]);
        assert_eq!(got, vec!["Red".to_string(), "Green".to_string()]);
    }

    #[test]
    fn parse_importance_maps_and_rejects() {
        assert_eq!(parse_importance(None).unwrap(), None);
        assert_eq!(parse_importance(Some("High")).unwrap(), Some(crate::constants::OL_IMPORTANCE_HIGH));
        assert!(parse_importance(Some("meh")).is_err());
    }
}

#[cfg(test)]
mod body_input_tests {
    use super::{read_text_file, require_mail_body, resolve_mail_body, resolve_text_input, BodyInputs, MailBody};

    fn temp_file(name: &str, bytes: &[u8]) -> String {
        let dir = std::env::temp_dir().join(format!("outlook-mcp-rs-body-{}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join(name);
        std::fs::write(&path, bytes).unwrap();
        path.to_string_lossy().into_owned()
    }

    fn inputs() -> BodyInputs {
        BodyInputs::default()
    }

    #[test]
    fn plain_and_html_strings() {
        let b = resolve_mail_body(BodyInputs { body: Some("hi".into()), ..inputs() }).unwrap();
        assert_eq!(b, Some(MailBody::Text("hi".into())));
        let b = resolve_mail_body(BodyInputs { html_body: Some("<p>hi</p>".into()), ..inputs() }).unwrap();
        assert_eq!(b, Some(MailBody::Html("<p>hi</p>".into())));
        assert_eq!(resolve_mail_body(inputs()).unwrap(), None);
    }

    #[test]
    fn deprecated_html_flag() {
        // html=true + body = HTML.
        let b = resolve_mail_body(BodyInputs { body: Some("<b>x</b>".into()), html: Some(true), ..inputs() }).unwrap();
        assert_eq!(b, Some(MailBody::Html("<b>x</b>".into())));
        // html=false + body = text; html=true + html_body is redundant but fine.
        let b = resolve_mail_body(BodyInputs { body: Some("x".into()), html: Some(false), ..inputs() }).unwrap();
        assert_eq!(b, Some(MailBody::Text("x".into())));
        let b = resolve_mail_body(BodyInputs { html_body: Some("<p/>".into()), html: Some(true), ..inputs() }).unwrap();
        assert_eq!(b, Some(MailBody::Html("<p/>".into())));
        // html=false contradicts html_body.
        let e = resolve_mail_body(BodyInputs { html_body: Some("<p/>".into()), html: Some(false), ..inputs() })
            .unwrap_err().to_string();
        assert!(e.contains("`html: false`") && e.contains("`html_body`"), "{e}");
    }

    #[test]
    fn sources_are_mutually_exclusive() {
        let e = resolve_mail_body(BodyInputs { body: Some("a".into()), html_body: Some("b".into()), ..inputs() })
            .unwrap_err().to_string();
        assert_eq!(e, "pass either `body` or `html_body`, not both");
        let e = resolve_mail_body(BodyInputs { html_body: Some("a".into()), html_body_file: Some("f".into()), ..inputs() })
            .unwrap_err().to_string();
        assert_eq!(e, "pass either `html_body` or `html_body_file`, not both");
        let e = resolve_mail_body(BodyInputs { body_file: Some("a".into()), html_body_file: Some("f".into()), ..inputs() })
            .unwrap_err().to_string();
        assert_eq!(e, "pass either `body_file` or `html_body_file`, not both");
    }

    #[test]
    fn files_are_read_as_utf8() {
        let html = temp_file("body.html", "\u{feff}<p>שלום</p>".as_bytes());
        let b = resolve_mail_body(BodyInputs { html_body_file: Some(html), ..inputs() }).unwrap();
        assert_eq!(b, Some(MailBody::Html("<p>שלום</p>".into())));
        let txt = temp_file("body.txt", b"plain");
        let b = resolve_mail_body(BodyInputs { body_file: Some(txt.clone()), ..inputs() }).unwrap();
        assert_eq!(b, Some(MailBody::Text("plain".into())));
        let b = resolve_mail_body(BodyInputs { body_file: Some(txt), html: Some(true), ..inputs() }).unwrap();
        assert_eq!(b, Some(MailBody::Html("plain".into())));
    }

    #[test]
    fn bad_files_error_with_the_parameter_name() {
        let e = read_text_file("/definitely/not/here.html", "html_body_file").unwrap_err().to_string();
        assert!(e.starts_with("html_body_file: could not read"), "{e}");
        let bin = temp_file("bin.dat", &[0xff, 0xfe, 0x00, 0x80]);
        let e = read_text_file(&bin, "body_file").unwrap_err().to_string();
        assert!(e.contains("not valid UTF-8"), "{e}");
    }

    #[test]
    fn required_body() {
        let e = require_mail_body(inputs(), "send_email").unwrap_err().to_string();
        assert!(e.starts_with("send_email needs a body"), "{e}");
        assert_eq!(require_mail_body(BodyInputs { body: Some(String::new()), ..inputs() }, "x").unwrap(),
            MailBody::Text(String::new()));
    }

    #[test]
    fn text_input_with_file_sibling() {
        assert_eq!(resolve_text_input(None, None, "body", "body_file").unwrap(), None);
        assert_eq!(resolve_text_input(Some("a".into()), None, "body", "body_file").unwrap(), Some("a".into()));
        let f = temp_file("note.txt", b"from file");
        assert_eq!(resolve_text_input(None, Some(f.clone()), "body", "body_file").unwrap(), Some("from file".into()));
        let e = resolve_text_input(Some("a".into()), Some(f), "body", "body_file").unwrap_err().to_string();
        assert_eq!(e, "pass either `body` or `body_file`, not both");
    }
}

#[cfg(test)]
mod data_uri_tests {
    use super::{
        extract_data_uri_images, prepare_html_images, prepare_mail_body, InlineImage,
        InlineImageSource, MailBody,
    };

    /// A 1x1 transparent PNG.
    const PNG: &str = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";
    /// "GIF89a" + filler.
    const GIF: &str = "R0lGODlhAQABAAAAACw=";

    fn png_bytes() -> Vec<u8> {
        use base64::Engine as _;
        base64::engine::general_purpose::STANDARD.decode(PNG).unwrap()
    }

    #[test]
    fn double_quoted_src_is_rewritten() {
        let html = format!("<p>Hi</p><img src=\"data:image/png;base64,{PNG}\" alt=\"x\">");
        let (out, imgs) = extract_data_uri_images(&html).unwrap();
        assert_eq!(imgs.len(), 1);
        let cid = &imgs[0].content_id;
        assert!(cid.starts_with("img-") && cid.len() == 20, "{cid}");
        assert_eq!(out, format!("<p>Hi</p><img src=\"cid:{cid}\" alt=\"x\">"));
        assert_eq!(imgs[0].mime_type, "image/png");
        assert_eq!(imgs[0].data, png_bytes());
    }

    #[test]
    fn single_quotes_whitespace_and_case() {
        let html = format!("<IMG SRC = '  DATA:Image/PNG;BASE64,{}\n{}  ' >", &PNG[..30], &PNG[30..]);
        let (out, imgs) = extract_data_uri_images(&html).unwrap();
        assert_eq!(imgs.len(), 1);
        assert_eq!(imgs[0].mime_type, "image/png");
        assert_eq!(imgs[0].data, png_bytes());
        assert_eq!(out, format!("<IMG SRC = '  cid:{}  ' >", imgs[0].content_id));
    }

    #[test]
    fn unquoted_attribute_and_css_url() {
        let html = format!(
            "<img src=data:image/png;base64,{PNG}><div style=\"background:url(data:image/gif;base64,{GIF})\">\
             <td style='background-image: url(\"data:image/png;base64,{PNG}\")'>"
        );
        let (out, imgs) = extract_data_uri_images(&html).unwrap();
        assert_eq!(imgs.len(), 2, "same PNG twice shares one entry");
        let (png, gif) = (&imgs[0].content_id, &imgs[1].content_id);
        assert_eq!(imgs[1].mime_type, "image/gif");
        assert_eq!(
            out,
            format!(
                "<img src=cid:{png}><div style=\"background:url(cid:{gif})\">\
                 <td style='background-image: url(\"cid:{png}\")'>"
            )
        );
    }

    #[test]
    fn duplicates_reuse_one_content_id() {
        let html = format!("<img src=\"data:image/png;base64,{PNG}\"><img src='data:image/png;base64,{PNG}'>");
        let (out, imgs) = extract_data_uri_images(&html).unwrap();
        assert_eq!(imgs.len(), 1);
        assert_eq!(out.matches(&format!("cid:{}", imgs[0].content_id)).count(), 2);
        // Same bytes under another MIME type is a different image.
        let html = format!("<img src=\"data:image/png;base64,{PNG}\"><img src=\"data:image/x-png;base64,{PNG}\">");
        assert_eq!(extract_data_uri_images(&html).unwrap().1.len(), 2);
        // Stable across calls.
        let a = extract_data_uri_images(&format!("<img src=\"data:image/png;base64,{PNG}\">")).unwrap().1;
        assert_eq!(a[0].content_id, imgs[0].content_id);
    }

    #[test]
    fn url_encoded_payload() {
        let html = "<img src=\"data:image/svg+xml,%3Csvg%20xmlns='http://www.w3.org/2000/svg'%3E%3C/svg%3E\">";
        // The ' inside the double-quoted value does not end it.
        let (out, imgs) = extract_data_uri_images(html).unwrap();
        assert_eq!(imgs[0].mime_type, "image/svg+xml");
        assert_eq!(imgs[0].data, b"<svg xmlns='http://www.w3.org/2000/svg'></svg>");
        assert_eq!(out, format!("<img src=\"cid:{}\">", imgs[0].content_id));
        // Percent-escaped base64 and missing padding are accepted; invalid
        // escapes are kept literally.
        let esc = PNG.replace('+', "%2B").trim_end_matches('=').to_string();
        let (_, imgs) = extract_data_uri_images(&format!("<img src=\"data:image/png;base64,{esc}\">")).unwrap();
        assert_eq!(imgs[0].data, png_bytes());
        let (_, imgs) = extract_data_uri_images("<img src=\"data:image/x-raw,a%zzb%4\">").unwrap();
        assert_eq!(imgs[0].data, b"a%zzb%4");
    }

    #[test]
    fn html_entities_in_the_value_are_decoded() {
        let (_, imgs) = extract_data_uri_images("<img src=\"data:image/x-raw,a&amp;b\">").unwrap();
        assert_eq!(imgs[0].data, b"a&b");
    }

    #[test]
    fn invalid_base64_is_an_error() {
        let e = extract_data_uri_images("<img src=\"data:image/png;base64,@@not-base64@@\">").unwrap_err().to_string();
        assert!(e.contains("not valid base64"), "{e}");
        let e = extract_data_uri_images("<img src=\"data:image/png;base64,\">").unwrap_err().to_string();
        assert!(e.contains("no data"), "{e}");
        let e = extract_data_uri_images("<img src=\"data:image/png;base64\">").unwrap_err().to_string();
        assert!(e.contains("missing ','"), "{e}");
        let e = extract_data_uri_images(&format!("<img src=\"data:image/png;base64,{PNG}")).unwrap_err().to_string();
        assert!(e.contains("unterminated"), "{e}");
    }

    #[test]
    fn non_image_and_non_attribute_data_uris_are_left_alone() {
        for html in [
            "<a href=\"data:text/plain;base64,aGk=\">x</a>",
            "<a href=\"data:application/pdf;base64,!!!\">x</a>",
            "<a href=\"data:;base64,aGk=\">x</a>",
            "<a href=\"data:,hello\">x</a>",
            "<p>Paste data:image/png;base64,iVBOR here</p>",
            "<p>He said \"data:image/png;base64,iVBOR\"</p>",
            "<p>func(data:image/png;base64,iVBOR)</p>",
            "<p>no uris at all — שלום</p>",
            "",
        ] {
            let (out, imgs) = extract_data_uri_images(html).unwrap();
            assert_eq!(out, html);
            assert!(imgs.is_empty(), "{html}");
        }
    }

    #[test]
    fn multibyte_text_around_uris_survives() {
        let html = format!("<p>שלום</p><img src=\"data:image/png;base64,{PNG}\"><p>עולם</p>");
        let (out, imgs) = extract_data_uri_images(&html).unwrap();
        assert_eq!(out, format!("<p>שלום</p><img src=\"cid:{}\"><p>עולם</p>", imgs[0].content_id));
    }

    #[test]
    fn prepare_html_images_merges_explicit_and_embedded() {
        let html = format!("<img src=\"cid:logo\"><img src=\"data:image/png;base64,{PNG}\">");
        let explicit = [InlineImage { content_id: "logo".into(), data_base64: Some(GIF.into()), ..Default::default() }];
        let (out, imgs) = prepare_html_images(&html, &explicit).unwrap();
        assert_eq!(imgs.len(), 2);
        assert_eq!(imgs[0].content_id, "logo");
        assert_eq!(imgs[1].mime_type, "image/png");
        assert_eq!(imgs[1].filename, format!("{}.png", imgs[1].content_id));
        assert!(matches!(&imgs[1].source, InlineImageSource::Data(d) if *d == png_bytes()));
        assert_eq!(out, format!("<img src=\"cid:logo\"><img src=\"cid:{}\">", imgs[1].content_id));
        // An explicit image that reuses a generated id is a duplicate.
        let cid = imgs[1].content_id.clone();
        let clash = [InlineImage { content_id: cid.to_uppercase(), data_base64: Some(GIF.into()), ..Default::default() }];
        assert!(prepare_html_images(&html, &clash).unwrap_err().to_string().contains("duplicate"));
    }

    #[test]
    fn prepare_mail_body_text_vs_html() {
        let img = [InlineImage { content_id: "a".into(), data_base64: Some(GIF.into()), ..Default::default() }];
        let e = prepare_mail_body(&MailBody::Text("x".into()), &img).unwrap_err().to_string();
        assert!(e.contains("requires an HTML body"), "{e}");
        // A text body is never scanned for data: URIs.
        let text = format!("src=\"data:image/png;base64,{PNG}\"");
        let (body, imgs) = prepare_mail_body(&MailBody::Text(text.clone()), &[]).unwrap();
        assert_eq!(body, MailBody::Text(text));
        assert!(imgs.is_empty());
        let (body, imgs) = prepare_mail_body(&MailBody::Html("<img src=\"cid:a\">".into()), &img).unwrap();
        assert_eq!(body, MailBody::Html("<img src=\"cid:a\">".into()));
        assert_eq!(imgs.len(), 1);
    }
}

#[cfg(test)]
mod inline_context_tests {
    use super::{text_before_cid, MAX_CONTEXT_LINES};

    fn ctx(html: &str, cid: &str, lines: u32) -> Option<String> {
        text_before_cid(html, cid, lines)
    }

    #[test]
    fn paragraph_before_an_image() {
        let html = "<html><body><p>Hello team,</p><p>Here is the new logo:</p>\
                    <p><img width=10 src=\"cid:image001.png@01D9ABCD\" alt=\"logo\"></p>\
                    <p>Thanks</p></body></html>";
        assert_eq!(ctx(html, "image001.png@01D9ABCD", 1).as_deref(), Some("Here is the new logo:"));
        assert_eq!(
            ctx(html, "image001.png@01D9ABCD", 5).as_deref(),
            Some("Hello team,\nHere is the new logo:")
        );
    }

    #[test]
    fn br_separated_lines_and_source_newlines_are_whitespace() {
        let html = "<div>line one<br>line\n   two<BR/>line three<br />\n<img src='cid:a@b'></div>";
        assert_eq!(ctx(html, "a@b", 2).as_deref(), Some("line two\nline three"));
        assert_eq!(ctx(html, "a@b", 3).as_deref(), Some("line one\nline two\nline three"));
    }

    #[test]
    fn decodes_entities() {
        let html = "<p>Fish&nbsp;&amp;&nbsp;chips &lt;b&gt; &quot;q&quot; it&#39;s &#65;&#x42; &bogus; &amp</p>\
                    <img src=\"cid:x@y\">";
        assert_eq!(
            ctx(html, "x@y", 1).as_deref(),
            Some("Fish & chips <b> \"q\" it's AB &bogus; &amp")
        );
    }

    #[test]
    fn ignores_style_script_and_comments() {
        let html = "<head><style>p { color: red; }\n.x { }</style>\
                    <script type=\"text/javascript\">var a = 1 < 2;</script></head>\
                    <!--[if gte mso 9]><xml><o:shapedefaults /></xml><![endif]-->\
                    <body><p>Caption<o:p></o:p></p><img src=\"cid:x@y\"></body>";
        assert_eq!(ctx(html, "x@y", 10).as_deref(), Some("Caption"));
    }

    #[test]
    fn block_closers_split_lines() {
        let html = "<h1>Title</h1><ul><li>one</li><li>two</li></ul>\
                    <table><tr><td>a</td><td>b</td></tr></table><div>last</div><img src=\"cid:x@y\">";
        assert_eq!(ctx(html, "x@y", 10).as_deref(), Some("Title\none\ntwo\na b\nlast"));
    }

    #[test]
    fn n_larger_than_available_returns_everything() {
        let html = "<p>only</p><img src=\"cid:x@y\">";
        assert_eq!(ctx(html, "x@y", 40).as_deref(), Some("only"));
        assert_eq!(ctx("<img src=\"cid:x@y\">", "x@y", 3).as_deref(), Some(""));
        assert_eq!(ctx(html, "x@y", 0).as_deref(), Some(""));
    }

    #[test]
    fn lines_are_clamped_to_the_maximum() {
        let body: String = (0..80).map(|i| format!("<p>l{i}</p>")).collect();
        let html = format!("{body}<img src=\"cid:x@y\">");
        let got = ctx(&html, "x@y", 1000).unwrap();
        assert_eq!(got.lines().count(), MAX_CONTEXT_LINES as usize);
        assert!(got.ends_with("l79"));
    }

    #[test]
    fn unreferenced_cid_is_none() {
        assert_eq!(ctx("<p>text</p><img src=\"cid:other@y\">", "x@y", 3), None);
        assert_eq!(ctx("", "x@y", 3), None);
        assert_eq!(ctx("<p>x@y</p>", "x@y", 3), None);
    }

    #[test]
    fn reference_matches_case_insensitively() {
        let html = "<p>Look:</p><IMG SRC=\"CID:Image001.PNG@01D9\">";
        assert_eq!(ctx(html, "image001.png@01d9", 1).as_deref(), Some("Look:"));
    }

    #[test]
    fn prefix_collisions_do_not_match() {
        let html = "<p>first</p><img src=\"cid:logo.png@01D9AB\"><p>second</p><img src=\"cid:logo.png@01D9\">\
                    <p>third</p><img src=\"cid:logo\">";
        assert_eq!(ctx(html, "logo.png@01D9", 1).as_deref(), Some("second"));
        assert_eq!(ctx(html, "logo", 1).as_deref(), Some("third"));
        assert_eq!(ctx(html, "logo.png@01D9A", 1), None);
        // Not part of a longer scheme-like word.
        assert_eq!(ctx("<p>a</p>xcid:z@y", "z@y", 1), None);
    }

    #[test]
    fn first_reference_wins() {
        let html = "<p>one</p><img src=\"cid:x@y\"><p>two</p><img src=\"cid:x@y\">";
        assert_eq!(ctx(html, "x@y", 5).as_deref(), Some("one"));
    }

    #[test]
    fn hebrew_text_survives() {
        let html = "<div dir=\"rtl\"><p>שלום לכולם,</p><p>הנה&nbsp;התמונה:</p></div><img src=\"cid:img@x\">";
        assert_eq!(ctx(html, "img@x", 2).as_deref(), Some("שלום לכולם,\nהנה התמונה:"));
        // An entity right before multi-byte text still decodes.
        let html = "<p>&amp;שלום&#x5D0;&nbsp;עולם</p><img src=\"cid:img@x\">";
        assert_eq!(ctx(html, "img@x", 1).as_deref(), Some("&שלוםא עולם"));
    }
}

#[cfg(test)]
mod inline_image_tests {
    use super::{
        decode_inline_base64, extension_for_mime, mime_from_filename, normalize_content_id,
        sniff_image_mime, validate_inline_images, InlineImage, InlineImageSource,
    };

    /// A 1x1 transparent PNG.
    const PNG_1X1_B64: &str = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

    fn b64(cid: &str, data: &str) -> InlineImage {
        InlineImage { content_id: cid.into(), data_base64: Some(data.into()), ..Default::default() }
    }

    fn temp_file(name: &str, bytes: &[u8]) -> std::path::PathBuf {
        let dir = std::env::temp_dir().join(format!("outlook-mcp-rs-unit-{}-{name}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join(name);
        std::fs::write(&path, bytes).unwrap();
        path
    }

    fn err(images: &[InlineImage], html: bool) -> String {
        validate_inline_images(images, html).unwrap_err().to_string()
    }

    #[test]
    fn normalize_content_id_strips_brackets_and_cid_prefix() {
        assert_eq!(normalize_content_id("logo"), "logo");
        assert_eq!(normalize_content_id("  <logo> "), "logo");
        assert_eq!(normalize_content_id("cid:logo"), "logo");
        assert_eq!(normalize_content_id("CID: logo"), "logo");
        assert_eq!(normalize_content_id("<cid:logo>"), "logo");
        assert_eq!(normalize_content_id("<>"), "");
        assert_eq!(normalize_content_id("cid:"), "");
    }

    #[test]
    fn decode_accepts_data_uri_prefix_and_whitespace() {
        let plain = decode_inline_base64(PNG_1X1_B64, "x").unwrap();
        let uri = decode_inline_base64(&format!("data:image/png;base64,{PNG_1X1_B64}"), "x").unwrap();
        let spaced = format!(" {}\n{} \r\n", &PNG_1X1_B64[..20], &PNG_1X1_B64[20..]);
        assert_eq!(plain, uri);
        assert_eq!(plain, decode_inline_base64(&spaced, "x").unwrap());
        assert_eq!(sniff_image_mime(&plain), Some("image/png"));
    }

    #[test]
    fn decode_rejects_invalid_or_empty_data() {
        assert!(decode_inline_base64("not base64!!", "x").unwrap_err().to_string().contains("not valid base64"));
        // Missing padding is rejected (strict decoding).
        assert!(decode_inline_base64("aGk", "x").is_err());
        assert!(decode_inline_base64("data:image/png,aGk=", "x").unwrap_err().to_string().contains("data URI"));
        assert!(decode_inline_base64("data:image/png;base64", "x").is_err());
        assert!(decode_inline_base64("data:image/png;base64,", "x").unwrap_err().to_string().contains("empty"));
    }

    #[test]
    fn sniff_and_extension_helpers() {
        assert_eq!(sniff_image_mime(b"\xff\xd8\xff\xe0rest"), Some("image/jpeg"));
        assert_eq!(sniff_image_mime(b"GIF89a..."), Some("image/gif"));
        assert_eq!(sniff_image_mime(b"GIF87a..."), Some("image/gif"));
        assert_eq!(sniff_image_mime(b"RIFF\0\0\0\0WEBPVP8 "), Some("image/webp"));
        assert_eq!(sniff_image_mime(b"RIFF"), None);
        assert_eq!(sniff_image_mime(b"hello"), None);
        assert_eq!(mime_from_filename("a.PNG"), Some("image/png"));
        assert_eq!(mime_from_filename("a.jpeg"), Some("image/jpeg"));
        assert_eq!(mime_from_filename("noext"), None);
        assert_eq!(mime_from_filename("a.txt"), None);
        assert_eq!(extension_for_mime("image/jpeg"), ".jpg");
        assert_eq!(extension_for_mime("application/octet-stream"), "");
    }

    #[test]
    fn requires_html() {
        assert!(err(&[b64("logo", PNG_1X1_B64)], false).contains("html_body"));
    }

    #[test]
    fn rejects_empty_content_id() {
        assert!(err(&[b64("  <cid:> ", PNG_1X1_B64)], true).contains("missing a content_id"));
    }

    #[test]
    fn requires_exactly_one_source() {
        let neither = InlineImage { content_id: "a".into(), ..Default::default() };
        assert!(err(&[neither], true).contains("exactly one"));
        let path = temp_file("both.png", b"x");
        let both = InlineImage {
            content_id: "a".into(),
            path: Some(path.to_string_lossy().into_owned()),
            data_base64: Some(PNG_1X1_B64.into()),
            ..Default::default()
        };
        assert!(err(&[both], true).contains("exactly one"));
        // Empty strings count as absent.
        let empty = InlineImage { content_id: "a".into(), path: Some("".into()), data_base64: Some(" ".into()), ..Default::default() };
        assert!(err(&[empty], true).contains("exactly one"));
    }

    #[test]
    fn rejects_missing_path() {
        let img = InlineImage { content_id: "a".into(), path: Some("/definitely/not/here.png".into()), ..Default::default() };
        assert!(err(&[img], true).contains("inline image not found"));
    }

    #[test]
    fn rejects_bad_base64() {
        assert!(err(&[b64("a", "@@@")], true).contains("not valid base64"));
    }

    #[test]
    fn rejects_duplicate_ids_case_insensitively() {
        let e = err(&[b64("Logo", PNG_1X1_B64), b64("<cid:LOGO>", PNG_1X1_B64)], true);
        assert!(e.contains("duplicate"), "{e}");
    }

    #[test]
    fn base64_image_gets_sniffed_mime_and_default_filename() {
        let v = validate_inline_images(&[b64("<cid:chart>", PNG_1X1_B64)], true).unwrap();
        assert_eq!(v.len(), 1);
        assert_eq!(v[0].content_id, "chart");
        assert_eq!(v[0].mime_type, "image/png");
        assert_eq!(v[0].filename, "chart.png");
        assert!(matches!(&v[0].source, InlineImageSource::Data(d) if d.starts_with(b"\x89PNG")));
    }

    #[test]
    fn unknown_data_falls_back_to_octet_stream() {
        let v = validate_inline_images(&[b64("blob", "aGVsbG8=")], true).unwrap();
        assert_eq!(v[0].mime_type, "application/octet-stream");
        assert_eq!(v[0].filename, "blob");
    }

    #[test]
    fn explicit_mime_and_filename_win_and_filename_is_sanitized() {
        let img = InlineImage {
            content_id: "x".into(),
            data_base64: Some(PNG_1X1_B64.into()),
            filename: Some("my:pic.gif".into()),
            mime_type: Some("image/custom".into()),
            ..Default::default()
        };
        let v = validate_inline_images(&[img], true).unwrap();
        assert_eq!(v[0].mime_type, "image/custom");
        assert_eq!(v[0].filename, "my_pic.gif");
        // Filename extension beats sniffing when no mime_type is given.
        let img = InlineImage { content_id: "y".into(), data_base64: Some(PNG_1X1_B64.into()), filename: Some("pic.gif".into()), ..Default::default() };
        assert_eq!(validate_inline_images(&[img], true).unwrap()[0].mime_type, "image/gif");
        // Unsafe characters in a cid-derived filename are sanitized too.
        let v = validate_inline_images(&[b64("a/b", PNG_1X1_B64)], true).unwrap();
        assert_eq!(v[0].content_id, "a/b");
        assert_eq!(v[0].filename, "a_b.png");
    }

    #[test]
    fn path_image_uses_file_name_and_extension_mime() {
        let path = temp_file("photo.JPG", b"\xff\xd8\xff");
        let img = InlineImage { content_id: "p".into(), path: Some(path.to_string_lossy().into_owned()), ..Default::default() };
        let v = validate_inline_images(&[img], true).unwrap();
        assert_eq!(v[0].filename, "photo.JPG");
        assert_eq!(v[0].mime_type, "image/jpeg");
        assert_eq!(v[0].source, InlineImageSource::Path(path.to_string_lossy().into_owned()));
    }

    #[test]
    fn empty_list_is_ok() {
        assert!(validate_inline_images(&[], true).unwrap().is_empty());
    }
}
