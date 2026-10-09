pub mod client;
pub mod com;
pub mod fake;
pub mod jet_dates;
pub mod types;

use crate::error::ToolError;
use serde_json::Value;
use types::*;

/// All filters for `list_emails`. All optional except `folder`/`count`/`offset`
/// (which the server fills with defaults). Supplying several ANDs them.
#[derive(Debug, Clone)]
pub struct EmailQuery {
    pub query: Option<String>,
    pub folder: String,
    pub count: i32,
    /// Matches to skip before the page starts (after every filter). Negative
    /// values are treated as 0.
    pub offset: i32,
    pub unread_only: bool,
    pub from: Option<String>,
    /// Recipient filter: caseless substring of any To/CC recipient's
    /// display name or address.
    pub to: Option<String>,
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

/// All changes `update_draft` can apply to one existing, unsent draft.
/// Every field except `draft_id` is optional; supplying several applies all
/// of them in field order and then saves once (the draft is never sent).
/// `subject`, `body` and `html_body` replace the current value (`body` and
/// `html_body` are mutually exclusive). `to`/`cc`/`bcc` replace that whole
/// recipient line, and `Some(vec![])` clears it. `attachments` are local file
/// paths appended to the existing attachments.
#[derive(Debug, Clone, Default)]
pub struct DraftUpdate {
    pub draft_id: String,
    pub subject: Option<String>,
    pub body: Option<String>,
    pub html_body: Option<String>,
    pub to: Option<Vec<String>>,
    pub cc: Option<Vec<String>>,
    pub bcc: Option<Vec<String>>,
    pub attachments: Option<Vec<String>>,
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

/// All filters for `list_tasks`. Every field is optional except
/// `include_completed`; supplying several ANDs them. `include_completed`
/// drives a server-side `Restrict`; the rest filter the streamed tasks
/// client-side (there's no established DASL text-search path for the Tasks
/// folder in this codebase, unlike email's `@SQL` queries — same approach
/// `EventQuery`'s `query`/`category` already use). `query` matches either
/// the subject or the real task body, read per-item — same as `NoteQuery`'s
/// `query` below.
#[derive(Debug, Clone, Default)]
pub struct TaskQuery {
    pub include_completed: bool,
    pub category: Option<String>,
    pub importance: Option<String>,
    pub query: Option<String>, // text match on subject OR body
}

/// All filters for `list_notes`. Both fields optional; supplying both ANDs
/// them. A note's *only* content is its body (it has no separate subject),
/// so `note_matches` reads the real body text to match `query` — the same
/// approach `TaskQuery`'s `query` above now uses alongside its subject.
#[derive(Debug, Clone, Default)]
pub struct NoteQuery {
    pub category: Option<String>,
    pub query: Option<String>,
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
    /// `max_body_chars`: `None` = the 100,000-char default; other values are
    /// clamped to 1,000..=5,000,000.
    fn get_email(&self, email_id: String, prefer_html: bool, max_body_chars: Option<u32>)
        -> Result<EmailDetail, ToolError>;
    fn send_email(&self, to: Vec<String>, subject: String, body: String,
        cc: Option<Vec<String>>, bcc: Option<Vec<String>>, html: bool,
        attachments: Option<Vec<String>>, inline_images: Option<Vec<InlineImage>>)
        -> Result<Value, ToolError>;
    fn create_draft(&self, to: Vec<String>, subject: String, body: String,
        cc: Option<Vec<String>>, bcc: Option<Vec<String>>, html: bool,
        attachments: Option<Vec<String>>, inline_images: Option<Vec<InlineImage>>)
        -> Result<Value, ToolError>;
    fn reply_email(&self, email_id: String, body: String, reply_all: bool,
        html: bool, send: bool, attachments: Option<Vec<String>>)
        -> Result<Value, ToolError>;
    fn update_email(&self, u: EmailUpdate) -> Result<Value, ToolError>;
    fn update_draft(&self, u: DraftUpdate) -> Result<Value, ToolError>;
    /// `permanent = false` moves the email to Deleted Items; `true`
    /// hard-deletes it like Outlook's shift+delete.
    fn delete_email(&self, email_id: String, permanent: bool) -> Result<Value, ToolError>;
    /// Permanently delete everything in the default store's Deleted Items.
    /// Refuses (via [`require_empty_confirm`]) unless `confirm` is true.
    fn empty_deleted_items(&self, confirm: bool) -> Result<Value, ToolError>;

    fn list_events(&self, q: EventQuery) -> Result<Vec<EventSummary>, ToolError>;
    fn get_event(&self, event_id: String) -> Result<EventDetail, ToolError>;
    fn create_event(&self, input: CreateEventInput) -> Result<Value, ToolError>;
    fn respond_to_meeting(&self, event_id: String, response: String,
        comment: Option<String>, send: bool) -> Result<Value, ToolError>;
    fn update_event(&self, u: EventUpdate) -> Result<Value, ToolError>;
    fn delete_event(&self, event_id: String, send_cancellation: bool) -> Result<Value, ToolError>;
    fn check_availability(&self, input: CheckAvailabilityInput) -> Result<AvailabilityResult, ToolError>;

    fn list_attachments(&self, email_id: String)
        -> Result<Vec<AttachmentInfo>, ToolError>;
    fn save_attachments(&self, email_id: String, save_dir: String,
        attachment_names: Option<Vec<String>>) -> Result<Vec<Value>, ToolError>;
    fn get_inline_image(&self, email_id: String, content_id: String,
        context_lines: Option<u32>) -> Result<InlineImageData, ToolError>;

    fn list_tasks(&self, q: TaskQuery) -> Result<Vec<TaskSummary>, ToolError>;
    fn create_task(&self, subject: String, body: Option<String>,
        due_date: Option<String>, importance: String, categories: Option<Vec<String>>,
        start_date: Option<String>, reminder_time: Option<String>) -> Result<Value, ToolError>;
    fn update_task(&self, u: TaskUpdate) -> Result<Value, ToolError>;
    fn delete_task(&self, task_id: String) -> Result<Value, ToolError>;

    fn list_notes(&self, q: NoteQuery) -> Result<Vec<NoteSummary>, ToolError>;
    fn get_note(&self, note_id: String) -> Result<NoteDetail, ToolError>;
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
/// (an empty `attachments` list counts as nothing), or an attachment path
/// that isn't an existing file. Called by both `update_draft` implementors.
pub fn validate_draft_update(u: &DraftUpdate) -> Result<(), ToolError> {
    if u.body.is_some() && u.html_body.is_some() {
        return Err(ToolError::new("pass either 'body' or 'html_body', not both"));
    }
    let attachments = u.attachments.as_deref().unwrap_or(&[]);
    if u.subject.is_none() && u.body.is_none() && u.html_body.is_none()
        && u.to.is_none() && u.cc.is_none() && u.bcc.is_none() && attachments.is_empty()
    {
        return Err(ToolError::new(
            "update_draft needs at least one of: subject, body, html_body, to, cc, bcc, attachments",
        ));
    }
    for p in attachments {
        if !std::path::Path::new(p).is_file() {
            return Err(ToolError::new(format!("attachment not found: {p}")));
        }
    }
    Ok(())
}

/// The `changed` list `update_draft` returns: the supplied fields, in the
/// order they are applied. Assumes `u` already passed `validate_draft_update`.
pub fn draft_update_changes(u: &DraftUpdate) -> Vec<&'static str> {
    let mut changed = Vec::new();
    if u.subject.is_some() { changed.push("subject"); }
    if u.body.is_some() { changed.push("body"); }
    if u.html_body.is_some() { changed.push("html_body"); }
    if u.to.is_some() { changed.push("to"); }
    if u.cc.is_some() { changed.push("cc"); }
    if u.bcc.is_some() { changed.push("bcc"); }
    if u.attachments.as_ref().is_some_and(|a| !a.is_empty()) { changed.push("attachments"); }
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
        return Err(ToolError::new(
            "inline_images requires html=true (reference them in the HTML body as <img src=\"cid:CONTENT_ID\">).",
        ));
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

#[cfg(test)]
mod tests {
    use super::{
        com_recurrence_interval, common_free, create_event_status, friendly_recurrence_interval,
        parse_freebusy_slots, take_page, validate_recurrence, validate_recurrence_update,
        EventUpdate, RecurrenceInput, permanent_delete_needs_move, require_empty_confirm,
        draft_update_changes, validate_draft_update, DraftUpdate,
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
        DraftUpdate { draft_id: id.to_string(), ..Default::default() }
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
        assert!(err(&[b64("logo", PNG_1X1_B64)], false).contains("html=true"));
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
