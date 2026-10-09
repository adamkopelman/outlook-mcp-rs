//! Win32 COM implementation of the `OutlookClient` trait.
//!
//! Direct port of `outlook_mcp/outlook/client.py`'s email section
//! (lines 133-336). Every public method wraps its body in [`with_com`],
//! which initializes COM on the current thread (mirroring the Python
//! `@_com` decorator) and maps `windows::core::Error` into [`ToolError`].
//!
//! All 20 `OutlookClient` trait methods are implemented (email, calendar,
//! attachments, tasks, and notes; Tasks 12-16) — no `todo!()` stubs remain.

use serde_json::{json, Value};
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Variant::VARIANT;

use crate::constants as c;
use crate::error::ToolError;
use crate::outlook::com::{
    call_method, clean_content_id, create_com_object, format_com_error, get_item_categories,
    get_mapi_prop, get_property, guess_mime, has_member, is_inline, make_item_id, normalize_cid_request, parse_item_id, put_property, safe_filename,
    set_item_categories, user_short_date_order, variant_from_bool, variant_from_datetime, variant_from_i32,
    variant_from_str, variant_to_bool, variant_to_datetime, variant_to_i32, variant_to_iso_string,
    variant_to_string, ComGuard,
};
use crate::outlook::jet_dates::{DateFilter, Placement};
use crate::outlook::types::*;
use chrono::Datelike;
use unicode_normalization::UnicodeNormalization;
use crate::outlook::{
    com_recurrence_interval, common_free, create_event_status, friendly_recurrence_interval,
    parse_freebusy_slots, take_page, validate_recurrence, validate_recurrence_update,
    CheckAvailabilityInput, permanent_delete_needs_move, require_empty_confirm, CreateEventInput,
    EmailQuery, EmailUpdate, EventQuery, EventUpdate, NoteQuery, NoteUpdate, OutlookClient,
    RecurrenceInput, TaskQuery, TaskUpdate, text_before_cid, draft_update_changes,
    validate_draft_update, DraftUpdate, InlineImage, InlineImageSource, ValidatedInlineImage,
    validate_inline_images,
};

/// Matches `MAX_EMAIL_COUNT` in `client.py`. Larger result sets are paged
/// with `EmailQuery::offset`.
const MAX_EMAIL_COUNT: i32 = 200;
/// How many items (newest first, after every other filter) `list_emails`
/// will open one by one when the `to` filter falls back to scanning each
/// item's `Recipients` collection. Bounds the cost of a per-item COM walk.
const RECIPIENT_SCAN_LIMIT: i32 = 2000;
/// MAPI `PR_SMTP_ADDRESS` (Unicode), read through `Recipient.PropertyAccessor`.
const PR_SMTP_ADDRESS: &str = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F";
/// Matches `MAX_CALENDAR_ITEMS` in `client.py`. Caps `list_events` because a
/// recurring appointment without an end date expands forever under
/// `IncludeRecurrences`.
const MAX_CALENDAR_ITEMS: usize = 250;
/// Matches `MAX_BODY_CHARS` in `client.py`. The default body cut, and the
/// fixed cut for `get_event`/`get_note`.
const MAX_BODY_CHARS: usize = 100_000;
/// Largest attachment `get_inline_image` returns inline (base64 in the tool
/// result); anything bigger belongs on disk via `save_attachments`.
const MAX_INLINE_IMAGE_BYTES: usize = 10 * 1024 * 1024;
/// Upper bound on items scanned client-side when the DASL text search for a
/// non-ASCII query comes back empty (see `list_emails`).
const MAX_SCAN_ITEMS: i32 = 2000;
/// Bounds for `get_email`'s caller-chosen `max_body_chars` (large HTML mail
/// with inline images can run to megabytes). Match `MIN_BODY_CHARS_LIMIT` /
/// `MAX_BODY_CHARS_LIMIT` in `client.py`.
const MIN_BODY_CHARS_LIMIT: usize = 1_000;
const MAX_BODY_CHARS_LIMIT: usize = 5_000_000;

/// Lets `?` turn a `windows::core::Error` into a [`ToolError`] anywhere in
/// this module, so COM-plumbing calls (`call_method`, `get_property`, …) and
/// context-carrying helpers (`get_item`, `resolve_folder`) can share one
/// `Result<_, ToolError>` error channel. Mirrors the Python `@_com`
/// decorator translating `pywintypes.com_error` into `ToolError`.
impl From<windows::core::Error> for ToolError {
    fn from(err: windows::core::Error) -> Self {
        ToolError::new(format_com_error(&err))
    }
}

pub struct WindowsOutlookClient;

impl Default for WindowsOutlookClient {
    fn default() -> Self {
        Self::new()
    }
}

impl WindowsOutlookClient {
    pub fn new() -> Self {
        Self
    }

    /// Wraps every public method body: initializes COM on the current
    /// (blocking-pool) thread for the duration of the call, then runs `f`.
    /// The closure returns `Result<T, ToolError>` (a small deviation from the
    /// task brief's `WinResult<T>`) so that `get_item`/`resolve_folder` can
    /// surface the exact, context-rich messages the Python client produces
    /// instead of routing every failure through `format_com_error`. COM
    /// plumbing errors still convert automatically via the `From` impl above.
    fn with_com<T>(&self, f: impl FnOnce() -> Result<T, ToolError>) -> Result<T, ToolError> {
        let _guard = ComGuard::new().map_err(|e| ToolError::new(format_com_error(&e)))?;
        f()
    }
}

// ---- module-level plumbing helpers (translated from client.py) ----------

/// `IDispatch`-returning `VARIANT` unwrap. `TryFrom<&VARIANT> for IDispatch`
/// borrows, so this takes the `VARIANT` by value and borrows it internally.
fn to_disp(v: VARIANT) -> Result<IDispatch, ToolError> {
    Ok(IDispatch::try_from(&v)?)
}

/// `client.py::_mapi`: the `Outlook.Application` object plus its MAPI namespace.
fn mapi() -> Result<(IDispatch, IDispatch), ToolError> {
    let app = create_com_object("Outlook.Application")?;
    let ns = to_disp(call_method(&app, "GetNamespace", &mut [variant_from_str("MAPI")])?)?;
    Ok((app, ns))
}

/// `client.py::_make_id`: `"{EntryID}|{Parent.StoreID}"`.
fn make_id(item: &IDispatch) -> Result<String, ToolError> {
    let entry_id = variant_to_string(&get_property(item, "EntryID")?);
    let parent = to_disp(get_property(item, "Parent")?)?;
    let store_id = variant_to_string(&get_property(&parent, "StoreID")?);
    Ok(make_item_id(&entry_id, &store_id))
}

/// `client.py::_get_item`: parse the opaque id, then `Namespace.GetItemFromID`.
fn get_item(ns: &IDispatch, item_id: &str) -> Result<IDispatch, ToolError> {
    let (entry_id, store_id) = parse_item_id(item_id)?;
    let item = call_method(
        ns,
        "GetItemFromID",
        &mut [variant_from_str(&entry_id), variant_from_str(&store_id)],
    )
    .map_err(|e| {
        ToolError::new(format!(
            "Item not found — it may have been moved or deleted (item ids change \
             when an item moves to another folder). {}",
            format_com_error(&e)
        ))
    })?;
    to_disp(item)
}

/// The Deleted Items folder of the item's own store, so a hard delete of an
/// item in a secondary mailbox/PST stays in that store. Falls back to the
/// default store's Deleted Items when the store can't be resolved.
fn deleted_items_for(ns: &IDispatch, item: &IDispatch) -> Result<IDispatch, ToolError> {
    let own = || -> Result<IDispatch, ToolError> {
        let parent = to_disp(get_property(item, "Parent")?)?;
        let store = to_disp(get_property(&parent, "Store")?)?;
        to_disp(call_method(
            &store, "GetDefaultFolder", &mut [variant_from_i32(c::OL_FOLDER_DELETED_ITEMS)],
        )?)
    };
    match own() {
        Ok(folder) => Ok(folder),
        Err(_) => to_disp(call_method(
            ns, "GetDefaultFolder", &mut [variant_from_i32(c::OL_FOLDER_DELETED_ITEMS)],
        )?),
    }
}

/// `client.py::_resolve_folder`: a well-known folder name maps to a default
/// folder id; otherwise walk a `Inbox/Sub/Sub` path from the store root.
fn resolve_folder(ns: &IDispatch, folder: Option<&str>) -> Result<IDispatch, ToolError> {
    let name = folder.unwrap_or("inbox").trim();
    if let Some(id) = c::folder_name_to_id(name) {
        return to_disp(call_method(ns, "GetDefaultFolder", &mut [variant_from_i32(id)])?);
    }
    let inbox = to_disp(call_method(
        ns,
        "GetDefaultFolder",
        &mut [variant_from_i32(c::OL_FOLDER_INBOX)],
    )?)?;
    let mut current = to_disp(get_property(&inbox, "Parent")?)?;
    for part in name.split(['/', '\\']).filter(|p| !p.is_empty()) {
        let folders = to_disp(get_property(&current, "Folders")?)?;
        let count = variant_to_i32(&get_property(&folders, "Count")?).unwrap_or(0);
        let mut found = None;
        for i in 1..=count {
            let sub = to_disp(call_method(&folders, "Item", &mut [variant_from_i32(i)])?)?;
            let sub_name = variant_to_string(&get_property(&sub, "Name")?);
            if sub_name.eq_ignore_ascii_case(part) {
                found = Some(sub);
                break;
            }
        }
        current = match found {
            Some(f) => f,
            None => {
                let cur_name = variant_to_string(&get_property(&current, "Name")?);
                return Err(ToolError::new(format!(
                    "Folder not found: {name:?} (no subfolder named {part:?} under {cur_name:?})"
                )));
            }
        };
    }
    Ok(current)
}

/// `os.path.abspath(os.path.expanduser(save_dir))`: expand a leading `~` to the
/// user's home directory, then make the path absolute. Mirrors the Python
/// `save_attachments` directory normalization so the returned `saved_to` paths
/// are absolute. Falls back to the un-absolutized path if `absolute` fails
/// (it does not touch the filesystem, so this only guards against odd inputs).
fn resolve_save_dir(save_dir: &str) -> std::path::PathBuf {
    let expanded = if save_dir == "~"
        || save_dir.starts_with("~/")
        || save_dir.starts_with("~\\")
    {
        match std::env::var_os("USERPROFILE").or_else(|| std::env::var_os("HOME")) {
            Some(home) => {
                let rest = save_dir[1..].trim_start_matches(['/', '\\']);
                std::path::Path::new(&home).join(rest)
            }
            None => std::path::PathBuf::from(save_dir),
        }
    } else {
        std::path::PathBuf::from(save_dir)
    };
    std::path::absolute(&expanded).unwrap_or(expanded)
}

/// `client.py::_truncate`: cap long bodies at `limit` *characters* (not
/// bytes) so multi-byte UTF-8 content is never split mid-codepoint. The cut
/// is a hard one and may land mid-tag or mid-base64 in HTML; the returned
/// flag says whether it happened so callers can re-request with a larger
/// limit. Returns `(text, truncated)`.
fn truncate(text: &str, limit: usize) -> (String, bool) {
    if text.chars().count() > limit {
        let head: String = text.chars().take(limit).collect();
        (format!("{head}\n\n[... truncated at {limit} characters]"), true)
    } else {
        (text.to_string(), false)
    }
}

/// `client.py::_clamp_body_limit`: `None` means `MAX_BODY_CHARS`; any other
/// value is clamped into `[MIN_BODY_CHARS_LIMIT, MAX_BODY_CHARS_LIMIT]`.
fn clamp_body_limit(max_body_chars: Option<u32>) -> usize {
    match max_body_chars {
        None => MAX_BODY_CHARS,
        Some(n) => (n as usize).clamp(MIN_BODY_CHARS_LIMIT, MAX_BODY_CHARS_LIMIT),
    }
}

/// Normalize text for caseless matching: NFC (so composed and decomposed
/// forms of the same character compare equal), then lowercase.
fn fold_text(text: &str) -> String {
    text.nfc().collect::<String>().to_lowercase()
}

/// True if `query` occurs (caseless, NFC-normalized) in any of `fields`.
/// Empty fields are skipped.
fn text_matches(query: &str, fields: &[&str]) -> bool {
    let needle = fold_text(query);
    fields.iter().filter(|f| !f.is_empty()).any(|f| fold_text(f).contains(&needle))
}

/// Client-side text match for the `list_emails` non-ASCII fallback: subject
/// and sender name first, then the body only if those miss (Body is the
/// expensive property to fetch).
fn email_text_matches(item: &IDispatch, query: &str) -> bool {
    let subject = variant_to_string(&get_property(item, "Subject").unwrap_or_default());
    let sender = variant_to_string(&get_property(item, "SenderName").unwrap_or_default());
    if text_matches(query, &[&subject, &sender]) {
        return true;
    }
    let body = variant_to_string(&get_property(item, "Body").unwrap_or_default());
    text_matches(query, &[&body])
}

/// `client.py::_parse_dt`: parse a user-supplied ISO date/datetime. Mirrors
/// Python's `datetime.fromisoformat`, accepting a bare date (`2026-06-10`) or a
/// date-time with `T` or space separator, with or without seconds/fractional
/// seconds. The error message mirrors the Python original (with this file's
/// `{:?}` quoting convention rather than Python's `!r`).
fn parse_dt(value: &str, field: &str) -> Result<chrono::NaiveDateTime, ToolError> {
    let trimmed = value.trim();
    // Normalize a single space separator to `T` so one set of formats covers
    // both `2026-06-10T14:30` and `2026-06-10 14:30` (Python 3.11+ accepts it).
    let normalized = trimmed.replacen(' ', "T", 1);
    for fmt in ["%Y-%m-%dT%H:%M:%S%.f", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M"] {
        if let Ok(dt) = chrono::NaiveDateTime::parse_from_str(&normalized, fmt) {
            return Ok(dt);
        }
    }
    if let Ok(d) = chrono::NaiveDate::parse_from_str(trimmed, "%Y-%m-%d") {
        return Ok(d.and_hms_opt(0, 0, 0).unwrap());
    }
    Err(ToolError::new(format!(
        "Invalid {field} {value:?}: expected ISO format like '2026-06-10' or '2026-06-10T14:30'"
    )))
}

/// Adds `address` to `recipients` and marks it required or optional. The
/// `Recipient` object `Recipients.Add()` returns must have its `.Type` set
/// explicitly — Outlook does not infer tier from call order.
fn add_meeting_recipient(recipients: &IDispatch, address: &str, role: i32) -> Result<(), ToolError> {
    let recipient = to_disp(call_method(recipients, "Add", &mut [variant_from_str(address)])?)?;
    put_property(&recipient, "Type", variant_from_i32(role))?;
    Ok(())
}

/// Sets an appointment's recurrence pattern via `GetRecurrencePattern()`.
/// Calling this on a non-recurring appointment converts it into a recurring
/// one (this is how `update_event` adds recurrence to an existing single
/// event, too — see Task 4). `"yearly"` derives its month/day from the
/// appointment's own `Start` property rather than a separate input field, so
/// this must run after `Start` is already set to its final value.
fn apply_recurrence(appt: &IDispatch, r: &RecurrenceInput) -> Result<(), ToolError> {
    let recurrence_type = validate_recurrence(r)?;
    let pattern = to_disp(call_method(appt, "GetRecurrencePattern", &mut [])?)?;
    put_property(&pattern, "RecurrenceType", variant_from_i32(recurrence_type))?;
    put_property(&pattern, "Interval", variant_from_i32(com_recurrence_interval(r)))?;
    match r.pattern.to_lowercase().as_str() {
        "weekly" => {
            let mask = crate::friendly::day_of_week_words_to_mask(
                r.days_of_week.as_deref().unwrap_or(&[]),
            )?;
            put_property(&pattern, "DayOfWeekMask", variant_from_i32(mask))?;
        }
        "monthly" => {
            put_property(&pattern, "DayOfMonth", variant_from_i32(r.day_of_month.unwrap()))?;
        }
        "yearly" => {
            let start_iso = variant_to_iso_string(&get_property(appt, "Start")?).ok_or_else(|| {
                ToolError::new("could not read Start to derive the yearly recurrence date")
            })?;
            let start_dt =
                chrono::NaiveDateTime::parse_from_str(&start_iso, "%Y-%m-%dT%H:%M:%S").map_err(|_| {
                    ToolError::new("could not parse Start to derive the yearly recurrence date")
                })?;
            put_property(&pattern, "MonthOfYear", variant_from_i32(start_dt.month() as i32))?;
            put_property(&pattern, "DayOfMonth", variant_from_i32(start_dt.day() as i32))?;
        }
        _ => {}
    }
    match (r.occurrences, r.until.as_deref()) {
        (Some(n), _) => {
            put_property(&pattern, "Occurrences", variant_from_i32(n))?;
        }
        (None, Some(until)) => {
            let until_dt = parse_dt(until, "recurrence.until")?;
            put_property(&pattern, "PatternEndDate", variant_from_datetime(&until_dt)?)?;
        }
        (None, None) => {
            put_property(&pattern, "NoEndDate", variant_from_bool(true))?;
        }
    }
    Ok(())
}

/// Removes every recipient whose `Name` or `Address` case-insensitively
/// matches any entry in `addresses`. Iterates from `Count` down to `1` —
/// `Recipients.Remove(index)` is 1-based and shifts every later index down
/// by one, so removing in reverse means an index we haven't visited yet is
/// never invalidated by an earlier removal.
fn remove_meeting_recipients(recipients: &IDispatch, addresses: &[String]) -> Result<(), ToolError> {
    let count = variant_to_i32(&get_property(recipients, "Count")?).unwrap_or(0);
    for i in (1..=count).rev() {
        let recipient = to_disp(call_method(recipients, "Item", &mut [variant_from_i32(i)])?)?;
        let name = variant_to_string(&get_property(&recipient, "Name").unwrap_or_default());
        let address = variant_to_string(&get_property(&recipient, "Address").unwrap_or_default());
        if addresses.iter().any(|a| a.eq_ignore_ascii_case(&name) || a.eq_ignore_ascii_case(&address)) {
            call_method(recipients, "Remove", &mut [variant_from_i32(i)])?;
        }
    }
    Ok(())
}

/// `client.py::_event_summary`, enriched for v2 with show_as/my_response and the
/// attendee strings so every calendar filter can operate on the built summary.
///
/// `calendar_store_id`: when `Some`, build the id from `EntryID` + this
/// caller-supplied store id instead of calling `make_id` (which reads
/// `item.Parent.StoreID`). `list_events`' enumeration passes this: items
/// returned by `Items.GetFirst()`/`GetNext()` after `Restrict()` with
/// `IncludeRecurrences = True` carry a `Parent` whose `StoreID` never
/// resolves (`DISP_E_UNKNOWNNAME`/"Unknown name", confirmed live —
/// deterministic on every enumerated item, not cleared by retrying the same
/// property read on the same object up to 5 times, nor by re-querying
/// minutes later — so it is a real object-model gap for this enumeration
/// path, not a Cached Exchange Mode sync-lag blip). The already-resolved
/// calendar folder (from `GetDefaultFolder`/`GetSharedDefaultFolder`, a
/// genuine `Folder` object, not a GetFirst/GetNext proxy) has a `StoreID`
/// that always resolves, and every item `list_events` enumerates belongs to
/// that same folder — so its `StoreID` is reused instead. `get_event`
/// (items fetched via `GetItemFromID`, unaffected by this) still passes
/// `None` and uses the normal `make_id` path.
fn event_summary(item: &IDispatch, calendar_store_id: Option<&str>) -> Result<EventSummary, ToolError> {
    let meeting_status = variant_to_i32(&get_property(item, "MeetingStatus").unwrap_or_default())
        .unwrap_or(c::OL_NONMEETING);
    let id = match calendar_store_id {
        Some(store_id) => {
            let entry_id = variant_to_string(&get_property(item, "EntryID")?);
            make_item_id(&entry_id, store_id)
        }
        None => make_id(item)?,
    };
    Ok(EventSummary {
        id,
        subject: variant_to_string(&get_property(item, "Subject").unwrap_or_default()),
        start: variant_to_iso_string(&get_property(item, "Start").unwrap_or_default()),
        end: variant_to_iso_string(&get_property(item, "End").unwrap_or_default()),
        location: variant_to_string(&get_property(item, "Location").unwrap_or_default()),
        organizer: variant_to_string(&get_property(item, "Organizer").unwrap_or_default()),
        all_day: variant_to_bool(&get_property(item, "AllDayEvent").unwrap_or_default())
            .unwrap_or(false),
        is_recurring: variant_to_bool(&get_property(item, "IsRecurring").unwrap_or_default())
            .unwrap_or(false),
        is_meeting: meeting_status != c::OL_NONMEETING,
        categories: get_item_categories(item),
        show_as: crate::friendly::busy_status_word(
            variant_to_i32(&get_property(item, "BusyStatus").unwrap_or_default())
                .unwrap_or(c::OL_BUSY),
        )
        .to_string(),
        my_response: crate::friendly::response_word(
            variant_to_i32(&get_property(item, "ResponseStatus").unwrap_or_default())
                .unwrap_or(c::OL_RESPONSE_NONE),
        )
        .to_string(),
        required_attendees: variant_to_string(
            &get_property(item, "RequiredAttendees").unwrap_or_default(),
        ),
        optional_attendees: variant_to_string(
            &get_property(item, "OptionalAttendees").unwrap_or_default(),
        ),
    })
}

/// Reads an appointment's recurrence pattern back via
/// `GetRecurrencePattern()`, or `None` if `IsRecurring` is false. Mirrors
/// `apply_recurrence`'s field set in the opposite direction.
fn recurrence_info(item: &IDispatch) -> Result<Option<RecurrenceInfo>, ToolError> {
    let is_recurring =
        variant_to_bool(&get_property(item, "IsRecurring").unwrap_or_default()).unwrap_or(false);
    if !is_recurring {
        return Ok(None);
    }
    let pattern = to_disp(call_method(item, "GetRecurrencePattern", &mut [])?)?;
    let recurrence_type =
        variant_to_i32(&get_property(&pattern, "RecurrenceType").unwrap_or_default())
            .unwrap_or(c::OL_RECURS_DAILY);
    let interval = friendly_recurrence_interval(
        recurrence_type,
        variant_to_i32(&get_property(&pattern, "Interval").unwrap_or_default()).unwrap_or(1),
    );
    let day_mask =
        variant_to_i32(&get_property(&pattern, "DayOfWeekMask").unwrap_or_default()).unwrap_or(0);
    let day_of_month = variant_to_i32(&get_property(&pattern, "DayOfMonth").unwrap_or_default());
    let no_end =
        variant_to_bool(&get_property(&pattern, "NoEndDate").unwrap_or_default()).unwrap_or(false);
    let until = if no_end {
        None
    } else {
        variant_to_iso_string(&get_property(&pattern, "PatternEndDate").unwrap_or_default())
    };
    // Confirmed live (see RecurrenceInfo's doc comment): once a series has a
    // finite end (`no_end` false), Outlook keeps `Occurrences` and
    // `PatternEndDate` mutually *consistent*, auto-computing whichever one
    // wasn't the caller's original input — e.g. an `until`-terminated series
    // still reports a real, correct `Occurrences` count, and vice versa.
    // There is no COM-level signal for which field the series was originally
    // created with, so both are reported together rather than one being
    // arbitrarily suppressed (suppressing either one on read-back was tried
    // and breaks the other, legitimately-populated direction).
    let occurrences = if no_end {
        None
    } else {
        variant_to_i32(&get_property(&pattern, "Occurrences").unwrap_or_default())
    };
    Ok(Some(RecurrenceInfo {
        pattern: crate::friendly::recurrence_pattern_word(recurrence_type).to_string(),
        interval,
        days_of_week: crate::friendly::day_of_week_mask_to_words(day_mask),
        day_of_month: day_of_month.filter(|d| *d > 0),
        until,
        occurrences,
        no_end,
    }))
}

/// True if `needle` is a case-insensitive substring of any of `candidates`
/// (a recipient's display name, address, SMTP address, …). Callers skip
/// the filter entirely for an empty needle.
fn recipient_matches(needle: &str, candidates: &[String]) -> bool {
    let needle = needle.to_lowercase();
    candidates.iter().any(|c| c.to_lowercase().contains(&needle))
}

/// Every name/address string for the To (Type 1) and CC (Type 2) recipients
/// of `item`: `Recipient.Name`, `Recipient.Address`, and the SMTP address
/// from `PropertyAccessor` (for an Exchange recipient `Address` is the X.500
/// legacy DN, so the SMTP form is only available through PR_SMTP_ADDRESS).
/// Best-effort: an item without `Recipients` (e.g. some report items) or a
/// recipient whose properties can't be read just contributes fewer strings.
fn recipient_strings(item: &IDispatch) -> Vec<String> {
    let mut out = Vec::new();
    let Some(recipients) = get_property(item, "Recipients").ok().and_then(|v| to_disp(v).ok()) else {
        return out;
    };
    let n = variant_to_i32(&get_property(&recipients, "Count").unwrap_or_default()).unwrap_or(0);
    for i in 1..=n {
        let Some(r) = call_method(&recipients, "Item", &mut [variant_from_i32(i)])
            .ok()
            .and_then(|v| to_disp(v).ok())
        else {
            continue;
        };
        let kind = variant_to_i32(&get_property(&r, "Type").unwrap_or_default()).unwrap_or(0);
        if kind != 1 && kind != 2 {
            continue; // BCC (3) and originator (0) aren't in displayto/displaycc either.
        }
        for prop in ["Name", "Address"] {
            out.push(variant_to_string(&get_property(&r, prop).unwrap_or_default()));
        }
        let smtp = get_property(&r, "PropertyAccessor")
            .ok()
            .and_then(|v| to_disp(v).ok())
            .and_then(|pa| {
                call_method(&pa, "GetProperty", &mut [variant_from_str(PR_SMTP_ADDRESS)]).ok()
            });
        if let Some(v) = smtp {
            out.push(variant_to_string(&v));
        }
    }
    out.retain(|s| !s.is_empty());
    out
}

/// True if `summary` passes every filter set on `q`. All comparisons are
/// case-insensitive. Attendee matching is a substring test against the
/// semicolon-separated `RequiredAttendees`/`OptionalAttendees` strings.
fn event_matches(summary: &EventSummary, q: &EventQuery) -> bool {
    if let Some(query) = q.query.as_deref().filter(|s| !s.is_empty()) {
        let needle = query.to_lowercase();
        if !summary.subject.to_lowercase().contains(&needle)
            && !summary.location.to_lowercase().contains(&needle)
        {
            return false;
        }
    }
    if let Some(cat) = q.category.as_deref().filter(|s| !s.is_empty()) {
        let want = cat.to_lowercase();
        if !summary.categories.iter().any(|c| c.to_lowercase() == want) {
            return false;
        }
    }
    if let Some(show_as) = q.show_as.as_deref().filter(|s| !s.is_empty()) {
        if !summary.show_as.eq_ignore_ascii_case(show_as) {
            return false;
        }
    }
    if let Some(resp) = q.my_response.as_deref().filter(|s| !s.is_empty()) {
        if !summary.my_response.eq_ignore_ascii_case(resp) {
            return false;
        }
    }
    if q.meetings_only && !summary.is_meeting {
        return false;
    }
    if let Some(want_all_day) = q.all_day {
        if summary.all_day != want_all_day {
            return false;
        }
    }
    if let Some(people) = q.attendees.as_ref().filter(|v| !v.is_empty()) {
        // Which attendee tier(s) to search, per attendee_role (default "any").
        let role = q.attendee_role.as_deref().unwrap_or("any").to_lowercase();
        let required = summary.required_attendees.to_lowercase();
        let optional = summary.optional_attendees.to_lowercase();
        let haystack = match role.as_str() {
            "required" => required,
            "optional" => optional,
            _ => format!("{required}; {optional}"), // "any"
        };
        if !people
            .iter()
            .any(|p| !p.is_empty() && haystack.contains(&p.to_lowercase()))
        {
            return false;
        }
    }
    true
}

/// Client-side filter for `list_tasks`'s `category`/`importance`/`query`.
/// `include_completed` is applied earlier via `Restrict`, not here. `query`
/// matches either the subject or the real task body — `body` is the task's
/// real, untruncated body text, read once per item by the caller (see
/// `list_tasks` below), the same pattern `note_matches` uses for notes.
fn task_matches(body: &str, summary: &TaskSummary, q: &TaskQuery) -> bool {
    if let Some(query) = q.query.as_deref().filter(|s| !s.is_empty()) {
        let needle = query.to_lowercase();
        if !summary.subject.to_lowercase().contains(&needle)
            && !body.to_lowercase().contains(&needle)
        {
            return false;
        }
    }
    if let Some(cat) = q.category.as_deref().filter(|s| !s.is_empty()) {
        let want = cat.to_lowercase();
        if !summary.categories.iter().any(|c| c.to_lowercase() == want) {
            return false;
        }
    }
    if let Some(imp) = q.importance.as_deref().filter(|s| !s.is_empty()) {
        if !summary.importance.eq_ignore_ascii_case(imp) {
            return false;
        }
    }
    true
}

/// Client-side filter for `list_notes`'s `category`/`query`. `body` is the
/// note's real, untruncated body text (read once per item by the caller —
/// see `list_notes` below). A note has no separate subject, so `query`
/// matches only this body, whereas `task_matches` also matches a task's
/// subject.
fn note_matches(body: &str, summary: &NoteSummary, q: &NoteQuery) -> bool {
    if let Some(query) = q.query.as_deref().filter(|s| !s.is_empty()) {
        if !body.to_lowercase().contains(&query.to_lowercase()) {
            return false;
        }
    }
    if let Some(cat) = q.category.as_deref().filter(|s| !s.is_empty()) {
        let want = cat.to_lowercase();
        if !summary.categories.iter().any(|c| c.to_lowercase() == want) {
            return false;
        }
    }
    true
}

/// `DISP_E_UNKNOWNNAME` ("Unknown name"), formatted the way `format_com_error`
/// renders it (`{:#010x}` on the HRESULT). `ToolError` only carries a
/// formatted string (no structured HRESULT), but `format_com_error` embeds
/// the raw code deterministically, so matching this substring reliably
/// isolates this one specific case from a genuine, differently-worded error
/// (e.g. an unresolvable `calendar_of` person, which carries no HRESULT in
/// its message at all and is raised earlier in `list_events` regardless).
const DISP_E_UNKNOWNNAME_HEX: &str = "0x80020006";

fn is_transient_unknown_name(err: &ToolError) -> bool {
    err.0.contains(DISP_E_UNKNOWNNAME_HEX)
}

/// One step of a lazy, filtered scan over an Outlook collection.
enum Scan<T> {
    Keep(T),
    Skip,
    /// Nothing further can match (the collection is sorted past the range).
    Stop,
}

/// Applies a date window to `items` with `Restrict`, trying each rendering in
/// `filters` (the user's short-date order, then ISO; see `jet_dates`) until
/// Outlook accepts one: `Restrict` and then `probe` (which forces the filter to
/// be evaluated) must both succeed. The dates are unambiguous, so a rendering
/// Outlook parses with a different day/month order is rejected, not misread.
/// If every rendering is rejected this is an error, never unfiltered results.
fn restrict_dates(
    items: &IDispatch,
    filters: &[String],
    probe: impl Fn(&IDispatch) -> Result<(), ToolError>,
) -> Result<IDispatch, ToolError> {
    if filters.is_empty() {
        return Ok(items.clone());
    }
    let mut last_err = String::new();
    for flt in filters {
        let attempt = (|| -> Result<IDispatch, ToolError> {
            let restricted =
                to_disp(call_method(items, "Restrict", &mut [variant_from_str(flt)])?)?;
            probe(&restricted)?;
            Ok(restricted)
        })();
        match attempt {
            Ok(restricted) => return Ok(restricted),
            Err(err) => last_err = err.0,
        }
    }
    Err(ToolError::new(format!(
        "Could not apply the date filter: Outlook rejected every date format tried ({}) \
         under this Windows regional format, so no results are returned rather than \
         unfiltered ones. Last error: {last_err}",
        filters.join(" | ")
    )))
}

/// The error for an item `Restrict` returned far outside the date window it
/// was given (`Placement::OutsideWindow`): Outlook misread the filter, so the
/// results may be incomplete and are not returned.
fn date_misread_error(field: &str, found: chrono::NaiveDateTime, f: &DateFilter) -> ToolError {
    let show = |d: Option<chrono::NaiveDate>| d.map_or("…".to_string(), |d| d.to_string());
    ToolError::new(format!(
        "Outlook misread the date filter under this Windows regional format: it returned an \
         item with {field} {found}, outside the requested window {} .. {}. Returning an \
         error rather than possibly incomplete results.",
        show(f.window_start),
        show(f.window_end)
    ))
}

/// Runs the `Restrict` + `GetFirst`/`GetNext` enumeration sequence for
/// `list_events`.
///
/// `calendar_store_id` is threaded through to `event_summary` to sidestep a
/// confirmed-live, deterministic bug (see `event_summary`'s doc comment):
/// every item this enumeration yields has a `Parent` whose `StoreID` never
/// resolves, so ids are built from the calendar folder's own (reliable)
/// `StoreID` instead of re-deriving it per item.
///
/// The whole sequence is also retried from scratch (never resumed
/// mid-enumeration — a partial `results` list from a run that threw partway
/// through isn't trustworthy) up to 3 attempts total, on the off chance a
/// *different* property throws this same transient-looking HRESULT for a
/// genuinely timing-related reason. Any other error propagates immediately,
/// unretried.
fn enumerate_events_with_retry(
    items: &IDispatch,
    filters: &[String],
    date_filter: &DateFilter,
    q: &EventQuery,
    calendar_store_id: &str,
) -> Result<Vec<EventSummary>, ToolError> {
    const MAX_ATTEMPTS: u32 = 3;
    for attempt in 1..=MAX_ATTEMPTS {
        let outcome = (|| -> Result<Vec<EventSummary>, ToolError> {
            // `GetFirst` forces Outlook to evaluate the filter.
            let restricted = restrict_dates(items, filters, |r| {
                call_method(r, "GetFirst", &mut [])?;
                Ok(())
            })?;
            // Enumerate with GetFirst/GetNext (not Count/Item): under
            // IncludeRecurrences the collection can expand without bound, so
            // we must stream it and stop at MAX_CALENDAR_ITEMS. The `Restrict`
            // window is up to 12 days wider than the request on each side, so
            // occurrences outside the exact range are skipped here, before
            // they count towards the cap.
            let mut results = Vec::new();
            let mut current = call_method(&restricted, "GetFirst", &mut [])?;
            while let Ok(item) = IDispatch::try_from(&current) {
                let start = get_property(&item, "Start").ok().and_then(|v| variant_to_datetime(&v));
                let in_range = match start.map(|t| (t, date_filter.classify(t))) {
                    Some((_, Placement::InRange)) => true,
                    Some((t, Placement::OutsideWindow)) => {
                        return Err(date_misread_error("Start", t, date_filter));
                    }
                    Some((_, Placement::Before | Placement::After)) | None => false,
                };
                if in_range {
                    let summary = event_summary(&item, Some(calendar_store_id))?;
                    if event_matches(&summary, q) {
                        results.push(summary);
                        if results.len() >= MAX_CALENDAR_ITEMS {
                            break;
                        }
                    }
                }
                current = call_method(&restricted, "GetNext", &mut [])?;
            }
            Ok(results)
        })();
        match outcome {
            Ok(results) => return Ok(results),
            Err(err) if attempt < MAX_ATTEMPTS && is_transient_unknown_name(&err) => {
                std::thread::sleep(std::time::Duration::from_millis(300));
            }
            Err(err) => return Err(err),
        }
    }
    unreachable!("loop always returns on the final attempt")
}

/// `client.py::_task_summary`. `status` and `importance` are read as raw
/// numeric COM properties (not name lookups), with a missing value falling
/// back to Outlook's defaults exactly like the Python `getattr(..., default)`,
/// then converted to the friendly words the MCP API exposes via
/// `friendly::task_status_word`/`friendly::importance_word`.
fn task_summary(item: &IDispatch) -> Result<TaskSummary, ToolError> {
    Ok(TaskSummary {
        id: make_id(item)?,
        subject: variant_to_string(&get_property(item, "Subject").unwrap_or_default()),
        due_date: variant_to_iso_string(&get_property(item, "DueDate").unwrap_or_default()),
        complete: variant_to_bool(&get_property(item, "Complete").unwrap_or_default())
            .unwrap_or(false),
        status: crate::friendly::task_status_word(
            variant_to_i32(&get_property(item, "Status").unwrap_or_default())
                .unwrap_or(c::OL_TASK_NOT_STARTED),
        )
        .to_string(),
        importance: crate::friendly::importance_word(
            variant_to_i32(&get_property(item, "Importance").unwrap_or_default())
                .unwrap_or(c::OL_IMPORTANCE_NORMAL),
        )
        .to_string(),
        categories: get_item_categories(item),
    })
}

/// `client.py::_note_summary`. Notes have no native `Subject` property, so the
/// subject is derived from the first non-empty line of the `Body`: strip the
/// body, take the first line if anything remains (else an empty string), then
/// truncate to 120 *characters* (`first_line[:120]`). `str::lines()` splits on
/// `\n`/`\r\n`, mirroring Python's `splitlines()[0]` for note bodies.
fn note_summary(item: &IDispatch) -> Result<NoteSummary, ToolError> {
    let body = variant_to_string(&get_property(item, "Body").unwrap_or_default());
    let trimmed = body.trim();
    let first_line = if trimmed.is_empty() {
        ""
    } else {
        trimmed.lines().next().unwrap_or("")
    };
    Ok(NoteSummary {
        id: make_id(item)?,
        subject: first_line.chars().take(120).collect(),
        created: variant_to_iso_string(&get_property(item, "CreationTime").unwrap_or_default()),
        categories: get_item_categories(item),
    })
}

/// `client.py::_email_summary`.
fn email_summary(item: &IDispatch) -> Result<EmailSummary, ToolError> {
    // `getattr(item, "Attachments", None)` then `attachments and attachments.Count > 0`:
    // a non-mail item may lack an `Attachments` collection entirely, so tolerate a
    // missing property (fall back to 0) rather than propagating the COM error.
    let att_count = (|| -> Result<i32, ToolError> {
        let attachments = to_disp(get_property(item, "Attachments")?)?;
        Ok(variant_to_i32(&get_property(&attachments, "Count")?).unwrap_or(0))
    })()
    .unwrap_or(0);
    Ok(EmailSummary {
        id: make_id(item)?,
        subject: variant_to_string(&get_property(item, "Subject").unwrap_or_default()),
        sender: variant_to_string(&get_property(item, "SenderName").unwrap_or_default()),
        sender_email: variant_to_string(&get_property(item, "SenderEmailAddress").unwrap_or_default()),
        to: variant_to_string(&get_property(item, "To").unwrap_or_default()),
        received: variant_to_iso_string(&get_property(item, "ReceivedTime").unwrap_or_default()),
        unread: variant_to_bool(&get_property(item, "UnRead").unwrap_or_default()).unwrap_or(false),
        has_attachments: att_count > 0,
        categories: get_item_categories(item),
    })
}

/// `client.py::_compose`: build a `MailItem`, set recipients/subject/body.
fn compose(
    app: &IDispatch,
    to: &[String],
    subject: &str,
    body: &str,
    cc: Option<&[String]>,
    bcc: Option<&[String]>,
    html: bool,
) -> Result<IDispatch, ToolError> {
    let mail = to_disp(call_method(app, "CreateItem", &mut [variant_from_i32(c::OL_MAIL_ITEM)])?)?;
    put_property(&mail, "To", variant_from_str(&to.join("; ")))?;
    if let Some(cc) = cc {
        if !cc.is_empty() {
            put_property(&mail, "CC", variant_from_str(&cc.join("; ")))?;
        }
    }
    if let Some(bcc) = bcc {
        if !bcc.is_empty() {
            put_property(&mail, "BCC", variant_from_str(&bcc.join("; ")))?;
        }
    }
    put_property(&mail, "Subject", variant_from_str(subject))?;
    if html {
        put_property(&mail, "BodyFormat", variant_from_i32(c::OL_FORMAT_HTML))?;
        put_property(&mail, "HTMLBody", variant_from_str(body))?;
    } else {
        put_property(&mail, "BodyFormat", variant_from_i32(c::OL_FORMAT_PLAIN))?;
        put_property(&mail, "Body", variant_from_str(body))?;
    }
    Ok(mail)
}

/// Metadata for one attachment (`index` is COM's 1-based position). Shared by
/// `list_attachments` and `save_attachments`. `FileName`/`Size` are required;
/// the MAPI properties are optional (absent -> `None`/`false`). `html_body` is
/// the owning item's `HTMLBody`, used for `is_inline`.
fn attachment_info(att: &IDispatch, index: i32, html_body: &str) -> Result<AttachmentInfo, ToolError> {
    let filename = variant_to_string(&get_property(att, "FileName")?);
    let size = variant_to_i32(&get_property(att, "Size")?).unwrap_or(0);
    let att_type = get_property(att, "Type")
        .ok()
        .and_then(|v| variant_to_i32(&v))
        .unwrap_or(c::OL_BY_VALUE);
    let content_id = get_mapi_prop(att, c::PR_ATTACH_CONTENT_ID).map(|v| variant_to_string(&v));
    let mime_tag = get_mapi_prop(att, c::PR_ATTACH_MIME_TAG).map(|v| variant_to_string(&v));
    let hidden = get_mapi_prop(att, c::PR_ATTACHMENT_HIDDEN)
        .and_then(|v| variant_to_bool(&v))
        .unwrap_or(false);
    let content_id = clean_content_id(content_id.as_deref());
    Ok(AttachmentInfo {
        index,
        mime_type: guess_mime(mime_tag.as_deref(), &filename),
        filename,
        size,
        att_type: c::attachment_type_name(att_type).to_string(),
        is_inline: is_inline(content_id.as_deref(), hidden, html_body),
        content_id,
        hidden,
    })
}

/// An item's `HTMLBody`, or "" when it has none or it can't be read.
fn item_html_body(item: &IDispatch) -> String {
    get_property(item, "HTMLBody")
        .map(|v| variant_to_string(&v))
        .unwrap_or_default()
}

/// Pick the attachment whose Content-ID matches `wanted` (already normalized
/// by `normalize_cid_request`), case-insensitively. `content_ids[i]` is the
/// cleaned Content-ID of the i-th attachment (`None` if it has none); returns
/// that position. Errors list the available Content-IDs, or point to
/// `list_attachments` when the email has none.
fn select_by_content_id(wanted: &str, content_ids: &[Option<String>]) -> Result<usize, ToolError> {
    let wanted_lower = wanted.to_lowercase();
    if let Some(pos) = content_ids
        .iter()
        .position(|cid| cid.as_deref().is_some_and(|c| c.to_lowercase() == wanted_lower))
    {
        return Ok(pos);
    }
    let available: Vec<&str> = content_ids.iter().flatten().map(String::as_str).collect();
    if available.is_empty() {
        Err(ToolError::new(format!(
            "Content-ID '{wanted}' not found: this email has no attachments with a \
             Content-ID (use list_attachments to see its attachments)."
        )))
    } else {
        Err(ToolError::new(format!(
            "Content-ID '{wanted}' not found. Available Content-IDs: {}",
            available.join(", ")
        )))
    }
}

/// `get_inline_image`'s error for an attachment over `MAX_INLINE_IMAGE_BYTES`.
fn inline_image_too_big(info: &AttachmentInfo) -> ToolError {
    let name = if info.filename.is_empty() {
        info.content_id.as_deref().unwrap_or_default()
    } else {
        &info.filename
    };
    ToolError::new(format!(
        "Attachment '{name}' exceeds the {} MB limit for get_inline_image; use \
         save_attachments to save it to disk instead.",
        MAX_INLINE_IMAGE_BYTES / (1024 * 1024)
    ))
}

/// `data:<mime>;base64,<payload>` for `bytes`.
fn data_uri(mime_type: &str, bytes: &[u8]) -> String {
    use base64::Engine as _;
    format!(
        "data:{mime_type};base64,{}",
        base64::engine::general_purpose::STANDARD.encode(bytes)
    )
}

/// A freshly created, uniquely named directory under `std::env::temp_dir()`,
/// removed (with its contents) when dropped, so every exit path cleans up.
struct TempDirGuard(std::path::PathBuf);

impl TempDirGuard {
    fn create() -> Result<Self, ToolError> {
        use std::sync::atomic::{AtomicU64, Ordering};
        static COUNTER: AtomicU64 = AtomicU64::new(0);
        let nanos = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let n = COUNTER.fetch_add(1, Ordering::Relaxed);
        let dir = std::env::temp_dir()
            .join(format!("outlook-mcp-rs-{}-{nanos}-{n}", std::process::id()));
        // `create_dir` (not `_all`) fails if the name is somehow taken, so we
        // never adopt (and later delete) someone else's directory.
        std::fs::create_dir(&dir).map_err(|e| {
            ToolError::new(format!("Could not create temp directory {:?}: {e}", dir.display()))
        })?;
        Ok(Self(dir))
    }

    fn path(&self) -> &std::path::Path {
        &self.0
    }
}

impl Drop for TempDirGuard {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}

/// Read an attachment's bytes: `SaveAsFile` into a private temp dir (always
/// removed afterwards), then read the file back.
fn read_attachment_bytes(att: &IDispatch, info: &AttachmentInfo) -> Result<Vec<u8>, ToolError> {
    let dir = TempDirGuard::create()?;
    let name = if info.filename.is_empty() {
        format!("attachment-{}", info.index)
    } else {
        info.filename.clone()
    };
    let target = dir.path().join(safe_filename(&name));
    call_method(att, "SaveAsFile", &mut [variant_from_str(&target.to_string_lossy())])?;
    std::fs::read(&target).map_err(|e| {
        ToolError::new(format!("Could not read the saved attachment {:?}: {e}", target.display()))
    })
}

/// Shallow-merge `extra`'s keys into `base` (both JSON objects; a non-object
/// `extra` is ignored). Used to append per-call fields to a serialized struct.
fn merge_json_objects(mut base: Value, extra: Value) -> Value {
    if let (Value::Object(fields), Value::Object(more)) = (&mut base, extra) {
        fields.extend(more);
    }
    base
}

/// Attach local files to a mail/reply item. Validates every path exists
/// FIRST (so a bad path fails before anything is sent), then adds each via
/// `MailItem.Attachments.Add(path)`.
fn attach_files(mail: &IDispatch, paths: &[String]) -> Result<(), ToolError> {
    for p in paths {
        if !std::path::Path::new(p).is_file() {
            return Err(ToolError::new(format!("attachment not found: {p}")));
        }
    }
    let atts = to_disp(get_property(mail, "Attachments")?)?;
    for p in paths {
        call_method(&atts, "Add", &mut [variant_from_str(p)])?;
    }
    Ok(())
}

/// Owns the temp dirs holding decoded inline-image data and removes them
/// when dropped, so they're cleaned up after `Save`/`Send` and on every
/// error path alike.
#[derive(Default)]
struct InlineTempDirs(Vec<std::path::PathBuf>);

impl InlineTempDirs {
    /// Create a fresh, uniquely named dir under the system temp dir (one per
    /// image, so equal filenames can't collide) and track it for cleanup.
    fn create(&mut self) -> Result<std::path::PathBuf, ToolError> {
        use std::sync::atomic::{AtomicU64, Ordering};
        static COUNTER: AtomicU64 = AtomicU64::new(0);
        let nanos = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let dir = std::env::temp_dir().join(format!(
            "outlook-mcp-rs-inline-{}-{nanos}-{}",
            std::process::id(),
            COUNTER.fetch_add(1, Ordering::Relaxed)
        ));
        std::fs::create_dir(&dir)
            .map_err(|e| ToolError::new(format!("could not create temp dir for inline image: {e}")))?;
        self.0.push(dir.clone());
        Ok(dir)
    }
}

impl Drop for InlineTempDirs {
    fn drop(&mut self) {
        for dir in &self.0 {
            let _ = std::fs::remove_dir_all(dir);
        }
    }
}

/// Validate `inline_images` (if any) before any COM item is created.
fn validated_inline_images(images: Option<&[InlineImage]>, html: bool)
    -> Result<Vec<ValidatedInlineImage>, ToolError> {
    match images {
        Some(images) if !images.is_empty() => validate_inline_images(images, html),
        _ => Ok(Vec::new()),
    }
}

/// Attach validated inline images as hidden Content-ID attachments the HTML
/// body references via `cid:CONTENT_ID`. Base64 data is written to a temp
/// file first (tracked in `temp` for cleanup), since `Attachments.Add` only
/// takes a path.
fn attach_inline_images(
    mail: &IDispatch,
    images: &[ValidatedInlineImage],
    temp: &mut InlineTempDirs,
) -> Result<(), ToolError> {
    if images.is_empty() {
        return Ok(());
    }
    let atts = to_disp(get_property(mail, "Attachments")?)?;
    for img in images {
        let path = match &img.source {
            InlineImageSource::Path(p) => p.clone(),
            InlineImageSource::Data(bytes) => {
                let file = temp.create()?.join(&img.filename);
                std::fs::write(&file, bytes).map_err(|e| {
                    ToolError::new(format!("could not write inline image {:?} to a temp file: {e}", img.content_id))
                })?;
                file.to_string_lossy().into_owned()
            }
        };
        // Position 0: don't render an attachment icon in the body.
        let att = to_disp(call_method(&atts, "Add", &mut [
            variant_from_str(&path),
            variant_from_i32(c::OL_BY_VALUE),
            variant_from_i32(0),
            variant_from_str(&img.filename),
        ])?)?;
        let pa = to_disp(get_property(&att, "PropertyAccessor")?)?;
        call_method(&pa, "SetProperty", &mut [
            variant_from_str(c::PR_ATTACH_CONTENT_ID),
            variant_from_str(&img.content_id),
        ])?;
        call_method(&pa, "SetProperty", &mut [
            variant_from_str(c::PR_ATTACH_MIME_TAG),
            variant_from_str(&img.mime_type),
        ])?;
        call_method(&pa, "SetProperty", &mut [
            variant_from_str(c::PR_ATTACHMENT_HIDDEN),
            variant_from_bool(true),
        ])?;
    }
    Ok(())
}

impl OutlookClient for WindowsOutlookClient {
    // ---- Email (implemented in Task 12) --------------------------------

    fn list_folders(&self) -> Result<Vec<FolderInfo>, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let inbox = to_disp(call_method(
                &ns,
                "GetDefaultFolder",
                &mut [variant_from_i32(c::OL_FOLDER_INBOX)],
            )?)?;
            let root = to_disp(get_property(&inbox, "Parent")?)?;

            fn walk(
                folder: &IDispatch,
                path: &str,
                depth: u32,
                results: &mut Vec<FolderInfo>,
            ) -> Result<(), ToolError> {
                let name = variant_to_string(&get_property(folder, "Name")?);
                // `folder.Items.Count` can raise for some special folders;
                // fall back to 0 like the Python try/except does.
                let item_count = (|| -> Result<i32, ToolError> {
                    let items = to_disp(get_property(folder, "Items")?)?;
                    Ok(variant_to_i32(&get_property(&items, "Count")?).unwrap_or(0))
                })()
                .unwrap_or(0);
                let unread = variant_to_i32(&get_property(folder, "UnReadItemCount")?).unwrap_or(0);
                results.push(FolderInfo {
                    name,
                    path: path.to_string(),
                    items: item_count,
                    unread,
                });
                if depth >= 3 {
                    return Ok(());
                }
                let subfolders = to_disp(get_property(folder, "Folders")?)?;
                let count = variant_to_i32(&get_property(&subfolders, "Count")?).unwrap_or(0);
                for i in 1..=count {
                    let sub =
                        to_disp(call_method(&subfolders, "Item", &mut [variant_from_i32(i)])?)?;
                    let sub_name = variant_to_string(&get_property(&sub, "Name")?);
                    walk(&sub, &format!("{path}/{sub_name}"), depth + 1, results)?;
                }
                Ok(())
            }

            let mut results = Vec::new();
            let root_folders = to_disp(get_property(&root, "Folders")?)?;
            let count = variant_to_i32(&get_property(&root_folders, "Count")?).unwrap_or(0);
            for i in 1..=count {
                let sub = to_disp(call_method(&root_folders, "Item", &mut [variant_from_i32(i)])?)?;
                let sub_name = variant_to_string(&get_property(&sub, "Name")?);
                walk(&sub, &sub_name, 1, &mut results)?;
            }
            Ok(results)
        })
    }

    // Cheap filters become sequential COM `Restrict` calls (they AND together);
    // `category`, `has_attachments`, and `flagged` are filtered client-side
    // while iterating.
    fn list_emails(&self, q: EmailQuery) -> Result<Vec<EmailSummary>, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let count = q.count.clamp(1, MAX_EMAIL_COUNT);
            let folder_obj = resolve_folder(&ns, Some(&q.folder))?;
            let mut items = to_disp(get_property(&folder_obj, "Items")?)?;

            // Sender: DASL @SQL against fromname + fromemail.
            if let Some(from) = q.from.as_deref().filter(|s| !s.is_empty()) {
                let e = from.replace('\'', "''");
                let dasl = format!(
                    "@SQL=(\"urn:schemas:httpmail:fromname\" LIKE '%{e}%' \
                     OR \"urn:schemas:httpmail:fromemail\" LIKE '%{e}%')"
                );
                items = to_disp(call_method(&items, "Restrict", &mut [variant_from_str(&dasl)])?)?;
            }
            if q.unread_only {
                items = to_disp(call_method(
                    &items,
                    "Restrict",
                    &mut [variant_from_str("[UnRead] = True")],
                )?)?;
            }
            // `flagged` is deliberately NOT a Restrict call: confirmed via a
            // raw PowerShell COM probe outside this codebase that
            // `Items.Restrict("[FlagStatus] = 2")` doesn't reliably match on
            // this account class, even though `Item.FlagStatus` reads
            // correctly when read directly per-item (modern Outlook/M365
            // flags sync through the To-Do integration rather than classic
            // MAPI, and the legacy DASL bracket filter doesn't see that
            // state reliably). Filtered client-side below instead, alongside
            // category/has_attachments.
            if q.high_importance {
                items = to_disp(call_method(
                    &items,
                    "Restrict",
                    &mut [variant_from_str("[Importance] = 2")],
                )?)?;
            }
            // Date filters: since_days (relative) and received_after/before
            // (absolute) combine into one exact, inclusive range. `Restrict`
            // only gets a slightly wider, unambiguous date-only window (see
            // `jet_dates`), so no regional day/month order can misread it;
            // the exact bounds are applied per item below, before paging.
            let mut after: Option<chrono::NaiveDateTime> = None;
            if let Some(days) = q.since_days.filter(|&d| d != 0) {
                after = Some(chrono::Local::now().naive_local() - chrono::Duration::days(days as i64));
            }
            if let Some(s) = q.received_after.as_deref().filter(|s| !s.is_empty()) {
                let dt = parse_dt(s, "received_after")?;
                after = Some(after.map_or(dt, |a| a.max(dt)));
            }
            let before = match q.received_before.as_deref().filter(|s| !s.is_empty()) {
                Some(s) => Some(parse_dt(s, "received_before")?),
                None => None,
            };
            let date_filter = DateFilter::new(after, before);
            if let Some(f) = &date_filter {
                let filters = f.restrict_filters("ReceivedTime", user_short_date_order().as_ref());
                // `Count` forces Outlook to evaluate the filter.
                items = restrict_dates(&items, &filters, |r| {
                    get_property(r, "Count")?;
                    Ok(())
                })?;
            }

            // Text query near-last, so the fallback scan below only walks
            // items that already passed every other filter. DASL @SQL across
            // subject/sender/body (escaped).
            let mut text_fallback: Option<String> = None;
            if let Some(query) = q.query.as_deref().filter(|s| !s.is_empty()) {
                let query: String = query.nfc().collect();
                let e = query.replace('\'', "''");
                let dasl = format!(
                    "@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%{e}%' \
                     OR \"urn:schemas:httpmail:fromname\" LIKE '%{e}%' \
                     OR \"urn:schemas:httpmail:textdescription\" LIKE '%{e}%')"
                );
                let matched =
                    to_disp(call_method(&items, "Restrict", &mut [variant_from_str(&dasl)])?)?;
                let matched_count = variant_to_i32(&get_property(&matched, "Count")?).unwrap_or(0);
                // DASL LIKE has been observed to return nothing for Hebrew
                // (and other non-Latin) terms even when matching mail exists
                // (issue #2). For a non-ASCII query, treat an empty DASL
                // result as unreliable and scan the pre-filtered items
                // client-side instead. ASCII queries stay DASL-only.
                if matched_count == 0 && !query.is_ascii() {
                    text_fallback = Some(query);
                } else {
                    items = matched;
                }
            }

            // Recipient: applied last so the fallback below scans an
            // already-narrowed set. First try DASL on the To/CC display
            // strings (fast, evaluated by the store, no per-item COM calls).
            // Those strings are what Outlook shows in the To/CC lines: for
            // resolved Exchange/contact recipients that's the display NAME
            // ("Ada Lovelace"), not the address, so an address needle can
            // match nothing even though the mail was sent to that address.
            // So when the DASL restrict comes back empty, fall back to
            // scanning each item's Recipients (Name, Address, SMTP address)
            // client-side, over at most RECIPIENT_SCAN_LIMIT newest items.
            // A non-empty DASL result is kept as-is: every hit genuinely has
            // the needle in its To/CC line, and it avoids the per-item walk.
            let mut to_scan: Option<&str> = None;
            if let Some(to) = q.to.as_deref().filter(|s| !s.is_empty()) {
                let e = to.replace('\'', "''");
                let dasl = format!(
                    "@SQL=(\"urn:schemas:httpmail:displayto\" LIKE '%{e}%' \
                     OR \"urn:schemas:httpmail:displaycc\" LIKE '%{e}%')"
                );
                let restricted =
                    to_disp(call_method(&items, "Restrict", &mut [variant_from_str(&dasl)])?)?;
                if variant_to_i32(&get_property(&restricted, "Count")?).unwrap_or(0) > 0 {
                    items = restricted;
                } else {
                    to_scan = Some(to);
                }
            }

            call_method(
                &items,
                "Sort",
                &mut [variant_from_str("[ReceivedTime]"), variant_from_bool(true)],
            )?;

            // Client-side filters: the exact date range, non-ASCII text
            // fallback and `to` recipient fallback (if active) + category +
            // has_attachments + flagged. Lazily build each summary and keep it
            // only if it passes; then `take_page` skips the first `offset`
            // matches (after these filters, so pages line up) and stops at
            // count. Newest first, so the first item older than the date range
            // ends the scan.
            let cat_want = q.category.as_deref().map(|c| c.to_lowercase());
            let mut total = variant_to_i32(&get_property(&items, "Count")?).unwrap_or(0);
            if text_fallback.is_some() {
                // The fallback reads Subject/SenderName (and maybe Body) per
                // item, so cap how far back it scans (newest first).
                total = total.min(MAX_SCAN_ITEMS);
            }
            if to_scan.is_some() {
                // Same for the per-item Recipients walk of the `to` fallback.
                total = total.min(RECIPIENT_SCAN_LIMIT);
            }
            let matches = (1..=total)
                .map(|i| (|| -> Result<Scan<EmailSummary>, ToolError> {
                    let item = to_disp(call_method(&items, "Item", &mut [variant_from_i32(i)])?)?;
                    if let Some(f) = &date_filter {
                        let received = get_property(&item, "ReceivedTime")
                            .ok()
                            .and_then(|v| variant_to_datetime(&v));
                        match received.map(|t| (t, f.classify(t))) {
                            Some((_, Placement::InRange)) => {}
                            Some((_, Placement::Before)) => return Ok(Scan::Stop),
                            Some((t, Placement::OutsideWindow)) => {
                                return Err(date_misread_error("ReceivedTime", t, f));
                            }
                            Some((_, Placement::After)) | None => return Ok(Scan::Skip),
                        }
                    }
                    if text_fallback.as_deref().is_some_and(|query| !email_text_matches(&item, query)) {
                        return Ok(Scan::Skip);
                    }
                    if to_scan.is_some_and(|to| !recipient_matches(to, &recipient_strings(&item))) {
                        return Ok(Scan::Skip);
                    }
                    let summary = email_summary(&item)?;
                    if let Some(want) = &cat_want {
                        if !summary.categories.iter().any(|c| c.to_lowercase() == *want) {
                            return Ok(Scan::Skip);
                        }
                    }
                    if let Some(want_att) = q.has_attachments {
                        if summary.has_attachments != want_att {
                            return Ok(Scan::Skip);
                        }
                    }
                    if q.flagged {
                        // "Flagged" means any non-zero FlagStatus: both a
                        // follow-up flag (OL_FLAG_MARKED = 2) and a completed
                        // flag (OL_FLAG_COMPLETE = 1) count; 0 = no flag/cleared.
                        let flag_status =
                            variant_to_i32(&get_property(&item, "FlagStatus").unwrap_or_default())
                                .unwrap_or(0);
                        if flag_status == 0 {
                            return Ok(Scan::Skip);
                        }
                    }
                    Ok(Scan::Keep(summary))
                })())
                .take_while(|r| !matches!(r, Ok(Scan::Stop)))
                .filter_map(|r| match r {
                    Ok(Scan::Keep(summary)) => Some(Ok(summary)),
                    Ok(_) => None,
                    Err(e) => Some(Err(e)),
                });
            take_page(matches, q.offset, count)
        })
    }

    fn get_email(&self, email_id: String, prefer_html: bool, max_body_chars: Option<u32>)
        -> Result<EmailDetail, ToolError> {
        let limit = clamp_body_limit(max_body_chars);
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &email_id)?;
            let summary = email_summary(&item)?;
            // Python's `get_email` reads these via `getattr(item, "X", "") or ""`
            // so a non-mail item (MeetingItem, ReportItem, …) that lacks CC/BCC/
            // Body/HTMLBody yields graceful partial detail rather than a COM error.
            let cc = variant_to_string(&get_property(&item, "CC").unwrap_or_default());
            let bcc = variant_to_string(&get_property(&item, "BCC").unwrap_or_default());
            // `*_length` is the full original length in chars, so a caller
            // that sees a `*_truncated` flag knows what limit to ask for.
            let full_body = variant_to_string(&get_property(&item, "Body").unwrap_or_default());
            let body_length = full_body.chars().count();
            let (body, body_truncated) = truncate(&full_body, limit);
            let (html_body, html_truncated, html_length) = if prefer_html {
                let full_html = variant_to_string(&get_property(&item, "HTMLBody").unwrap_or_default());
                let (html, truncated) = truncate(&full_html, limit);
                (Some(html), Some(truncated), Some(full_html.chars().count()))
            } else {
                (None, None, None)
            };
            // `getattr(item, "Attachments", None)` then `attachments and attachments.Count`:
            // tolerate an item that has no `Attachments` collection at all (falls back
            // to an empty list) rather than propagating the COM error, mirroring
            // `email_summary`'s already-fixed handling.
            let attachments = (|| -> Result<Vec<String>, ToolError> {
                let attachments_obj = to_disp(get_property(&item, "Attachments")?)?;
                let att_count =
                    variant_to_i32(&get_property(&attachments_obj, "Count")?).unwrap_or(0);
                let mut names = Vec::new();
                for i in 1..=att_count {
                    let att =
                        to_disp(call_method(&attachments_obj, "Item", &mut [variant_from_i32(i)])?)?;
                    names.push(variant_to_string(&get_property(&att, "FileName")?));
                }
                Ok(names)
            })()
            .unwrap_or_default();
            let message_class = variant_to_string(&get_property(&item, "MessageClass").unwrap_or_default());
            let item_type = crate::friendly::item_type_from_class(&message_class).to_string();

            // A MeetingItem exposes GetAssociatedAppointment; a plain MailItem
            // does not. Build the meeting block from the associated appointment.
            let (is_meeting, meeting) = if has_member(&item, "GetAssociatedAppointment") {
                let appt = to_disp(call_method(
                    &item, "GetAssociatedAppointment", &mut [variant_from_bool(false)],
                )?)?;
                let info = MeetingInfo {
                    meeting_type: crate::friendly::meeting_type_from_class(&message_class).to_string(),
                    start: variant_to_iso_string(&get_property(&appt, "Start").unwrap_or_default()),
                    end: variant_to_iso_string(&get_property(&appt, "End").unwrap_or_default()),
                    location: variant_to_string(&get_property(&appt, "Location").unwrap_or_default()),
                    organizer: variant_to_string(&get_property(&appt, "Organizer").unwrap_or_default()),
                    required_attendees: variant_to_string(&get_property(&appt, "RequiredAttendees").unwrap_or_default()),
                    optional_attendees: variant_to_string(&get_property(&appt, "OptionalAttendees").unwrap_or_default()),
                    is_recurring: variant_to_bool(&get_property(&appt, "IsRecurring").unwrap_or_default()).unwrap_or(false),
                };
                (true, Some(info))
            } else {
                (false, None)
            };
            Ok(EmailDetail {
                summary,
                cc,
                bcc,
                body,
                body_truncated,
                body_length,
                html_body,
                html_truncated,
                html_length,
                attachments,
                item_type,
                is_meeting,
                meeting,
            })
        })
    }

    fn send_email(
        &self,
        to: Vec<String>,
        subject: String,
        body: String,
        cc: Option<Vec<String>>,
        bcc: Option<Vec<String>>,
        html: bool,
        attachments: Option<Vec<String>>,
        inline_images: Option<Vec<InlineImage>>,
    ) -> Result<Value, ToolError> {
        if to.is_empty() {
            return Err(ToolError::new(
                "send_email requires at least one recipient in 'to'.",
            ));
        }
        let images = validated_inline_images(inline_images.as_deref(), html)?;
        self.with_com(|| {
            let (app, _ns) = mapi()?;
            let mail = compose(&app, &to, &subject, &body, cc.as_deref(), bcc.as_deref(), html)?;
            if let Some(atts) = attachments.as_deref() {
                attach_files(&mail, atts)?;
            }
            // Dropped at the end of this closure: temp files outlive Send.
            let mut temp = InlineTempDirs::default();
            attach_inline_images(&mail, &images, &mut temp)?;
            call_method(&mail, "Send", &mut [])?;
            Ok(json!({"status": "sent", "to": to.join("; "), "subject": subject}))
        })
    }

    fn create_draft(
        &self,
        to: Vec<String>,
        subject: String,
        body: String,
        cc: Option<Vec<String>>,
        bcc: Option<Vec<String>>,
        html: bool,
        attachments: Option<Vec<String>>,
        inline_images: Option<Vec<InlineImage>>,
    ) -> Result<Value, ToolError> {
        let images = validated_inline_images(inline_images.as_deref(), html)?;
        self.with_com(|| {
            let (app, _ns) = mapi()?;
            let mail = compose(&app, &to, &subject, &body, cc.as_deref(), bcc.as_deref(), html)?;
            if let Some(atts) = attachments.as_deref() {
                attach_files(&mail, atts)?;
            }
            // Dropped at the end of this closure: temp files outlive Save.
            let mut temp = InlineTempDirs::default();
            attach_inline_images(&mail, &images, &mut temp)?;
            call_method(&mail, "Save", &mut [])?; // Save first so EntryID exists
            let id = make_id(&mail)?;
            Ok(json!({"status": "draft_saved", "id": id, "subject": subject}))
        })
    }

    fn reply_email(
        &self,
        email_id: String,
        body: String,
        reply_all: bool,
        html: bool,
        send: bool,
        attachments: Option<Vec<String>>,
    ) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &email_id)?;
            let reply = to_disp(call_method(
                &item,
                if reply_all { "ReplyAll" } else { "Reply" },
                &mut [],
            )?)?;
            if html {
                let existing = variant_to_string(&get_property(&reply, "HTMLBody")?);
                put_property(&reply, "HTMLBody", variant_from_str(&format!("{body}{existing}")))?;
            } else {
                let existing = variant_to_string(&get_property(&reply, "Body")?);
                put_property(&reply, "Body", variant_from_str(&format!("{body}\n\n{existing}")))?;
            }
            if let Some(atts) = attachments.as_deref() {
                attach_files(&reply, atts)?;
            }
            if send {
                // Read Subject *before* Send() — Outlook invalidates the COM
                // item once sent (a well-known lifecycle rule), so reading a
                // property off `reply` afterward throws "The item has been
                // moved or deleted." (0x80020009) even though the send itself
                // succeeded. Mirrors send_email's pattern of never touching
                // the item post-Send.
                let subject = variant_to_string(&get_property(&reply, "Subject")?);
                call_method(&reply, "Send", &mut [])?;
                Ok(json!({"status": "sent", "subject": subject}))
            } else {
                call_method(&reply, "Save", &mut [])?;
                let id = make_id(&reply)?;
                let subject = variant_to_string(&get_property(&reply, "Subject")?);
                Ok(json!({"status": "draft_saved", "id": id, "subject": subject}))
            }
        })
    }

    fn update_email(&self, u: EmailUpdate) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &u.email_id)?;
            let mut changed: Vec<&str> = Vec::new();

            // ---- state changes first (they address the item by its current id) ----

            if let Some(read) = u.mark_read {
                // UnRead is the inverse of "read". Save so a mark_read-only
                // update persists even when no later Save/Move follows.
                put_property(&item, "UnRead", variant_from_bool(!read))?;
                call_method(&item, "Save", &mut [])?;
                changed.push("mark_read");
            }

            if let Some(flag) = &u.flag {
                match flag.to_lowercase().as_str() {
                    "follow_up" => {
                        // MarkAsTask flags for follow-up with no due date.
                        call_method(&item, "MarkAsTask", &mut [variant_from_i32(c::OL_MARK_NO_DATE)])?;
                    }
                    "complete" => {
                        put_property(&item, "FlagStatus", variant_from_i32(c::OL_FLAG_COMPLETE))?;
                    }
                    "clear" => {
                        // ClearTaskFlag removes the follow-up flag entirely.
                        call_method(&item, "ClearTaskFlag", &mut [])?;
                    }
                    other => {
                        return Err(ToolError::new(format!(
                            "invalid flag {other:?}: expected \"follow_up\", \"complete\", or \"clear\""
                        )));
                    }
                }
                call_method(&item, "Save", &mut [])?;
                changed.push("flag");
            }

            // Categories: read the current set once, then add/remove against it,
            // so tagging never wipes existing categories.
            if u.add_categories.is_some() || u.remove_categories.is_some() {
                let mut cats = get_item_categories(&item);
                if let Some(add) = &u.add_categories {
                    for a in add {
                        if !cats.iter().any(|c| c.eq_ignore_ascii_case(a)) {
                            cats.push(a.clone());
                        }
                    }
                    changed.push("add_categories");
                }
                if let Some(remove) = &u.remove_categories {
                    cats.retain(|c| !remove.iter().any(|r| r.eq_ignore_ascii_case(c)));
                    changed.push("remove_categories");
                }
                set_item_categories(&item, &cats)?;
                call_method(&item, "Save", &mut [])?;
            }

            if let Some(imp) = &u.importance {
                let id = c::importance_name_to_id(imp).ok_or_else(|| {
                    ToolError::new(format!(
                        "invalid importance {imp:?}: expected \"low\", \"normal\", or \"high\""
                    ))
                })?;
                put_property(&item, "Importance", variant_from_i32(id))?;
                call_method(&item, "Save", &mut [])?;
                changed.push("importance");
            }

            // ---- move last (Move changes the EntryID) ----

            let id = if let Some(dest) = &u.move_to {
                let target = resolve_folder(&ns, Some(dest))?;
                let moved = to_disp(call_method(
                    &item, "Move", &mut [VARIANT::from(target.clone())],
                )?)?;
                changed.push("move_to");
                make_id(&moved)? // EntryID changed — return the new id.
            } else {
                u.email_id.clone()
            };

            Ok(json!({"status": "updated", "id": id, "changed": changed}))
        })
    }

    fn update_draft(&self, u: DraftUpdate) -> Result<Value, ToolError> {
        // Validate everything before the item is touched.
        validate_draft_update(&u)?;
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &u.draft_id)?;
            // `Sent` is true for anything sent or received; only an unsent
            // draft may be edited.
            if variant_to_bool(&get_property(&item, "Sent")?).unwrap_or(false) {
                return Err(ToolError::new(
                    "only unsent drafts can be edited; this item has already been sent or \
                     received. Use reply_email or create_draft instead.",
                ));
            }

            if let Some(subject) = &u.subject {
                put_property(&item, "Subject", variant_from_str(subject))?;
            }
            // Set BodyFormat before the body so Outlook doesn't re-convert it.
            if let Some(body) = &u.body {
                put_property(&item, "BodyFormat", variant_from_i32(c::OL_FORMAT_PLAIN))?;
                put_property(&item, "Body", variant_from_str(body))?;
            }
            if let Some(html) = &u.html_body {
                put_property(&item, "BodyFormat", variant_from_i32(c::OL_FORMAT_HTML))?;
                put_property(&item, "HTMLBody", variant_from_str(html))?;
            }
            // Recipients replace that whole line; an empty list clears it.
            if let Some(to) = &u.to {
                put_property(&item, "To", variant_from_str(&to.join("; ")))?;
            }
            if let Some(cc) = &u.cc {
                put_property(&item, "CC", variant_from_str(&cc.join("; ")))?;
            }
            if let Some(bcc) = &u.bcc {
                put_property(&item, "BCC", variant_from_str(&bcc.join("; ")))?;
            }
            // Appended; existing attachments are kept.
            if let Some(atts) = u.attachments.as_deref().filter(|a| !a.is_empty()) {
                attach_files(&item, atts)?;
            }

            call_method(&item, "Save", &mut [])?; // one Save for all changes; never Send
            let id = make_id(&item)?;
            Ok(json!({"status": "draft_updated", "id": id, "changed": draft_update_changes(&u)}))
        })
    }

    fn delete_email(&self, email_id: String, permanent: bool) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &email_id)?;
            let subject = variant_to_string(&get_property(&item, "Subject")?);
            if !permanent {
                call_method(&item, "Delete", &mut [])?;
                return Ok(json!({
                    "status": "deleted", "subject": subject, "permanent": false,
                    "note": "Moved to Deleted Items.",
                }));
            }

            // OOM has no hard-delete call; Delete() on an item already in
            // Deleted Items is permanent, so move it there first (Outlook's
            // shift+delete).
            let deleted = deleted_items_for(&ns, &item)?;
            let deleted_id = variant_to_string(&get_property(&deleted, "EntryID")?);
            // An unreadable parent id just means "move first" (always safe).
            let parent_id = (|| -> Result<String, ToolError> {
                let parent = to_disp(get_property(&item, "Parent")?)?;
                Ok(variant_to_string(&get_property(&parent, "EntryID")?))
            })()
            .unwrap_or_default();
            if permanent_delete_needs_move(&parent_id, &deleted_id) {
                // Move returns the item in its new home, but Delete() on that
                // returned object is a silent no-op (confirmed with a raw
                // PowerShell COM probe on an Outlook.com store, with or
                // without a delay): the item stays in Deleted Items. Re-open
                // the moved item by EntryID and delete that fresh object.
                let moved = to_disp(call_method(
                    &item, "Move", &mut [VARIANT::from(deleted.clone())],
                )?)?;
                let moved_id = variant_to_string(&get_property(&moved, "EntryID")?);
                let store_id = variant_to_string(&get_property(&deleted, "StoreID")?);
                let fresh = to_disp(call_method(
                    &ns,
                    "GetItemFromID",
                    &mut [variant_from_str(&moved_id), variant_from_str(&store_id)],
                )?)?;
                call_method(&fresh, "Delete", &mut [])?;
            } else {
                call_method(&item, "Delete", &mut [])?;
            }
            Ok(json!({
                "status": "deleted", "subject": subject, "permanent": true,
                "note": "Permanently deleted (not recoverable from Deleted Items).",
            }))
        })
    }

    fn empty_deleted_items(&self, confirm: bool) -> Result<Value, ToolError> {
        require_empty_confirm(confirm)?;
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let folder = to_disp(call_method(
                &ns,
                "GetDefaultFolder",
                &mut [variant_from_i32(c::OL_FOLDER_DELETED_ITEMS)],
            )?)?;
            let (mut items_deleted, mut folders_deleted, mut failed) = (0, 0, 0);

            // Walk backwards: deleting during forward iteration skips items
            // as the collection re-indexes. One stuck item must not abort
            // the rest, so per-item errors are only counted.
            let items = to_disp(get_property(&folder, "Items")?)?;
            let count = variant_to_i32(&get_property(&items, "Count")?).unwrap_or(0);
            for i in (1..=count).rev() {
                let deleted = (|| -> Result<(), ToolError> {
                    let it = to_disp(call_method(&items, "Item", &mut [variant_from_i32(i)])?)?;
                    call_method(&it, "Delete", &mut [])?;
                    Ok(())
                })();
                match deleted {
                    Ok(_) => items_deleted += 1,
                    Err(_) => failed += 1,
                }
            }

            let subfolders = to_disp(get_property(&folder, "Folders")?)?;
            let count = variant_to_i32(&get_property(&subfolders, "Count")?).unwrap_or(0);
            for i in (1..=count).rev() {
                let deleted = (|| -> Result<(), ToolError> {
                    let f = to_disp(call_method(&subfolders, "Item", &mut [variant_from_i32(i)])?)?;
                    call_method(&f, "Delete", &mut [])?;
                    Ok(())
                })();
                match deleted {
                    Ok(_) => folders_deleted += 1,
                    Err(_) => failed += 1,
                }
            }

            Ok(json!({
                "status": "emptied", "items_deleted": items_deleted,
                "folders_deleted": folders_deleted, "failed": failed,
            }))
        })
    }

    // ---- Calendar (Task 13) --------------------------------------------

    fn list_events(&self, q: EventQuery) -> Result<Vec<EventSummary>, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let start = match &q.start_date {
                Some(s) => parse_dt(s, "start_date")?,
                None => chrono::Local::now()
                    .date_naive()
                    .and_hms_opt(0, 0, 0)
                    .unwrap(),
            };
            let mut end = match &q.end_date {
                Some(s) => parse_dt(s, "end_date")?,
                None => start + chrono::Duration::days(7),
            };
            // If only a bare date was given for the end, treat it as the whole
            // end day (Python: `end.time() == time.min and "T" not in end_date`).
            if let Some(ed) = &q.end_date {
                if end.time() == chrono::NaiveTime::MIN && !ed.contains('T') {
                    end = end.date().and_hms_micro_opt(23, 59, 59, 999_999).unwrap();
                }
            }
            // `calendar_of`: open another person's shared calendar; otherwise
            // our own default calendar (current behavior).
            let calendar = match q.calendar_of.as_deref().filter(|s| !s.is_empty()) {
                Some(person) => {
                    let recipient = to_disp(call_method(
                        &ns, "CreateRecipient", &mut [variant_from_str(person)],
                    )?)?;
                    let resolved = variant_to_bool(&call_method(&recipient, "Resolve", &mut [])?)
                        .unwrap_or(false);
                    if !resolved {
                        return Err(ToolError::new(format!(
                            "Could not resolve {person:?} to a person — check the name/email."
                        )));
                    }
                    // olFolderCalendar = 9. Requires that person to have shared
                    // their calendar with you; otherwise COM errors with a
                    // permission message, surfaced as-is.
                    to_disp(call_method(
                        &ns,
                        "GetSharedDefaultFolder",
                        &mut [
                            VARIANT::from(recipient),
                            variant_from_i32(c::OL_FOLDER_CALENDAR),
                        ],
                    ).map_err(|e| ToolError::new(format!(
                        "Could not open {person:?}'s calendar — they may not have shared it with you. {}",
                        format_com_error(&e)
                    )))?)?
                }
                None => to_disp(call_method(
                    &ns,
                    "GetDefaultFolder",
                    &mut [variant_from_i32(c::OL_FOLDER_CALENDAR)],
                )?)?,
            };
            // Read once, up front, while `calendar` is still a genuine
            // `Folder` object (not a GetFirst/GetNext-returned occurrence
            // proxy) — see `event_summary`'s doc comment for why this is
            // needed instead of reading `StoreID` per enumerated item.
            let calendar_store_id = variant_to_string(&get_property(&calendar, "StoreID")?);
            let items = to_disp(get_property(&calendar, "Items")?)?;
            // Must precede Sort/Restrict — setting it afterwards has no effect.
            put_property(&items, "IncludeRecurrences", variant_from_bool(true))?;
            call_method(&items, "Sort", &mut [variant_from_str("[Start]")])?;
            // `Restrict` is still needed to bound the recurrence expansion; it
            // gets an unambiguous date-only window a little wider than
            // [start, end] (see `jet_dates`), and the exact range is applied
            // per occurrence in `enumerate_events_with_retry`.
            let date_filter = DateFilter::new(Some(start), Some(end))
                .expect("list_events always has both date bounds");
            let filters = date_filter.restrict_filters("Start", user_short_date_order().as_ref());
            if filters.is_empty() {
                return Err(ToolError::new(format!(
                    "Invalid date range {start} .. {end}: start_date and end_date must be between the years 1000 and 9999."
                )));
            }
            enumerate_events_with_retry(&items, &filters, &date_filter, &q, &calendar_store_id)
        })
    }

    fn get_event(&self, event_id: String) -> Result<EventDetail, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &event_id)?;
            let summary = event_summary(&item, None)?;
            let recurrence = recurrence_info(&item)?;
            let (body, body_truncated) = truncate(
                &variant_to_string(&get_property(&item, "Body").unwrap_or_default()),
                MAX_BODY_CHARS,
            );
            Ok(EventDetail {
                summary,
                body,
                body_truncated,
                recurrence,
            })
        })
    }

    fn create_event(&self, input: CreateEventInput) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (app, _ns) = mapi()?;
            let appt = to_disp(call_method(
                &app,
                "CreateItem",
                &mut [variant_from_i32(c::OL_APPOINTMENT_ITEM)],
            )?)?;
            put_property(&appt, "Subject", variant_from_str(&input.subject))?;
            put_property(&appt, "Start", variant_from_datetime(&parse_dt(&input.start, "start")?)?)?;
            put_property(&appt, "End", variant_from_datetime(&parse_dt(&input.end, "end")?)?)?;
            if input.all_day {
                put_property(&appt, "AllDayEvent", variant_from_bool(true))?;
            }
            if let Some(body) = input.body.as_deref().filter(|b| !b.is_empty()) {
                put_property(&appt, "Body", variant_from_str(body))?;
            }
            if let Some(location) = input.location.as_deref().filter(|l| !l.is_empty()) {
                put_property(&appt, "Location", variant_from_str(location))?;
            }
            if let Some(minutes) = input.reminder_minutes {
                put_property(&appt, "ReminderSet", variant_from_bool(true))?;
                put_property(&appt, "ReminderMinutesBeforeStart", variant_from_i32(minutes))?;
            }
            if let Some(categories) = input.categories.as_ref().filter(|c| !c.is_empty()) {
                set_item_categories(&appt, categories)?;
            }
            if let Some(show_as) = input.show_as.as_deref().filter(|s| !s.is_empty()) {
                let busy_status = crate::friendly::busy_status_to_id(show_as).ok_or_else(|| {
                    ToolError::new(format!(
                        "invalid show_as {show_as:?}: expected \"free\", \"tentative\", \"busy\", \"out_of_office\", or \"working_elsewhere\""
                    ))
                })?;
                put_property(&appt, "BusyStatus", variant_from_i32(busy_status))?;
            }
            if let Some(recurrence) = input.recurrence.as_ref() {
                apply_recurrence(&appt, recurrence)?;
            }
            let required = input.required_attendees.unwrap_or_default();
            let optional = input.optional_attendees.unwrap_or_default();
            let has_attendees = !required.is_empty() || !optional.is_empty();
            if has_attendees {
                put_property(&appt, "MeetingStatus", variant_from_i32(c::OL_MEETING))?;
                let recipients = to_disp(get_property(&appt, "Recipients")?)?;
                for address in &required {
                    add_meeting_recipient(&recipients, address, c::OL_RECIPIENT_REQUIRED)?;
                }
                for address in &optional {
                    add_meeting_recipient(&recipients, address, c::OL_RECIPIENT_OPTIONAL)?;
                }
                call_method(&recipients, "ResolveAll", &mut [])?;
                if input.send {
                    call_method(&appt, "Send", &mut [])?;
                } else {
                    call_method(&appt, "Save", &mut [])?;
                }
            } else {
                call_method(&appt, "Save", &mut [])?;
            }
            let status = create_event_status(has_attendees, input.send);
            Ok(json!({"status": status, "id": make_id(&appt)?, "subject": input.subject}))
        })
    }

    fn respond_to_meeting(
        &self,
        event_id: String,
        response: String,
        comment: Option<String>,
        send: bool,
    ) -> Result<Value, ToolError> {
        let response_key = response.trim().to_lowercase();
        let response_id = c::meeting_response_to_id(&response_key).ok_or_else(|| {
            ToolError::new(format!(
                "Invalid response {response:?}: use 'accept', 'decline' or 'tentative'."
            ))
        })?;
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let mut item = get_item(&ns, &event_id)?;
            // A meeting request from the inbox resolves to a MeetingItem; get
            // its appointment. Calendar ids resolve straight to appointments.
            if has_member(&item, "GetAssociatedAppointment") {
                item = to_disp(call_method(
                    &item,
                    "GetAssociatedAppointment",
                    &mut [variant_from_bool(true)],
                )?)?;
            }
            let resp = call_method(
                &item,
                "Respond",
                &mut [variant_from_i32(response_id), variant_from_bool(true)],
            )?;
            if let Ok(resp) = IDispatch::try_from(&resp) {
                if let Some(comment) = comment.as_deref().filter(|c| !c.is_empty()) {
                    put_property(&resp, "Body", variant_from_str(comment))?;
                }
                if send {
                    call_method(&resp, "Send", &mut [])?;
                } else {
                    call_method(&resp, "Save", &mut [])?;
                }
            }
            let subject = variant_to_string(&get_property(&item, "Subject")?);
            let status = format!("{response_key}{}", if send { "_sent" } else { "_saved" });
            Ok(json!({"status": status, "subject": subject}))
        })
    }

    fn update_event(&self, u: EventUpdate) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &u.event_id)?;
            let mut changed: Vec<&str> = Vec::new();

            if let Some(subject) = &u.subject {
                put_property(&item, "Subject", variant_from_str(subject))?;
                changed.push("subject");
            }
            if let Some(start) = &u.start {
                put_property(&item, "Start", variant_from_datetime(&parse_dt(start, "start")?)?)?;
                changed.push("start");
            }
            if let Some(end) = &u.end {
                put_property(&item, "End", variant_from_datetime(&parse_dt(end, "end")?)?)?;
                changed.push("end");
            }
            if let Some(location) = &u.location {
                put_property(&item, "Location", variant_from_str(location))?;
                changed.push("location");
            }
            if let Some(body) = &u.body {
                put_property(&item, "Body", variant_from_str(body))?;
                changed.push("body");
            }
            if let Some(all_day) = u.all_day {
                put_property(&item, "AllDayEvent", variant_from_bool(all_day))?;
                changed.push("all_day");
            }
            if let Some(minutes) = u.reminder_minutes {
                put_property(&item, "ReminderSet", variant_from_bool(true))?;
                put_property(&item, "ReminderMinutesBeforeStart", variant_from_i32(minutes))?;
                changed.push("reminder_minutes");
            }
            if let Some(show_as) = u.show_as.as_deref().filter(|s| !s.is_empty()) {
                let busy_status = crate::friendly::busy_status_to_id(show_as).ok_or_else(|| {
                    ToolError::new(format!(
                        "invalid show_as {show_as:?}: expected \"free\", \"tentative\", \"busy\", \"out_of_office\", or \"working_elsewhere\""
                    ))
                })?;
                put_property(&item, "BusyStatus", variant_from_i32(busy_status))?;
                changed.push("show_as");
            }
            if u.add_categories.is_some() || u.remove_categories.is_some() {
                let mut cats = get_item_categories(&item);
                if let Some(add) = &u.add_categories {
                    for a in add {
                        if !cats.iter().any(|c| c.eq_ignore_ascii_case(a)) {
                            cats.push(a.clone());
                        }
                    }
                    changed.push("add_categories");
                }
                if let Some(remove) = &u.remove_categories {
                    cats.retain(|c| !remove.iter().any(|r| r.eq_ignore_ascii_case(c)));
                    changed.push("remove_categories");
                }
                set_item_categories(&item, &cats)?;
            }

            // Adding either tier converts a personal appointment into a
            // meeting; MeetingStatus must be set before Recipients.Add for a
            // previously-non-meeting item.
            let adding_attendees = u.add_required_attendees.as_ref().is_some_and(|v| !v.is_empty())
                || u.add_optional_attendees.as_ref().is_some_and(|v| !v.is_empty());
            if adding_attendees {
                let current_status =
                    variant_to_i32(&get_property(&item, "MeetingStatus")?).unwrap_or(c::OL_NONMEETING);
                if current_status == c::OL_NONMEETING {
                    put_property(&item, "MeetingStatus", variant_from_i32(c::OL_MEETING))?;
                }
                let recipients = to_disp(get_property(&item, "Recipients")?)?;
                for address in u.add_required_attendees.as_deref().unwrap_or(&[]) {
                    add_meeting_recipient(&recipients, address, c::OL_RECIPIENT_REQUIRED)?;
                }
                for address in u.add_optional_attendees.as_deref().unwrap_or(&[]) {
                    add_meeting_recipient(&recipients, address, c::OL_RECIPIENT_OPTIONAL)?;
                }
                call_method(&recipients, "ResolveAll", &mut [])?;
                if u.add_required_attendees.is_some() { changed.push("add_required_attendees"); }
                if u.add_optional_attendees.is_some() { changed.push("add_optional_attendees"); }
            }
            if let Some(remove) = u.remove_attendees.as_ref().filter(|v| !v.is_empty()) {
                let recipients = to_disp(get_property(&item, "Recipients")?)?;
                remove_meeting_recipients(&recipients, remove)?;
                changed.push("remove_attendees");
            }

            validate_recurrence_update(&u)?;
            if let Some(recurrence) = u.recurrence.as_ref() {
                apply_recurrence(&item, recurrence)?;
                changed.push("recurrence");
            }
            if u.clear_recurrence {
                call_method(&item, "ClearRecurrencePattern", &mut [])?;
                changed.push("clear_recurrence");
            }

            // Save vs Send: only a meeting can notify attendees; a personal
            // appointment always just saves, regardless of send_update.
            let is_meeting =
                variant_to_i32(&get_property(&item, "MeetingStatus")?).unwrap_or(c::OL_NONMEETING)
                    != c::OL_NONMEETING;
            if is_meeting && u.send_update {
                call_method(&item, "Send", &mut [])?;
            } else {
                call_method(&item, "Save", &mut [])?;
            }

            Ok(json!({"status": "updated", "id": u.event_id, "changed": changed}))
        })
    }

    fn delete_event(&self, event_id: String, send_cancellation: bool) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &event_id)?;
            let subject = variant_to_string(&get_property(&item, "Subject")?);
            let meeting_status =
                variant_to_i32(&get_property(&item, "MeetingStatus")?).unwrap_or(c::OL_NONMEETING);
            let note = if meeting_status == c::OL_MEETING {
                // You organize this meeting: mark it canceled, optionally
                // notify attendees, then remove your own copy.
                put_property(&item, "MeetingStatus", variant_from_i32(c::OL_MEETING_CANCELED))?;
                if send_cancellation {
                    call_method(&item, "Send", &mut [])?;
                    "Meeting canceled; attendees notified. Moved to Deleted Items."
                } else {
                    "Meeting canceled without notifying attendees. Moved to Deleted Items."
                }
            } else {
                "Moved to Deleted Items."
            };
            call_method(&item, "Delete", &mut [])?;
            Ok(json!({"status": "deleted", "subject": subject, "note": note}))
        })
    }

    fn check_availability(&self, input: CheckAvailabilityInput) -> Result<AvailabilityResult, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let start = parse_dt(&input.start, "start")?;
            let end = parse_dt(&input.end, "end")?;
            if end <= start {
                return Err(ToolError::new(format!(
                    "check_availability: end ({}) must be after start ({})",
                    input.end, input.start
                )));
            }
            let interval = input.interval_minutes.max(1);
            // FreeBusy has no "end" parameter — it returns a string covering a
            // fixed range from `start`. Compute how many of its slots fall
            // within [start, end) and truncate to that.
            let total_minutes = (end - start).num_minutes().max(0);
            let max_slots = ((total_minutes + interval as i64 - 1) / interval as i64) as usize;

            let mut people = Vec::new();
            for person in &input.people {
                let recipient = to_disp(call_method(
                    &ns, "CreateRecipient", &mut [variant_from_str(person)],
                )?)?;
                let resolved = variant_to_bool(&call_method(&recipient, "Resolve", &mut [])?)
                    .unwrap_or(false);
                if !resolved {
                    people.push(PersonAvailability { person: person.clone(), resolved: false, slots: Vec::new() });
                    continue;
                }
                // `Resolve()` succeeds trivially for any syntactically valid SMTP
                // address — Outlook does no existence/deliverability check at that
                // point, only format/GAL-lookup. A made-up-but-well-formed address
                // (or a real address with no free/busy published) resolves fine but
                // then fails here, in `FreeBusy()` itself. Treat that failure the
                // same as an unresolved person — record it and move on — rather
                // than letting `?` abort the whole multi-person call over one bad
                // address.
                let raw = match call_method(
                    &recipient,
                    "FreeBusy",
                    &mut [
                        variant_from_datetime(&start)?,
                        variant_from_i32(interval),
                        variant_from_bool(true),
                    ],
                ) {
                    Ok(v) => variant_to_string(&v),
                    Err(_) => {
                        people.push(PersonAvailability { person: person.clone(), resolved: false, slots: Vec::new() });
                        continue;
                    }
                };
                let slots = parse_freebusy_slots(&raw, &start, interval, max_slots);
                people.push(PersonAvailability { person: person.clone(), resolved: true, slots });
            }
            let common = common_free(&people, &input.treat_as_free);
            Ok(AvailabilityResult { people, common_free: common })
        })
    }

    // ---- Attachments (Task 14) -----------------------------------------

    fn list_attachments(&self, email_id: String) -> Result<Vec<AttachmentInfo>, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &email_id)?;
            // `getattr(item, "Attachments", None)` then `if attachments:` — an item
            // with no `Attachments` collection yields an empty list, not a COM error.
            let attachments = match get_property(&item, "Attachments").ok().map(to_disp) {
                Some(Ok(a)) => a,
                _ => return Ok(Vec::new()),
            };
            let count = variant_to_i32(&get_property(&attachments, "Count")?).unwrap_or(0);
            if count == 0 {
                return Ok(Vec::new());
            }
            // Read the HTML body once (only when there are attachments) for
            // `is_inline`; a plain-text item or an unreadable body counts as "".
            let html_body = item_html_body(&item);
            let mut results = Vec::new();
            for i in 1..=count {
                // COM collections are 1-based.
                let att = to_disp(call_method(&attachments, "Item", &mut [variant_from_i32(i)])?)?;
                results.push(attachment_info(&att, i, &html_body)?);
            }
            Ok(results)
        })
    }

    fn save_attachments(
        &self,
        email_id: String,
        save_dir: String,
        attachment_names: Option<Vec<String>>,
    ) -> Result<Vec<Value>, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &email_id)?;
            // Python: `if not attachments or attachments.Count == 0`. Tolerate an item
            // that has no `Attachments` collection at all (missing property) the same
            // way as a present-but-empty collection: the clear "no attachments" error.
            let attachments = match get_property(&item, "Attachments").ok().map(to_disp) {
                Some(Ok(a)) => a,
                _ => return Err(ToolError::new("This email has no attachments.")),
            };
            let count = variant_to_i32(&get_property(&attachments, "Count")?).unwrap_or(0);
            if count == 0 {
                return Err(ToolError::new("This email has no attachments."));
            }
            let dir = resolve_save_dir(&save_dir);
            std::fs::create_dir_all(&dir).map_err(|e| {
                ToolError::new(format!(
                    "Could not create save directory {:?}: {e}",
                    dir.display()
                ))
            })?;
            // `{n.lower() for n in attachment_names}`: case-insensitive set membership.
            let wanted: Option<std::collections::HashSet<String>> = attachment_names
                .map(|names| names.iter().map(|n| n.to_lowercase()).collect());
            // Read the HTML body once (only when there are attachments) for
            // `is_inline`; a plain-text item or an unreadable body counts as "".
            let html_body = item_html_body(&item);
            let mut results = Vec::new();
            for i in 1..=count {
                let att = to_disp(call_method(&attachments, "Item", &mut [variant_from_i32(i)])?)?;
                let mut info = attachment_info(&att, i, &html_body)?;
                if info.filename.is_empty() {
                    info.filename = format!("attachment-{i}");
                }
                if let Some(wanted) = &wanted {
                    if !wanted.contains(&info.filename.to_lowercase()) {
                        continue;
                    }
                }
                let target = dir.join(safe_filename(&info.filename));
                let target_str = target.to_string_lossy().into_owned();
                // Each entry is the attachment's metadata (`index` stays its
                // original COM position even when filtering) plus the outcome.
                // A COM failure saving one file is collected per-file and does
                // NOT abort the batch (mirrors the per-file try/except in Python).
                let outcome = match call_method(&att, "SaveAsFile", &mut [variant_from_str(&target_str)]) {
                    Ok(_) => json!({"saved_to": target_str, "status": "saved"}),
                    Err(e) => json!({"status": "failed", "error": format_com_error(&e)}),
                };
                results.push(merge_json_objects(json!(info), outcome));
            }
            if results.is_empty() {
                return Err(ToolError::new(
                    "No attachments matched attachment_names; use list_attachments \
                     to see the exact file names.",
                ));
            }
            Ok(results)
        })
    }

    fn get_inline_image(&self, email_id: String, content_id: String, context_lines: Option<u32>)
        -> Result<InlineImageData, ToolError> {
        let wanted = normalize_cid_request(&content_id).ok_or_else(|| {
            ToolError::new("content_id must be a non-empty Content-ID (e.g. \"image001.png@01D9...\", optionally prefixed with cid:).")
        })?;
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &email_id)?;
            // As in list_attachments, an item without an `Attachments`
            // collection is treated as having none.
            let mut candidates: Vec<(IDispatch, AttachmentInfo)> = Vec::new();
            if let Some(Ok(attachments)) = get_property(&item, "Attachments").ok().map(to_disp) {
                let count = variant_to_i32(&get_property(&attachments, "Count")?).unwrap_or(0);
                for i in 1..=count {
                    let att = to_disp(call_method(&attachments, "Item", &mut [variant_from_i32(i)])?)?;
                    // `is_inline` isn't part of the result, so skip reading
                    // HTMLBody just to compute it.
                    let info = attachment_info(&att, i, "")?;
                    candidates.push((att, info));
                }
            }
            let content_ids: Vec<Option<String>> =
                candidates.iter().map(|(_, info)| info.content_id.clone()).collect();
            let (att, info) = &candidates[select_by_content_id(&wanted, &content_ids)?];

            // `Size` is the MAPI attachment size (a bit above the payload), so
            // it only pre-screens; the real byte count is checked after reading.
            if usize::try_from(info.size).unwrap_or(0) > MAX_INLINE_IMAGE_BYTES {
                return Err(inline_image_too_big(info));
            }
            let bytes = read_attachment_bytes(att, info)?;
            if bytes.len() > MAX_INLINE_IMAGE_BYTES {
                return Err(inline_image_too_big(info));
            }
            let mime_type = info
                .mime_type
                .clone()
                .unwrap_or_else(|| "application/octet-stream".to_string());
            // As in get_email, an item without `HTMLBody` reads as empty, so
            // the image just counts as unreferenced.
            let context = context_lines.map(|n| {
                let html = variant_to_string(&get_property(&item, "HTMLBody").unwrap_or_default());
                text_before_cid(&html, &wanted, n).unwrap_or_default()
            });
            Ok(InlineImageData {
                content_id: info.content_id.clone().unwrap_or_default(),
                filename: info.filename.clone(),
                data_uri: data_uri(&mime_type, &bytes),
                mime_type,
                size: bytes.len(),
                context,
            })
        })
    }

    // ---- Tasks (Task 15) -----------------------------------------------

    fn list_tasks(&self, q: TaskQuery) -> Result<Vec<TaskSummary>, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let tasks = to_disp(call_method(
                &ns,
                "GetDefaultFolder",
                &mut [variant_from_i32(c::OL_FOLDER_TASKS)],
            )?)?;
            let mut items = to_disp(get_property(&tasks, "Items")?)?;
            if !q.include_completed {
                items = to_disp(call_method(
                    &items,
                    "Restrict",
                    &mut [variant_from_str("[Complete] = False")],
                )?)?;
            }
            let count = variant_to_i32(&get_property(&items, "Count")?).unwrap_or(0);
            let mut results = Vec::new();
            for i in 1..=count {
                let item = to_disp(call_method(&items, "Item", &mut [variant_from_i32(i)])?)?;
                let summary = task_summary(&item)?;
                // Read the real body directly for query matching — `task_summary`
                // doesn't expose the body at all, so this is a second, deliberate
                // property read (same pattern `list_notes` already uses).
                let body = variant_to_string(&get_property(&item, "Body").unwrap_or_default());
                if task_matches(&body, &summary, &q) {
                    results.push(summary);
                }
            }
            Ok(results)
        })
    }

    fn create_task(
        &self,
        subject: String,
        body: Option<String>,
        due_date: Option<String>,
        importance: String,
        categories: Option<Vec<String>>,
        start_date: Option<String>,
        reminder_time: Option<String>,
    ) -> Result<Value, ToolError> {
        let importance_key = importance.trim().to_lowercase();
        let importance_id = c::importance_name_to_id(&importance_key).ok_or_else(|| {
            ToolError::new(format!(
                "Invalid importance {importance:?}: use 'low', 'normal' or 'high'."
            ))
        })?;
        self.with_com(|| {
            let (app, _ns) = mapi()?;
            let task = to_disp(call_method(
                &app,
                "CreateItem",
                &mut [variant_from_i32(c::OL_TASK_ITEM)],
            )?)?;
            put_property(&task, "Subject", variant_from_str(&subject))?;
            if let Some(body) = body.as_deref().filter(|b| !b.is_empty()) {
                put_property(&task, "Body", variant_from_str(body))?;
            }
            if let Some(due) = due_date.as_deref().filter(|d| !d.is_empty()) {
                put_property(
                    &task,
                    "DueDate",
                    variant_from_datetime(&parse_dt(due, "due_date")?)?,
                )?;
            }
            if let Some(start) = start_date.as_deref().filter(|d| !d.is_empty()) {
                put_property(
                    &task,
                    "StartDate",
                    variant_from_datetime(&parse_dt(start, "start_date")?)?,
                )?;
            }
            if let Some(reminder) = reminder_time.as_deref().filter(|d| !d.is_empty()) {
                put_property(&task, "ReminderSet", variant_from_bool(true))?;
                put_property(
                    &task,
                    "ReminderTime",
                    variant_from_datetime(&parse_dt(reminder, "reminder_time")?)?,
                )?;
            }
            put_property(&task, "Importance", variant_from_i32(importance_id))?;
            if let Some(cats) = categories.as_ref().filter(|c| !c.is_empty()) {
                set_item_categories(&task, cats)?;
            }
            call_method(&task, "Save", &mut [])?;
            Ok(json!({"status": "created", "id": make_id(&task)?, "subject": subject}))
        })
    }

    fn update_task(&self, u: TaskUpdate) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let task = get_item(&ns, &u.task_id)?;
            let mut changed: Vec<&str> = Vec::new();

            if let Some(subject) = &u.subject {
                put_property(&task, "Subject", variant_from_str(subject))?;
                changed.push("subject");
            }
            if let Some(body) = &u.body {
                put_property(&task, "Body", variant_from_str(body))?;
                changed.push("body");
            }
            if let Some(due) = &u.due_date {
                put_property(&task, "DueDate", variant_from_datetime(&parse_dt(due, "due_date")?)?)?;
                changed.push("due_date");
            }
            if let Some(start) = &u.start_date {
                put_property(&task, "StartDate", variant_from_datetime(&parse_dt(start, "start_date")?)?)?;
                changed.push("start_date");
            }
            if let Some(imp) = &u.importance {
                let id = c::importance_name_to_id(imp).ok_or_else(|| {
                    ToolError::new(format!(
                        "invalid importance {imp:?}: expected \"low\", \"normal\", or \"high\""
                    ))
                })?;
                put_property(&task, "Importance", variant_from_i32(id))?;
                changed.push("importance");
            }
            if u.add_categories.is_some() || u.remove_categories.is_some() {
                let mut cats = get_item_categories(&task);
                if let Some(add) = &u.add_categories {
                    for a in add {
                        if !cats.iter().any(|c| c.eq_ignore_ascii_case(a)) {
                            cats.push(a.clone());
                        }
                    }
                    changed.push("add_categories");
                }
                if let Some(remove) = &u.remove_categories {
                    cats.retain(|c| !remove.iter().any(|r| r.eq_ignore_ascii_case(c)));
                    changed.push("remove_categories");
                }
                set_item_categories(&task, &cats)?;
            }
            if let Some(pct) = u.percent_complete {
                put_property(&task, "PercentComplete", variant_from_i32(pct))?;
                changed.push("percent_complete");
            }
            if let Some(reminder) = &u.reminder_time {
                put_property(&task, "ReminderSet", variant_from_bool(true))?;
                put_property(&task, "ReminderTime", variant_from_datetime(&parse_dt(reminder, "reminder_time")?)?)?;
                changed.push("reminder_time");
            }
            // mark_complete last: MarkComplete() is Outlook's dedicated
            // "finish this task" method (it also sets PercentComplete=100
            // and Status=olTaskComplete), so apply any field edits above
            // to the task's live state first, then finish/reopen it.
            if let Some(complete) = u.mark_complete {
                if complete {
                    call_method(&task, "MarkComplete", &mut [])?;
                } else {
                    put_property(&task, "Complete", variant_from_bool(false))?;
                    put_property(&task, "Status", variant_from_i32(c::OL_TASK_NOT_STARTED))?;
                    put_property(&task, "PercentComplete", variant_from_i32(0))?;
                }
                changed.push("mark_complete");
            }

            call_method(&task, "Save", &mut [])?;
            Ok(json!({"status": "updated", "id": u.task_id, "changed": changed}))
        })
    }

    fn delete_task(&self, task_id: String) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &task_id)?;
            let subject = variant_to_string(&get_property(&item, "Subject")?);
            call_method(&item, "Delete", &mut [])?;
            Ok(json!({"status": "deleted", "subject": subject, "note": "Moved to Deleted Items."}))
        })
    }

    // ---- Notes (Task 16) -----------------------------------------------

    fn list_notes(&self, q: NoteQuery) -> Result<Vec<NoteSummary>, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let notes = to_disp(call_method(
                &ns,
                "GetDefaultFolder",
                &mut [variant_from_i32(c::OL_FOLDER_NOTES)],
            )?)?;
            let items = to_disp(get_property(&notes, "Items")?)?;
            let count = variant_to_i32(&get_property(&items, "Count")?).unwrap_or(0);
            let mut results = Vec::new();
            for i in 1..=count {
                let item = to_disp(call_method(&items, "Item", &mut [variant_from_i32(i)])?)?;
                let summary = note_summary(&item)?;
                // Read the real body directly for query matching — `note_summary`
                // only exposes the derived (120-char-truncated) subject, not the
                // full body, so this is a second, deliberate property read (same
                // pattern `get_note` already uses: it re-reads `Body` outside
                // `note_summary` too, for its own untruncated-body purpose).
                let body = variant_to_string(&get_property(&item, "Body").unwrap_or_default());
                if note_matches(&body, &summary, &q) {
                    results.push(summary);
                }
            }
            Ok(results)
        })
    }

    fn get_note(&self, note_id: String) -> Result<NoteDetail, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let note = get_item(&ns, &note_id)?;
            let summary = note_summary(&note)?;
            let (body, body_truncated) =
                truncate(&variant_to_string(&get_property(&note, "Body")?), MAX_BODY_CHARS);
            Ok(NoteDetail {
                summary,
                body,
                body_truncated,
                modified: variant_to_iso_string(&get_property(&note, "LastModificationTime").unwrap_or_default()),
            })
        })
    }

    fn create_note(&self, body: String, categories: Option<Vec<String>>, color: Option<String>) -> Result<Value, ToolError> {
        // Validate before touching COM (fail-fast, like `create_task`).
        if body.is_empty() {
            return Err(ToolError::new("create_note requires a non-empty body."));
        }
        let color_id = color.as_deref().map(|c| {
            c::note_color_to_id(c).ok_or_else(|| {
                ToolError::new(format!(
                    "invalid color {c:?}: expected \"blue\", \"green\", \"pink\", \"yellow\", or \"white\""
                ))
            })
        }).transpose()?;
        self.with_com(|| {
            let (app, _ns) = mapi()?;
            let note = to_disp(call_method(
                &app,
                "CreateItem",
                &mut [variant_from_i32(c::OL_NOTE_ITEM)],
            )?)?;
            put_property(&note, "Body", variant_from_str(&body))?;
            if let Some(id) = color_id {
                put_property(&note, "Color", variant_from_i32(id))?;
            }
            call_method(&note, "Save", &mut [])?;
            if let Some(cats) = categories.as_ref().filter(|c| !c.is_empty()) {
                set_item_categories(&note, cats)?;
                call_method(&note, "Save", &mut [])?;
            }
            Ok(json!({"status": "created", "id": make_id(&note)?}))
        })
    }

    fn update_note(&self, u: NoteUpdate) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let note = get_item(&ns, &u.note_id)?;
            let mut changed: Vec<&str> = Vec::new();

            if let Some(body) = &u.body {
                put_property(&note, "Body", variant_from_str(body))?;
                changed.push("body");
            }
            if u.add_categories.is_some() || u.remove_categories.is_some() {
                let mut cats = get_item_categories(&note);
                if let Some(add) = &u.add_categories {
                    for a in add {
                        if !cats.iter().any(|c| c.eq_ignore_ascii_case(a)) {
                            cats.push(a.clone());
                        }
                    }
                    changed.push("add_categories");
                }
                if let Some(remove) = &u.remove_categories {
                    cats.retain(|c| !remove.iter().any(|r| r.eq_ignore_ascii_case(c)));
                    changed.push("remove_categories");
                }
                set_item_categories(&note, &cats)?;
            }
            if let Some(color) = &u.color {
                let id = c::note_color_to_id(color).ok_or_else(|| {
                    ToolError::new(format!(
                        "invalid color {color:?}: expected \"blue\", \"green\", \"pink\", \"yellow\", or \"white\""
                    ))
                })?;
                put_property(&note, "Color", variant_from_i32(id))?;
                changed.push("color");
            }

            call_method(&note, "Save", &mut [])?;
            Ok(json!({"status": "updated", "id": u.note_id, "changed": changed}))
        })
    }

    fn delete_note(&self, note_id: String) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &note_id)?;
            call_method(&item, "Delete", &mut [])?;
            Ok(json!({"status": "deleted", "note": "Moved to Deleted Items."}))
        })
    }
}

#[cfg(test)]
mod event_filter_tests {
    use super::*;

    /// Create a representative base EventSummary for testing.
    fn base() -> EventSummary {
        EventSummary {
            id: "test-id|store-id".to_string(),
            subject: "Weekly Review".to_string(),
            start: Some("2026-06-10T14:00:00".to_string()),
            end: Some("2026-06-10T15:00:00".to_string()),
            location: "Room A".to_string(),
            organizer: "Alice Smith; alice@example.com".to_string(),
            all_day: false,
            is_recurring: false,
            is_meeting: true,
            categories: vec!["Work".to_string()],
            show_as: "busy".to_string(),
            my_response: "accepted".to_string(),
            required_attendees: "Alice Smith; alice@example.com".to_string(),
            optional_attendees: "Bob Jones; bob@example.com".to_string(),
        }
    }

    #[test]
    fn empty_query_matches_any_summary() {
        let summary = base();
        let query = EventQuery::default();
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn query_substring_matches_subject() {
        let summary = base();
        let query = EventQuery {
            query: Some("weekly".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn query_substring_matches_location() {
        let summary = base();
        let query = EventQuery {
            query: Some("room".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn query_substring_no_match() {
        let summary = base();
        let query = EventQuery {
            query: Some("nonexistent".to_string()),
            ..Default::default()
        };
        assert!(!event_matches(&summary, &query));
    }

    #[test]
    fn query_substring_case_insensitive() {
        let summary = base();
        let query = EventQuery {
            query: Some("WEEKLY REVIEW".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn empty_query_string_is_noop() {
        let summary = base();
        let query = EventQuery {
            query: Some("".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn category_present_case_insensitive() {
        let summary = base();
        let query = EventQuery {
            category: Some("work".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn category_absent() {
        let summary = base();
        let query = EventQuery {
            category: Some("Personal".to_string()),
            ..Default::default()
        };
        assert!(!event_matches(&summary, &query));
    }

    #[test]
    fn category_multiple_matches_one() {
        let mut summary = base();
        summary.categories = vec!["Work".to_string(), "Meeting".to_string()];
        let query = EventQuery {
            category: Some("MEETING".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn show_as_match_case_insensitive() {
        let summary = base();
        let query = EventQuery {
            show_as: Some("BUSY".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn show_as_mismatch() {
        let summary = base();
        let query = EventQuery {
            show_as: Some("free".to_string()),
            ..Default::default()
        };
        assert!(!event_matches(&summary, &query));
    }

    #[test]
    fn my_response_match_case_insensitive() {
        let summary = base();
        let query = EventQuery {
            my_response: Some("ACCEPTED".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn my_response_mismatch() {
        let summary = base();
        let query = EventQuery {
            my_response: Some("declined".to_string()),
            ..Default::default()
        };
        assert!(!event_matches(&summary, &query));
    }

    #[test]
    fn meetings_only_true_with_meeting() {
        let summary = base();
        let query = EventQuery {
            meetings_only: true,
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn meetings_only_true_without_meeting() {
        let mut summary = base();
        summary.is_meeting = false;
        let query = EventQuery {
            meetings_only: true,
            ..Default::default()
        };
        assert!(!event_matches(&summary, &query));
    }

    #[test]
    fn all_day_true_matches_all_day_event() {
        let mut summary = base();
        summary.all_day = true;
        let query = EventQuery {
            all_day: Some(true),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn all_day_true_rejects_non_all_day_event() {
        let summary = base();
        let query = EventQuery {
            all_day: Some(true),
            ..Default::default()
        };
        assert!(!event_matches(&summary, &query));
    }

    #[test]
    fn all_day_false_matches_non_all_day_event() {
        let summary = base();
        let query = EventQuery {
            all_day: Some(false),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn all_day_false_rejects_all_day_event() {
        let mut summary = base();
        summary.all_day = true;
        let query = EventQuery {
            all_day: Some(false),
            ..Default::default()
        };
        assert!(!event_matches(&summary, &query));
    }

    #[test]
    fn all_day_none_is_noop() {
        let summary = base();
        let query = EventQuery {
            all_day: None,
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn attendees_required_role_substring_match() {
        let summary = base();
        let query = EventQuery {
            attendees: Some(vec!["alice".to_string()]),
            attendee_role: Some("required".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn attendees_required_role_no_match() {
        let summary = base();
        let query = EventQuery {
            attendees: Some(vec!["bob".to_string()]),
            attendee_role: Some("required".to_string()),
            ..Default::default()
        };
        assert!(!event_matches(&summary, &query));
    }

    #[test]
    fn attendees_optional_role_substring_match() {
        let summary = base();
        let query = EventQuery {
            attendees: Some(vec!["bob".to_string()]),
            attendee_role: Some("optional".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn attendees_optional_role_no_match() {
        let summary = base();
        let query = EventQuery {
            attendees: Some(vec!["alice".to_string()]),
            attendee_role: Some("optional".to_string()),
            ..Default::default()
        };
        assert!(!event_matches(&summary, &query));
    }

    #[test]
    fn attendees_any_role_matches_required() {
        let summary = base();
        let query = EventQuery {
            attendees: Some(vec!["alice".to_string()]),
            attendee_role: Some("any".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn attendees_any_role_matches_optional() {
        let summary = base();
        let query = EventQuery {
            attendees: Some(vec!["bob".to_string()]),
            attendee_role: Some("any".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn attendees_no_role_defaults_to_any() {
        let summary = base();
        let query = EventQuery {
            attendees: Some(vec!["alice".to_string()]),
            attendee_role: None,
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn attendees_case_insensitive_search() {
        let summary = base();
        let query = EventQuery {
            attendees: Some(vec!["ALICE".to_string()]),
            attendee_role: Some("required".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn attendees_multiple_in_list_any_matches() {
        let summary = base();
        let query = EventQuery {
            attendees: Some(vec!["nonexistent".to_string(), "bob".to_string()]),
            attendee_role: Some("optional".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn attendees_empty_string_does_not_match() {
        let summary = base();
        let query = EventQuery {
            attendees: Some(vec!["".to_string()]),
            attendee_role: None,
            ..Default::default()
        };
        assert!(!event_matches(&summary, &query));
    }

    #[test]
    fn attendees_empty_list_is_noop() {
        let summary = base();
        let query = EventQuery {
            attendees: Some(vec![]),
            attendee_role: None,
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn multiple_filters_all_satisfied() {
        let summary = base();
        let query = EventQuery {
            query: Some("weekly".to_string()),
            category: Some("work".to_string()),
            show_as: Some("busy".to_string()),
            meetings_only: true,
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn multiple_filters_one_fails() {
        let summary = base();
        let query = EventQuery {
            query: Some("weekly".to_string()),
            category: Some("work".to_string()),
            show_as: Some("free".to_string()),
            meetings_only: true,
            ..Default::default()
        };
        assert!(!event_matches(&summary, &query));
    }

    #[test]
    fn complex_scenario_meeting_with_required_attendee_and_category() {
        let summary = base();
        let query = EventQuery {
            meetings_only: true,
            category: Some("Work".to_string()),
            attendees: Some(vec!["alice".to_string()]),
            attendee_role: Some("required".to_string()),
            query: Some("review".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn complex_scenario_wrong_attendee_tier() {
        let summary = base();
        let query = EventQuery {
            meetings_only: true,
            category: Some("Work".to_string()),
            attendees: Some(vec!["bob".to_string()]),
            attendee_role: Some("required".to_string()),
            query: Some("review".to_string()),
            ..Default::default()
        };
        assert!(!event_matches(&summary, &query));
    }

    #[test]
    fn attendees_full_email_substring_match() {
        let summary = base();
        let query = EventQuery {
            attendees: Some(vec!["example.com".to_string()]),
            attendee_role: Some("required".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn query_location_exact_match_case_insensitive() {
        let summary = base();
        let query = EventQuery {
            query: Some("ROOM A".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }
}

#[cfg(test)]
mod task_filter_tests {
    use super::*;

    fn base() -> TaskSummary {
        TaskSummary {
            id: "test-id|store-id".to_string(),
            subject: "Quarterly Report".to_string(),
            due_date: Some("2026-06-10T00:00:00".to_string()),
            complete: false,
            status: "not_started".to_string(),
            importance: "normal".to_string(),
            categories: vec!["Work".to_string()],
        }
    }

    #[test]
    fn empty_query_matches_any_summary() {
        let summary = base();
        let query = TaskQuery::default();
        assert!(task_matches("some body text", &summary, &query));
    }

    #[test]
    fn query_substring_matches_subject() {
        let summary = base();
        let query = TaskQuery { query: Some("quarterly".to_string()), ..Default::default() };
        assert!(task_matches("unrelated body", &summary, &query));
    }

    #[test]
    fn query_substring_matches_body() {
        let summary = base();
        let query = TaskQuery { query: Some("budget numbers".to_string()), ..Default::default() };
        assert!(task_matches("here are the budget numbers for review", &summary, &query));
    }

    #[test]
    fn query_substring_no_match_in_subject_or_body() {
        let summary = base();
        let query = TaskQuery { query: Some("nonexistent".to_string()), ..Default::default() };
        assert!(!task_matches("also nothing here", &summary, &query));
    }

    #[test]
    fn query_substring_case_insensitive_against_body() {
        let summary = base();
        let query = TaskQuery { query: Some("BUDGET NUMBERS".to_string()), ..Default::default() };
        assert!(task_matches("here are the budget numbers", &summary, &query));
    }

    #[test]
    fn empty_query_string_is_noop() {
        let summary = base();
        let query = TaskQuery { query: Some("".to_string()), ..Default::default() };
        assert!(task_matches("anything", &summary, &query));
    }

    #[test]
    fn category_filter_still_applies_alongside_body_query() {
        let summary = base();
        let query = TaskQuery {
            query: Some("budget".to_string()),
            category: Some("Personal".to_string()),
            ..Default::default()
        };
        assert!(!task_matches("budget numbers", &summary, &query));
    }

    #[test]
    fn importance_filter_still_applies_alongside_body_query() {
        let summary = base();
        let query = TaskQuery {
            query: Some("budget".to_string()),
            importance: Some("high".to_string()),
            ..Default::default()
        };
        assert!(!task_matches("budget numbers", &summary, &query));
    }
}

#[cfg(test)]
mod attachment_tests {
    use super::*;

    fn info() -> AttachmentInfo {
        AttachmentInfo {
            index: 3,
            filename: "logo.png".to_string(),
            size: 42,
            att_type: "file".to_string(),
            content_id: Some("logo@x".to_string()),
            mime_type: Some("image/png".to_string()),
            hidden: true,
            is_inline: true,
        }
    }

    #[test]
    fn attachment_info_serializes_type_key() {
        let v = json!(info());
        assert_eq!(v["type"], "file");
        assert!(v.get("att_type").is_none());
        assert_eq!(v["content_id"], "logo@x");
        assert_eq!(v["mime_type"], "image/png");
        assert_eq!(v["hidden"], true);
        assert_eq!(v["is_inline"], true);
    }

    #[test]
    fn save_entry_is_info_plus_outcome() {
        let v = merge_json_objects(json!(info()), json!({"saved_to": "C:/x/logo.png", "status": "saved"}));
        assert_eq!(v["index"], 3);
        assert_eq!(v["filename"], "logo.png");
        assert_eq!(v["type"], "file");
        assert_eq!(v["saved_to"], "C:/x/logo.png");
        assert_eq!(v["status"], "saved");
    }

    #[test]
    fn merge_ignores_non_object_extra() {
        let v = merge_json_objects(json!({"a": 1}), json!("nope"));
        assert_eq!(v, json!({"a": 1}));
    }

    fn cids(ids: &[Option<&str>]) -> Vec<Option<String>> {
        ids.iter().map(|c| c.map(str::to_string)).collect()
    }

    #[test]
    fn select_by_content_id_matches_case_insensitively() {
        let ids = cids(&[None, Some("logo@01D9"), Some("Image001.PNG@01D9ABCD")]);
        assert_eq!(select_by_content_id("image001.png@01d9abcd", &ids).unwrap(), 2);
        assert_eq!(select_by_content_id("LOGO@01d9", &ids).unwrap(), 1);
    }

    #[test]
    fn select_by_content_id_lists_available_ids_when_missing() {
        let ids = cids(&[Some("a@x"), None, Some("b@y")]);
        let msg = select_by_content_id("missing@z", &ids).unwrap_err().0;
        assert!(msg.contains("missing@z"), "{msg}");
        assert!(msg.contains("Available Content-IDs: a@x, b@y"), "{msg}");
    }

    #[test]
    fn select_by_content_id_points_to_list_attachments_when_none() {
        for ids in [cids(&[]), cids(&[None, None])] {
            let msg = select_by_content_id("x@y", &ids).unwrap_err().0;
            assert!(msg.contains("no attachments with a Content-ID"), "{msg}");
            assert!(msg.contains("list_attachments"), "{msg}");
        }
    }

    #[test]
    fn inline_image_too_big_points_to_save_attachments() {
        let msg = inline_image_too_big(&info()).0;
        assert!(msg.contains("logo.png"), "{msg}");
        assert!(msg.contains("10 MB"), "{msg}");
        assert!(msg.contains("save_attachments"), "{msg}");
        let unnamed = AttachmentInfo { filename: String::new(), ..info() };
        assert!(inline_image_too_big(&unnamed).0.contains("logo@x"));
    }

    #[test]
    fn data_uri_base64_encodes_with_the_mime_header() {
        assert_eq!(data_uri("image/png", b"\x89PNG"), "data:image/png;base64,iVBORw==");
        assert_eq!(data_uri("application/octet-stream", b""), "data:application/octet-stream;base64,");
    }

    #[test]
    fn temp_dir_guard_removes_its_directory_on_drop() {
        let guard = TempDirGuard::create().unwrap();
        let dir = guard.path().to_path_buf();
        std::fs::write(dir.join("f.bin"), b"abc").unwrap();
        let other = TempDirGuard::create().unwrap();
        assert_ne!(other.path(), dir.as_path());
        drop(guard);
        assert!(!dir.exists());
    }
}

#[cfg(test)]
mod text_match_tests {
    use super::*;

    #[test]
    fn hebrew_matches_subject_or_later_field() {
        assert!(text_matches("מייל שיקוף", &["Re: מייל שיקוף שבועי"]));
        assert!(text_matches("סיכום עשייה", &["", "Dana", "גוף: סיכום עשייה Q3"]));
        assert!(!text_matches("מייל שיקוף", &["Weekly report", "Dana"]));
    }

    #[test]
    fn mixed_hebrew_and_english() {
        assert!(text_matches("Q3 סיכום", &["Weekly Q3 סיכום עשייה"]));
        assert!(text_matches("q3 סיכום", &["Weekly Q3 סיכום עשייה"]));
        assert!(!text_matches("Q4 סיכום", &["Weekly Q3 סיכום עשייה"]));
    }

    #[test]
    fn matching_is_caseless() {
        assert!(text_matches("weekly", &["WEEKLY Report"]));
        assert!(text_matches("ÉCOLE", &["notes from école today"]));
        assert!(text_matches("ΣΟΦΊΑ", &["σοφία"]));
    }

    #[test]
    fn matching_ignores_normalization_form() {
        // Latin with an accent: composed (NFC) vs decomposed (NFD).
        let nfc: String = "café".nfc().collect();
        let nfd: String = "café".nfd().collect();
        assert_ne!(nfc, nfd);
        assert!(text_matches(&nfd, &[&format!("subject {nfc}")]));
        assert!(text_matches(&nfc, &[&format!("subject {nfd}")]));
        // Hebrew niqqud has no precomposed forms, but the same marks typed
        // in a different order (shin dot + qamats vs qamats + shin dot) are
        // canonically equivalent and must still match.
        let a = "\u{05E9}\u{05C1}\u{05B8}לום";
        let b = "\u{05E9}\u{05B8}\u{05C1}לום";
        assert_ne!(a, b);
        assert!(text_matches(a, &[&format!("subject {b}")]));
        assert!(text_matches(b, &[&format!("subject {a}")]));
    }

    #[test]
    fn empty_fields_never_match() {
        assert!(!text_matches("שלום", &[]));
        assert!(!text_matches("שלום", &["", ""]));
    }
}

#[cfg(test)]
mod recipient_filter_tests {
    use super::recipient_matches;

    fn strs(v: &[&str]) -> Vec<String> {
        v.iter().map(|s| s.to_string()).collect()
    }

    #[test]
    fn matches_name_substring_caselessly() {
        assert!(recipient_matches("lovelace", &strs(&["Ada Lovelace", "/o=ExchangeLabs/cn=ada"])));
        assert!(recipient_matches("ADA", &strs(&["Ada Lovelace"])));
    }

    #[test]
    fn matches_smtp_address_when_name_differs() {
        let c = strs(&["Ada Lovelace", "/o=ExchangeLabs/cn=ada", "Ada.Lovelace@Example.com"]);
        assert!(recipient_matches("ada.lovelace@example.com", &c));
        assert!(recipient_matches("@example.com", &c));
    }

    #[test]
    fn rejects_when_no_candidate_contains_needle() {
        let c = strs(&["nobody@example.invalid"]);
        assert!(!recipient_matches("someone-else@example.invalid", &c));
        assert!(!recipient_matches("ada", &[]));
    }

    #[test]
    fn empty_needle_matches_anything_with_a_candidate() {
        assert!(recipient_matches("", &strs(&["x"])));
    }
}

#[cfg(test)]
mod body_limit_tests {
    use super::*;

    #[test]
    fn truncate_leaves_short_text_untouched() {
        assert_eq!(truncate("hello", 10), ("hello".to_string(), false));
        // Exactly at the limit is not truncated.
        assert_eq!(truncate("hello", 5), ("hello".to_string(), false));
        assert_eq!(truncate("", 5), (String::new(), false));
    }

    #[test]
    fn truncate_cuts_long_text_and_flags_it() {
        let (text, truncated) = truncate("abcdefghij", 4);
        assert!(truncated);
        assert_eq!(text, "abcd\n\n[... truncated at 4 characters]");
    }

    #[test]
    fn truncate_counts_chars_not_bytes() {
        // Each "é" is 2 bytes; a byte-based cut would split a codepoint.
        let (text, truncated) = truncate("ééééé", 3);
        assert!(truncated);
        assert!(text.starts_with("ééé\n\n"));
        assert_eq!(truncate("ééé", 3), ("ééé".to_string(), false));
    }

    #[test]
    fn clamp_body_limit_defaults_and_clamps() {
        assert_eq!(clamp_body_limit(None), MAX_BODY_CHARS);
        assert_eq!(clamp_body_limit(Some(0)), MIN_BODY_CHARS_LIMIT);
        assert_eq!(clamp_body_limit(Some(999)), MIN_BODY_CHARS_LIMIT);
        assert_eq!(clamp_body_limit(Some(1_000)), 1_000);
        assert_eq!(clamp_body_limit(Some(250_000)), 250_000);
        assert_eq!(clamp_body_limit(Some(5_000_000)), MAX_BODY_CHARS_LIMIT);
        assert_eq!(clamp_body_limit(Some(u32::MAX)), MAX_BODY_CHARS_LIMIT);
    }
}
