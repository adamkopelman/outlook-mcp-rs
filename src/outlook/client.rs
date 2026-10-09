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
    self,
    call_method, clean_content_id, create_com_object, format_com_error, get_item_categories,
    get_mapi_prop, get_property, guess_mime, has_member, is_inline, jet_datetime, make_item_id, normalize_cid_request, parse_item_id, put_property, safe_filename,
    set_item_categories, variant_from_bool, variant_from_datetime, variant_from_i32, variant_from_str,
    variant_to_bool, variant_to_i32, variant_to_iso_string, variant_to_string, ComGuard,
};
use crate::outlook::types::*;
use chrono::Datelike;
use crate::outlook::filters::{self, DateRange};
use crate::outlook::text_query::TextQuery;
use crate::outlook::{MAX_EMAIL_COUNT, MAX_EVENT_COUNT, MAX_NOTE_COUNT, MAX_TASK_COUNT};
use crate::outlook::{
    com_recurrence_interval, common_free, create_event_status, friendly_recurrence_interval,
    parse_freebusy_slots, take_page, validate_recurrence, validate_recurrence_update,
    CheckAvailabilityInput, permanent_delete_needs_move, require_empty_confirm, CreateEventInput,
    EmailQuery, EmailUpdate, EventQuery, EventUpdate, NoteQuery, NoteUpdate, OutlookClient,
    RecurrenceInput, TaskQuery, TaskUpdate, text_before_cid, draft_update_changes,
    validate_draft_update, DraftUpdate, InlineImageSource, ValidatedInlineImage, MailBody,
    NewEmail, ReplyInput, merge_categories, parse_importance, prepare_html_images,
    prepare_mail_body,
};
use crate::outlook::read::{
    cid_references, image_extension, outlook_date, output_file_path, prepare_output_dir,
    replace_cid_references, resolve_save_dir, shape_field, write_output_file, BatchResult, BodyOut,
    CidRewrite, ReadOptions, ReadTool,
};

/// How many items (newest first, after every other filter) `list_emails`
/// will open one by one when the `to` filter falls back to scanning each
/// item's `Recipients` collection. Bounds the cost of a per-item COM walk.
const RECIPIENT_SCAN_LIMIT: i32 = 2000;
/// MAPI `PR_SMTP_ADDRESS` (Unicode), read through `Recipient.PropertyAccessor`.
const PR_SMTP_ADDRESS: &str = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F";
/// Largest attachment `get_inline_image` returns inline (base64 in the tool
/// result); anything bigger belongs on disk via `save_attachments`.
const MAX_INLINE_IMAGE_BYTES: usize = 10 * 1024 * 1024;
/// Upper bound on items scanned client-side when the DASL text search for a
/// non-ASCII query comes back empty (see `list_emails`).
const MAX_SCAN_ITEMS: i32 = 2000;

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

/// Parses a user-supplied date parameter with the shared grammar in
/// [`crate::outlook::dates`] (ISO date/datetime, keywords like `today` or
/// `start_of_week`, offsets like `-14d`), relative to the current local time
/// and with weeks starting on the Windows user's first day of the week.
/// Every single date parameter goes through here; `*_after`/`*_before` ranges
/// go through `DateRange::parse` with the same `com::user_first_day_of_week`.
fn parse_dt(value: &str, field: &str) -> Result<chrono::NaiveDateTime, ToolError> {
    Ok(crate::outlook::dates::parse_date_param(value, field, com::user_first_day_of_week())?.at)
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

/// One `list_emails` query field of `item`, for the client-side query
/// fallback (see [`filters::EMAIL_QUERY_FIELDS`]). Read lazily per field, so
/// the body is only fetched when the subject/sender didn't already match.
fn email_query_field(item: &IDispatch, field: &str) -> String {
    let prop = |name: &str| variant_to_string(&get_property(item, name).unwrap_or_default());
    match field {
        "subject" => prop("Subject"),
        "from" => format!("{} {}", prop("SenderName"), prop("SenderEmailAddress")),
        "to" => format!("{}; {}; {}", prop("To"), prop("CC"), recipient_strings(item).join("; ")),
        "body" => prop("Body"),
        _ => String::new(),
    }
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

/// Runs the `Restrict` + `GetFirst`/`GetNext` enumeration sequence for
/// `list_events`, keeping events that pass [`filters::event_matches`] and
/// returning the page `offset..offset + count` of them (`count` already
/// clamped to `MAX_EVENT_COUNT`).
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
    flt: &str,
    q: &EventQuery,
    text: &TextQuery,
    calendar_store_id: &str,
    count: usize,
) -> Result<Vec<EventSummary>, ToolError> {
    const MAX_ATTEMPTS: u32 = 3;
    let offset = q.offset.max(0) as usize;
    for attempt in 1..=MAX_ATTEMPTS {
        let outcome = (|| -> Result<Vec<EventSummary>, ToolError> {
            let restricted =
                to_disp(call_method(items, "Restrict", &mut [variant_from_str(flt)])?)?;
            // Enumerate with GetFirst/GetNext (not Count/Item): under
            // IncludeRecurrences the collection can expand without bound, so
            // we must stream it and stop once the page is full.
            let mut skipped = 0;
            let mut results = Vec::new();
            let mut current = call_method(&restricted, "GetFirst", &mut [])?;
            while let Ok(item) = IDispatch::try_from(&current) {
                let summary = event_summary(&item, Some(calendar_store_id))?;
                if filters::event_matches(&summary, q, text) {
                    if skipped < offset {
                        skipped += 1;
                    } else {
                        results.push(summary);
                        if results.len() >= count {
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
    body: &MailBody,
    cc: Option<&[String]>,
    bcc: Option<&[String]>,
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
    set_mail_body(&mail, body)?;
    Ok(mail)
}

/// Replace a mail item's body. `BodyFormat` is set first so Outlook doesn't
/// re-convert the new body.
fn set_mail_body(mail: &IDispatch, body: &MailBody) -> Result<(), ToolError> {
    match body {
        MailBody::Html(html) => {
            put_property(mail, "BodyFormat", variant_from_i32(c::OL_FORMAT_HTML))?;
            put_property(mail, "HTMLBody", variant_from_str(html))?;
        }
        MailBody::Text(text) => {
            put_property(mail, "BodyFormat", variant_from_i32(c::OL_FORMAT_PLAIN))?;
            put_property(mail, "Body", variant_from_str(text))?;
        }
    }
    Ok(())
}

/// Set categories (when given and non-empty) and importance (an already
/// validated `OlImportance` id) on a new mail item.
fn set_mail_metadata(mail: &IDispatch, categories: Option<&[String]>, importance: Option<i32>)
    -> Result<(), ToolError> {
    if let Some(cats) = categories.filter(|c| !c.is_empty()) {
        set_item_categories(mail, cats)?;
    }
    if let Some(id) = importance {
        put_property(mail, "Importance", variant_from_i32(id))?;
    }
    Ok(())
}

/// Validate a [`NewEmail`] before any COM call: importance, attachment
/// paths, inline images and the body's `data:` URIs. Returns the final body,
/// all inline images to attach and the importance id.
fn prepare_new_email(m: &NewEmail)
    -> Result<(MailBody, Vec<ValidatedInlineImage>, Option<i32>), ToolError> {
    let importance = parse_importance(m.importance.as_deref())?;
    check_attachment_paths(m.attachments.as_deref().unwrap_or(&[]))?;
    let (body, images) = prepare_mail_body(&m.body, m.inline_images.as_deref().unwrap_or(&[]))?;
    Ok((body, images, importance))
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

/// Every attachment of `item` with its metadata (`is_inline` is not
/// computed, so the HTML body isn't read for it). An item without an
/// `Attachments` collection is treated as having none.
fn attachment_candidates(item: &IDispatch) -> Result<Vec<(IDispatch, AttachmentInfo)>, ToolError> {
    let mut candidates = Vec::new();
    if let Some(Ok(attachments)) = get_property(item, "Attachments").ok().map(to_disp) {
        let count = variant_to_i32(&get_property(&attachments, "Count")?).unwrap_or(0);
        for i in 1..=count {
            let att = to_disp(call_method(&attachments, "Item", &mut [variant_from_i32(i)])?)?;
            let info = attachment_info(&att, i, "")?;
            candidates.push((att, info));
        }
    }
    Ok(candidates)
}

/// An attachment's bytes for `get_inline_image` / `resolve_inline_images`,
/// refused over `MAX_INLINE_IMAGE_BYTES`. `Size` is the MAPI attachment size
/// (a bit above the payload), so it only pre-screens; the real byte count is
/// checked after reading.
fn read_inline_image_bytes(att: &IDispatch, info: &AttachmentInfo) -> Result<Vec<u8>, ToolError> {
    if usize::try_from(info.size).unwrap_or(0) > MAX_INLINE_IMAGE_BYTES {
        return Err(inline_image_too_big(info));
    }
    let bytes = read_attachment_bytes(att, info)?;
    if bytes.len() > MAX_INLINE_IMAGE_BYTES {
        return Err(inline_image_too_big(info));
    }
    Ok(bytes)
}

/// The MIME type an inline image's data URI carries.
fn inline_mime(info: &AttachmentInfo) -> String {
    info.mime_type.clone().unwrap_or_else(|| "application/octet-stream".to_string())
}

/// `get_email`'s `resolve_inline_images`: replace each `cid:` reference in
/// `html` with the matching attachment's `data:` URI. A reference with no
/// matching attachment, or whose attachment is over `MAX_INLINE_IMAGE_BYTES`
/// or can't be read, is left as is and reported in `unresolved`. Only
/// referenced attachments are read.
fn resolve_inline_images(item: &IDispatch, html: &str) -> CidRewrite {
    let wanted = cid_references(html);
    let mut images = Vec::new();
    if !wanted.is_empty() {
        let candidates = attachment_candidates(item).unwrap_or_default();
        for cid in &wanted {
            let key = cid.to_lowercase();
            let found = candidates.iter().find(|(_, info)| {
                info.content_id.as_deref().is_some_and(|c| c.to_lowercase() == key)
            });
            if let Some((att, info)) = found
                && let Ok(bytes) = read_inline_image_bytes(att, info)
            {
                images.push((cid.clone(), data_uri(&inline_mime(info), &bytes)));
            }
        }
    }
    replace_cid_references(html, &images)
}

/// One `get_email` result (see `OutlookClient::get_email`).
fn read_email(ns: &IDispatch, email_id: &str, opts: &ReadOptions,
    out_dir: Option<&std::path::Path>) -> Result<EmailDetail, ToolError> {
    let item = get_item(ns, email_id)?;
    let summary = email_summary(&item)?;
    // Python's `get_email` reads these via `getattr(item, "X", "") or ""`
    // so a non-mail item (MeetingItem, ReportItem, …) that lacks CC/BCC/
    // Body/HTMLBody yields graceful partial detail rather than a COM error.
    let cc = variant_to_string(&get_property(&item, "CC").unwrap_or_default());
    let bcc = variant_to_string(&get_property(&item, "BCC").unwrap_or_default());
    let shape = |full: &str, field: &str, ext: &str| {
        shape_field(full, ReadTool::Email, email_id, field, ext, opts, out_dir)
    };
    // `*_length` is the full original length in chars, so a caller that
    // sees a `*_truncated` flag knows what limit to ask for.
    let body = if opts.body {
        let full = variant_to_string(&get_property(&item, "Body").unwrap_or_default());
        Some(shape(&full, "body", ".txt")?)
    } else {
        None
    };
    let (mut inline_images_resolved, mut inline_images_unresolved) = (None, None);
    let html = if opts.html_body {
        let mut full = variant_to_string(&get_property(&item, "HTMLBody").unwrap_or_default());
        if opts.resolve_inline_images {
            let rewrite = resolve_inline_images(&item, &full);
            inline_images_resolved = Some(rewrite.resolved.len());
            inline_images_unresolved = Some(rewrite.unresolved);
            full = rewrite.html;
        }
        Some(shape(&full, "html_body", ".html")?)
    } else {
        None
    };
    let (body, body_file, body_truncated, body_length) = BodyOut::into_parts(body);
    let (html_body, html_body_file, html_truncated, html_length) = BodyOut::into_parts(html);
    // `getattr(item, "Attachments", None)` then `attachments and attachments.Count`:
    // tolerate an item that has no `Attachments` collection at all (falls back
    // to an empty list) rather than propagating the COM error, mirroring
    // `email_summary`'s already-fixed handling.
    let attachments = opts.attachments.then(|| {
        (|| -> Result<Vec<String>, ToolError> {
            let attachments_obj = to_disp(get_property(&item, "Attachments")?)?;
            let att_count = variant_to_i32(&get_property(&attachments_obj, "Count")?).unwrap_or(0);
            let mut names = Vec::new();
            for i in 1..=att_count {
                let att = to_disp(call_method(&attachments_obj, "Item", &mut [variant_from_i32(i)])?)?;
                names.push(variant_to_string(&get_property(&att, "FileName")?));
            }
            Ok(names)
        })()
        .unwrap_or_default()
    });
    let message_class = variant_to_string(&get_property(&item, "MessageClass").unwrap_or_default());
    let item_type = crate::friendly::item_type_from_class(&message_class).to_string();

    // A MeetingItem exposes GetAssociatedAppointment; a plain MailItem
    // does not. Build the meeting block from the associated appointment.
    let is_meeting = has_member(&item, "GetAssociatedAppointment");
    let meeting = if is_meeting && opts.meeting {
        let appt = to_disp(call_method(
            &item, "GetAssociatedAppointment", &mut [variant_from_bool(false)],
        )?)?;
        Some(MeetingInfo {
            meeting_type: crate::friendly::meeting_type_from_class(&message_class).to_string(),
            start: variant_to_iso_string(&get_property(&appt, "Start").unwrap_or_default()),
            end: variant_to_iso_string(&get_property(&appt, "End").unwrap_or_default()),
            location: variant_to_string(&get_property(&appt, "Location").unwrap_or_default()),
            organizer: variant_to_string(&get_property(&appt, "Organizer").unwrap_or_default()),
            required_attendees: variant_to_string(&get_property(&appt, "RequiredAttendees").unwrap_or_default()),
            optional_attendees: variant_to_string(&get_property(&appt, "OptionalAttendees").unwrap_or_default()),
            is_recurring: variant_to_bool(&get_property(&appt, "IsRecurring").unwrap_or_default()).unwrap_or(false),
        })
    } else {
        None
    };
    Ok(EmailDetail {
        summary,
        cc,
        bcc,
        body,
        body_file,
        body_truncated,
        body_length,
        html_body,
        html_body_file,
        html_truncated,
        html_length,
        inline_images_resolved,
        inline_images_unresolved,
        attachments,
        item_type,
        is_meeting,
        meeting,
    })
}

/// `opts.body` -> the item's plain-text `Body`, shaped for `tool`.
fn read_plain_body(item: &IDispatch, tool: ReadTool, id: &str, opts: &ReadOptions,
    out_dir: Option<&std::path::Path>) -> Result<Option<BodyOut>, ToolError> {
    if !opts.body {
        return Ok(None);
    }
    let full = variant_to_string(&get_property(item, "Body").unwrap_or_default());
    Ok(Some(shape_field(&full, tool, id, "body", ".txt", opts, out_dir)?))
}

/// One `get_event` result.
fn read_event(ns: &IDispatch, event_id: &str, opts: &ReadOptions,
    out_dir: Option<&std::path::Path>) -> Result<EventDetail, ToolError> {
    let item = get_item(ns, event_id)?;
    let summary = event_summary(&item, None)?;
    let recurrence = recurrence_info(&item)?;
    let body = read_plain_body(&item, ReadTool::Event, event_id, opts, out_dir)?;
    let (body, body_file, body_truncated, body_length) = BodyOut::into_parts(body);
    Ok(EventDetail { summary, body, body_file, body_truncated, body_length, recurrence })
}

/// One `get_note` result.
fn read_note(ns: &IDispatch, note_id: &str, opts: &ReadOptions,
    out_dir: Option<&std::path::Path>) -> Result<NoteDetail, ToolError> {
    let note = get_item(ns, note_id)?;
    let summary = note_summary(&note)?;
    let body = read_plain_body(&note, ReadTool::Note, note_id, opts, out_dir)?;
    let (body, body_file, body_truncated, body_length) = BodyOut::into_parts(body);
    Ok(NoteDetail {
        summary,
        body,
        body_file,
        body_truncated,
        body_length,
        modified: variant_to_iso_string(&get_property(&note, "LastModificationTime").unwrap_or_default()),
    })
}

/// One `get_task` result: the `list_tasks` summary plus body and details.
fn read_task(ns: &IDispatch, task_id: &str, opts: &ReadOptions,
    out_dir: Option<&std::path::Path>) -> Result<TaskDetail, ToolError> {
    let task = get_item(ns, task_id)?;
    let summary = task_summary(&task)?;
    let body = read_plain_body(&task, ReadTool::Task, task_id, opts, out_dir)?;
    let (body, body_file, body_truncated, body_length) = BodyOut::into_parts(body);
    let date = |name: &str| {
        outlook_date(variant_to_iso_string(&get_property(&task, name).unwrap_or_default()))
    };
    let reminder_set =
        variant_to_bool(&get_property(&task, "ReminderSet").unwrap_or_default()).unwrap_or(false);
    Ok(TaskDetail {
        summary,
        body,
        body_file,
        body_truncated,
        body_length,
        start_date: date("StartDate"),
        date_completed: date("DateCompleted"),
        percent_complete: variant_to_i32(&get_property(&task, "PercentComplete").unwrap_or_default())
            .unwrap_or(0),
        reminder_set,
        reminder_time: if reminder_set { date("ReminderTime") } else { None },
        created: date("CreationTime"),
        modified: date("LastModificationTime"),
    })
}

/// One `list_attachments` result.
fn list_item_attachments(ns: &IDispatch, email_id: &str) -> Result<Vec<AttachmentInfo>, ToolError> {
    let item = get_item(ns, email_id)?;
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
    check_attachment_paths(paths)?;
    let atts = to_disp(get_property(mail, "Attachments")?)?;
    for p in paths {
        call_method(&atts, "Add", &mut [variant_from_str(p)])?;
    }
    Ok(())
}

/// Every attachment path must be an existing file.
fn check_attachment_paths(paths: &[String]) -> Result<(), ToolError> {
    for p in paths {
        if !std::path::Path::new(p).is_file() {
            return Err(ToolError::new(format!("attachment not found: {p}")));
        }
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
        add_inline_image(&atts, img, temp)?;
    }
    Ok(())
}

/// Add one validated inline image to an `Attachments` collection.
fn add_inline_image(atts: &IDispatch, img: &ValidatedInlineImage, temp: &mut InlineTempDirs)
    -> Result<(), ToolError> {
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
    let att = to_disp(call_method(atts, "Add", &mut [
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
    Ok(())
}

/// `update_draft`'s inline images, added to a draft that may already hold
/// some. An explicit image (`explicit_count` first entries of `images`)
/// replaces any existing attachment with the same Content-ID. An image taken
/// from a `data:` URI is skipped when the draft already has its Content-ID:
/// that id is a hash of the bytes, so it's the same image (e.g. an earlier
/// update with the same HTML).
fn attach_draft_inline_images(
    item: &IDispatch,
    images: &[ValidatedInlineImage],
    explicit_count: usize,
    temp: &mut InlineTempDirs,
) -> Result<(), ToolError> {
    if images.is_empty() {
        return Ok(());
    }
    let atts = to_disp(get_property(item, "Attachments")?)?;
    for (i, img) in images.iter().enumerate() {
        // Walk backwards so deleting keeps the remaining indexes valid.
        let count = variant_to_i32(&get_property(&atts, "Count")?).unwrap_or(0);
        let mut present = false;
        for idx in (1..=count).rev() {
            let att = to_disp(call_method(&atts, "Item", &mut [variant_from_i32(idx)])?)?;
            let cid = get_mapi_prop(&att, c::PR_ATTACH_CONTENT_ID).map(|v| variant_to_string(&v));
            let same = clean_content_id(cid.as_deref())
                .is_some_and(|cid| cid.eq_ignore_ascii_case(&img.content_id));
            if !same {
                continue;
            }
            if i < explicit_count {
                call_method(&att, "Delete", &mut [])?;
            } else {
                present = true;
            }
        }
        if !present {
            add_inline_image(&atts, img, temp)?;
        }
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
    // `category`, `has_attachments`, `flag`, `item_type` and the `query`/`to`
    // fallbacks are filtered client-side while iterating, before paging.
    fn list_emails(&self, q: EmailQuery) -> Result<Vec<EmailSummary>, ToolError> {
        // Resolve and validate every input before touching Outlook.
        let received = DateRange::parse(
            q.received_after.as_deref(), q.received_before.as_deref(),
            "received_after", "received_before", chrono::Local::now().naive_local(),
            com::user_first_day_of_week(),
        )?;
        let importance_ids = q.importance.iter()
            .map(|i| c::importance_name_to_id(i.trim()).ok_or_else(|| {
                ToolError::new(format!("Invalid importance {i:?}: use one of low, normal, high"))
            }))
            .collect::<Result<Vec<i32>, ToolError>>()?;
        let flag_ids = q.flag.iter()
            .map(|f| filters::flag_status_id(f.trim()).ok_or_else(|| {
                ToolError::new(format!("Invalid flag {f:?}: use one of {}", filters::FLAGS.join(", ")))
            }))
            .collect::<Result<Vec<i32>, ToolError>>()?;
        let item_types = filters::normalize_choices(q.item_type.clone(), "item_type", filters::ITEM_TYPES)?;
        let from = filters::clean_values(q.from.clone());
        let to = filters::clean_values(q.to.clone());
        let text = TextQuery::parse(q.query.as_deref().unwrap_or(""), filters::EMAIL_QUERY_FIELDS);
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let count = q.count.clamp(1, MAX_EMAIL_COUNT);
            let folder_obj = resolve_folder(&ns, Some(&q.folder))?;
            let mut items = to_disp(get_property(&folder_obj, "Items")?)?;
            let restrict = |items: &IDispatch, filter: &str| -> Result<IDispatch, ToolError> {
                to_disp(call_method(items, "Restrict", &mut [variant_from_str(filter)])?)
            };
            let count_of = |items: &IDispatch| -> Result<i32, ToolError> {
                Ok(variant_to_i32(&get_property(items, "Count")?).unwrap_or(0))
            };

            // Sender: DASL @SQL against fromname + fromemail, ORed over the needles.
            if !from.is_empty() {
                let ors: Vec<String> = from.iter().map(|f| {
                    let e = f.replace('\'', "''");
                    format!("\"urn:schemas:httpmail:fromname\" LIKE '%{e}%' \
                             OR \"urn:schemas:httpmail:fromemail\" LIKE '%{e}%'")
                }).collect();
                items = restrict(&items, &format!("@SQL=({})", ors.join(" OR ")))?;
            }
            if q.unread_only {
                items = restrict(&items, "[UnRead] = True")?;
            }
            // `flag` is deliberately NOT a Restrict call: confirmed via a
            // raw PowerShell COM probe outside this codebase that
            // `Items.Restrict("[FlagStatus] = 2")` doesn't reliably match on
            // this account class, even though `Item.FlagStatus` reads
            // correctly when read directly per-item (modern Outlook/M365
            // flags sync through the To-Do integration rather than classic
            // MAPI, and the legacy DASL bracket filter doesn't see that
            // state reliably). Filtered client-side below instead, alongside
            // category/has_attachments.
            if !importance_ids.is_empty() {
                let ors: Vec<String> =
                    importance_ids.iter().map(|id| format!("[Importance] = {id}")).collect();
                items = restrict(&items, &ors.join(" OR "))?;
            }
            // Date range, already resolved by the shared date grammar. The
            // JET string format (and its day-first-locale bug) is issue #1.
            if let Some(after) = received.after {
                items = restrict(&items, &format!("[ReceivedTime] >= '{}'", jet_datetime(&after)))?;
            }
            if let Some(before) = received.before {
                items = restrict(&items, &format!("[ReceivedTime] <= '{}'", jet_datetime(&before)))?;
            }

            // Text query near-last, so the fallback scan below only walks
            // items that already passed every other filter: the shared query
            // syntax as one DASL @SQL filter (see `TextQuery::to_dasl`).
            let mut text_fallback = false;
            if let Some(dasl) = text.to_dasl(filters::EMAIL_QUERY_DEFAULTS, filters::EMAIL_QUERY_DASL) {
                let matched = restrict(&items, &dasl)?;
                // DASL LIKE has been observed to return nothing for Hebrew
                // (and other non-Latin) terms even when matching mail exists
                // (issue #2). For a non-ASCII query, treat an empty DASL
                // result as unreliable and scan the pre-filtered items
                // client-side instead. ASCII queries stay DASL-only.
                if count_of(&matched)? == 0 && !text.is_ascii() {
                    text_fallback = true;
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
            // So when the DASL restrict for any one needle comes back empty,
            // fall back to scanning each item's Recipients (Name, Address,
            // SMTP address) client-side for every needle, over at most
            // RECIPIENT_SCAN_LIMIT newest items. When every needle has DASL
            // hits, their ORed DASL result is kept as-is: every hit
            // genuinely has a needle in its To/CC line, and it avoids the
            // per-item walk.
            let mut to_scan = false;
            if !to.is_empty() {
                let clause = |needle: &str| {
                    let e = needle.replace('\'', "''");
                    format!("\"urn:schemas:httpmail:displayto\" LIKE '%{e}%' \
                             OR \"urn:schemas:httpmail:displaycc\" LIKE '%{e}%'")
                };
                for needle in &to {
                    if count_of(&restrict(&items, &format!("@SQL=({})", clause(needle)))?)? == 0 {
                        to_scan = true;
                        break;
                    }
                }
                if !to_scan {
                    let ors: Vec<String> = to.iter().map(|n| clause(n)).collect();
                    items = restrict(&items, &format!("@SQL=({})", ors.join(" OR ")))?;
                }
            }

            call_method(
                &items,
                "Sort",
                &mut [variant_from_str("[ReceivedTime]"), variant_from_bool(true)],
            )?;

            // Client-side filters: item_type + flag (one property read each),
            // the non-ASCII query fallback and `to` recipient fallback (if
            // active), then category + has_attachments on the summary.
            // Lazily build each summary and keep it only if it passes; then
            // `take_page` skips the first `offset` matches (after these
            // filters, so pages line up) and stops at count.
            let mut total = count_of(&items)?;
            if text_fallback {
                // The fallback reads Subject/SenderName (and maybe Body) per
                // item, so cap how far back it scans (newest first).
                total = total.min(MAX_SCAN_ITEMS);
            }
            if to_scan {
                // Same for the per-item Recipients walk of the `to` fallback.
                total = total.min(RECIPIENT_SCAN_LIMIT);
            }
            let matches = (1..=total).filter_map(|i| {
                (|| -> Result<Option<EmailSummary>, ToolError> {
                    let item = to_disp(call_method(&items, "Item", &mut [variant_from_i32(i)])?)?;
                    if !item_types.is_empty() {
                        let class = variant_to_string(&get_property(&item, "MessageClass").unwrap_or_default());
                        let kind = crate::friendly::item_type_from_class(&class);
                        if !item_types.iter().any(|t| t == kind) {
                            return Ok(None);
                        }
                    }
                    if !flag_ids.is_empty() {
                        let status = variant_to_i32(&get_property(&item, "FlagStatus").unwrap_or_default())
                            .unwrap_or(c::OL_NO_FLAG);
                        if !flag_ids.contains(&status) {
                            return Ok(None);
                        }
                    }
                    if text_fallback
                        && !text.matches(filters::EMAIL_QUERY_DEFAULTS, |field| email_query_field(&item, field))
                    {
                        return Ok(None);
                    }
                    if to_scan {
                        let candidates = recipient_strings(&item);
                        if !to.iter().any(|needle| recipient_matches(needle, &candidates)) {
                            return Ok(None);
                        }
                    }
                    let summary = email_summary(&item)?;
                    if !filters::has_category(&summary.categories, &q.category) {
                        return Ok(None);
                    }
                    if q.has_attachments.is_some_and(|want| summary.has_attachments != want) {
                        return Ok(None);
                    }
                    Ok(Some(summary))
                })()
                .transpose()
            });
            take_page(matches, q.offset, count)
        })
    }

    fn get_email(&self, email_ids: Vec<String>, opts: &ReadOptions) -> BatchResult<EmailDetail> {
        let out_dir = opts.output_dir.as_deref().map(prepare_output_dir).transpose()?;
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            Ok(email_ids.iter().map(|id| read_email(&ns, id, opts, out_dir.as_deref())).collect())
        })
    }

    fn send_email(&self, m: NewEmail) -> Result<Value, ToolError> {
        if m.to.is_empty() {
            return Err(ToolError::new(
                "send_email requires at least one recipient in 'to'.",
            ));
        }
        // Validate everything before the item is created.
        let (body, images, importance) = prepare_new_email(&m)?;
        self.with_com(|| {
            let (app, _ns) = mapi()?;
            let mail = compose(&app, &m.to, &m.subject, &body, m.cc.as_deref(), m.bcc.as_deref())?;
            set_mail_metadata(&mail, m.categories.as_deref(), importance)?;
            if let Some(atts) = m.attachments.as_deref() {
                attach_files(&mail, atts)?;
            }
            // Dropped at the end of this closure: temp files outlive Send.
            let mut temp = InlineTempDirs::default();
            attach_inline_images(&mail, &images, &mut temp)?;
            call_method(&mail, "Send", &mut [])?;
            Ok(json!({"status": "sent", "to": m.to.join("; "), "subject": m.subject}))
        })
    }

    fn create_draft(&self, m: NewEmail) -> Result<Value, ToolError> {
        let (body, images, importance) = prepare_new_email(&m)?;
        self.with_com(|| {
            let (app, _ns) = mapi()?;
            let mail = compose(&app, &m.to, &m.subject, &body, m.cc.as_deref(), m.bcc.as_deref())?;
            set_mail_metadata(&mail, m.categories.as_deref(), importance)?;
            if let Some(atts) = m.attachments.as_deref() {
                attach_files(&mail, atts)?;
            }
            // Dropped at the end of this closure: temp files outlive Save.
            let mut temp = InlineTempDirs::default();
            attach_inline_images(&mail, &images, &mut temp)?;
            call_method(&mail, "Save", &mut [])?; // Save first so EntryID exists
            let id = make_id(&mail)?;
            Ok(json!({"status": "draft_saved", "id": id, "subject": m.subject}))
        })
    }

    fn reply_email(&self, r: ReplyInput) -> Result<Value, ToolError> {
        // Validate everything before the reply is created.
        let importance = parse_importance(r.importance.as_deref())?;
        check_attachment_paths(r.attachments.as_deref().unwrap_or(&[]))?;
        let (body, images) = prepare_mail_body(&r.body, r.inline_images.as_deref().unwrap_or(&[]))?;
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &r.email_id)?;
            let reply = to_disp(call_method(
                &item,
                if r.reply_all { "ReplyAll" } else { "Reply" },
                &mut [],
            )?)?;
            match &body {
                MailBody::Html(html) => {
                    let existing = variant_to_string(&get_property(&reply, "HTMLBody")?);
                    put_property(&reply, "HTMLBody", variant_from_str(&format!("{html}{existing}")))?;
                }
                MailBody::Text(text) => {
                    let existing = variant_to_string(&get_property(&reply, "Body")?);
                    put_property(&reply, "Body", variant_from_str(&format!("{text}\n\n{existing}")))?;
                }
            }
            set_mail_metadata(&reply, r.categories.as_deref(), importance)?;
            if let Some(atts) = r.attachments.as_deref() {
                attach_files(&reply, atts)?;
            }
            // Dropped at the end of this closure: temp files outlive Send/Save.
            let mut temp = InlineTempDirs::default();
            attach_inline_images(&reply, &images, &mut temp)?;
            if r.send {
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
        let importance = parse_importance(u.importance.as_deref())?;
        let explicit = u.inline_images.as_deref().unwrap_or(&[]);
        let (html_body, images) = match &u.html_body {
            Some(html) => {
                let (html, images) = prepare_html_images(html, explicit)?;
                (Some(html), images)
            }
            None => (None, Vec::new()),
        };
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &u.email_id)?;
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
            if let Some(body) = &u.body {
                set_mail_body(&item, &MailBody::Text(body.clone()))?;
            }
            if let Some(html) = html_body {
                set_mail_body(&item, &MailBody::Html(html))?;
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
            // Dropped at the end of this closure: temp files outlive Save/Send.
            let mut temp = InlineTempDirs::default();
            attach_draft_inline_images(&item, &images, explicit.len(), &mut temp)?;
            if u.add_categories.is_some() || u.remove_categories.is_some() {
                let cats = merge_categories(
                    get_item_categories(&item),
                    u.add_categories.as_deref().unwrap_or(&[]),
                    u.remove_categories.as_deref().unwrap_or(&[]),
                );
                set_item_categories(&item, &cats)?;
            }
            if let Some(id) = importance {
                put_property(&item, "Importance", variant_from_i32(id))?;
            }

            call_method(&item, "Save", &mut [])?; // one Save for all changes
            let id = make_id(&item)?;
            let changed = draft_update_changes(&u);
            if !u.send {
                return Ok(json!({"status": "draft_updated", "id": id, "changed": changed}));
            }
            let recipients = to_disp(get_property(&item, "Recipients")?)?;
            if variant_to_i32(&get_property(&recipients, "Count")?).unwrap_or(0) == 0 {
                return Err(ToolError::new(format!(
                    "the draft has no recipients, so it was saved but not sent (id {id}); \
                     pass to/cc/bcc to address it"
                )));
            }
            // Read Subject before Send(): the item is invalid afterwards.
            let subject = variant_to_string(&get_property(&item, "Subject")?);
            call_method(&item, "Send", &mut [])?;
            Ok(json!({"status": "sent", "subject": subject, "changed": changed}))
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
        // Resolve the scan window with the shared date grammar (a bare ISO
        // `start_before` date includes that whole day). Defaults: from today
        // 00:00, for 7 days.
        let now = chrono::Local::now().naive_local();
        let range = DateRange::parse(
            q.start_after.as_deref(), q.start_before.as_deref(), "start_after", "start_before", now,
            com::user_first_day_of_week(),
        )?;
        let start = range.after.unwrap_or_else(|| now.date().and_hms_opt(0, 0, 0).unwrap());
        let end = range.before.unwrap_or(start + chrono::Duration::days(7));
        if start > end {
            return Err(ToolError::new(format!(
                "start_before ({end}) is earlier than the default start_after (today, {start}): pass start_after too"
            )));
        }
        let text = TextQuery::parse(q.query.as_deref().unwrap_or(""), filters::EVENT_QUERY_FIELDS);
        let count = q.count.clamp(1, MAX_EVENT_COUNT) as usize;
        self.with_com(|| {
            let (_app, ns) = mapi()?;
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
            let flt = format!(
                "[Start] >= '{}' AND [Start] <= '{}'",
                jet_datetime(&start),
                jet_datetime(&end)
            );
            enumerate_events_with_retry(&items, &flt, &q, &text, &calendar_store_id, count)
        })
    }

    fn get_event(&self, event_ids: Vec<String>, opts: &ReadOptions) -> BatchResult<EventDetail> {
        let out_dir = opts.output_dir.as_deref().map(prepare_output_dir).transpose()?;
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            Ok(event_ids.iter().map(|id| read_event(&ns, id, opts, out_dir.as_deref())).collect())
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

    fn list_attachments(&self, email_ids: Vec<String>) -> BatchResult<Vec<AttachmentInfo>> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            Ok(email_ids.iter().map(|id| list_item_attachments(&ns, id)).collect())
        })
    }

    fn save_attachments(
        &self,
        email_id: String,
        save_dir: String,
        attachment_names: Option<Vec<String>>,
        inline: Option<bool>,
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
                if inline.is_some_and(|want| info.is_inline != want) {
                    continue;
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
                    "No attachments matched attachment_names / inline; use \
                     list_attachments to see the exact file names and `is_inline`.",
                ));
            }
            Ok(results)
        })
    }

    fn get_inline_image(&self, email_id: String, content_ids: Vec<String>,
        context_lines: Option<u32>, output_dir: Option<String>) -> BatchResult<InlineImageData> {
        let out_dir = output_dir.as_deref().map(prepare_output_dir).transpose()?;
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &email_id)?;
            let candidates = attachment_candidates(&item)?;
            let available: Vec<Option<String>> =
                candidates.iter().map(|(_, info)| info.content_id.clone()).collect();
            // As in get_email, an item without `HTMLBody` reads as empty, so
            // an image just counts as unreferenced. Read once, only when
            // `context` was asked for.
            let html = context_lines
                .map(|_| variant_to_string(&get_property(&item, "HTMLBody").unwrap_or_default()));
            let fetch = |raw: &String| -> Result<InlineImageData, ToolError> {
                let wanted = normalize_cid_request(raw).ok_or_else(|| {
                    ToolError::new("content_id must be a non-empty Content-ID (e.g. \"image001.png@01D9...\", optionally prefixed with cid:).")
                })?;
                let (att, info) = &candidates[select_by_content_id(&wanted, &available)?];
                let bytes = read_inline_image_bytes(att, info)?;
                let mime_type = inline_mime(info);
                let context = context_lines
                    .zip(html.as_deref())
                    .map(|(n, html)| text_before_cid(html, &wanted, n).unwrap_or_default());
                let (data_uri, data_file) = match out_dir.as_deref() {
                    Some(dir) => {
                        let key = format!("{email_id}\n{}", wanted.to_lowercase());
                        let ext = image_extension(&mime_type, &info.filename);
                        let path = output_file_path(dir, "image", &key, "data", &ext);
                        (None, Some(write_output_file(&path, &bytes)?))
                    }
                    None => (Some(data_uri(&mime_type, &bytes)), None),
                };
                Ok(InlineImageData {
                    content_id: info.content_id.clone().unwrap_or_default(),
                    filename: info.filename.clone(),
                    data_uri,
                    data_file,
                    mime_type,
                    size: bytes.len(),
                    context,
                })
            };
            Ok(content_ids.iter().map(fetch).collect())
        })
    }

    // ---- Tasks (Task 15) -----------------------------------------------

    fn list_tasks(&self, q: TaskQuery) -> Result<Vec<TaskSummary>, ToolError> {
        let due = DateRange::parse(
            q.due_after.as_deref(), q.due_before.as_deref(), "due_after", "due_before",
            chrono::Local::now().naive_local(), com::user_first_day_of_week(),
        )?;
        let text = TextQuery::parse(q.query.as_deref().unwrap_or(""), filters::TASK_QUERY_FIELDS);
        let count = q.count.clamp(1, MAX_TASK_COUNT);
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
            let total = variant_to_i32(&get_property(&items, "Count")?).unwrap_or(0);
            let matches = (1..=total).filter_map(|i| {
                (|| -> Result<Option<TaskSummary>, ToolError> {
                    let item = to_disp(call_method(&items, "Item", &mut [variant_from_i32(i)])?)?;
                    let summary = task_summary(&item)?;
                    // `task_summary` doesn't expose the body, so a query that
                    // needs it reads the real body here, lazily.
                    let body = || variant_to_string(&get_property(&item, "Body").unwrap_or_default());
                    Ok(filters::task_matches(&summary, &q, &due, &text, body).then_some(summary))
                })()
                .transpose()
            });
            take_page(matches, q.offset, count)
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

    fn get_task(&self, task_ids: Vec<String>, opts: &ReadOptions) -> BatchResult<TaskDetail> {
        let out_dir = opts.output_dir.as_deref().map(prepare_output_dir).transpose()?;
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            Ok(task_ids.iter().map(|id| read_task(&ns, id, opts, out_dir.as_deref())).collect())
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
        let created = DateRange::parse(
            q.created_after.as_deref(), q.created_before.as_deref(), "created_after",
            "created_before", chrono::Local::now().naive_local(), com::user_first_day_of_week(),
        )?;
        let text = TextQuery::parse(q.query.as_deref().unwrap_or(""), filters::NOTE_QUERY_FIELDS);
        let count = q.count.clamp(1, MAX_NOTE_COUNT);
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let notes = to_disp(call_method(
                &ns,
                "GetDefaultFolder",
                &mut [variant_from_i32(c::OL_FOLDER_NOTES)],
            )?)?;
            let items = to_disp(get_property(&notes, "Items")?)?;
            let total = variant_to_i32(&get_property(&items, "Count")?).unwrap_or(0);
            let matches = (1..=total).filter_map(|i| {
                (|| -> Result<Option<NoteSummary>, ToolError> {
                    let item = to_disp(call_method(&items, "Item", &mut [variant_from_i32(i)])?)?;
                    let summary = note_summary(&item)?;
                    // `note_summary` only exposes the derived (120-char)
                    // subject, so a query on the body re-reads the real,
                    // untruncated body here, lazily.
                    let body = || variant_to_string(&get_property(&item, "Body").unwrap_or_default());
                    Ok(filters::note_matches(&summary, &q, &created, &text, body).then_some(summary))
                })()
                .transpose()
            });
            take_page(matches, q.offset, count)
        })
    }

    fn get_note(&self, note_ids: Vec<String>, opts: &ReadOptions) -> BatchResult<NoteDetail> {
        let out_dir = opts.output_dir.as_deref().map(prepare_output_dir).transpose()?;
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            Ok(note_ids.iter().map(|id| read_note(&ns, id, opts, out_dir.as_deref())).collect())
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
