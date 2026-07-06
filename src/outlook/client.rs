//! Win32 COM implementation of the `OutlookClient` trait.
//!
//! Direct port of `outlook_mcp/outlook/client.py`'s email section
//! (lines 133-336). Every public method wraps its body in [`with_com`],
//! which initializes COM on the current thread (mirroring the Python
//! `@_com` decorator) and maps `windows::core::Error` into [`ToolError`].
//!
//! All 21 `OutlookClient` trait methods are implemented (email, calendar,
//! attachments, tasks, and notes; Tasks 12-16) — no `todo!()` stubs remain.

use serde_json::{json, Value};
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Variant::VARIANT;

use crate::constants as c;
use crate::error::ToolError;
use crate::outlook::com::{
    call_method, create_com_object, format_com_error, get_property, has_member, jet_datetime,
    make_item_id, parse_item_id, put_property, safe_filename, variant_from_bool,
    variant_from_datetime, variant_from_i32, variant_from_str, variant_to_bool, variant_to_i32,
    variant_to_iso_string, variant_to_string, ComGuard,
};
use crate::outlook::types::*;
use crate::outlook::OutlookClient;

/// Matches `MAX_EMAIL_COUNT` in `client.py`.
const MAX_EMAIL_COUNT: i32 = 50;
/// Matches `MAX_CALENDAR_ITEMS` in `client.py`. Caps `list_events` because a
/// recurring appointment without an end date expands forever under
/// `IncludeRecurrences`.
const MAX_CALENDAR_ITEMS: usize = 250;
/// Matches `MAX_BODY_CHARS` in `client.py`.
const MAX_BODY_CHARS: usize = 100_000;

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

/// `client.py::_truncate`: cap long bodies at `MAX_BODY_CHARS` *characters*
/// (not bytes) so multi-byte UTF-8 content is never split mid-codepoint.
fn truncate(text: &str) -> String {
    if text.chars().count() > MAX_BODY_CHARS {
        let head: String = text.chars().take(MAX_BODY_CHARS).collect();
        format!("{head}\n\n[... truncated at {MAX_BODY_CHARS} characters]")
    } else {
        text.to_string()
    }
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

/// `client.py::_event_summary`.
fn event_summary(item: &IDispatch) -> Result<EventSummary, ToolError> {
    let meeting_status = variant_to_i32(&get_property(item, "MeetingStatus").unwrap_or_default())
        .unwrap_or(c::OL_NONMEETING);
    Ok(EventSummary {
        id: make_id(item)?,
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
    })
}

/// `client.py::_task_summary`. `status` and `importance` are the raw numeric
/// COM properties (not name lookups); a missing value falls back to Outlook's
/// defaults exactly like the Python `getattr(..., default)`.
fn task_summary(item: &IDispatch) -> Result<TaskSummary, ToolError> {
    Ok(TaskSummary {
        id: make_id(item)?,
        subject: variant_to_string(&get_property(item, "Subject").unwrap_or_default()),
        due_date: variant_to_iso_string(&get_property(item, "DueDate").unwrap_or_default()),
        complete: variant_to_bool(&get_property(item, "Complete").unwrap_or_default())
            .unwrap_or(false),
        status: variant_to_i32(&get_property(item, "Status").unwrap_or_default())
            .unwrap_or(c::OL_TASK_NOT_STARTED),
        importance: variant_to_i32(&get_property(item, "Importance").unwrap_or_default())
            .unwrap_or(c::OL_IMPORTANCE_NORMAL),
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
    })
}

/// Shared tail of `list_emails`/`search_emails`: iterate a (sorted, possibly
/// restricted) `Items` collection by 1-based index, building summaries until
/// `count` is reached. Item order reflects the prior `Sort` call.
fn collect_summaries(items: &IDispatch, count: i32) -> Result<Vec<EmailSummary>, ToolError> {
    let total = variant_to_i32(&get_property(items, "Count")?).unwrap_or(0);
    let mut results = Vec::new();
    for i in 1..=total {
        let item = to_disp(call_method(items, "Item", &mut [variant_from_i32(i)])?)?;
        results.push(email_summary(&item)?);
        if results.len() as i32 >= count {
            break;
        }
    }
    Ok(results)
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

    fn list_emails(
        &self,
        folder: String,
        count: i32,
        unread_only: bool,
    ) -> Result<Vec<EmailSummary>, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let count = count.clamp(1, MAX_EMAIL_COUNT);
            let folder_obj = resolve_folder(&ns, Some(&folder))?;
            let mut items = to_disp(get_property(&folder_obj, "Items")?)?;
            if unread_only {
                items = to_disp(call_method(
                    &items,
                    "Restrict",
                    &mut [variant_from_str("[UnRead] = True")],
                )?)?;
            }
            call_method(
                &items,
                "Sort",
                &mut [variant_from_str("[ReceivedTime]"), variant_from_bool(true)],
            )?;
            collect_summaries(&items, count)
        })
    }

    fn search_emails(
        &self,
        query: String,
        folder: String,
        count: i32,
        since_days: Option<i32>,
    ) -> Result<Vec<EmailSummary>, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let count = count.clamp(1, MAX_EMAIL_COUNT);
            // Escape single quotes by doubling them, exactly like the Python
            // `query.replace("'", "''")` before interpolating into the DASL.
            let q = query.replace('\'', "''");
            let dasl = format!(
                "@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%{q}%' \
                 OR \"urn:schemas:httpmail:fromname\" LIKE '%{q}%' \
                 OR \"urn:schemas:httpmail:textdescription\" LIKE '%{q}%')"
            );
            let folder_obj = resolve_folder(&ns, Some(&folder))?;
            let base_items = to_disp(get_property(&folder_obj, "Items")?)?;
            let mut items =
                to_disp(call_method(&base_items, "Restrict", &mut [variant_from_str(&dasl)])?)?;
            if since_days.is_some_and(|d| d != 0) {
                let days = since_days.unwrap();
                let cutoff = chrono::Local::now().naive_local() - chrono::Duration::days(days as i64);
                let filter = format!("[ReceivedTime] >= '{}'", jet_datetime(&cutoff));
                items = to_disp(call_method(&items, "Restrict", &mut [variant_from_str(&filter)])?)?;
            }
            call_method(
                &items,
                "Sort",
                &mut [variant_from_str("[ReceivedTime]"), variant_from_bool(true)],
            )?;
            collect_summaries(&items, count)
        })
    }

    fn get_email(&self, email_id: String, prefer_html: bool) -> Result<EmailDetail, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &email_id)?;
            let summary = email_summary(&item)?;
            // Python's `get_email` reads these via `getattr(item, "X", "") or ""`
            // so a non-mail item (MeetingItem, ReportItem, …) that lacks CC/BCC/
            // Body/HTMLBody yields graceful partial detail rather than a COM error.
            let cc = variant_to_string(&get_property(&item, "CC").unwrap_or_default());
            let bcc = variant_to_string(&get_property(&item, "BCC").unwrap_or_default());
            let body = truncate(&variant_to_string(&get_property(&item, "Body").unwrap_or_default()));
            let html_body = if prefer_html {
                Some(truncate(&variant_to_string(
                    &get_property(&item, "HTMLBody").unwrap_or_default(),
                )))
            } else {
                None
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
            Ok(EmailDetail {
                summary,
                cc,
                bcc,
                body,
                html_body,
                attachments,
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
    ) -> Result<Value, ToolError> {
        if to.is_empty() {
            return Err(ToolError::new(
                "send_email requires at least one recipient in 'to'.",
            ));
        }
        self.with_com(|| {
            let (app, _ns) = mapi()?;
            let mail = compose(&app, &to, &subject, &body, cc.as_deref(), bcc.as_deref(), html)?;
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
    ) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (app, _ns) = mapi()?;
            let mail = compose(&app, &to, &subject, &body, cc.as_deref(), bcc.as_deref(), html)?;
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
            if send {
                call_method(&reply, "Send", &mut [])?;
                let subject = variant_to_string(&get_property(&reply, "Subject")?);
                Ok(json!({"status": "sent", "subject": subject}))
            } else {
                call_method(&reply, "Save", &mut [])?;
                let id = make_id(&reply)?;
                let subject = variant_to_string(&get_property(&reply, "Subject")?);
                Ok(json!({"status": "draft_saved", "id": id, "subject": subject}))
            }
        })
    }

    fn move_email(&self, email_id: String, target_folder: String) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &email_id)?;
            let target = resolve_folder(&ns, Some(&target_folder))?;
            // `Move` takes the destination folder as its argument; wrap the
            // IDispatch in a VARIANT (clone — `From<IDispatch>` consumes it).
            let moved = to_disp(call_method(
                &item,
                "Move",
                &mut [VARIANT::from(target.clone())],
            )?)?;
            let folder_name = variant_to_string(&get_property(&target, "Name")?);
            let id = make_id(&moved)?; // EntryID changes on Move — return the new id.
            Ok(json!({"status": "moved", "folder": folder_name, "id": id}))
        })
    }

    fn delete_email(&self, email_id: String) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &email_id)?;
            let subject = variant_to_string(&get_property(&item, "Subject")?);
            call_method(&item, "Delete", &mut [])?;
            Ok(json!({"status": "deleted", "subject": subject, "note": "Moved to Deleted Items."}))
        })
    }

    // ---- Calendar (Task 13) --------------------------------------------

    fn list_events(
        &self,
        start_date: Option<String>,
        end_date: Option<String>,
    ) -> Result<Vec<EventSummary>, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let start = match &start_date {
                Some(s) => parse_dt(s, "start_date")?,
                None => chrono::Local::now()
                    .date_naive()
                    .and_hms_opt(0, 0, 0)
                    .unwrap(),
            };
            let mut end = match &end_date {
                Some(s) => parse_dt(s, "end_date")?,
                None => start + chrono::Duration::days(7),
            };
            // If only a bare date was given for the end, treat it as the whole
            // end day (Python: `end.time() == time.min and "T" not in end_date`).
            if let Some(ed) = &end_date {
                if end.time() == chrono::NaiveTime::MIN && !ed.contains('T') {
                    end = end.date().and_hms_micro_opt(23, 59, 59, 999_999).unwrap();
                }
            }
            let calendar = to_disp(call_method(
                &ns,
                "GetDefaultFolder",
                &mut [variant_from_i32(c::OL_FOLDER_CALENDAR)],
            )?)?;
            let items = to_disp(get_property(&calendar, "Items")?)?;
            // Must precede Sort/Restrict — setting it afterwards has no effect.
            put_property(&items, "IncludeRecurrences", variant_from_bool(true))?;
            call_method(&items, "Sort", &mut [variant_from_str("[Start]")])?;
            let flt = format!(
                "[Start] >= '{}' AND [Start] <= '{}'",
                jet_datetime(&start),
                jet_datetime(&end)
            );
            let restricted =
                to_disp(call_method(&items, "Restrict", &mut [variant_from_str(&flt)])?)?;
            // Enumerate with GetFirst/GetNext (not Count/Item): under
            // IncludeRecurrences the collection can expand without bound, so we
            // must stream it and stop at MAX_CALENDAR_ITEMS.
            let mut results = Vec::new();
            let mut current = call_method(&restricted, "GetFirst", &mut [])?;
            while let Ok(item) = IDispatch::try_from(&current) {
                results.push(event_summary(&item)?);
                if results.len() >= MAX_CALENDAR_ITEMS {
                    break;
                }
                current = call_method(&restricted, "GetNext", &mut [])?;
            }
            Ok(results)
        })
    }

    fn get_event(&self, event_id: String) -> Result<EventDetail, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &event_id)?;
            let summary = event_summary(&item)?;
            // Python's `get_event` reads all four via `getattr(..., default)`, so an
            // item lacking these appointment-only properties yields partial detail
            // instead of a COM error. An empty VARIANT for ResponseStatus decodes to
            // None, matching Python's `getattr(item, "ResponseStatus", None)` default.
            Ok(EventDetail {
                summary,
                body: truncate(&variant_to_string(&get_property(&item, "Body").unwrap_or_default())),
                required_attendees: variant_to_string(
                    &get_property(&item, "RequiredAttendees").unwrap_or_default(),
                ),
                optional_attendees: variant_to_string(
                    &get_property(&item, "OptionalAttendees").unwrap_or_default(),
                ),
                response_status: variant_to_i32(
                    &get_property(&item, "ResponseStatus").unwrap_or_default(),
                ),
            })
        })
    }

    fn create_event(
        &self,
        subject: String,
        start: String,
        end: String,
        body: Option<String>,
        location: Option<String>,
        attendees: Option<Vec<String>>,
        all_day: bool,
        reminder_minutes: Option<i32>,
    ) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (app, _ns) = mapi()?;
            let appt = to_disp(call_method(
                &app,
                "CreateItem",
                &mut [variant_from_i32(c::OL_APPOINTMENT_ITEM)],
            )?)?;
            put_property(&appt, "Subject", variant_from_str(&subject))?;
            put_property(&appt, "Start", variant_from_datetime(&parse_dt(&start, "start")?)?)?;
            put_property(&appt, "End", variant_from_datetime(&parse_dt(&end, "end")?)?)?;
            if all_day {
                put_property(&appt, "AllDayEvent", variant_from_bool(true))?;
            }
            if let Some(body) = body.as_deref().filter(|b| !b.is_empty()) {
                put_property(&appt, "Body", variant_from_str(body))?;
            }
            if let Some(location) = location.as_deref().filter(|l| !l.is_empty()) {
                put_property(&appt, "Location", variant_from_str(location))?;
            }
            if let Some(minutes) = reminder_minutes {
                put_property(&appt, "ReminderSet", variant_from_bool(true))?;
                put_property(&appt, "ReminderMinutesBeforeStart", variant_from_i32(minutes))?;
            }
            let status = match attendees {
                Some(addresses) if !addresses.is_empty() => {
                    put_property(&appt, "MeetingStatus", variant_from_i32(c::OL_MEETING))?;
                    let recipients = to_disp(get_property(&appt, "Recipients")?)?;
                    for address in &addresses {
                        call_method(&recipients, "Add", &mut [variant_from_str(address)])?;
                    }
                    call_method(&recipients, "ResolveAll", &mut [])?;
                    call_method(&appt, "Send", &mut [])?;
                    "meeting_sent"
                }
                _ => {
                    call_method(&appt, "Save", &mut [])?;
                    "saved"
                }
            };
            Ok(json!({"status": status, "id": make_id(&appt)?, "subject": subject}))
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
            let mut results = Vec::new();
            for i in 1..=count {
                // COM collections are 1-based.
                let att = to_disp(call_method(&attachments, "Item", &mut [variant_from_i32(i)])?)?;
                results.push(AttachmentInfo {
                    index: i,
                    filename: variant_to_string(&get_property(&att, "FileName")?),
                    size: variant_to_i32(&get_property(&att, "Size")?).unwrap_or(0),
                });
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
            let mut results = Vec::new();
            for i in 1..=count {
                let att = to_disp(call_method(&attachments, "Item", &mut [variant_from_i32(i)])?)?;
                let raw = variant_to_string(&get_property(&att, "FileName")?);
                let filename = if raw.is_empty() {
                    format!("attachment-{i}")
                } else {
                    raw
                };
                if let Some(wanted) = &wanted {
                    if !wanted.contains(&filename.to_lowercase()) {
                        continue;
                    }
                }
                let target = dir.join(safe_filename(&filename));
                let target_str = target.to_string_lossy().into_owned();
                // A COM failure saving one file is collected per-file and does
                // NOT abort the batch (mirrors the per-file try/except in Python).
                match call_method(&att, "SaveAsFile", &mut [variant_from_str(&target_str)]) {
                    Ok(_) => results.push(json!({
                        "filename": filename,
                        "saved_to": target_str,
                        "status": "saved",
                    })),
                    Err(e) => results.push(json!({
                        "filename": filename,
                        "status": "failed",
                        "error": format_com_error(&e),
                    })),
                }
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

    // ---- Tasks (Task 15) -----------------------------------------------

    fn list_tasks(&self, include_completed: bool) -> Result<Vec<TaskSummary>, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let tasks = to_disp(call_method(
                &ns,
                "GetDefaultFolder",
                &mut [variant_from_i32(c::OL_FOLDER_TASKS)],
            )?)?;
            let mut items = to_disp(get_property(&tasks, "Items")?)?;
            if !include_completed {
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
                results.push(task_summary(&item)?);
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
            put_property(&task, "Importance", variant_from_i32(importance_id))?;
            call_method(&task, "Save", &mut [])?;
            Ok(json!({"status": "created", "id": make_id(&task)?, "subject": subject}))
        })
    }

    fn complete_task(&self, task_id: String) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let task = get_item(&ns, &task_id)?;
            call_method(&task, "MarkComplete", &mut [])?;
            let subject = variant_to_string(&get_property(&task, "Subject")?);
            Ok(json!({"status": "completed", "subject": subject}))
        })
    }

    // ---- Notes (Task 16) -----------------------------------------------

    fn list_notes(&self) -> Result<Vec<NoteSummary>, ToolError> {
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
                results.push(note_summary(&item)?);
            }
            Ok(results)
        })
    }

    fn get_note(&self, note_id: String) -> Result<NoteDetail, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let note = get_item(&ns, &note_id)?;
            let summary = note_summary(&note)?;
            Ok(NoteDetail {
                summary,
                body: truncate(&variant_to_string(&get_property(&note, "Body")?)),
            })
        })
    }

    fn create_note(&self, body: String) -> Result<Value, ToolError> {
        // Validate before touching COM (fail-fast, like `create_task`).
        if body.is_empty() {
            return Err(ToolError::new("create_note requires a non-empty body."));
        }
        self.with_com(|| {
            let (app, _ns) = mapi()?;
            let note = to_disp(call_method(
                &app,
                "CreateItem",
                &mut [variant_from_i32(c::OL_NOTE_ITEM)],
            )?)?;
            put_property(&note, "Body", variant_from_str(&body))?;
            call_method(&note, "Save", &mut [])?;
            Ok(json!({"status": "created", "id": make_id(&note)?}))
        })
    }
}
