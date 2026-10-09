use std::sync::Mutex;

use serde_json::{json, Value};

use crate::error::ToolError;
use super::types::*;
use super::filters::DateRange;
use super::{
    require_empty_confirm, validate_recurrence_update, CheckAvailabilityInput, CreateEventInput,
    EmailQuery, EmailUpdate, EventQuery, EventUpdate, NoteQuery, NoteUpdate, OutlookClient,
    TaskQuery, TaskUpdate, draft_update_changes, validate_draft_update, DraftUpdate, NewEmail,
    ReplyInput, parse_importance, prepare_html_images, prepare_mail_body, ValidatedInlineImage,
};
use super::read::{
    output_file_path, prepare_output_dir, replace_cid_references, shape_field, write_output_file,
    BatchResult, BodyOut, ReadOptions, ReadTool,
};

/// The first day of the week the fake's date parsing uses: always Monday,
/// so tests are deterministic. The real client reads it from the Windows
/// user's regional settings (`com::user_first_day_of_week`).
pub const FAKE_WEEK_START: chrono::Weekday = chrono::Weekday::Mon;

pub const EMAIL_ID: &str = "entry-1|store-1";
pub const EVENT_ID: &str = "entry-2|store-1";
pub const TASK_ID: &str = "entry-3|store-1";
pub const NOTE_ID: &str = "entry-4|store-1";
/// An id every batch read reports as not found (a per-item error).
pub const MISSING_ID: &str = "missing|store-1";
/// The one inline image the fake email has (see `get_inline_image` and
/// `get_email`'s `resolve_inline_images`).
pub const FAKE_IMAGE_CID: &str = "logo@example";
pub const FAKE_IMAGE_DATA_URI: &str = "data:image/png;base64,iVBORw==";
/// The bytes `FAKE_IMAGE_DATA_URI` encodes.
pub const FAKE_IMAGE_BYTES: &[u8] = b"\x89PNG";

fn not_found(id: &str) -> ToolError {
    ToolError::new(format!("Item not found: {id}"))
}

/// What a batch read records: the ids plus the resolved options.
fn read_args(ids_key: &str, ids: &[String], opts: &ReadOptions) -> Value {
    json!({
        ids_key: ids, "body": opts.body, "html_body": opts.html_body,
        "attachments": opts.attachments, "meeting": opts.meeting,
        "resolve_inline_images": opts.resolve_inline_images,
        "max_body_chars": opts.max_body_chars, "output_dir": opts.output_dir,
    })
}

/// In-memory stand-in for COM Outlook; records every call. Mirrors
/// `tests/conftest.py::FakeOutlookClient` in the Python project.
pub struct FakeOutlookClient {
    calls: Mutex<Vec<(String, Value)>>,
    fail_with: Mutex<Option<String>>,
    email_text: Mutex<Option<EmailText>>,
    email_html: Mutex<Option<String>>,
}

/// Custom subject/sender/body returned by `list_emails` and `get_email`
/// (see `FakeOutlookClient::set_email_text`). Lets tests feed non-ASCII
/// text through the tool layer.
#[derive(Clone)]
struct EmailText {
    subject: String,
    sender: String,
    body: String,
}

impl FakeOutlookClient {
    pub fn new() -> Self {
        Self {
            calls: Mutex::new(Vec::new()),
            fail_with: Mutex::new(None),
            email_text: Mutex::new(None),
            email_html: Mutex::new(None),
        }
    }

    /// Make `get_email` return this HTML body instead of `<p>Hi there</p>`.
    pub fn set_email_html(&self, html: impl Into<String>) {
        *self.email_html.lock().unwrap() = Some(html.into());
    }

    /// Make `list_emails` and `get_email` return this subject/sender/body
    /// instead of their canned values.
    pub fn set_email_text(&self, subject: impl Into<String>, sender: impl Into<String>,
        body: impl Into<String>) {
        *self.email_text.lock().unwrap() = Some(EmailText {
            subject: subject.into(), sender: sender.into(), body: body.into(),
        });
    }

    /// The custom email text if one was set, else the given canned defaults.
    fn email_text(&self, subject: &str, sender: &str, body: &str) -> EmailText {
        self.email_text.lock().unwrap().clone().unwrap_or_else(|| EmailText {
            subject: subject.into(), sender: sender.into(), body: body.into(),
        })
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

/// The Content-IDs of every inline image a mail tool would attach.
fn content_ids(images: &[ValidatedInlineImage]) -> Vec<String> {
    images.iter().map(|i| i.content_id.clone()).collect()
}

/// Validate a [`NewEmail`] like the real client does (body, inline images,
/// `data:` URIs, importance) and build the recorded arguments: `body` is the
/// final body (data URIs rewritten to `cid:`), `html` whether it's HTML, and
/// `inline_content_ids` every inline image that would be attached.
fn new_email_args(m: &NewEmail) -> Result<Value, ToolError> {
    parse_importance(m.importance.as_deref())?;
    let (body, images) = prepare_mail_body(&m.body, m.inline_images.as_deref().unwrap_or(&[]))?;
    Ok(json!({
        "to": m.to, "subject": m.subject, "body": body.as_str(), "html": body.is_html(),
        "cc": m.cc, "bcc": m.bcc, "attachments": m.attachments,
        "inline_images": m.inline_images, "inline_content_ids": content_ids(&images),
        "categories": m.categories, "importance": m.importance,
    }))
}

impl OutlookClient for FakeOutlookClient {
    fn list_folders(&self) -> Result<Vec<FolderInfo>, ToolError> {
        self.record("list_folders", json!({}))?;
        Ok(vec![FolderInfo {
            name: "Inbox".into(), path: "Inbox".into(), items: 2, unread: 1,
        }])
    }

    fn list_emails(&self, q: EmailQuery) -> Result<Vec<EmailSummary>, ToolError> {
        // Same up-front date validation as the real client.
        DateRange::parse(q.received_after.as_deref(), q.received_before.as_deref(),
            "received_after", "received_before", chrono::Local::now().naive_local(), FAKE_WEEK_START)?;
        self.record("list_emails", json!({
            "query": q.query, "folder": q.folder, "count": q.count, "offset": q.offset,
            "unread_only": q.unread_only, "from": q.from, "to": q.to, "category": q.category,
            "received_after": q.received_after, "received_before": q.received_before,
            "has_attachments": q.has_attachments, "item_type": q.item_type,
            "importance": q.importance, "flag": q.flag,
        }))?;
        let text = self.email_text("Hello", "Ada", "");
        Ok(vec![EmailSummary {
            id: EMAIL_ID.into(), subject: text.subject, sender: text.sender,
            sender_email: "".into(), to: "".into(), received: None,
            unread: true, has_attachments: false,
            categories: vec!["Work".to_string()],
        }])
    }

    fn get_email(&self, email_ids: Vec<String>, opts: &ReadOptions) -> BatchResult<EmailDetail> {
        self.record("get_email", read_args("email_ids", &email_ids, opts))?;
        let out_dir = opts.output_dir.as_deref().map(prepare_output_dir).transpose()?;
        let text = self.email_text("Hello", "", "Hi there");
        let html_source =
            self.email_html.lock().unwrap().clone().unwrap_or_else(|| "<p>Hi there</p>".into());
        Ok(email_ids.into_iter().map(|id| {
            if id == MISSING_ID {
                return Err(not_found(&id));
            }
            let shape = |full: &str, field: &str, ext: &str| {
                shape_field(full, ReadTool::Email, &id, field, ext, opts, out_dir.as_deref())
            };
            let body = opts.body.then(|| shape(&text.body, "body", ".txt")).transpose()?;
            let (mut inline_images_resolved, mut inline_images_unresolved) = (None, None);
            let html = if opts.html_body {
                let mut full = html_source.clone();
                if opts.resolve_inline_images {
                    let images = [(FAKE_IMAGE_CID.to_string(), FAKE_IMAGE_DATA_URI.to_string())];
                    let rewrite = replace_cid_references(&full, &images);
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
            Ok(EmailDetail {
                summary: EmailSummary {
                    id: id.clone(), subject: text.subject.clone(), sender: text.sender.clone(),
                    sender_email: "".into(), to: "".into(), received: None,
                    unread: false, has_attachments: false, categories: vec![],
                },
                cc: "".into(), bcc: "".into(),
                body, body_file, body_truncated, body_length,
                html_body, html_body_file, html_truncated, html_length,
                inline_images_resolved, inline_images_unresolved,
                attachments: opts.attachments.then(Vec::new),
                item_type: "email".to_string(),
                is_meeting: false,
                meeting: None,
            })
        }).collect())
    }

    fn send_email(&self, m: NewEmail) -> Result<Value, ToolError> {
        let args = new_email_args(&m)?;
        self.record("send_email", args)?;
        Ok(json!({"status": "sent", "to": m.to.join("; "), "subject": m.subject}))
    }

    fn create_draft(&self, m: NewEmail) -> Result<Value, ToolError> {
        let args = new_email_args(&m)?;
        self.record("create_draft", args)?;
        Ok(json!({"status": "draft_saved", "id": EMAIL_ID, "subject": m.subject}))
    }

    fn reply_email(&self, r: ReplyInput) -> Result<Value, ToolError> {
        // Same up-front validation as the real client.
        parse_importance(r.importance.as_deref())?;
        let (body, images) = prepare_mail_body(&r.body, r.inline_images.as_deref().unwrap_or(&[]))?;
        self.record("reply_email", json!({
            "email_id": r.email_id, "body": body.as_str(), "html": body.is_html(),
            "reply_all": r.reply_all, "send": r.send, "attachments": r.attachments,
            "inline_images": r.inline_images, "inline_content_ids": content_ids(&images),
            "categories": r.categories, "importance": r.importance,
        }))?;
        Ok(json!({"status": if r.send { "sent" } else { "draft_saved" }}))
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

    fn update_draft(&self, u: DraftUpdate) -> Result<Value, ToolError> {
        // Same up-front validation as the real client.
        validate_draft_update(&u)?;
        let (html_body, images) = match &u.html_body {
            Some(html) => {
                let (html, images) = prepare_html_images(html, u.inline_images.as_deref().unwrap_or(&[]))?;
                (Some(html), images)
            }
            None => (None, Vec::new()),
        };
        self.record("update_draft", json!({
            "email_id": u.email_id, "subject": u.subject, "body": u.body,
            "html_body": html_body, "to": u.to, "cc": u.cc, "bcc": u.bcc,
            "attachments": u.attachments, "inline_images": u.inline_images,
            "inline_content_ids": content_ids(&images),
            "add_categories": u.add_categories, "remove_categories": u.remove_categories,
            "importance": u.importance, "send": u.send,
        }))?;
        let changed = draft_update_changes(&u);
        if u.send {
            return Ok(json!({"status": "sent", "subject": u.subject.unwrap_or_default(), "changed": changed}));
        }
        Ok(json!({"status": "draft_updated", "id": u.email_id, "changed": changed}))
    }

    fn delete_email(&self, email_id: String, permanent: bool) -> Result<Value, ToolError> {
        self.record("delete_email", json!({"email_id": email_id, "permanent": permanent}))?;
        Ok(json!({"status": "deleted", "permanent": permanent}))
    }

    fn empty_deleted_items(&self, confirm: bool) -> Result<Value, ToolError> {
        // Record first so tests can see the call even when it's refused.
        self.record("empty_deleted_items", json!({"confirm": confirm}))?;
        require_empty_confirm(confirm)?;
        Ok(json!({"status": "emptied", "items_deleted": 2, "folders_deleted": 1, "failed": 0}))
    }

    fn list_events(&self, q: EventQuery) -> Result<Vec<EventSummary>, ToolError> {
        DateRange::parse(q.start_after.as_deref(), q.start_before.as_deref(),
            "start_after", "start_before", chrono::Local::now().naive_local(), FAKE_WEEK_START)?;
        self.record("list_events", json!({
            "start_after": q.start_after, "start_before": q.start_before, "query": q.query,
            "category": q.category, "show_as": q.show_as, "my_response": q.my_response,
            "attendees": q.attendees, "attendee_role": q.attendee_role,
            "meetings_only": q.meetings_only, "all_day": q.all_day,
            "calendar_of": q.calendar_of, "count": q.count, "offset": q.offset,
        }))?;
        Ok(vec![EventSummary {
            id: EVENT_ID.into(), subject: "Standup".into(), start: None, end: None,
            location: "".into(), organizer: "".into(), all_day: false,
            is_recurring: false, is_meeting: false, categories: vec![],
            show_as: "busy".into(), my_response: "accepted".into(),
            required_attendees: "".into(), optional_attendees: "".into(),
        }])
    }

    fn get_event(&self, event_ids: Vec<String>, opts: &ReadOptions) -> BatchResult<EventDetail> {
        self.record("get_event", read_args("event_ids", &event_ids, opts))?;
        let out_dir = opts.output_dir.as_deref().map(prepare_output_dir).transpose()?;
        Ok(event_ids.into_iter().map(|id| {
            if id == MISSING_ID {
                return Err(not_found(&id));
            }
            let body = opts.body
                .then(|| shape_field("", ReadTool::Event, &id, "body", ".txt", opts, out_dir.as_deref()))
                .transpose()?;
            let (body, body_file, body_truncated, body_length) = BodyOut::into_parts(body);
            Ok(EventDetail {
                summary: EventSummary {
                    id, subject: "Standup".into(), start: None, end: None,
                    location: "".into(), organizer: "".into(), all_day: false,
                    is_recurring: false, is_meeting: false, categories: vec![],
                    show_as: "busy".into(), my_response: "accepted".into(),
                    required_attendees: "".into(), optional_attendees: "".into(),
                },
                body, body_file, body_truncated, body_length,
                recurrence: None,
            })
        }).collect())
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
            "recurrence": input.recurrence.as_ref().map(|r| json!({
                "pattern": r.pattern, "interval": r.interval, "days_of_week": r.days_of_week,
                "day_of_month": r.day_of_month, "until": r.until, "occurrences": r.occurrences,
            })),
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

    fn update_event(&self, u: EventUpdate) -> Result<Value, ToolError> {
        validate_recurrence_update(&u)?;
        self.record("update_event", json!({
            "event_id": u.event_id, "subject": u.subject, "start": u.start, "end": u.end,
            "location": u.location, "body": u.body, "all_day": u.all_day,
            "reminder_minutes": u.reminder_minutes, "show_as": u.show_as,
            "add_categories": u.add_categories, "remove_categories": u.remove_categories,
            "add_required_attendees": u.add_required_attendees,
            "add_optional_attendees": u.add_optional_attendees,
            "remove_attendees": u.remove_attendees, "send_update": u.send_update,
            "recurrence": u.recurrence.as_ref().map(|r| json!({
                "pattern": r.pattern, "interval": r.interval, "days_of_week": r.days_of_week,
                "day_of_month": r.day_of_month, "until": r.until, "occurrences": r.occurrences,
            })),
            "clear_recurrence": u.clear_recurrence,
        }))?;
        let mut changed: Vec<&str> = Vec::new();
        if u.subject.is_some() { changed.push("subject"); }
        if u.start.is_some() { changed.push("start"); }
        if u.end.is_some() { changed.push("end"); }
        if u.location.is_some() { changed.push("location"); }
        if u.body.is_some() { changed.push("body"); }
        if u.all_day.is_some() { changed.push("all_day"); }
        if u.reminder_minutes.is_some() { changed.push("reminder_minutes"); }
        if u.show_as.is_some() { changed.push("show_as"); }
        if u.add_categories.is_some() { changed.push("add_categories"); }
        if u.remove_categories.is_some() { changed.push("remove_categories"); }
        if u.add_required_attendees.is_some() { changed.push("add_required_attendees"); }
        if u.add_optional_attendees.is_some() { changed.push("add_optional_attendees"); }
        if u.remove_attendees.is_some() { changed.push("remove_attendees"); }
        if u.recurrence.is_some() { changed.push("recurrence"); }
        if u.clear_recurrence { changed.push("clear_recurrence"); }
        Ok(json!({"status": "updated", "id": u.event_id, "changed": changed}))
    }

    fn delete_event(&self, event_id: String, send_cancellation: bool) -> Result<Value, ToolError> {
        self.record("delete_event", json!({"event_id": event_id, "send_cancellation": send_cancellation}))?;
        Ok(json!({"status": "deleted", "note": "Moved to Deleted Items."}))
    }

    fn check_availability(&self, input: CheckAvailabilityInput) -> Result<AvailabilityResult, ToolError> {
        self.record("check_availability", json!({
            "people": input.people, "start": input.start, "end": input.end,
            "interval_minutes": input.interval_minutes, "treat_as_free": input.treat_as_free,
        }))?;
        // Deterministic fake: every person resolves and is free for the
        // whole requested window (one slot spanning [start, end)), so
        // common_free tests can assert a single window without needing
        // real FreeBusy parsing in the fake.
        let people = input.people.iter().map(|p| PersonAvailability {
            person: p.clone(),
            resolved: true,
            slots: vec![AvailabilitySlot {
                start: input.start.clone(),
                end: input.end.clone(),
                status: "free".to_string(),
            }],
        }).collect::<Vec<_>>();
        let common_free = if people.is_empty() {
            Vec::new()
        } else {
            vec![FreeWindow { start: input.start.clone(), end: input.end.clone() }]
        };
        Ok(AvailabilityResult { people, common_free })
    }

    fn list_attachments(&self, email_ids: Vec<String>) -> BatchResult<Vec<AttachmentInfo>> {
        self.record("list_attachments", json!({"email_ids": email_ids}))?;
        Ok(email_ids.into_iter().map(|id| {
            if id == MISSING_ID {
                return Err(not_found(&id));
            }
            Ok(vec![
                AttachmentInfo {
                    index: 1, filename: "report.pdf".into(), size: 1234, att_type: "file".into(),
                    content_id: None, mime_type: Some("application/pdf".into()), hidden: false,
                    is_inline: false,
                },
                AttachmentInfo {
                    index: 2, filename: "logo.png".into(), size: 512, att_type: "file".into(),
                    content_id: Some(FAKE_IMAGE_CID.into()), mime_type: Some("image/png".into()), hidden: true,
                    is_inline: true,
                },
            ])
        }).collect())
    }

    fn save_attachments(&self, email_id: String, save_dir: String,
        attachment_names: Option<Vec<String>>, inline: Option<bool>) -> Result<Vec<Value>, ToolError> {
        self.record("save_attachments", json!({
            "email_id": email_id, "save_dir": save_dir, "attachment_names": attachment_names,
            "inline": inline,
        }))?;
        // Same two attachments as list_attachments: a regular file and an
        // inline image. Only `inline` filters here (names are just recorded).
        let all = vec![
            json!({
                "index": 1, "filename": "report.pdf", "size": 1234, "type": "file",
                "content_id": null, "mime_type": "application/pdf", "hidden": false,
                "is_inline": false, "saved_to": save_dir, "status": "saved",
            }),
            json!({
                "index": 2, "filename": "logo.png", "size": 512, "type": "file",
                "content_id": "logo@example", "mime_type": "image/png", "hidden": true,
                "is_inline": true, "saved_to": save_dir, "status": "saved",
            }),
        ];
        let results: Vec<Value> = all.into_iter()
            .filter(|a| inline.is_none_or(|want| a["is_inline"] == want))
            .collect();
        if results.is_empty() {
            return Err(ToolError::new("No attachments matched attachment_names / inline."));
        }
        Ok(results)
    }

    fn get_inline_image(&self, email_id: String, content_ids: Vec<String>,
        context_lines: Option<u32>, output_dir: Option<String>) -> BatchResult<InlineImageData> {
        self.record("get_inline_image", json!({
            "email_id": email_id, "content_ids": content_ids, "context_lines": context_lines,
            "output_dir": output_dir,
        }))?;
        if email_id == MISSING_ID {
            return Err(not_found(&email_id));
        }
        let out_dir = output_dir.as_deref().map(prepare_output_dir).transpose()?;
        Ok(content_ids.into_iter().map(|raw| {
            let wanted = super::normalize_content_id(&raw);
            if !wanted.eq_ignore_ascii_case(FAKE_IMAGE_CID) {
                return Err(ToolError::new(format!(
                    "Content-ID '{wanted}' not found. Available Content-IDs: {FAKE_IMAGE_CID}"
                )));
            }
            let (data_uri, data_file) = match out_dir.as_deref() {
                Some(dir) => {
                    let key = format!("{email_id}\n{}", wanted.to_lowercase());
                    let path = output_file_path(dir, "image", &key, "data", ".png");
                    (None, Some(write_output_file(&path, FAKE_IMAGE_BYTES)?))
                }
                None => (Some(FAKE_IMAGE_DATA_URI.to_string()), None),
            };
            Ok(InlineImageData {
                content_id: FAKE_IMAGE_CID.into(), filename: "logo.png".into(),
                mime_type: "image/png".into(), size: FAKE_IMAGE_BYTES.len(),
                data_uri, data_file,
                context: context_lines.map(|_| "Here is our new logo:".to_string()),
            })
        }).collect())
    }

    fn list_tasks(&self, q: TaskQuery) -> Result<Vec<TaskSummary>, ToolError> {
        DateRange::parse(q.due_after.as_deref(), q.due_before.as_deref(),
            "due_after", "due_before", chrono::Local::now().naive_local(), FAKE_WEEK_START)?;
        self.record("list_tasks", json!({
            "include_completed": q.include_completed, "category": q.category,
            "importance": q.importance, "query": q.query, "due_after": q.due_after,
            "due_before": q.due_before, "count": q.count, "offset": q.offset,
        }))?;
        Ok(vec![TaskSummary {
            id: TASK_ID.into(), subject: "Buy milk".into(), due_date: None,
            complete: false, status: "not_started".to_string(), importance: "normal".to_string(), categories: vec![],
        }])
    }

    fn create_task(&self, subject: String, body: Option<String>,
        due_date: Option<String>, importance: String, categories: Option<Vec<String>>,
        start_date: Option<String>, reminder_time: Option<String>) -> Result<Value, ToolError> {
        self.record("create_task", json!({
            "subject": subject, "body": body, "due_date": due_date, "importance": importance,
            "categories": categories, "start_date": start_date, "reminder_time": reminder_time,
        }))?;
        Ok(json!({"status": "created", "id": TASK_ID, "subject": subject}))
    }

    fn get_task(&self, task_ids: Vec<String>, opts: &ReadOptions) -> BatchResult<TaskDetail> {
        self.record("get_task", read_args("task_ids", &task_ids, opts))?;
        let out_dir = opts.output_dir.as_deref().map(prepare_output_dir).transpose()?;
        Ok(task_ids.into_iter().map(|id| {
            if id == MISSING_ID {
                return Err(not_found(&id));
            }
            let body = opts.body
                .then(|| shape_field("2 litres, semi-skimmed", ReadTool::Task, &id, "body", ".txt",
                    opts, out_dir.as_deref()))
                .transpose()?;
            let (body, body_file, body_truncated, body_length) = BodyOut::into_parts(body);
            Ok(TaskDetail {
                summary: TaskSummary {
                    id, subject: "Buy milk".into(), due_date: None, complete: false,
                    status: "not_started".to_string(), importance: "normal".to_string(),
                    categories: vec![],
                },
                body, body_file, body_truncated, body_length,
                start_date: None, date_completed: None, percent_complete: 0,
                reminder_set: false, reminder_time: None, created: None, modified: None,
            })
        }).collect())
    }

    fn update_task(&self, u: TaskUpdate) -> Result<Value, ToolError> {
        self.record("update_task", json!({
            "task_id": u.task_id, "mark_complete": u.mark_complete, "subject": u.subject,
            "body": u.body, "due_date": u.due_date, "start_date": u.start_date,
            "importance": u.importance, "add_categories": u.add_categories,
            "remove_categories": u.remove_categories, "percent_complete": u.percent_complete,
            "reminder_time": u.reminder_time,
        }))?;
        let mut changed: Vec<&str> = Vec::new();
        if u.mark_complete.is_some() { changed.push("mark_complete"); }
        if u.subject.is_some() { changed.push("subject"); }
        if u.body.is_some() { changed.push("body"); }
        if u.due_date.is_some() { changed.push("due_date"); }
        if u.start_date.is_some() { changed.push("start_date"); }
        if u.importance.is_some() { changed.push("importance"); }
        if u.add_categories.is_some() { changed.push("add_categories"); }
        if u.remove_categories.is_some() { changed.push("remove_categories"); }
        if u.percent_complete.is_some() { changed.push("percent_complete"); }
        if u.reminder_time.is_some() { changed.push("reminder_time"); }
        Ok(json!({"status": "updated", "id": u.task_id, "changed": changed}))
    }

    fn delete_task(&self, task_id: String) -> Result<Value, ToolError> {
        self.record("delete_task", json!({"task_id": task_id}))?;
        Ok(json!({"status": "deleted", "note": "Moved to Deleted Items."}))
    }

    fn list_notes(&self, q: NoteQuery) -> Result<Vec<NoteSummary>, ToolError> {
        DateRange::parse(q.created_after.as_deref(), q.created_before.as_deref(),
            "created_after", "created_before", chrono::Local::now().naive_local(), FAKE_WEEK_START)?;
        self.record("list_notes", json!({
            "category": q.category, "query": q.query, "created_after": q.created_after,
            "created_before": q.created_before, "count": q.count, "offset": q.offset,
        }))?;
        Ok(vec![NoteSummary { id: NOTE_ID.into(), subject: "Ideas".into(), created: None, categories: vec![] }])
    }

    fn get_note(&self, note_ids: Vec<String>, opts: &ReadOptions) -> BatchResult<NoteDetail> {
        self.record("get_note", read_args("note_ids", &note_ids, opts))?;
        let out_dir = opts.output_dir.as_deref().map(prepare_output_dir).transpose()?;
        Ok(note_ids.into_iter().map(|id| {
            if id == MISSING_ID {
                return Err(not_found(&id));
            }
            let body = opts.body
                .then(|| shape_field("Ideas\n- one", ReadTool::Note, &id, "body", ".txt", opts,
                    out_dir.as_deref()))
                .transpose()?;
            let (body, body_file, body_truncated, body_length) = BodyOut::into_parts(body);
            Ok(NoteDetail {
                summary: NoteSummary { id, subject: "Ideas".into(), created: None, categories: vec![] },
                body, body_file, body_truncated, body_length,
                modified: None,
            })
        }).collect())
    }

    fn create_note(&self, body: String, categories: Option<Vec<String>>, color: Option<String>) -> Result<Value, ToolError> {
        self.record("create_note", json!({"body": body, "categories": categories, "color": color}))?;
        Ok(json!({"status": "created", "id": NOTE_ID}))
    }

    fn update_note(&self, u: NoteUpdate) -> Result<Value, ToolError> {
        self.record("update_note", json!({
            "note_id": u.note_id, "body": u.body, "add_categories": u.add_categories,
            "remove_categories": u.remove_categories, "color": u.color,
        }))?;
        let mut changed: Vec<&str> = Vec::new();
        if u.body.is_some() { changed.push("body"); }
        if u.add_categories.is_some() { changed.push("add_categories"); }
        if u.remove_categories.is_some() { changed.push("remove_categories"); }
        if u.color.is_some() { changed.push("color"); }
        Ok(json!({"status": "updated", "id": u.note_id, "changed": changed}))
    }

    fn delete_note(&self, note_id: String) -> Result<Value, ToolError> {
        self.record("delete_note", json!({"note_id": note_id}))?;
        Ok(json!({"status": "deleted", "note": "Moved to Deleted Items."}))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn basic_query() -> EmailQuery {
        EmailQuery::default()
    }

    #[test]
    fn records_calls_in_order() {
        let fake = FakeOutlookClient::new();
        fake.list_folders().unwrap();
        fake.list_emails(basic_query()).unwrap();
        assert_eq!(fake.calls(), vec![
            ("list_folders".to_string(), json!({})),
            ("list_emails".to_string(), json!({
                "query": null, "folder": "inbox", "count": 10, "offset": 0, "unread_only": false,
                "from": [], "to": [], "category": [], "received_after": null,
                "received_before": null, "has_attachments": null,
                "item_type": [], "importance": [], "flag": [],
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
