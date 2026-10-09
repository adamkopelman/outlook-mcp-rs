//! Pure (COM-free) filter logic shared by the `list_*` tools (issues #30,
//! #32): validating enum filter values, resolving `*_after`/`*_before`
//! ranges, the per-tool `query` field tables, and the client-side matchers
//! for events, tasks and notes.

use chrono::{Datelike, NaiveDateTime, Weekday};

use super::dates::parse_date_param_at;
use super::text_query::TextQuery;
use super::types::{EventSummary, NoteSummary, TaskSummary};
use super::{EventQuery, NoteQuery, TaskQuery};
use crate::error::ToolError;

/// `item_type` values, as `get_email` reports them.
pub const ITEM_TYPES: &[&str] = &["email", "meeting", "bounce", "read_receipt", "other"];
/// `importance` values, as `list_tasks` and `update_email` use them.
pub const IMPORTANCES: &[&str] = &["low", "normal", "high"];
/// Flag states, the same values as `update_email.flag` ("clear" = no flag).
pub const FLAGS: &[&str] = &["follow_up", "complete", "clear"];
/// `show_as` values, as `list_events` reports them.
pub const SHOW_AS: &[&str] = &["free", "tentative", "busy", "out_of_office", "working_elsewhere"];
/// `my_response` values, as `list_events` reports them.
pub const MY_RESPONSES: &[&str] =
    &["organizer", "accepted", "declined", "tentative", "not_responded", "none"];

/// `query` field scopes per tool, and which fields an unscoped term searches.
pub const EMAIL_QUERY_FIELDS: &[&str] = &["subject", "from", "to", "body"];
pub const EMAIL_QUERY_DEFAULTS: &[&str] = &["subject", "from", "body"];
/// DASL columns behind each `list_emails` query field.
pub const EMAIL_QUERY_DASL: &[(&str, &[&str])] = &[
    ("subject", &["urn:schemas:httpmail:subject"]),
    ("from", &["urn:schemas:httpmail:fromname", "urn:schemas:httpmail:fromemail"]),
    ("to", &["urn:schemas:httpmail:displayto", "urn:schemas:httpmail:displaycc"]),
    ("body", &["urn:schemas:httpmail:textdescription"]),
];
pub const EVENT_QUERY_FIELDS: &[&str] = &["subject", "location", "organizer", "attendees"];
pub const EVENT_QUERY_DEFAULTS: &[&str] = &["subject", "location"];
pub const TASK_QUERY_FIELDS: &[&str] = &["subject", "body"];
pub const TASK_QUERY_DEFAULTS: &[&str] = &["subject", "body"];
pub const NOTE_QUERY_FIELDS: &[&str] = &["subject", "body"];
pub const NOTE_QUERY_DEFAULTS: &[&str] = &["body"];

/// The `MailItem.FlagStatus` value for a flag word from [`FLAGS`]
/// (caseless): follow_up = flagged, complete = flag completed, clear = none.
pub fn flag_status_id(flag: &str) -> Option<i32> {
    match flag.to_lowercase().as_str() {
        "follow_up" => Some(crate::constants::OL_FLAG_MARKED),
        "complete" => Some(crate::constants::OL_FLAG_COMPLETE),
        "clear" => Some(crate::constants::OL_NO_FLAG),
        _ => None,
    }
}

/// Trims and lowercases each value, drops empty ones and duplicates, and
/// rejects any value not in `allowed`.
pub fn normalize_choices(values: Vec<String>, field: &str, allowed: &[&str])
    -> Result<Vec<String>, ToolError> {
    let mut out: Vec<String> = Vec::new();
    for v in values {
        let v = v.trim().to_lowercase();
        if v.is_empty() || out.contains(&v) {
            continue;
        }
        if !allowed.contains(&v.as_str()) {
            return Err(ToolError::new(format!(
                "Invalid {field} {v:?}: use one of {}", allowed.join(", ")
            )));
        }
        out.push(v);
    }
    Ok(out)
}

/// Trims free-text filter values and drops empty ones.
pub fn clean_values(values: Vec<String>) -> Vec<String> {
    values.into_iter().map(|v| v.trim().to_string()).filter(|v| !v.is_empty()).collect()
}

/// True when `wants` is empty or `value` equals one of them (caseless).
pub fn one_of(value: &str, wants: &[String]) -> bool {
    wants.is_empty() || wants.iter().any(|w| w.to_lowercase() == value.to_lowercase())
}

/// True when `wants` is empty or any of `categories` equals any of them
/// (caseless, whole name).
pub fn has_category(categories: &[String], wants: &[String]) -> bool {
    wants.is_empty() || categories.iter().any(|c| one_of(c, wants))
}

/// True when `needles` is empty or any needle is a caseless substring of any
/// of `haystacks`. Empty needles never match.
pub fn contains_any(haystacks: &[&str], needles: &[String]) -> bool {
    if needles.is_empty() {
        return true;
    }
    let hay: Vec<String> = haystacks.iter().map(|h| h.to_lowercase()).collect();
    needles
        .iter()
        .filter(|n| !n.is_empty())
        .any(|n| { let n = n.to_lowercase(); hay.iter().any(|h| h.contains(&n)) })
}

/// A resolved `*_after` / `*_before` pair; both bounds are inclusive.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DateRange {
    pub after: Option<NaiveDateTime>,
    pub before: Option<NaiveDateTime>,
}

impl DateRange {
    /// Parses both bounds with the shared date grammar (empty strings count
    /// as absent), relative to `now` and with weeks starting on `week_start`.
    /// A bare ISO date in `before` includes that whole day.
    /// Errors name the parameter, and an `after` later than `before` is
    /// rejected rather than silently matching nothing.
    pub fn parse(after: Option<&str>, before: Option<&str>, after_name: &str, before_name: &str,
        now: NaiveDateTime, week_start: Weekday) -> Result<DateRange, ToolError> {
        let after = after
            .filter(|s| !s.trim().is_empty())
            .map(|s| parse_date_param_at(s, after_name, now, week_start).map(|v| v.at))
            .transpose()?;
        let before = before
            .filter(|s| !s.trim().is_empty())
            .map(|s| parse_date_param_at(s, before_name, now, week_start).map(|v| v.upper_bound()))
            .transpose()?;
        if let (Some(a), Some(b)) = (after, before)
            && a > b
        {
            return Err(ToolError::new(format!(
                "{after_name} ({a}) is later than {before_name} ({b}): nothing can match"
            )));
        }
        Ok(DateRange { after, before })
    }

    pub fn is_set(&self) -> bool {
        self.after.is_some() || self.before.is_some()
    }

    /// True when the range is unset, or `value` (an ISO timestamp as the
    /// summaries report it) lies inside it. A missing or unparsable value
    /// never matches a set range; nor does Outlook's "no date" sentinel
    /// (year 4501, e.g. a task without a due date).
    pub fn contains_iso(&self, value: Option<&str>) -> bool {
        if !self.is_set() {
            return true;
        }
        let Some(dt) = value.and_then(parse_iso_timestamp) else {
            return false;
        };
        if dt.year() >= 4500 {
            return false;
        }
        self.after.is_none_or(|a| dt >= a) && self.before.is_none_or(|b| dt <= b)
    }
}

fn parse_iso_timestamp(s: &str) -> Option<NaiveDateTime> {
    NaiveDateTime::parse_from_str(s, "%Y-%m-%dT%H:%M:%S%.f").ok()
}

/// True if `summary` passes every client-side filter on `q` (everything but
/// the date range, which `Restrict` applies). `text` is `q.query` parsed
/// with [`EVENT_QUERY_FIELDS`]. Attendee needles are caseless substrings of
/// the semicolon-separated `required_attendees`/`optional_attendees`.
pub fn event_matches(summary: &EventSummary, q: &EventQuery, text: &TextQuery) -> bool {
    if !has_category(&summary.categories, &q.category)
        || !one_of(&summary.show_as, &q.show_as)
        || !one_of(&summary.my_response, &q.my_response)
    {
        return false;
    }
    if q.meetings_only && !summary.is_meeting {
        return false;
    }
    if q.all_day.is_some_and(|want| summary.all_day != want) {
        return false;
    }
    // An empty needle never matches (so `[""]` matches nothing, unlike `[]`).
    if !q.attendees.is_empty() {
        // Which attendee tier(s) to search, per attendee_role (default "any").
        let role = q.attendee_role.as_deref().unwrap_or("any").to_lowercase();
        let tiers: Vec<&str> = match role.as_str() {
            "required" => vec![summary.required_attendees.as_str()],
            "optional" => vec![summary.optional_attendees.as_str()],
            _ => vec![summary.required_attendees.as_str(), summary.optional_attendees.as_str()],
        };
        if !contains_any(&tiers, &q.attendees) {
            return false;
        }
    }
    text.matches(EVENT_QUERY_DEFAULTS, |field| match field {
        "subject" => summary.subject.clone(),
        "location" => summary.location.clone(),
        "organizer" => summary.organizer.clone(),
        "attendees" => format!("{}; {}", summary.required_attendees, summary.optional_attendees),
        _ => String::new(),
    })
}

/// True if a task passes every client-side filter on `q`
/// (`include_completed` is a `Restrict`). `due` is `q`'s resolved
/// `due_after`/`due_before`; `text` is `q.query` parsed with
/// [`TASK_QUERY_FIELDS`]. `body` is called only if the query needs it.
pub fn task_matches(summary: &TaskSummary, q: &TaskQuery, due: &DateRange, text: &TextQuery,
    mut body: impl FnMut() -> String) -> bool {
    has_category(&summary.categories, &q.category)
        && one_of(&summary.importance, &q.importance)
        && due.contains_iso(summary.due_date.as_deref())
        && text.matches(TASK_QUERY_DEFAULTS, |field| match field {
            "subject" => summary.subject.clone(),
            "body" => body(),
            _ => String::new(),
        })
}

/// True if a note passes every filter on `q`. `created` is `q`'s resolved
/// `created_after`/`created_before`; `text` is `q.query` parsed with
/// [`NOTE_QUERY_FIELDS`]. `body` is called only if the query needs it.
pub fn note_matches(summary: &NoteSummary, q: &NoteQuery, created: &DateRange, text: &TextQuery,
    mut body: impl FnMut() -> String) -> bool {
    has_category(&summary.categories, &q.category)
        && created.contains_iso(summary.created.as_deref())
        && text.matches(NOTE_QUERY_DEFAULTS, |field| match field {
            "subject" => summary.subject.clone(),
            "body" => body(),
            _ => String::new(),
        })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn s(v: &[&str]) -> Vec<String> {
        v.iter().map(|x| x.to_string()).collect()
    }

    fn now() -> NaiveDateTime {
        NaiveDateTime::parse_from_str("2026-10-08T14:30:00", "%Y-%m-%dT%H:%M:%S").unwrap()
    }

    fn dt(v: &str) -> NaiveDateTime {
        NaiveDateTime::parse_from_str(v, "%Y-%m-%dT%H:%M:%S").unwrap()
    }

    #[test]
    fn normalize_choices_lowercases_dedupes_and_validates() {
        assert_eq!(
            normalize_choices(s(&[" High", "high", "", "LOW"]), "importance", IMPORTANCES).unwrap(),
            s(&["high", "low"])
        );
        let err = normalize_choices(s(&["urgent"]), "importance", IMPORTANCES).unwrap_err();
        assert!(err.to_string().contains("Invalid importance \"urgent\""));
        assert!(err.to_string().contains("low, normal, high"));
        assert!(normalize_choices(vec![], "flag", FLAGS).unwrap().is_empty());
    }

    #[test]
    fn list_helpers_or_their_values() {
        assert!(one_of("busy", &[]));
        assert!(one_of("busy", &s(&["free", "BUSY"])));
        assert!(!one_of("busy", &s(&["free"])));
        assert!(has_category(&s(&["Work", "Red Category"]), &s(&["red category", "x"])));
        assert!(!has_category(&s(&["Work"]), &s(&["Wor"])));
        assert!(has_category(&[], &[]));
        assert!(contains_any(&["Ada Lovelace; Bob"], &s(&["zed", "bob"])));
        assert!(!contains_any(&["Ada"], &s(&["", "zed"])));
        assert!(contains_any(&["דנה כהן"], &s(&["כהן"])));
        assert_eq!(clean_values(s(&[" a ", "", "  "])), s(&["a"]));
    }

    #[test]
    fn date_range_parses_inclusive_bounds() {
        let r = DateRange::parse(Some("2026-06-01"), Some("2026-06-30"), "a", "b", now(), Weekday::Mon).unwrap();
        assert_eq!(r.after, Some(dt("2026-06-01T00:00:00")));
        assert!(r.contains_iso(Some("2026-06-30T23:00:00")));
        assert!(r.contains_iso(Some("2026-06-01T00:00:00")));
        assert!(!r.contains_iso(Some("2026-07-01T00:00:00")));
        assert!(!r.contains_iso(None));
        let r = DateRange::parse(Some("-14d"), None, "a", "b", now(), Weekday::Mon).unwrap();
        assert_eq!(r.after, Some(dt("2026-09-24T14:30:00")));
        assert_eq!(r.before, None);
        let r = DateRange::parse(Some(""), Some("  "), "a", "b", now(), Weekday::Mon).unwrap();
        assert!(!r.is_set());
        assert!(r.contains_iso(None));
    }

    #[test]
    fn date_range_uses_the_given_week_start() {
        // 2026-10-08 is a Thursday.
        let r = DateRange::parse(Some("start_of_week"), Some("end_of_week"), "a", "b", now(), Weekday::Sun).unwrap();
        assert_eq!(r.after, Some(dt("2026-10-04T00:00:00")));
        assert_eq!(r.before, Some(dt("2026-10-10T23:59:59")));
        let r = DateRange::parse(Some("start_of_week"), None, "a", "b", now(), Weekday::Mon).unwrap();
        assert_eq!(r.after, Some(dt("2026-10-05T00:00:00")));
    }

    #[test]
    fn date_range_errors_name_the_parameter() {
        let e = DateRange::parse(Some("soon"), None, "due_after", "due_before", now(), Weekday::Mon).unwrap_err();
        assert!(e.to_string().starts_with("Invalid due_after \"soon\""));
        let e = DateRange::parse(Some("today"), Some("yesterday"), "x_after", "x_before", now(), Weekday::Mon).unwrap_err();
        assert!(e.to_string().contains("x_after"));
        assert!(e.to_string().contains("later than x_before"));
    }

    #[test]
    fn outlook_none_date_never_matches_a_range() {
        let r = DateRange::parse(Some("2026-01-01"), None, "a", "b", now(), Weekday::Mon).unwrap();
        assert!(!r.contains_iso(Some("4501-01-01T00:00:00")));
    }

    fn event() -> EventSummary {
        EventSummary {
            id: "e".into(), subject: "Quarterly planning".into(), start: None, end: None,
            location: "Room 4".into(), organizer: "Ada Lovelace".into(), all_day: false,
            is_recurring: false, is_meeting: true, categories: s(&["Work"]),
            show_as: "busy".into(), my_response: "accepted".into(),
            required_attendees: "Ada Lovelace; Bob Smith".into(), optional_attendees: "Carol".into(),
        }
    }

    fn eq(f: impl FnOnce(&mut EventQuery)) -> (EventQuery, TextQuery) {
        let mut q = EventQuery::default();
        f(&mut q);
        let text = TextQuery::parse(q.query.as_deref().unwrap_or(""), EVENT_QUERY_FIELDS);
        (q, text)
    }

    #[test]
    fn event_matches_list_valued_filters_with_or() {
        let e = event();
        let check = |f: fn(&mut EventQuery)| { let (q, t) = eq(f); event_matches(&e, &q, &t) };
        assert!(check(|_| {}));
        assert!(check(|q| q.show_as = s(&["free", "busy"])));
        assert!(!check(|q| q.show_as = s(&["free"])));
        assert!(check(|q| q.my_response = s(&["declined", "accepted"])));
        assert!(check(|q| q.category = s(&["home", "work"])));
        assert!(!check(|q| q.category = s(&["home"])));
        assert!(check(|q| q.attendees = s(&["nobody", "carol"])));
        assert!(!check(|q| { q.attendees = s(&["carol"]); q.attendee_role = Some("required".into()) }));
        assert!(check(|q| { q.attendees = s(&["bob"]); q.attendee_role = Some("required".into()) }));
        assert!(!check(|q| q.all_day = Some(true)));
        assert!(check(|q| q.meetings_only = true));
    }

    #[test]
    fn event_query_uses_the_shared_syntax() {
        let e = event();
        let check = |f: fn(&mut EventQuery)| { let (q, t) = eq(f); event_matches(&e, &q, &t) };
        assert!(check(|q| q.query = Some("planning room".into())));
        assert!(check(|q| q.query = Some("quart*plan".into())));
        assert!(!check(|q| q.query = Some("lovelace".into()))); // organizer isn't a default field
        assert!(check(|q| q.query = Some("organizer:lovelace".into())));
        assert!(check(|q| q.query = Some("attendees:carol".into())));
        assert!(!check(|q| q.query = Some("\"room planning\"".into())));
    }

    fn task() -> TaskSummary {
        TaskSummary {
            id: "t".into(), subject: "Send report".into(), due_date: Some("2026-10-10T00:00:00".into()),
            complete: false, status: "not_started".into(), importance: "high".into(),
            categories: s(&["Work"]),
        }
    }

    #[test]
    fn task_matches_filters_and_reads_body_lazily() {
        let t = task();
        let none = DateRange::default();
        let mut q = TaskQuery { importance: s(&["low", "high"]), ..Default::default() };
        let text = TextQuery::parse("report", TASK_QUERY_FIELDS);
        let mut reads = 0;
        assert!(task_matches(&t, &q, &none, &text, || { reads += 1; String::new() }));
        assert_eq!(reads, 0); // subject matched first
        let text = TextQuery::parse("body:budget", TASK_QUERY_FIELDS);
        assert!(task_matches(&t, &q, &none, &text, || "the budget".into()));
        assert!(!task_matches(&t, &q, &none, &text, || "nothing".into()));
        q.importance = s(&["low"]);
        assert!(!task_matches(&t, &q, &none, &TextQuery::default(), String::new));
        q.importance.clear();
        let due = DateRange::parse(None, Some("2026-10-09"), "due_after", "due_before", now(), Weekday::Mon).unwrap();
        assert!(!task_matches(&t, &q, &due, &TextQuery::default(), String::new));
        let due = DateRange::parse(None, Some("end_of_week"), "due_after", "due_before", now(), Weekday::Mon).unwrap();
        assert!(task_matches(&t, &q, &due, &TextQuery::default(), String::new));
    }

    #[test]
    fn note_matches_body_by_default() {
        let n = NoteSummary {
            id: "n".into(), subject: "Ideas".into(), created: Some("2026-10-01T09:00:00".into()),
            categories: vec![],
        };
        let q = NoteQuery::default();
        let none = DateRange::default();
        let body = || "Ideas\n- zephyrling".to_string();
        assert!(note_matches(&n, &q, &none, &TextQuery::parse("zephyr*", NOTE_QUERY_FIELDS), body));
        assert!(!note_matches(&n, &q, &none, &TextQuery::parse("subject:zephyrling", NOTE_QUERY_FIELDS), body));
        let created = DateRange::parse(Some("start_of_month"), None, "a", "b", now(), Weekday::Mon).unwrap();
        assert!(note_matches(&n, &q, &created, &TextQuery::default(), body));
        let created = DateRange::parse(Some("today"), None, "a", "b", now(), Weekday::Mon).unwrap();
        assert!(!note_matches(&n, &q, &created, &TextQuery::default(), body));
        let q = NoteQuery { category: s(&["x"]), ..Default::default() };
        assert!(!note_matches(&n, &q, &none, &TextQuery::default(), body));
    }
}

#[cfg(test)]
mod event_filter_tests {
    use super::*;

    /// The pre-#32 two-argument form, parsing `q.query` like `list_events` does.
    fn event_matches(summary: &EventSummary, q: &EventQuery) -> bool {
        let text = TextQuery::parse(q.query.as_deref().unwrap_or(""), EVENT_QUERY_FIELDS);
        super::event_matches(summary, q, &text)
    }

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
            category: vec!["work".to_string()],
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn category_absent() {
        let summary = base();
        let query = EventQuery {
            category: vec!["Personal".to_string()],
            ..Default::default()
        };
        assert!(!event_matches(&summary, &query));
    }

    #[test]
    fn category_multiple_matches_one() {
        let mut summary = base();
        summary.categories = vec!["Work".to_string(), "Meeting".to_string()];
        let query = EventQuery {
            category: vec!["MEETING".to_string()],
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn show_as_match_case_insensitive() {
        let summary = base();
        let query = EventQuery {
            show_as: vec!["BUSY".to_string()],
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn show_as_mismatch() {
        let summary = base();
        let query = EventQuery {
            show_as: vec!["free".to_string()],
            ..Default::default()
        };
        assert!(!event_matches(&summary, &query));
    }

    #[test]
    fn my_response_match_case_insensitive() {
        let summary = base();
        let query = EventQuery {
            my_response: vec!["ACCEPTED".to_string()],
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn my_response_mismatch() {
        let summary = base();
        let query = EventQuery {
            my_response: vec!["declined".to_string()],
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
            attendees: vec!["alice".to_string()],
            attendee_role: Some("required".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn attendees_required_role_no_match() {
        let summary = base();
        let query = EventQuery {
            attendees: vec!["bob".to_string()],
            attendee_role: Some("required".to_string()),
            ..Default::default()
        };
        assert!(!event_matches(&summary, &query));
    }

    #[test]
    fn attendees_optional_role_substring_match() {
        let summary = base();
        let query = EventQuery {
            attendees: vec!["bob".to_string()],
            attendee_role: Some("optional".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn attendees_optional_role_no_match() {
        let summary = base();
        let query = EventQuery {
            attendees: vec!["alice".to_string()],
            attendee_role: Some("optional".to_string()),
            ..Default::default()
        };
        assert!(!event_matches(&summary, &query));
    }

    #[test]
    fn attendees_any_role_matches_required() {
        let summary = base();
        let query = EventQuery {
            attendees: vec!["alice".to_string()],
            attendee_role: Some("any".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn attendees_any_role_matches_optional() {
        let summary = base();
        let query = EventQuery {
            attendees: vec!["bob".to_string()],
            attendee_role: Some("any".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn attendees_no_role_defaults_to_any() {
        let summary = base();
        let query = EventQuery {
            attendees: vec!["alice".to_string()],
            attendee_role: None,
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn attendees_case_insensitive_search() {
        let summary = base();
        let query = EventQuery {
            attendees: vec!["ALICE".to_string()],
            attendee_role: Some("required".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn attendees_multiple_in_list_any_matches() {
        let summary = base();
        let query = EventQuery {
            attendees: vec!["nonexistent".to_string(), "bob".to_string()],
            attendee_role: Some("optional".to_string()),
            ..Default::default()
        };
        assert!(event_matches(&summary, &query));
    }

    #[test]
    fn attendees_empty_string_does_not_match() {
        let summary = base();
        let query = EventQuery {
            attendees: vec!["".to_string()],
            attendee_role: None,
            ..Default::default()
        };
        assert!(!event_matches(&summary, &query));
    }

    #[test]
    fn attendees_empty_list_is_noop() {
        let summary = base();
        let query = EventQuery {
            attendees: vec![],
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
            category: vec!["work".to_string()],
            show_as: vec!["busy".to_string()],
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
            category: vec!["work".to_string()],
            show_as: vec!["free".to_string()],
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
            category: vec!["Work".to_string()],
            attendees: vec!["alice".to_string()],
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
            category: vec!["Work".to_string()],
            attendees: vec!["bob".to_string()],
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
            attendees: vec!["example.com".to_string()],
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

    /// The pre-#32 form: no due range, `q.query` parsed like `list_tasks` does.
    fn task_matches(body: &str, summary: &TaskSummary, q: &TaskQuery) -> bool {
        let text = TextQuery::parse(q.query.as_deref().unwrap_or(""), TASK_QUERY_FIELDS);
        super::task_matches(summary, q, &DateRange::default(), &text, || body.to_string())
    }

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
            category: vec!["Personal".to_string()],
            ..Default::default()
        };
        assert!(!task_matches("budget numbers", &summary, &query));
    }

    #[test]
    fn importance_filter_still_applies_alongside_body_query() {
        let summary = base();
        let query = TaskQuery {
            query: Some("budget".to_string()),
            importance: vec!["high".to_string()],
            ..Default::default()
        };
        assert!(!task_matches("budget numbers", &summary, &query));
    }
}
