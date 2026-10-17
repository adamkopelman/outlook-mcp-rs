//! Ids for single occurrences of a recurring series (issue #45). Pure, so
//! it's tested without COM.
//!
//! Every occurrence `list_events` enumerates carries the series master's
//! `EntryID`, so the plain item id (`"{EntryID}|{StoreID}"`) names the whole
//! series. An occurrence id appends the occurrence's start:
//! `"{EntryID}|{StoreID}|{YYYY-MM-DDTHH:MM:SS}"`. The COM client resolves it
//! with `GetRecurrencePattern().GetOccurrence(start)`, falling back to the
//! series' `Exceptions` for an occurrence that was moved. It stays valid
//! until that occurrence is moved again (its start, and so its id, changes;
//! `update_event` and `respond_to_meeting` return the new one).

use chrono::NaiveDateTime;

use crate::error::ToolError;
use crate::outlook::com::make_item_id;

const START_FORMAT: &str = "%Y-%m-%dT%H:%M:%S";

/// A parsed `event_id`: a series (or one-off event), or one occurrence.
#[derive(Debug, Clone, PartialEq)]
pub struct EventId {
    pub entry_id: String,
    pub store_id: String,
    /// The occurrence's start, for an occurrence id.
    pub occurrence_start: Option<NaiveDateTime>,
}

impl EventId {
    /// The id of the series (or one-off event) this id belongs to.
    pub fn series_id(&self) -> String {
        make_item_id(&self.entry_id, &self.store_id)
    }

    pub fn is_occurrence(&self) -> bool {
        self.occurrence_start.is_some()
    }
}

/// The occurrence id for the occurrence of `series_id` starting at `start`
/// (an ISO datetime as `list_events` reports it).
pub fn make_occurrence_id(series_id: &str, start: &str) -> String {
    format!("{series_id}|{start}")
}

/// The occurrence id for `start` as a datetime.
pub fn occurrence_id_at(series_id: &str, start: &NaiveDateTime) -> String {
    make_occurrence_id(series_id, &start.format(START_FORMAT).to_string())
}

/// The occurrence start a datetime is compared by (seconds precision).
pub fn format_start(start: &NaiveDateTime) -> String {
    start.format(START_FORMAT).to_string()
}

/// Parses an event id from any calendar tool: `"{EntryID}|{StoreID}"` or
/// `"{EntryID}|{StoreID}|{start}"`.
pub fn parse_event_id(id: &str) -> Result<EventId, ToolError> {
    let parts: Vec<&str> = id.split('|').collect();
    let invalid = || {
        ToolError::new(format!(
            "Invalid event id {id:?}: expected an id or occurrence_id returned by list_events."
        ))
    };
    if parts.iter().any(|p| p.is_empty()) {
        return Err(invalid());
    }
    match parts.as_slice() {
        [entry, store] => Ok(EventId {
            entry_id: entry.to_string(), store_id: store.to_string(), occurrence_start: None,
        }),
        [entry, store, start] => {
            let start = NaiveDateTime::parse_from_str(start, START_FORMAT).map_err(|_| invalid())?;
            Ok(EventId {
                entry_id: entry.to_string(), store_id: store.to_string(),
                occurrence_start: Some(start),
            })
        }
        _ => Err(invalid()),
    }
}

/// The ids a calendar write returns for the item it changed, read back
/// after the change (accepting a meeting or moving an occurrence changes
/// them): `{"id": <series or event id>}`, plus `"occurrence_id"` when the
/// call acted on one occurrence (`occurrence_start` is its start now).
pub fn current_ids(entry_id: &str, store_id: &str, occurrence_start: Option<&str>) -> serde_json::Value {
    let id = make_item_id(entry_id, store_id);
    match occurrence_start {
        Some(start) => serde_json::json!({"occurrence_id": make_occurrence_id(&id, start), "id": id}),
        None => serde_json::json!({"id": id}),
    }
}

/// The error for an occurrence id that no longer resolves.
pub fn occurrence_not_found(id: &str) -> ToolError {
    ToolError::new(format!(
        "Occurrence not found for {id:?}: it was moved, deleted, or the series changed. \
         Call list_events again for that day and use the new occurrence_id \
         (the series id plus the occurrence's current start)."
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::outlook::com::parse_item_id;

    fn dt(s: &str) -> NaiveDateTime {
        NaiveDateTime::parse_from_str(s, START_FORMAT).unwrap()
    }

    #[test]
    fn series_ids_parse_without_a_start() {
        let id = parse_event_id("E1|S1").unwrap();
        assert_eq!(id, EventId { entry_id: "E1".into(), store_id: "S1".into(), occurrence_start: None });
        assert_eq!(id.series_id(), "E1|S1");
        assert!(!id.is_occurrence());
    }

    #[test]
    fn occurrence_ids_round_trip() {
        let occ = make_occurrence_id("E1|S1", "2026-10-14T09:30:00");
        assert_eq!(occ, "E1|S1|2026-10-14T09:30:00");
        let id = parse_event_id(&occ).unwrap();
        assert_eq!(id.occurrence_start, Some(dt("2026-10-14T09:30:00")));
        assert_eq!(id.series_id(), "E1|S1");
        assert_eq!(occurrence_id_at("E1|S1", &dt("2026-10-14T09:30:00")), occ);
    }

    #[test]
    fn malformed_ids_are_rejected() {
        for bad in ["", "E1", "E1|", "|S1", "E1|S1|", "E1|S1|tomorrow", "E1|S1|2026-10-14", "a|b|2026-10-14T09:30:00|x"] {
            let e = parse_event_id(bad).unwrap_err().to_string();
            assert!(e.contains("Invalid event id"), "{bad}: {e}");
        }
    }

    #[test]
    fn non_calendar_tools_reject_an_occurrence_id() {
        let e = parse_item_id("E1|S1|2026-10-14T09:30:00").unwrap_err().to_string();
        assert!(e.contains("occurrence"), "{e}");
        assert!(parse_item_id("E1|S1").is_ok());
    }

    #[test]
    fn current_ids_name_the_series_and_the_occurrence() {
        assert_eq!(current_ids("E2", "S1", None), serde_json::json!({"id": "E2|S1"}));
        assert_eq!(current_ids("E2", "S1", Some("2026-10-14T11:00:00")),
            serde_json::json!({"id": "E2|S1", "occurrence_id": "E2|S1|2026-10-14T11:00:00"}));
    }

    #[test]
    fn not_found_says_to_relist() {
        assert!(occurrence_not_found("x").to_string().contains("list_events"));
    }
}
