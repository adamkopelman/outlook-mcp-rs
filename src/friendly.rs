//! Convert Outlook enum integers to/from the lowercase friendly words the
//! MCP API exposes, so callers see "accepted" / "busy" rather than 3 / 2.

use crate::constants as c;

pub fn importance_word(v: i32) -> &'static str {
    match v {
        c::OL_IMPORTANCE_LOW => "low",
        c::OL_IMPORTANCE_HIGH => "high",
        _ => "normal",
    }
}

pub fn response_word(v: i32) -> &'static str {
    match v {
        c::OL_RESPONSE_ORGANIZED => "organizer",
        c::OL_RESPONSE_TENTATIVE => "tentative",
        c::OL_RESPONSE_ACCEPTED => "accepted",
        c::OL_RESPONSE_DECLINED => "declined",
        c::OL_RESPONSE_NOT_RESPONDED => "not_responded",
        _ => "none",
    }
}

pub fn busy_status_word(v: i32) -> &'static str {
    match v {
        c::OL_FREE => "free",
        c::OL_TENTATIVE => "tentative",
        c::OL_OUT_OF_OFFICE => "out_of_office",
        c::OL_WORKING_ELSEWHERE => "working_elsewhere",
        _ => "busy",
    }
}

pub fn task_status_word(v: i32) -> &'static str {
    match v {
        c::OL_TASK_IN_PROGRESS => "in_progress",
        c::OL_TASK_COMPLETE => "complete",
        c::OL_TASK_WAITING => "waiting",
        c::OL_TASK_DEFERRED => "deferred",
        _ => "not_started",
    }
}

pub fn busy_status_to_id(name: &str) -> Option<i32> {
    match name.to_lowercase().as_str() {
        "free" => Some(c::OL_FREE),
        "tentative" => Some(c::OL_TENTATIVE),
        "busy" => Some(c::OL_BUSY),
        "out_of_office" => Some(c::OL_OUT_OF_OFFICE),
        "working_elsewhere" => Some(c::OL_WORKING_ELSEWHERE),
        _ => None,
    }
}

pub fn task_status_to_id(name: &str) -> Option<i32> {
    match name.to_lowercase().as_str() {
        "not_started" => Some(c::OL_TASK_NOT_STARTED),
        "in_progress" => Some(c::OL_TASK_IN_PROGRESS),
        "complete" => Some(c::OL_TASK_COMPLETE),
        "waiting" => Some(c::OL_TASK_WAITING),
        "deferred" => Some(c::OL_TASK_DEFERRED),
        _ => None,
    }
}

/// Map an Outlook `MessageClass` to a coarse item type.
pub fn item_type_from_class(class: &str) -> &'static str {
    let c = class.to_ascii_uppercase();
    if c.starts_with("IPM.SCHEDULE.MEETING") {
        "meeting"
    } else if c.contains("NDR") || c.starts_with("REPORT.") && c.contains("NDR") {
        "bounce"
    } else if c.contains("RN") && c.starts_with("REPORT.") {
        "read_receipt"
    } else if c.starts_with("IPM.NOTE") {
        "email"
    } else {
        "other"
    }
}

/// Map a meeting-item `MessageClass` to a meeting type. Updates are delivered
/// with the same class as requests, so they map to "request".
pub fn meeting_type_from_class(class: &str) -> &'static str {
    let c = class.to_ascii_uppercase();
    if c.contains("CANCELED") || c.contains("CANCELLED") {
        "cancellation"
    } else if c.contains("RESP") {
        "response"
    } else {
        "request"
    }
}

#[cfg(test)]
mod class_tests {
    use super::*;

    #[test]
    fn item_type_mapping() {
        assert_eq!(item_type_from_class("IPM.Note"), "email");
        assert_eq!(item_type_from_class("IPM.Schedule.Meeting.Request"), "meeting");
        assert_eq!(item_type_from_class("IPM.Schedule.Meeting.Canceled"), "meeting");
        assert_eq!(item_type_from_class("REPORT.IPM.Note.NDR"), "bounce");
        assert_eq!(item_type_from_class("REPORT.IPM.Note.IPNRN"), "read_receipt");
        assert_eq!(item_type_from_class("IPM.Contact"), "other");
    }

    #[test]
    fn meeting_type_mapping() {
        assert_eq!(meeting_type_from_class("IPM.Schedule.Meeting.Request"), "request");
        assert_eq!(meeting_type_from_class("IPM.Schedule.Meeting.Canceled"), "cancellation");
        assert_eq!(meeting_type_from_class("IPM.Schedule.Meeting.Resp.Pos"), "response");
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn words_map_known_and_unknown_values() {
        assert_eq!(importance_word(2), "high");
        assert_eq!(importance_word(99), "normal"); // unknown → default
        assert_eq!(response_word(3), "accepted");
        assert_eq!(response_word(5), "not_responded");
        assert_eq!(busy_status_word(0), "free");
        assert_eq!(busy_status_word(3), "out_of_office");
        assert_eq!(busy_status_word(99), "busy"); // unknown → default
        assert_eq!(task_status_word(1), "in_progress");
        assert_eq!(task_status_word(99), "not_started");
    }

    #[test]
    fn reverse_lookups_are_case_insensitive_and_reject_garbage() {
        assert_eq!(busy_status_to_id("Out_Of_Office"), Some(3));
        assert_eq!(busy_status_to_id("nope"), None);
        assert_eq!(task_status_to_id("COMPLETE"), Some(2));
        assert_eq!(task_status_to_id("nope"), None);
    }
}
