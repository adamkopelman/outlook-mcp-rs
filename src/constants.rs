// OlDefaultFolders
pub const OL_FOLDER_DELETED_ITEMS: i32 = 3;
pub const OL_FOLDER_OUTBOX: i32 = 4;
pub const OL_FOLDER_SENT_MAIL: i32 = 5;
pub const OL_FOLDER_INBOX: i32 = 6;
pub const OL_FOLDER_CALENDAR: i32 = 9;
pub const OL_FOLDER_CONTACTS: i32 = 10;
pub const OL_FOLDER_JOURNAL: i32 = 11;
pub const OL_FOLDER_NOTES: i32 = 12;
pub const OL_FOLDER_TASKS: i32 = 13;
pub const OL_FOLDER_DRAFTS: i32 = 16;

// OlItemType (Application.CreateItem)
pub const OL_MAIL_ITEM: i32 = 0;
pub const OL_APPOINTMENT_ITEM: i32 = 1;
pub const OL_TASK_ITEM: i32 = 3;
pub const OL_NOTE_ITEM: i32 = 5;

// OlBodyFormat
pub const OL_FORMAT_PLAIN: i32 = 1;
pub const OL_FORMAT_HTML: i32 = 2;

// OlMeetingResponse (AppointmentItem.Respond)
pub const OL_MEETING_TENTATIVE: i32 = 2;
pub const OL_MEETING_ACCEPTED: i32 = 3;
pub const OL_MEETING_DECLINED: i32 = 4;

// OlMeetingStatus
pub const OL_NONMEETING: i32 = 0;
pub const OL_MEETING: i32 = 1;
pub const OL_MEETING_CANCELED: i32 = 5;

// OlMeetingRecipientType (Recipient.Type on an AppointmentItem)
pub const OL_RECIPIENT_REQUIRED: i32 = 1;
pub const OL_RECIPIENT_OPTIONAL: i32 = 2;

// OlTaskStatus
pub const OL_TASK_NOT_STARTED: i32 = 0;
pub const OL_TASK_IN_PROGRESS: i32 = 1;
pub const OL_TASK_COMPLETE: i32 = 2;
pub const OL_TASK_WAITING: i32 = 3;
pub const OL_TASK_DEFERRED: i32 = 4;

// OlImportance
pub const OL_IMPORTANCE_LOW: i32 = 0;
pub const OL_IMPORTANCE_NORMAL: i32 = 1;
pub const OL_IMPORTANCE_HIGH: i32 = 2;

// OlFlagStatus (MailItem.FlagStatus)
pub const OL_NO_FLAG: i32 = 0;
pub const OL_FLAG_COMPLETE: i32 = 1;
pub const OL_FLAG_MARKED: i32 = 2;

// OlMarkInterval (MailItem.MarkAsTask)
pub const OL_MARK_NO_DATE: i32 = 0;

// OlBusyStatus (AppointmentItem.BusyStatus)
pub const OL_FREE: i32 = 0;
pub const OL_TENTATIVE: i32 = 1;
pub const OL_BUSY: i32 = 2;
pub const OL_OUT_OF_OFFICE: i32 = 3;
pub const OL_WORKING_ELSEWHERE: i32 = 4;

// OlResponseStatus (AppointmentItem.ResponseStatus)
pub const OL_RESPONSE_NONE: i32 = 0;
pub const OL_RESPONSE_ORGANIZED: i32 = 1;
pub const OL_RESPONSE_TENTATIVE: i32 = 2;
pub const OL_RESPONSE_ACCEPTED: i32 = 3;
pub const OL_RESPONSE_DECLINED: i32 = 4;
pub const OL_RESPONSE_NOT_RESPONDED: i32 = 5;

// OlRecurrenceType (RecurrencePattern.RecurrenceType)
pub const OL_RECURS_DAILY: i32 = 0;
pub const OL_RECURS_WEEKLY: i32 = 1;
pub const OL_RECURS_MONTHLY: i32 = 2;
pub const OL_RECURS_MONTH_NTH: i32 = 3;
pub const OL_RECURS_YEARLY: i32 = 5;
pub const OL_RECURS_YEAR_NTH: i32 = 6;

// OlDaysOfWeek (RecurrencePattern.DayOfWeekMask, a bitmask — OR the bits you want)
pub const OL_SUNDAY: i32 = 1;
pub const OL_MONDAY: i32 = 2;
pub const OL_TUESDAY: i32 = 4;
pub const OL_WEDNESDAY: i32 = 8;
pub const OL_THURSDAY: i32 = 16;
pub const OL_FRIDAY: i32 = 32;
pub const OL_SATURDAY: i32 = 64;

// OlNoteColor (NoteItem.Color) — confirmed against Microsoft's official
// enum reference; note the declared order (blue, green, pink, yellow,
// white) is NOT alphabetical.
pub const OL_NOTE_COLOR_BLUE: i32 = 0;
pub const OL_NOTE_COLOR_GREEN: i32 = 1;
pub const OL_NOTE_COLOR_PINK: i32 = 2;
pub const OL_NOTE_COLOR_YELLOW: i32 = 3;
pub const OL_NOTE_COLOR_WHITE: i32 = 4;

pub fn folder_name_to_id(name: &str) -> Option<i32> {
    match name.to_lowercase().as_str() {
        "inbox" => Some(OL_FOLDER_INBOX),
        "sent" | "sent items" => Some(OL_FOLDER_SENT_MAIL),
        "drafts" => Some(OL_FOLDER_DRAFTS),
        "deleted" | "deleted items" | "trash" => Some(OL_FOLDER_DELETED_ITEMS),
        "outbox" => Some(OL_FOLDER_OUTBOX),
        _ => None,
    }
}

pub fn importance_name_to_id(name: &str) -> Option<i32> {
    match name.to_lowercase().as_str() {
        "low" => Some(OL_IMPORTANCE_LOW),
        "normal" => Some(OL_IMPORTANCE_NORMAL),
        "high" => Some(OL_IMPORTANCE_HIGH),
        _ => None,
    }
}

pub fn note_color_to_id(name: &str) -> Option<i32> {
    match name.to_lowercase().as_str() {
        "blue" => Some(OL_NOTE_COLOR_BLUE),
        "green" => Some(OL_NOTE_COLOR_GREEN),
        "pink" => Some(OL_NOTE_COLOR_PINK),
        "yellow" => Some(OL_NOTE_COLOR_YELLOW),
        "white" => Some(OL_NOTE_COLOR_WHITE),
        _ => None,
    }
}

pub fn meeting_response_to_id(name: &str) -> Option<i32> {
    match name.to_lowercase().as_str() {
        "accept" => Some(OL_MEETING_ACCEPTED),
        "decline" => Some(OL_MEETING_DECLINED),
        "tentative" => Some(OL_MEETING_TENTATIVE),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn folder_name_lookup_is_case_insensitive() {
        assert_eq!(folder_name_to_id("Sent Items"), Some(OL_FOLDER_SENT_MAIL));
        assert_eq!(folder_name_to_id("nonexistent"), None);
    }

    #[test]
    fn importance_and_meeting_response_lookups() {
        assert_eq!(importance_name_to_id("HIGH"), Some(OL_IMPORTANCE_HIGH));
        assert_eq!(meeting_response_to_id("Accept"), Some(OL_MEETING_ACCEPTED));
        assert_eq!(meeting_response_to_id("maybe"), None);
    }

    #[test]
    fn note_color_lookup_is_case_insensitive() {
        assert_eq!(note_color_to_id("BLUE"), Some(OL_NOTE_COLOR_BLUE));
        assert_eq!(note_color_to_id("Yellow"), Some(OL_NOTE_COLOR_YELLOW));
        assert_eq!(note_color_to_id("purple"), None);
    }
}
