//! Output shaping shared by every `list_*` tool (issue #43): `fields` picks
//! the fields each row carries, `compact` is each tool's preset of the
//! fields needed to triage a list, and `output_dir` writes the rows to a
//! JSON Lines file instead of returning them. No COM, so it's tested anywhere.

use serde_json::{json, Value};

use super::read::{prepare_output_dir, write_output_file};
use crate::error::ToolError;

/// The fields one `list_*` tool's rows can carry, and its `compact` preset.
#[derive(Debug, Clone, Copy)]
pub struct RowSpec {
    /// Every field, in output order.
    pub all: &'static [&'static str],
    /// `compact: true`.
    pub compact: &'static [&'static str],
}

pub const FOLDER_ROWS: RowSpec = RowSpec {
    all: &["name", "path", "items", "unread"],
    compact: &["path", "unread"],
};

pub const EMAIL_ROWS: RowSpec = RowSpec {
    all: &["id", "subject", "sender", "sender_email", "to", "received", "unread",
        "has_attachments", "categories"],
    compact: &["id", "subject", "sender", "received", "unread", "has_attachments"],
};

pub const EVENT_ROWS: RowSpec = RowSpec {
    all: &["id", "occurrence_id", "subject", "start", "end", "location", "organizer", "all_day", "is_recurring",
        "is_meeting", "is_canceled", "categories", "show_as", "my_response", "user_is_required",
        "user_is_optional", "required_count", "optional_count", "required_attendees",
        "optional_attendees"],
    compact: &["id", "occurrence_id", "subject", "start", "end", "organizer", "is_recurring", "my_response",
        "show_as", "user_is_required", "user_is_optional", "required_count", "optional_count"],
};

pub const TASK_ROWS: RowSpec = RowSpec {
    all: &["id", "subject", "due_date", "complete", "status", "importance", "categories"],
    compact: &["id", "subject", "due_date", "status", "importance"],
};

pub const NOTE_ROWS: RowSpec = RowSpec {
    all: &["id", "subject", "created", "categories"],
    compact: &["id", "subject", "created"],
};

pub const ATTACHMENT_ROWS: RowSpec = RowSpec {
    all: &["index", "filename", "size", "type", "content_id", "mime_type", "hidden", "is_inline"],
    compact: &["index", "filename", "size", "is_inline"],
};

/// Event fields that need a walk of each event's `Recipients` (and maybe
/// list expansion); `list_events` skips that work when none is asked for.
pub const USER_ROLE_FIELDS: &[&str] = &["user_is_required", "user_is_optional"];

/// The field set a `list_*` tool returns, from its `fields` / `compact`
/// parameters: `None` means every field (the default). Names are trimmed
/// and lowercased; duplicates are dropped; an unknown name is an error that
/// lists the valid ones. Passing both `fields` and `compact` is an error.
pub fn row_fields(spec: RowSpec, fields: Option<Vec<String>>, compact: bool)
    -> Result<Option<Vec<String>>, ToolError> {
    match (fields, compact) {
        (Some(_), true) => Err(ToolError::new(
            "Pass either fields or compact, not both: compact is a preset field list.",
        )),
        (None, true) => Ok(Some(spec.compact.iter().map(|f| f.to_string()).collect())),
        (None, false) => Ok(None),
        (Some(fields), false) => {
            let mut out: Vec<String> = Vec::new();
            for f in fields {
                let f = f.trim().to_lowercase();
                if !spec.all.contains(&f.as_str()) {
                    return Err(ToolError::new(format!(
                        "Unknown field {f:?} in fields: valid fields are {}.",
                        spec.all.join(", ")
                    )));
                }
                if !out.contains(&f) {
                    out.push(f);
                }
            }
            if out.is_empty() {
                return Err(ToolError::new(format!(
                    "fields must name at least one field: {}.",
                    spec.all.join(", ")
                )));
            }
            Ok(Some(out))
        }
    }
}

/// Whether the rows need any of `wanted` (e.g. [`USER_ROLE_FIELDS`]).
pub fn wants_any(fields: Option<&[String]>, wanted: &[&str]) -> bool {
    fields.is_none_or(|f| f.iter().any(|f| wanted.contains(&f.as_str())))
}

/// Keeps only `fields` in each row object of `rows` (a JSON array, or one
/// object); `None` leaves the rows unchanged.
pub fn select_fields(rows: Value, fields: Option<&[String]>) -> Value {
    let Some(fields) = fields else { return rows };
    match rows {
        Value::Array(items) => Value::Array(
            items.into_iter().map(|row| select_row(row, fields)).collect(),
        ),
        other => select_row(other, fields),
    }
}

fn select_row(row: Value, fields: &[String]) -> Value {
    match row {
        Value::Object(mut map) => {
            map.retain(|k, _| fields.iter().any(|f| f == k));
            Value::Object(map)
        }
        other => other,
    }
}

/// The rows as JSON Lines (one row per line, trailing newline), and how
/// many there are. A non-array value counts as one row.
pub fn to_json_lines(rows: &Value) -> (String, usize) {
    let items: Vec<&Value> = match rows {
        Value::Array(items) => items.iter().collect(),
        other => vec![other],
    };
    let text: String = items.iter().map(|r| format!("{r}\n")).collect();
    (text, items.len())
}

/// `output_dir` on a `list_*` tool: writes the rows to a new
/// `<tool>-<timestamp>.jsonl` in that directory (created if missing) and
/// returns `{"rows_file": <absolute path>, "count": <rows>}`. Each call gets
/// its own file, so paging with `offset` leaves one file per page.
pub fn write_rows_file(output_dir: &str, tool: &str, rows: &Value) -> Result<Value, ToolError> {
    let dir = prepare_output_dir(output_dir)?;
    let (text, count) = to_json_lines(rows);
    let stamp = chrono::Local::now().format("%Y%m%d-%H%M%S-%f");
    let path = write_output_file(&dir.join(format!("{tool}-{stamp}.jsonl")), text.as_bytes())?;
    Ok(json!({"rows_file": path, "count": count}))
}

#[cfg(test)]
mod tests {
    use super::*;

    const ALL: [RowSpec; 6] = [FOLDER_ROWS, EMAIL_ROWS, EVENT_ROWS, TASK_ROWS, NOTE_ROWS, ATTACHMENT_ROWS];

    fn s(v: &[&str]) -> Vec<String> {
        v.iter().map(|x| x.to_string()).collect()
    }

    #[test]
    fn default_is_every_field() {
        for spec in ALL {
            assert_eq!(row_fields(spec, None, false).unwrap(), None);
        }
        assert!(wants_any(None, USER_ROLE_FIELDS));
    }

    #[test]
    fn compact_is_each_tools_preset() {
        for spec in ALL {
            let f = row_fields(spec, None, true).unwrap().unwrap();
            assert_eq!(f, s(spec.compact));
            for c in spec.compact {
                assert!(spec.all.contains(c), "{c}");
            }
            assert!(spec.compact.len() < spec.all.len());
        }
        let f = row_fields(EVENT_ROWS, None, true).unwrap().unwrap();
        assert!(!f.contains(&"required_attendees".to_string()));
        assert!(wants_any(Some(&f), USER_ROLE_FIELDS));
    }

    #[test]
    fn fields_are_normalized_and_validated() {
        let f = row_fields(EVENT_ROWS, Some(s(&[" Subject", "id", "subject"])), false).unwrap().unwrap();
        assert_eq!(f, s(&["subject", "id"]));
        assert!(!wants_any(Some(&f), USER_ROLE_FIELDS));
        let e = row_fields(EMAIL_ROWS, Some(s(&["subject", "body"])), false).unwrap_err().to_string();
        assert!(e.contains("\"body\"") && e.contains("sender_email"), "{e}");
        assert!(row_fields(TASK_ROWS, Some(vec![]), false).is_err());
        assert!(row_fields(NOTE_ROWS, Some(s(&["id"])), true).is_err());
    }

    #[test]
    fn rows_become_json_lines() {
        let (text, n) = to_json_lines(&json!([{"id": "a", "subject": "טקסט"}, {"id": "b"}]));
        assert_eq!(text, "{\"id\":\"a\",\"subject\":\"טקסט\"}\n{\"id\":\"b\"}\n");
        assert_eq!(n, 2);
        assert_eq!(to_json_lines(&json!([])), (String::new(), 0));
        assert_eq!(to_json_lines(&json!({"id": "a"})).1, 1);
    }

    #[test]
    fn write_rows_file_writes_one_file_per_call() {
        let dir = std::env::temp_dir().join(format!("omcp-rows-{}", std::process::id()));
        let rows = json!([{"id": "a"}, {"id": "b"}]);
        let first = write_rows_file(dir.to_str().unwrap(), "list_emails", &rows).unwrap();
        let second = write_rows_file(dir.to_str().unwrap(), "list_emails", &rows).unwrap();
        assert_eq!(first["count"], 2);
        assert_ne!(first["rows_file"], second["rows_file"]);
        let path = first["rows_file"].as_str().unwrap();
        assert!(path.ends_with(".jsonl") && path.contains("list_emails-"), "{path}");
        assert_eq!(std::fs::read_to_string(path).unwrap(), "{\"id\":\"a\"}\n{\"id\":\"b\"}\n");
        std::fs::remove_dir_all(dir).unwrap();
    }

    #[test]
    fn select_fields_keeps_only_the_named_keys() {
        let rows = json!([{"id": "a", "subject": "s", "location": "l"}, {"id": "b"}]);
        assert_eq!(select_fields(rows.clone(), None), rows);
        let f = s(&["id", "subject"]);
        assert_eq!(select_fields(rows, Some(&f)), json!([{"id": "a", "subject": "s"}, {"id": "b"}]));
        assert_eq!(select_fields(json!({"id": "a", "x": 1}), Some(&f)), json!({"id": "a"}));
    }
}
