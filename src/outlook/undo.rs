//! Undo for the `update_*` tools (issue #46). Each update result carries
//! `undo`, the update that puts back what it changed, in the tool's own
//! parameter vocabulary and addressed by the item's current id, so passing
//! a batch's `undo` objects as `updates` rolls the batch back. Changes that
//! can't be reverted that way (bodies, attachments, recurrence, removed
//! attendees, a send) are listed in `not_undoable` instead. Pure, so it's
//! tested without COM; the client reads the old values before writing.

use serde_json::{json, Map, Value};

/// Collects one update's undo while the update is applied.
#[derive(Debug, Default)]
pub struct Undo {
    fields: Map<String, Value>,
    not_undoable: Vec<String>,
}

impl Undo {
    pub fn new() -> Self {
        Self::default()
    }

    /// The undo sets `param` back to `previous`.
    pub fn set(&mut self, param: &str, previous: impl Into<Value>) {
        self.fields.insert(param.to_string(), previous.into());
    }

    /// `param`'s change can't be reverted by another update.
    pub fn not_undoable(&mut self, param: &str) {
        if !self.not_undoable.iter().any(|p| p == param) {
            self.not_undoable.push(param.to_string());
        }
    }

    /// Category changes: the undo removes the categories this update added
    /// and adds back the ones it removed (compared caselessly with `before`).
    pub fn categories(&mut self, before: &[String], add: Option<&[String]>, remove: Option<&[String]>) {
        let has = |list: &[String], c: &str| list.iter().any(|b| b.eq_ignore_ascii_case(c));
        let added: Vec<String> = add
            .unwrap_or(&[])
            .iter()
            .filter(|a| !has(before, a))
            .cloned()
            .collect();
        let removed: Vec<String> = remove
            .unwrap_or(&[])
            .iter()
            .filter_map(|r| before.iter().find(|b| b.eq_ignore_ascii_case(r)).cloned())
            .collect();
        if !added.is_empty() {
            self.set("remove_categories", added);
        }
        if !removed.is_empty() {
            self.set("add_categories", removed);
        }
    }

    /// The result fields: `undo` (with `id_param: id`, or null if nothing can
    /// be undone) and `not_undoable` (omitted when empty). Flags such as
    /// `send_update` in `keep` are copied into a non-empty undo so it is
    /// applied the same way the update was.
    pub fn finish(self, id_param: &str, id: &str, keep: &[(&str, Value)]) -> Map<String, Value> {
        let mut out = Map::new();
        if self.fields.is_empty() {
            out.insert("undo".into(), Value::Null);
        } else {
            let mut undo = Map::new();
            undo.insert(id_param.to_string(), json!(id));
            undo.extend(self.fields);
            for (k, v) in keep {
                undo.insert(k.to_string(), v.clone());
            }
            out.insert("undo".into(), Value::Object(undo));
        }
        if !self.not_undoable.is_empty() {
            out.insert("not_undoable".into(), json!(self.not_undoable));
        }
        out
    }
}

/// Adds `extra`'s fields to the object `result`.
pub fn merged(mut result: Value, extra: Map<String, Value>) -> Value {
    if let Some(o) = result.as_object_mut() {
        o.extend(extra);
    }
    result
}

/// `update_email.flag`'s word for a `MailItem.FlagStatus` value.
pub fn flag_word(status: i32) -> &'static str {
    match status {
        s if s == crate::constants::OL_FLAG_MARKED => "follow_up",
        s if s == crate::constants::OL_FLAG_COMPLETE => "complete",
        _ => "clear",
    }
}

/// `update_email.move_to` for a folder, from its `Folder.FolderPath`
/// (`\\mailbox\Inbox\Sub`): the path below the store root, `Inbox/Sub`.
pub fn move_to_path(folder_path: &str) -> String {
    let trimmed = folder_path.trim_start_matches('\\');
    let below_root = trimmed.split_once('\\').map_or("", |(_, rest)| rest);
    below_root.replace('\\', "/")
}

/// A recipients line (`"A; B"`) as `to` / `cc` / `bcc` take it.
pub fn recipient_list(line: &str) -> Vec<String> {
    line.split(';').map(str::trim).filter(|s| !s.is_empty()).map(String::from).collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn s(v: &[&str]) -> Vec<String> {
        v.iter().map(|x| x.to_string()).collect()
    }

    #[test]
    fn undo_restores_fields_and_carries_the_id_and_flags() {
        let mut u = Undo::new();
        u.set("subject", "Old");
        u.set("start", "2026-10-14T09:00:00");
        u.not_undoable("body");
        u.not_undoable("body");
        let out = u.finish("event_id", "E|S", &[("send_update", json!(false))]);
        assert_eq!(out["undo"], json!({"event_id": "E|S", "subject": "Old",
            "start": "2026-10-14T09:00:00", "send_update": false}));
        assert_eq!(out["not_undoable"], json!(["body"]));
    }

    #[test]
    fn nothing_to_undo_is_null() {
        let out = Undo::new().finish("task_id", "T|S", &[]);
        assert_eq!(out["undo"], Value::Null);
        assert!(!out.contains_key("not_undoable"));
    }

    #[test]
    fn category_undo_reverts_only_what_changed() {
        let mut u = Undo::new();
        let before = s(&["Work", "Blue"]);
        u.categories(&before, Some(&s(&["work", "New"])), Some(&s(&["blue", "Missing"])));
        let out = u.finish("email_id", "M|S", &[]);
        assert_eq!(out["undo"]["remove_categories"], json!(["New"]), "Work was already there");
        assert_eq!(out["undo"]["add_categories"], json!(["Blue"]), "Missing wasn't there");
    }

    #[test]
    fn flag_words_match_update_email() {
        use crate::constants::*;
        assert_eq!(flag_word(OL_FLAG_MARKED), "follow_up");
        assert_eq!(flag_word(OL_FLAG_COMPLETE), "complete");
        assert_eq!(flag_word(OL_NO_FLAG), "clear");
    }

    #[test]
    fn folder_paths_become_move_to_paths() {
        assert_eq!(move_to_path("\\\\me@example.com\\Inbox"), "Inbox");
        assert_eq!(move_to_path("\\\\me@example.com\\Inbox\\Projects\\תיקייה"), "Inbox/Projects/תיקייה");
        assert_eq!(move_to_path("\\\\Mailbox"), "");
    }

    #[test]
    fn recipient_lines_split() {
        assert_eq!(recipient_list("Ada; bob@x.com;  "), s(&["Ada", "bob@x.com"]));
        assert!(recipient_list("").is_empty());
    }

    #[test]
    fn merged_adds_fields() {
        let mut m = Map::new();
        m.insert("undo".into(), Value::Null);
        assert_eq!(merged(json!({"id": "x"}), m), json!({"id": "x", "undo": null}));
    }
}
