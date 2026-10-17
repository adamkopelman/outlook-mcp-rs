//! Parameter types shared by several tools.

use serde::Deserialize;

/// A tool parameter that accepts either a single value or a list of values,
/// e.g. `"a@x.com"` or `["a@x.com", "b@x.com"]`.
#[derive(Debug, Clone, PartialEq, Deserialize, schemars::JsonSchema)]
#[serde(untagged)]
pub enum OneOrMany<T> {
    One(T),
    Many(Vec<T>),
}

impl<T> OneOrMany<T> {
    /// True when the caller passed a list (even a one-element list), so the
    /// result should be returned as a list too.
    pub fn is_many(&self) -> bool {
        matches!(self, OneOrMany::Many(_))
    }

    pub fn into_vec(self) -> Vec<T> {
        match self {
            OneOrMany::One(v) => vec![v],
            OneOrMany::Many(v) => v,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn accepts_single_value() {
        let v: OneOrMany<String> = serde_json::from_str("\"a\"").unwrap();
        assert!(!v.is_many());
        assert_eq!(v.into_vec(), vec!["a".to_string()]);
    }

    #[test]
    fn accepts_list() {
        let v: OneOrMany<String> = serde_json::from_str("[\"a\",\"b\"]").unwrap();
        assert!(v.is_many());
        assert_eq!(v.into_vec(), vec!["a".to_string(), "b".to_string()]);
    }

    #[test]
    fn one_element_list_is_still_many() {
        let v: OneOrMany<String> = serde_json::from_str("[\"a\"]").unwrap();
        assert!(v.is_many());
    }
}
