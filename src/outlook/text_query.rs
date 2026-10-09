//! The `query` syntax shared by every `list_*` tool (issue #32).
//!
//! ```text
//! query  := term (whitespace term)*
//! term   := [field ':'] (word | '"' phrase '"')
//! ```
//!
//! - Every term must match (AND). A term matches when its text occurs in the
//!   field(s) it targets: case-insensitive, after Unicode NFC normalization,
//!   so non-ASCII text (Hebrew, accented Latin, Greek, ...) works.
//! - `"quoted phrase"` keeps its words together, in order, as one term. A
//!   missing closing quote runs to the end of the query.
//! - `*` inside a word or phrase matches any run of characters (including
//!   none), e.g. `status*update`. Terms are substring matches, so a leading
//!   or trailing `*` changes nothing.
//! - `field:term` restricts a term to one field. Each tool defines its own
//!   field names (e.g. `subject`, `from`, `body`); a `name:` prefix that is
//!   not one of them is ordinary text, so `10:30` or `re:` still match
//!   literally. An unscoped term searches the tool's default fields.
//!
//! Pure (no COM) so it can be unit tested: [`TextQuery::matches`] is the
//! client-side matcher and [`TextQuery::to_dasl`] builds the equivalent
//! DASL `@SQL` filter for `list_emails`.

use unicode_normalization::UnicodeNormalization;

/// Normalize text for caseless matching: NFC (so composed and decomposed
/// forms of the same character compare equal), then lowercase.
pub fn fold_text(text: &str) -> String {
    text.nfc().collect::<String>().to_lowercase()
}

/// One parsed term.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Term {
    /// The field it is scoped to (`None` = the tool's default fields).
    pub field: Option<String>,
    /// The text as typed (NFC-normalized, quotes removed, case kept).
    pub text: String,
}

/// A parsed `query`. Empty when the query had no terms (matches everything).
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct TextQuery {
    pub terms: Vec<Term>,
}

impl TextQuery {
    /// Parses `query`, recognizing `name:` scopes only for `fields`
    /// (compared case-insensitively; the stored field name is lowercase).
    pub fn parse(query: &str, fields: &[&str]) -> TextQuery {
        let query: String = query.nfc().collect();
        let mut terms = Vec::new();
        let mut chars = query.chars().peekable();
        loop {
            while chars.next_if(|c| c.is_whitespace()).is_some() {}
            if chars.peek().is_none() {
                break;
            }
            // Read one raw token up to whitespace; quotes may embed spaces.
            let mut field = None;
            let mut text = String::new();
            let mut in_quotes = false;
            while let Some(&c) = chars.peek() {
                if c.is_whitespace() && !in_quotes {
                    break;
                }
                chars.next();
                match c {
                    '"' => in_quotes = !in_quotes,
                    ':' if field.is_none() && !in_quotes && !text.is_empty() => {
                        let name = text.to_lowercase();
                        if fields.iter().any(|f| f.eq_ignore_ascii_case(&name)) {
                            field = Some(name);
                            text.clear();
                        } else {
                            text.push(c);
                        }
                    }
                    _ => text.push(c),
                }
            }
            if !text.is_empty() {
                terms.push(Term { field, text });
            }
        }
        TextQuery { terms }
    }

    pub fn is_empty(&self) -> bool {
        self.terms.is_empty()
    }

    /// True when every term is ASCII (the DASL path is only trusted then;
    /// see `list_emails`'s non-ASCII fallback, issue #2).
    pub fn is_ascii(&self) -> bool {
        self.terms.iter().all(|t| t.text.is_ascii())
    }

    /// True if every term matches. `defaults` are the fields an unscoped
    /// term searches; `value_of(field)` returns a field's text and is called
    /// lazily, at most once per field, so expensive fields (a body) are only
    /// read when an earlier field didn't already match.
    pub fn matches(&self, defaults: &[&str], mut value_of: impl FnMut(&str) -> String) -> bool {
        let mut cache: Vec<(String, String)> = Vec::new();
        let mut folded = |field: &str| -> String {
            if let Some((_, v)) = cache.iter().find(|(f, _)| f == field) {
                return v.clone();
            }
            let v = fold_text(&value_of(field));
            cache.push((field.to_string(), v.clone()));
            v
        };
        self.terms.iter().all(|term| {
            let pattern = fold_text(&term.text);
            match &term.field {
                Some(f) => glob_contains(&folded(f), &pattern),
                None => defaults.iter().any(|f| glob_contains(&folded(f), &pattern)),
            }
        })
    }

    /// The DASL `@SQL=` filter equivalent to [`TextQuery::matches`]: each
    /// term becomes `(col LIKE '%text%' OR ...)` over the DASL columns its
    /// field maps to in `columns`, and the terms are ANDed. `*` becomes
    /// `%`; `'` is doubled. `None` when there are no terms.
    pub fn to_dasl(&self, defaults: &[&str], columns: &[(&str, &[&str])]) -> Option<String> {
        let cols_for = |field: &str| -> Vec<&str> {
            columns.iter().filter(|(f, _)| *f == field).flat_map(|(_, c)| c.iter().copied()).collect()
        };
        let clauses: Vec<String> = self
            .terms
            .iter()
            .map(|term| {
                let like = term.text.replace('\'', "''").replace('*', "%");
                let cols: Vec<&str> = match &term.field {
                    Some(f) => cols_for(f),
                    None => defaults.iter().flat_map(|f| cols_for(f)).collect(),
                };
                let ors: Vec<String> =
                    cols.iter().map(|c| format!("\"{c}\" LIKE '%{like}%'")).collect();
                format!("({})", ors.join(" OR "))
            })
            .collect();
        if clauses.is_empty() {
            None
        } else {
            Some(format!("@SQL={}", clauses.join(" AND ")))
        }
    }
}

/// True if `pattern` (already folded) occurs in `haystack` (already folded),
/// where `*` in `pattern` matches any run of characters. Unanchored: the
/// pieces between `*`s must appear in order, anywhere.
pub fn glob_contains(haystack: &str, pattern: &str) -> bool {
    let mut rest = haystack;
    for piece in pattern.split('*').filter(|p| !p.is_empty()) {
        match rest.find(piece) {
            Some(i) => rest = &rest[i + piece.len()..],
            None => return false,
        }
    }
    true
}

#[cfg(test)]
mod tests {
    use super::*;

    const FIELDS: &[&str] = &["subject", "from", "body"];
    const DEFAULTS: &[&str] = &["subject", "from", "body"];

    fn q(s: &str) -> TextQuery {
        TextQuery::parse(s, FIELDS)
    }

    fn term(field: Option<&str>, text: &str) -> Term {
        Term { field: field.map(str::to_string), text: text.to_string() }
    }

    fn mail<'a>(subject: &'a str, from: &'a str, body: &'a str) -> impl Fn(&str) -> String + 'a {
        move |f| match f {
            "subject" => subject.to_string(),
            "from" => from.to_string(),
            "body" => body.to_string(),
            _ => String::new(),
        }
    }

    #[test]
    fn parses_words_phrases_and_scopes() {
        assert_eq!(q("weekly  update").terms, vec![term(None, "weekly"), term(None, "update")]);
        assert_eq!(q("\"weekly update\" q3").terms, vec![term(None, "weekly update"), term(None, "q3")]);
        assert_eq!(q("subject:report").terms, vec![term(Some("subject"), "report")]);
        assert_eq!(q("SUBJECT:\"team lead\"").terms, vec![term(Some("subject"), "team lead")]);
        assert_eq!(q("from:ada subject:\"status update\" body:q3*").terms, vec![
            term(Some("from"), "ada"), term(Some("subject"), "status update"), term(Some("body"), "q3*"),
        ]);
    }

    #[test]
    fn unknown_scopes_and_colons_are_literal_text() {
        assert_eq!(q("10:30").terms, vec![term(None, "10:30")]);
        assert_eq!(q("re:").terms, vec![term(None, "re:")]);
        assert_eq!(q("cc:bob").terms, vec![term(None, "cc:bob")]);
        assert_eq!(q(":x").terms, vec![term(None, ":x")]);
        assert_eq!(q("subject:a:b").terms, vec![term(Some("subject"), "a:b")]);
    }

    #[test]
    fn empty_and_degenerate_queries() {
        assert!(q("").is_empty());
        assert!(q("   ").is_empty());
        assert!(q("\"\"").is_empty());
        assert!(q("subject:").is_empty());
        // An unterminated quote runs to the end.
        assert_eq!(q("\"weekly upd").terms, vec![term(None, "weekly upd")]);
    }

    #[test]
    fn all_terms_must_match_anywhere_in_default_fields() {
        let m = mail("Weekly update", "Ada Lovelace", "numbers for Q3");
        assert!(q("weekly ada").matches(DEFAULTS, &m));
        assert!(q("WEEKLY q3").matches(DEFAULTS, &m));
        assert!(!q("weekly bob").matches(DEFAULTS, &m));
        assert!(q("").matches(DEFAULTS, &m));
    }

    #[test]
    fn phrases_keep_words_together() {
        let m = mail("Update: weekly status", "", "");
        assert!(q("weekly status").matches(DEFAULTS, &m));
        assert!(q("\"weekly status\"").matches(DEFAULTS, &m));
        assert!(!q("\"status weekly\"").matches(DEFAULTS, &m));
    }

    #[test]
    fn wildcards_match_any_run() {
        let m = mail("Status report for the week", "", "");
        assert!(q("status*week").matches(DEFAULTS, &m));
        assert!(q("\"status*the week\"").matches(DEFAULTS, &m));
        assert!(q("*report*").matches(DEFAULTS, &m));
        assert!(q("*").matches(DEFAULTS, &m));
        assert!(!q("week*status").matches(DEFAULTS, &m));
    }

    #[test]
    fn scoped_terms_only_search_their_field() {
        let m = mail("Budget", "Ada", "please review the budget");
        assert!(q("subject:budget").matches(DEFAULTS, &m));
        assert!(!q("subject:review").matches(DEFAULTS, &m));
        assert!(q("body:review from:ada").matches(DEFAULTS, &m));
        assert!(!q("from:budget").matches(DEFAULTS, &m));
    }

    #[test]
    fn non_ascii_and_normalization() {
        let m = mail("Re: מייל שיקוף שבועי", "דנה", "גוף: סיכום עשייה Q3");
        assert!(q("שיקוף").matches(DEFAULTS, &m));
        assert!(q("\"סיכום עשייה\" q3").matches(DEFAULTS, &m));
        assert!(q("subject:\"מייל שיקוף\"").matches(DEFAULTS, &m));
        assert!(!q("subject:סיכום").matches(DEFAULTS, &m));
        assert!(q("ΣΟΦΊΑ").matches(DEFAULTS, mail("σοφία", "", "")));
        let nfc = "caf\u{e9}";
        let nfd = "cafe\u{301}";
        assert!(q(nfd).matches(DEFAULTS, mail(nfc, "", "")));
        assert!(q(nfc).matches(DEFAULTS, mail(nfd, "", "")));
        assert!(!q("עברית").is_ascii());
        assert!(q("plain words").is_ascii());
    }

    #[test]
    fn fields_are_read_lazily_and_at_most_once() {
        use std::cell::RefCell;
        let reads = RefCell::new(Vec::new());
        let value_of = |f: &str| {
            reads.borrow_mut().push(f.to_string());
            match f {
                "subject" => "alpha beta".to_string(),
                "from" => "x".to_string(),
                _ => "gamma".to_string(),
            }
        };
        assert!(q("alpha beta").matches(DEFAULTS, value_of));
        assert_eq!(*reads.borrow(), vec!["subject"]);
        reads.borrow_mut().clear();
        assert!(q("gamma alpha gamma").matches(DEFAULTS, value_of));
        assert_eq!(*reads.borrow(), vec!["subject", "from", "body"]);
    }

    #[test]
    fn builds_dasl() {
        let cols: &[(&str, &[&str])] = &[
            ("subject", &["urn:subject"]),
            ("from", &["urn:fromname", "urn:fromemail"]),
            ("body", &["urn:body"]),
        ];
        assert_eq!(q("").to_dasl(DEFAULTS, cols), None);
        assert_eq!(
            q("subject:\"it's done\"").to_dasl(DEFAULTS, cols).unwrap(),
            "@SQL=(\"urn:subject\" LIKE '%it''s done%')"
        );
        assert_eq!(
            q("q3*plan from:ada").to_dasl(&["subject", "body"], cols).unwrap(),
            "@SQL=(\"urn:subject\" LIKE '%q3%plan%' OR \"urn:body\" LIKE '%q3%plan%') AND \
             (\"urn:fromname\" LIKE '%ada%' OR \"urn:fromemail\" LIKE '%ada%')"
        );
    }

    #[test]
    fn glob_contains_basics() {
        assert!(glob_contains("abc", ""));
        assert!(glob_contains("abc", "b"));
        assert!(glob_contains("abcabc", "c*a"));
        assert!(!glob_contains("abc", "c*a"));
        assert!(glob_contains("שלום עולם", "של*לם"));
    }
}

#[cfg(test)]
mod text_match_tests {
    use super::*;

    /// The pre-#32 `client.rs` helper: does `query` match any of `fields`?
    fn text_matches(query: &str, fields: &[&str]) -> bool {
        let names: Vec<String> = (0..fields.len()).map(|i| i.to_string()).collect();
        let defaults: Vec<&str> = names.iter().map(String::as_str).collect();
        TextQuery::parse(query, &[])
            .matches(&defaults, |f| fields[f.parse::<usize>().unwrap()].to_string())
    }

    #[test]
    fn hebrew_matches_subject_or_later_field() {
        assert!(text_matches("מייל שיקוף", &["Re: מייל שיקוף שבועי"]));
        assert!(text_matches("סיכום עשייה", &["", "Dana", "גוף: סיכום עשייה Q3"]));
        assert!(!text_matches("מייל שיקוף", &["Weekly report", "Dana"]));
    }

    #[test]
    fn mixed_hebrew_and_english() {
        assert!(text_matches("Q3 סיכום", &["Weekly Q3 סיכום עשייה"]));
        assert!(text_matches("q3 סיכום", &["Weekly Q3 סיכום עשייה"]));
        assert!(!text_matches("Q4 סיכום", &["Weekly Q3 סיכום עשייה"]));
    }

    #[test]
    fn matching_is_caseless() {
        assert!(text_matches("weekly", &["WEEKLY Report"]));
        assert!(text_matches("ÉCOLE", &["notes from école today"]));
        assert!(text_matches("ΣΟΦΊΑ", &["σοφία"]));
    }

    #[test]
    fn matching_ignores_normalization_form() {
        // Latin with an accent: composed (NFC) vs decomposed (NFD).
        let nfc: String = "café".nfc().collect();
        let nfd: String = "café".nfd().collect();
        assert_ne!(nfc, nfd);
        assert!(text_matches(&nfd, &[&format!("subject {nfc}")]));
        assert!(text_matches(&nfc, &[&format!("subject {nfd}")]));
        // Hebrew niqqud has no precomposed forms, but the same marks typed
        // in a different order (shin dot + qamats vs qamats + shin dot) are
        // canonically equivalent and must still match.
        let a = "\u{05E9}\u{05C1}\u{05B8}לום";
        let b = "\u{05E9}\u{05B8}\u{05C1}לום";
        assert_ne!(a, b);
        assert!(text_matches(a, &[&format!("subject {b}")]));
        assert!(text_matches(b, &[&format!("subject {a}")]));
    }

    #[test]
    fn empty_fields_never_match() {
        assert!(!text_matches("שלום", &[]));
        assert!(!text_matches("שלום", &["", ""]));
    }
}
