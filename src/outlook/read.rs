//! Output control shared by the read tools (`get_email`, `get_event`,
//! `get_note`, `get_task`, and the file output of `get_inline_image`):
//! which optional fields come back (`include`), the body format, the body
//! size cap (`max_body_chars`), writing large fields to files
//! (`output_dir`), and rewriting `cid:` inline-image references to `data:`
//! URIs (`resolve_inline_images`). Everything here is plain logic (plus the
//! file writes for `output_dir`) shared by the real COM client and the fake.

use std::path::{Path, PathBuf};

use crate::error::ToolError;

use super::{extension_for_mime, is_cid_char, normalize_content_id};

/// The default body cut (characters) for every body a read tool returns.
/// Matches `MAX_BODY_CHARS` in `client.py`.
pub const DEFAULT_BODY_CHARS: usize = 100_000;
/// Bounds for a caller-chosen `max_body_chars` (large HTML mail with inline
/// images can run to megabytes). Match `MIN_BODY_CHARS_LIMIT` /
/// `MAX_BODY_CHARS_LIMIT` in `client.py`.
pub const MIN_BODY_CHARS_LIMIT: usize = 1_000;
pub const MAX_BODY_CHARS_LIMIT: usize = 5_000_000;

/// Result of a batch read: one entry per requested id, in request order.
/// The outer error is for failures of the whole call (COM unavailable, an
/// unusable `output_dir`); a failure of one id is its own entry's error.
pub type BatchResult<T> = Result<Vec<Result<T, ToolError>>, ToolError>;

/// The first (only) entry of a one-id [`BatchResult`], as a plain result.
pub fn single<T>(batch: BatchResult<T>) -> Result<T, ToolError> {
    batch?
        .into_iter()
        .next()
        .unwrap_or_else(|| Err(ToolError::new("no result returned for the requested id")))
}

/// `client.py::_truncate`: cap long bodies at `limit` *characters* (not
/// bytes) so multi-byte UTF-8 content is never split mid-codepoint. The cut
/// is a hard one and may land mid-tag or mid-base64 in HTML; the returned
/// flag says whether it happened so callers can re-request with a larger
/// limit. Returns `(text, truncated)`.
pub fn truncate(text: &str, limit: usize) -> (String, bool) {
    if text.chars().count() > limit {
        let head: String = text.chars().take(limit).collect();
        (format!("{head}\n\n[... truncated at {limit} characters]"), true)
    } else {
        (text.to_string(), false)
    }
}

/// `client.py::_clamp_body_limit`: `None` means [`DEFAULT_BODY_CHARS`]; any
/// other value is clamped into `[MIN_BODY_CHARS_LIMIT, MAX_BODY_CHARS_LIMIT]`.
pub fn clamp_body_limit(max_body_chars: Option<u32>) -> usize {
    match max_body_chars {
        None => DEFAULT_BODY_CHARS,
        Some(n) => (n as usize).clamp(MIN_BODY_CHARS_LIMIT, MAX_BODY_CHARS_LIMIT),
    }
}

// ---- include / body_format ----------------------------------------------

/// Which read tool a request is for; decides the valid `include` fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ReadTool {
    Email,
    Event,
    Note,
    Task,
}

impl ReadTool {
    pub fn tool_name(self) -> &'static str {
        match self {
            ReadTool::Email => "get_email",
            ReadTool::Event => "get_event",
            ReadTool::Note => "get_note",
            ReadTool::Task => "get_task",
        }
    }

    /// The optional (heavy) fields this tool's `include` accepts.
    pub fn include_fields(self) -> &'static [&'static str] {
        match self {
            ReadTool::Email => &["body", "html_body", "attachments", "meeting"],
            _ => &["body"],
        }
    }

    /// What `include` means when omitted: exactly what the tool returned
    /// before `include` existed.
    pub fn default_include(self) -> &'static [&'static str] {
        match self {
            ReadTool::Email => &["body", "attachments", "meeting"],
            _ => &["body"],
        }
    }

    /// Prefix of the file names `output_dir` writes for this tool.
    pub fn file_prefix(self) -> &'static str {
        match self {
            ReadTool::Email => "email",
            ReadTool::Event => "event",
            ReadTool::Note => "note",
            ReadTool::Task => "task",
        }
    }
}

/// Resolved output options for one read call (all ids in a batch share
/// them). Fields that a tool doesn't have (`html_body`, `attachments`,
/// `meeting`, `resolve_inline_images` outside `get_email`) are ignored.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReadOptions {
    pub body: bool,
    pub html_body: bool,
    pub attachments: bool,
    pub meeting: bool,
    /// Replace `cid:` references in `html_body` with `data:` URIs.
    pub resolve_inline_images: bool,
    /// `None` = [`DEFAULT_BODY_CHARS`]; see [`clamp_body_limit`].
    pub max_body_chars: Option<u32>,
    /// Write bodies to files in this directory instead of returning them.
    pub output_dir: Option<String>,
}

impl Default for ReadOptions {
    /// The pre-`include` output: plain body, attachment names and the
    /// meeting block; no HTML.
    fn default() -> Self {
        Self {
            body: true,
            html_body: false,
            attachments: true,
            meeting: true,
            resolve_inline_images: false,
            max_body_chars: None,
            output_dir: None,
        }
    }
}

impl ReadOptions {
    pub fn body_limit(&self) -> usize {
        clamp_body_limit(self.max_body_chars)
    }
}

/// A read tool's output parameters as the caller sent them (ids aside).
#[derive(Debug, Clone, Default)]
pub struct ReadRequest {
    pub include: Option<Vec<String>>,
    pub body_format: Option<String>,
    /// `get_email`'s deprecated `prefer_html` (= `body_format: "html"`).
    pub prefer_html: Option<bool>,
    pub resolve_inline_images: bool,
    pub max_body_chars: Option<u32>,
    pub output_dir: Option<String>,
}

/// Validate a read request and resolve it to [`ReadOptions`].
///
/// - `include` omitted = [`ReadTool::default_include`]; `[]` = metadata
///   only. Field names are case-insensitive; unknown ones are an error.
/// - `body_format: "html"` (or the deprecated `prefer_html: true`) adds
///   `html_body`, so the default `get_email` output plus `body_format: "html"`
///   is exactly what `prefer_html: true` returned. Only `get_email` has an
///   HTML body; the other tools accept only `"text"`.
/// - `resolve_inline_images` (`get_email` only) also adds `html_body`.
pub fn read_options(tool: ReadTool, req: ReadRequest) -> Result<ReadOptions, ToolError> {
    let name = tool.tool_name();
    let format_html = match req.body_format.as_deref().map(|f| f.trim().to_ascii_lowercase()) {
        None => None,
        Some(f) if f == "text" => Some(false),
        Some(f) if f == "html" => Some(true),
        Some(other) => {
            return Err(ToolError::new(format!(
                "{name}: body_format must be \"text\" or \"html\", got {other:?}."
            )))
        }
    };
    let html = match (req.prefer_html, format_html) {
        (Some(prefer), Some(format)) if prefer != format => {
            return Err(ToolError::new(format!(
                "{name}: `prefer_html` (deprecated) is {prefer} but `body_format` is \"{}\"; \
                 they contradict. Pass only body_format.",
                if format { "html" } else { "text" }
            )))
        }
        (prefer, format) => format.or(prefer).unwrap_or(false),
    };
    if tool != ReadTool::Email {
        if html {
            return Err(ToolError::new(format!(
                "{name}: this item only has a plain-text body; body_format must be \"text\"."
            )));
        }
        if req.resolve_inline_images {
            return Err(ToolError::new(format!(
                "{name}: resolve_inline_images is only supported by get_email."
            )));
        }
    }
    let allowed = tool.include_fields();
    let fields: Vec<String> = match req.include {
        None => tool.default_include().iter().map(|f| f.to_string()).collect(),
        Some(list) => {
            let mut fields = Vec::new();
            for raw in list {
                let field = raw.trim().to_ascii_lowercase();
                if !allowed.contains(&field.as_str()) {
                    return Err(ToolError::new(format!(
                        "{name}: unknown include field {raw:?}; valid fields are: {}.",
                        allowed.join(", ")
                    )));
                }
                fields.push(field);
            }
            fields
        }
    };
    let has = |f: &str| fields.iter().any(|x| x == f);
    let output_dir = match req.output_dir {
        Some(dir) if dir.trim().is_empty() => {
            return Err(ToolError::new(format!(
                "{name}: output_dir must be a directory path (or omitted)."
            )))
        }
        other => other,
    };
    Ok(ReadOptions {
        body: has("body"),
        html_body: has("html_body") || html || req.resolve_inline_images,
        attachments: has("attachments"),
        meeting: has("meeting"),
        resolve_inline_images: req.resolve_inline_images,
        max_body_chars: req.max_body_chars,
        output_dir,
    })
}

// ---- output_dir ------------------------------------------------------------

/// `os.path.abspath(os.path.expanduser(dir))`: expand a leading `~` to the
/// user's home directory, then make the path absolute. Used for
/// `save_attachments`' `save_dir` and the read tools' `output_dir`, so the
/// returned paths are absolute. Falls back to the un-absolutized path if
/// `absolute` fails (it does not touch the filesystem, so this only guards
/// against odd inputs).
pub fn resolve_save_dir(save_dir: &str) -> PathBuf {
    let expanded = if save_dir == "~"
        || save_dir.starts_with("~/")
        || save_dir.starts_with("~\\")
    {
        match std::env::var_os("USERPROFILE").or_else(|| std::env::var_os("HOME")) {
            Some(home) => {
                let rest = save_dir[1..].trim_start_matches(['/', '\\']);
                Path::new(&home).join(rest)
            }
            None => PathBuf::from(save_dir),
        }
    } else {
        PathBuf::from(save_dir)
    };
    std::path::absolute(&expanded).unwrap_or(expanded)
}

/// Resolve `output_dir` (see [`resolve_save_dir`]) and create it if needed.
pub fn prepare_output_dir(raw: &str) -> Result<PathBuf, ToolError> {
    let dir = resolve_save_dir(raw.trim());
    if dir.exists() && !dir.is_dir() {
        return Err(ToolError::new(format!(
            "output_dir {:?} exists and is not a directory.",
            dir.display()
        )));
    }
    std::fs::create_dir_all(&dir).map_err(|e| {
        ToolError::new(format!("Could not create output_dir {:?}: {e}", dir.display()))
    })?;
    Ok(dir)
}

/// 64-bit FNV-1a: a stable (across builds and platforms) hash used to give
/// each item a short, deterministic file name.
pub fn fnv1a64(text: &str) -> u64 {
    let mut hash: u64 = 0xcbf2_9ce4_8422_2325;
    for byte in text.bytes() {
        hash ^= u64::from(byte);
        hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
    hash
}

/// `<dir>/<prefix>-<16 hex digits of fnv1a64(key)>-<field><ext>`, e.g.
/// `email-1f2e…-html_body.html`. The same item and field always map to the
/// same name, so reading again overwrites the earlier file.
pub fn output_file_path(dir: &Path, prefix: &str, key: &str, field: &str, ext: &str) -> PathBuf {
    dir.join(format!("{prefix}-{:016x}-{field}{ext}", fnv1a64(key)))
}

/// Write `bytes` to `path`; returns the path as a string for `<field>_file`.
pub fn write_output_file(path: &Path, bytes: &[u8]) -> Result<String, ToolError> {
    std::fs::write(path, bytes)
        .map_err(|e| ToolError::new(format!("Could not write {:?}: {e}", path.display())))?;
    Ok(path.to_string_lossy().into_owned())
}

/// One body field as returned: inline (possibly cut at the limit) or as a
/// file path. `length` is always the full original length in characters.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BodyOut {
    pub text: Option<String>,
    pub file: Option<String>,
    pub truncated: bool,
    pub length: usize,
}

impl BodyOut {
    /// `(<field>, <field>_file, *_truncated, *_length)` for an output struct.
    pub fn into_parts(this: Option<BodyOut>)
        -> (Option<String>, Option<String>, Option<bool>, Option<usize>) {
        match this {
            Some(b) => (b.text, b.file, Some(b.truncated), Some(b.length)),
            None => (None, None, None, None),
        }
    }
}

/// Shape one body field. With `file`, the FULL text is written there as
/// UTF-8 (never truncated: `max_body_chars` only limits inline text) and
/// `text` is `None`; otherwise `text` is cut at `limit` characters.
pub fn shape_body(full: &str, limit: usize, file: Option<&Path>) -> Result<BodyOut, ToolError> {
    let length = full.chars().count();
    if let Some(path) = file {
        let file = write_output_file(path, full.as_bytes())?;
        return Ok(BodyOut { text: None, file: Some(file), truncated: false, length });
    }
    let (text, truncated) = truncate(full, limit);
    Ok(BodyOut { text: Some(text), file: None, truncated, length })
}

/// [`shape_body`] for a read tool's field: the file (when `out_dir` is set)
/// is named by [`output_file_path`] from the tool, the item id and the field.
pub fn shape_field(full: &str, tool: ReadTool, id: &str, field: &str, ext: &str,
    opts: &ReadOptions, out_dir: Option<&Path>) -> Result<BodyOut, ToolError> {
    let file = out_dir.map(|dir| output_file_path(dir, tool.file_prefix(), id, field, ext));
    shape_body(full, opts.body_limit(), file.as_deref())
}

/// File extension (with the dot) for an image written to `output_dir`:
/// from the MIME type, else the attachment's own extension, else `.bin`.
pub fn image_extension(mime: &str, filename: &str) -> String {
    let ext = extension_for_mime(mime);
    if !ext.is_empty() {
        return ext.to_string();
    }
    Path::new(filename)
        .extension()
        .and_then(|e| e.to_str())
        .filter(|e| !e.is_empty() && e.chars().all(|c| c.is_ascii_alphanumeric()))
        .map_or_else(|| ".bin".to_string(), |e| format!(".{}", e.to_ascii_lowercase()))
}

/// Outlook's "no date" sentinel (`4501-01-01`) as `None`, other ISO dates
/// unchanged.
pub fn outlook_date(iso: Option<String>) -> Option<String> {
    iso.filter(|s| !s.starts_with("4501-"))
}

// ---- resolve_inline_images ------------------------------------------------

/// One `cid:` reference in HTML: the byte span of the whole reference
/// (`cid:` and any `<>` included) and the Content-ID it names
/// (percent-decoded, brackets removed).
#[derive(Debug, Clone, PartialEq, Eq)]
struct CidRef {
    start: usize,
    end: usize,
    cid: String,
}

/// Decode `%XX` escapes (RFC 2392 `cid:` URLs are URL-encoded). Input that
/// doesn't decode to valid UTF-8 is returned unchanged.
fn percent_decode(raw: &str) -> String {
    if !raw.contains('%') {
        return raw.to_string();
    }
    let bytes = raw.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 2 < bytes.len() {
            let hex = |b: u8| char::from(b).to_digit(16);
            if let (Some(hi), Some(lo)) = (hex(bytes[i + 1]), hex(bytes[i + 2])) {
                out.push((hi * 16 + lo) as u8);
                i += 3;
                continue;
            }
        }
        out.push(bytes[i]);
        i += 1;
    }
    String::from_utf8(out).unwrap_or_else(|_| raw.to_string())
}

/// Every `cid:` reference in `html`, in order. `cid:` is matched in any
/// case and only at a word start (not inside e.g. `xcid:`). The id is
/// either `<...>` (no whitespace or quotes inside) or the run of Content-ID
/// characters that follows ([`is_cid_char`]).
fn find_cid_refs(html: &str) -> Vec<CidRef> {
    // ASCII-only lowercasing keeps byte offsets identical to `html`.
    let lower = html.to_ascii_lowercase();
    let mut refs = Vec::new();
    let mut from = 0;
    while let Some(rel) = lower[from..].find("cid:") {
        let pos = from + rel;
        from = pos + 4;
        if html[..pos].chars().next_back().is_some_and(|c| c.is_alphanumeric()) {
            continue;
        }
        let rest = &html[pos + 4..];
        let (raw, len) = if let Some(inner) = rest.strip_prefix('<') {
            match inner.find('>') {
                Some(gt) if gt > 0
                    && !inner[..gt].contains(|c: char| c.is_whitespace() || "\"'<".contains(c)) =>
                {
                    (&inner[..gt], gt + 2)
                }
                _ => continue,
            }
        } else {
            let len = rest
                .char_indices()
                .find(|&(_, c)| !is_cid_char(c))
                .map_or(rest.len(), |(i, _)| i);
            if len == 0 {
                continue;
            }
            (&rest[..len], len)
        };
        refs.push(CidRef { start: pos, end: pos + 4 + len, cid: percent_decode(raw) });
        from = pos + 4 + len;
    }
    refs
}

/// The distinct Content-IDs `html` references as `cid:...`, in order of
/// first appearance (compared case-insensitively; the first spelling wins).
pub fn cid_references(html: &str) -> Vec<String> {
    let mut seen: Vec<String> = Vec::new();
    for r in find_cid_refs(html) {
        if !seen.iter().any(|s| s.to_lowercase() == r.cid.to_lowercase()) {
            seen.push(r.cid);
        }
    }
    seen
}

/// `html` with its `cid:` references rewritten, plus what happened.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CidRewrite {
    pub html: String,
    /// Distinct Content-IDs that were replaced by a `data:` URI.
    pub resolved: Vec<String>,
    /// Distinct referenced Content-IDs with no entry in `images` (unknown,
    /// over the size cap, or unreadable); their references are untouched.
    pub unresolved: Vec<String>,
}

/// Replace each `cid:<id>` reference in `html` whose id matches an entry of
/// `images` (`(content_id, data_uri)`; the content id may carry `<>` or a
/// `cid:` prefix) with that data URI. Matching is case-insensitive and on
/// the whole id; a `cid:<id>` reference loses its brackets too. References
/// to ids not in `images` are left exactly as they were.
pub fn replace_cid_references(html: &str, images: &[(String, String)]) -> CidRewrite {
    let keyed: Vec<(String, &str)> = images
        .iter()
        // Twice, so both `<cid:x>` and `cid:<x>` come down to `x`.
        .map(|(cid, uri)| {
            (normalize_content_id(&normalize_content_id(cid)).to_lowercase(), uri.as_str())
        })
        .collect();
    let mut out = String::with_capacity(html.len());
    let mut resolved: Vec<String> = Vec::new();
    let mut unresolved: Vec<String> = Vec::new();
    let mut last = 0;
    for r in find_cid_refs(html) {
        let key = r.cid.to_lowercase();
        let note = |list: &mut Vec<String>| {
            if !list.iter().any(|s| s.to_lowercase() == key) {
                list.push(r.cid.clone());
            }
        };
        match keyed.iter().find(|(k, _)| *k == key) {
            Some((_, uri)) => {
                out.push_str(&html[last..r.start]);
                out.push_str(uri);
                last = r.end;
                note(&mut resolved);
            }
            None => note(&mut unresolved),
        }
    }
    out.push_str(&html[last..]);
    CidRewrite { html: out, resolved, unresolved }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn req() -> ReadRequest {
        ReadRequest::default()
    }

    // ---- truncate / clamp -------------------------------------------------

    #[test]
    fn truncate_leaves_short_text_untouched() {
        assert_eq!(truncate("hello", 10), ("hello".to_string(), false));
        // Exactly at the limit is not truncated.
        assert_eq!(truncate("hello", 5), ("hello".to_string(), false));
        assert_eq!(truncate("", 5), (String::new(), false));
    }

    #[test]
    fn truncate_cuts_long_text_and_flags_it() {
        let (text, truncated) = truncate("abcdefghij", 4);
        assert!(truncated);
        assert_eq!(text, "abcd\n\n[... truncated at 4 characters]");
    }

    #[test]
    fn truncate_counts_chars_not_bytes() {
        // 'é' is 2 bytes in UTF-8; a byte-based cut would panic or split it.
        let (text, truncated) = truncate("ééééé", 3);
        assert!(truncated);
        assert!(text.starts_with("ééé\n"));
        assert_eq!(truncate("ééé", 3), ("ééé".to_string(), false));
    }

    #[test]
    fn clamp_body_limit_defaults_and_clamps() {
        assert_eq!(clamp_body_limit(None), DEFAULT_BODY_CHARS);
        assert_eq!(clamp_body_limit(Some(0)), MIN_BODY_CHARS_LIMIT);
        assert_eq!(clamp_body_limit(Some(999)), MIN_BODY_CHARS_LIMIT);
        assert_eq!(clamp_body_limit(Some(1_000)), 1_000);
        assert_eq!(clamp_body_limit(Some(250_000)), 250_000);
        assert_eq!(clamp_body_limit(Some(5_000_000)), MAX_BODY_CHARS_LIMIT);
        assert_eq!(clamp_body_limit(Some(u32::MAX)), MAX_BODY_CHARS_LIMIT);
    }

    // ---- read_options -----------------------------------------------------

    #[test]
    fn defaults_match_the_pre_include_output() {
        assert_eq!(read_options(ReadTool::Email, req()).unwrap(), ReadOptions::default());
        let ev = read_options(ReadTool::Event, req()).unwrap();
        assert!(ev.body && !ev.html_body && !ev.attachments && !ev.meeting);
    }

    #[test]
    fn body_format_html_and_prefer_html_add_html_body() {
        let html = read_options(ReadTool::Email, ReadRequest {
            body_format: Some(" HTML ".into()), ..req()
        }).unwrap();
        assert_eq!(html, ReadOptions { html_body: true, ..ReadOptions::default() });
        let legacy = read_options(ReadTool::Email, ReadRequest {
            prefer_html: Some(true), ..req()
        }).unwrap();
        assert_eq!(legacy, html);
        let text = read_options(ReadTool::Email, ReadRequest {
            body_format: Some("text".into()), prefer_html: Some(false), ..req()
        }).unwrap();
        assert_eq!(text, ReadOptions::default());
    }

    #[test]
    fn contradicting_prefer_html_and_body_format_is_an_error() {
        let err = read_options(ReadTool::Email, ReadRequest {
            body_format: Some("text".into()), prefer_html: Some(true), ..req()
        }).unwrap_err();
        assert!(err.0.contains("prefer_html") && err.0.contains("body_format"), "{err}");
        assert!(read_options(ReadTool::Email, ReadRequest {
            body_format: Some("html".into()), prefer_html: Some(false), ..req()
        }).is_err());
    }

    #[test]
    fn unknown_body_format_is_an_error() {
        let err = read_options(ReadTool::Note, ReadRequest {
            body_format: Some("rtf".into()), ..req()
        }).unwrap_err();
        assert!(err.0.contains("get_note") && err.0.contains("\"rtf\""), "{err}");
    }

    #[test]
    fn non_email_tools_reject_html_and_resolve() {
        for tool in [ReadTool::Event, ReadTool::Note, ReadTool::Task] {
            assert!(read_options(tool, ReadRequest { body_format: Some("html".into()), ..req() }).is_err());
            assert!(read_options(tool, ReadRequest { resolve_inline_images: true, ..req() }).is_err());
            assert!(read_options(tool, ReadRequest { body_format: Some("text".into()), ..req() }).is_ok());
        }
    }

    #[test]
    fn include_selects_fields_and_empty_means_metadata_only() {
        let meta = read_options(ReadTool::Email, ReadRequest { include: Some(vec![]), ..req() }).unwrap();
        assert!(!meta.body && !meta.html_body && !meta.attachments && !meta.meeting);
        let some = read_options(ReadTool::Email, ReadRequest {
            include: Some(vec!["HTML_BODY".into(), " attachments".into()]), ..req()
        }).unwrap();
        assert!(!some.body && some.html_body && some.attachments && !some.meeting);
        // body_format html adds html_body to an explicit include too.
        let html_only = read_options(ReadTool::Email, ReadRequest {
            include: Some(vec![]), body_format: Some("html".into()), ..req()
        }).unwrap();
        assert!(!html_only.body && html_only.html_body);
        let task = read_options(ReadTool::Task, ReadRequest { include: Some(vec![]), ..req() }).unwrap();
        assert!(!task.body);
    }

    #[test]
    fn unknown_include_field_names_the_valid_ones() {
        let err = read_options(ReadTool::Email, ReadRequest {
            include: Some(vec!["body".into(), "headers".into()]), ..req()
        }).unwrap_err();
        assert!(err.0.contains("\"headers\""), "{err}");
        assert!(err.0.contains("body, html_body, attachments, meeting"), "{err}");
        // html_body is an email-only field.
        assert!(read_options(ReadTool::Event, ReadRequest {
            include: Some(vec!["html_body".into()]), ..req()
        }).is_err());
    }

    #[test]
    fn resolve_inline_images_implies_html_body() {
        let o = read_options(ReadTool::Email, ReadRequest {
            include: Some(vec![]), resolve_inline_images: true, ..req()
        }).unwrap();
        assert!(o.html_body && o.resolve_inline_images && !o.body);
    }

    #[test]
    fn blank_output_dir_is_an_error() {
        assert!(read_options(ReadTool::Email, ReadRequest { output_dir: Some("  ".into()), ..req() }).is_err());
        let o = read_options(ReadTool::Email, ReadRequest { output_dir: Some("out".into()), ..req() }).unwrap();
        assert_eq!(o.output_dir.as_deref(), Some("out"));
    }

    // ---- output_dir -------------------------------------------------------

    fn temp_dir(name: &str) -> PathBuf {
        let dir = std::env::temp_dir().join(format!("outlook-mcp-read-{}-{name}", std::process::id()));
        let _ = std::fs::remove_dir_all(&dir);
        dir
    }

    #[test]
    fn fnv1a64_is_the_reference_fnv() {
        assert_eq!(fnv1a64(""), 0xcbf2_9ce4_8422_2325);
        assert_eq!(fnv1a64("a"), 0xaf63_dc4c_8601_ec8c);
        assert_ne!(fnv1a64("entry-1|store-1"), fnv1a64("entry-2|store-1"));
    }

    #[test]
    fn output_file_path_is_deterministic_per_item_and_field() {
        let dir = Path::new("out");
        let a = output_file_path(dir, "email", "e1|s1", "body", ".txt");
        assert_eq!(a, output_file_path(dir, "email", "e1|s1", "body", ".txt"));
        assert_ne!(a, output_file_path(dir, "email", "e2|s1", "body", ".txt"));
        let name = a.file_name().unwrap().to_str().unwrap().to_string();
        assert!(name.starts_with("email-") && name.ends_with("-body.txt"), "{name}");
        assert_eq!(name.len(), "email-".len() + 16 + "-body.txt".len());
    }

    #[test]
    fn prepare_output_dir_creates_and_rejects_files() {
        let dir = temp_dir("prep");
        let made = prepare_output_dir(dir.join("nested").to_str().unwrap()).unwrap();
        assert!(made.is_dir() && made.is_absolute());
        let file = dir.join("f.txt");
        std::fs::write(&file, b"x").unwrap();
        assert!(prepare_output_dir(file.to_str().unwrap()).is_err());
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn shape_body_inline_truncates_and_reports_length() {
        let b = shape_body("abcdefghij", 4, None).unwrap();
        assert_eq!(b.length, 10);
        assert!(b.truncated && b.file.is_none());
        assert!(b.text.unwrap().starts_with("abcd\n"));
        let parts = BodyOut::into_parts(Some(shape_body("hi", 4, None).unwrap()));
        assert_eq!(parts, (Some("hi".into()), None, Some(false), Some(2)));
        assert_eq!(BodyOut::into_parts(None), (None, None, None, None));
    }

    #[test]
    fn shape_body_to_file_writes_the_full_text() {
        let dir = temp_dir("shape");
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("b.txt");
        let b = shape_body("שלום abcdefghij", 4, Some(&path)).unwrap();
        assert_eq!(b.text, None);
        assert!(!b.truncated);
        assert_eq!(b.length, 15);
        assert_eq!(b.file.as_deref(), Some(path.to_str().unwrap()));
        assert_eq!(std::fs::read_to_string(&path).unwrap(), "שלום abcdefghij");
        let _ = std::fs::remove_dir_all(&dir);
    }

    #[test]
    fn image_extension_prefers_mime_then_filename() {
        assert_eq!(image_extension("image/png", "x.jpg"), ".png");
        assert_eq!(image_extension("application/octet-stream", "Chart.EMF"), ".emf");
        assert_eq!(image_extension("application/octet-stream", "noext"), ".bin");
        assert_eq!(image_extension("", "bad.ex t"), ".bin");
    }

    #[test]
    fn outlook_date_drops_the_none_sentinel() {
        assert_eq!(outlook_date(Some("4501-01-01T00:00:00".into())), None);
        assert_eq!(outlook_date(Some("2026-10-09T00:00:00".into())).as_deref(), Some("2026-10-09T00:00:00"));
        assert_eq!(outlook_date(None), None);
    }

    // ---- cid rewriting ----------------------------------------------------

    fn img(cid: &str, uri: &str) -> (String, String) {
        (cid.to_string(), uri.to_string())
    }

    #[test]
    fn finds_references_in_order_without_duplicates() {
        let html = r#"<img src="cid:a@x"><img src='CID:b@x'><img src="cid:A@X">"#;
        assert_eq!(cid_references(html), vec!["a@x".to_string(), "b@x".to_string()]);
        assert!(cid_references("<p>no images</p>").is_empty());
    }

    #[test]
    fn replaces_references_case_insensitively() {
        let html = r#"<img src="cid:Logo@X"> and <img src="CID:logo@x">"#;
        let r = replace_cid_references(html, &[img("logo@x", "data:image/png;base64,AA==")]);
        assert_eq!(
            r.html,
            r#"<img src="data:image/png;base64,AA=="> and <img src="data:image/png;base64,AA==">"#
        );
        assert_eq!(r.resolved, vec!["Logo@X".to_string()]);
        assert!(r.unresolved.is_empty());
    }

    #[test]
    fn angle_brackets_on_either_side_are_accepted() {
        // The attachment's Content-ID carries <> and a cid: prefix...
        let r = replace_cid_references(r#"<img src="cid:a@x">"#, &[img("<a@x>", "data:x")]);
        assert_eq!(r.html, r#"<img src="data:x">"#);
        let r = replace_cid_references(r#"<img src="cid:a@x">"#, &[img("cid:<A@X>", "data:x")]);
        assert_eq!(r.html, r#"<img src="data:x">"#);
        // ...and so does the reference: the brackets go with it.
        let r = replace_cid_references(r#"<img src="cid:<a@x>">"#, &[img("a@x", "data:x")]);
        assert_eq!(r.html, r#"<img src="data:x">"#);
        // An unclosed or empty bracket is not a reference.
        assert!(cid_references(r#"cid:<a@x "#).is_empty());
        assert!(cid_references("cid:<>").is_empty());
    }

    #[test]
    fn unknown_cids_are_left_untouched_and_reported() {
        let html = r#"<img src="cid:known@x"><img src="cid:missing@x"><img src="cid:missing@x">"#;
        let r = replace_cid_references(html, &[img("known@x", "data:k")]);
        assert_eq!(r.html, r#"<img src="data:k"><img src="cid:missing@x"><img src="cid:missing@x">"#);
        assert_eq!(r.resolved, vec!["known@x".to_string()]);
        assert_eq!(r.unresolved, vec!["missing@x".to_string()]);
        // No images at all: the HTML comes back byte-for-byte.
        let r = replace_cid_references(html, &[]);
        assert_eq!(r.html, html);
    }

    #[test]
    fn whole_id_matching_and_word_boundary() {
        // `cid:logo` is not a reference to `logo.png` and vice versa.
        let html = r#"<img src="cid:logo.png"><img src="cid:logo">"#;
        let r = replace_cid_references(html, &[img("logo", "data:l")]);
        assert_eq!(r.html, r#"<img src="cid:logo.png"><img src="data:l">"#);
        // `xcid:` is not a cid reference.
        let r = replace_cid_references("xcid:logo", &[img("logo", "data:l")]);
        assert_eq!(r.html, "xcid:logo");
    }

    #[test]
    fn percent_encoded_references_match() {
        let r = replace_cid_references(r#"<img src="cid:a%40x">"#, &[img("a@x", "data:p")]);
        assert_eq!(r.html, r#"<img src="data:p">"#);
        assert_eq!(percent_decode("a%4"), "a%4");
        assert_eq!(percent_decode("%zz"), "%zz");
        assert_eq!(percent_decode("%C3%A9"), "é");
        assert_eq!(percent_decode("%FF"), "%FF");
    }

    #[test]
    fn non_ascii_html_keeps_its_bytes() {
        let html = "<p>שלום</p><img src=\"cid:a@x\"><p>עוד</p>";
        let r = replace_cid_references(html, &[img("a@x", "data:h")]);
        assert_eq!(r.html, "<p>שלום</p><img src=\"data:h\"><p>עוד</p>");
    }
}
