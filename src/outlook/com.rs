use crate::error::ToolError;

/// `"{EntryID}|{StoreID}"`, matching the Python client's opaque item id format.
pub fn make_item_id(entry_id: &str, store_id: &str) -> String {
    format!("{entry_id}|{store_id}")
}

pub fn parse_item_id(item_id: &str) -> Result<(String, String), ToolError> {
    match item_id.split_once('|') {
        Some((entry, store)) if !entry.is_empty() && !store.is_empty() => {
            Ok((entry.to_string(), store.to_string()))
        }
        _ => Err(ToolError::new(format!(
            "Invalid item id {item_id:?}: expected the opaque id returned by a list/search tool."
        ))),
    }
}

/// JET `Restrict` filters want `MM/DD/YYYY HH:MM AM/PM` (US format, no
/// seconds) — anything else silently misfilters. Mirrors `_jet_dt` in
/// `outlook_mcp/outlook/client.py`.
pub fn jet_datetime(dt: &chrono::NaiveDateTime) -> String {
    dt.format("%m/%d/%Y %I:%M %p").to_string()
}

pub fn safe_filename(name: &str) -> String {
    let cleaned: String = name
        .chars()
        .map(|c| if "\\/:*?\"<>|".contains(c) || (c as u32) < 0x20 { '_' } else { c })
        .collect();
    let trimmed = cleaned.trim_matches(|c| c == '.' || c == ' ');
    if trimmed.is_empty() { "attachment".to_string() } else { trimmed.to_string() }
}

/// Outlook stores categories as one `", "`-joined string. Split it into names,
/// trimming whitespace and dropping empties. Mirrors how the Python client
/// would `.split(", ")`.
pub fn parse_categories(raw: &str) -> Vec<String> {
    raw.split(',')
        .map(|s| s.trim())
        .filter(|s| !s.is_empty())
        .map(|s| s.to_string())
        .collect()
}

/// Join category names back into the `", "`-separated string Outlook expects.
pub fn join_categories(cats: &[String]) -> String {
    cats.join(", ")
}

/// Normalize an attachment Content-ID (`PR_ATTACH_CONTENT_ID`): trim whitespace
/// and one pair of surrounding `<>`, so it matches the `cid:` references in an
/// HTML body. Empty (or missing) becomes `None`.
pub fn clean_content_id(raw: Option<&str>) -> Option<String> {
    let trimmed = raw?.trim();
    let inner = trimmed
        .strip_prefix('<')
        .and_then(|s| s.strip_suffix('>'))
        .unwrap_or(trimmed)
        .trim();
    if inner.is_empty() { None } else { Some(inner.to_string()) }
}

/// Normalize a caller-supplied Content-ID the way an HTML body might spell it:
/// trim whitespace, drop a leading `cid:` (any case), then strip `<>` as
/// `clean_content_id` does. `None` when nothing is left. Compare the result
/// case-insensitively.
pub fn normalize_cid_request(raw: &str) -> Option<String> {
    let trimmed = raw.trim();
    let rest = match trimmed.get(..4) {
        Some(prefix) if prefix.eq_ignore_ascii_case("cid:") => &trimmed[4..],
        _ => trimmed,
    };
    clean_content_id(Some(rest))
}

/// An attachment's MIME type: its own MIME tag (`PR_ATTACH_MIME_TAG`),
/// lowercased, when present; otherwise a guess from the file extension.
/// `None` when neither yields anything.
pub fn guess_mime(mime_tag: Option<&str>, filename: &str) -> Option<String> {
    if let Some(tag) = mime_tag.map(str::trim).filter(|t| !t.is_empty()) {
        return Some(tag.to_lowercase());
    }
    let (_, ext) = filename.rsplit_once('.')?;
    let mime = match ext.to_lowercase().as_str() {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "webp" => "image/webp",
        "bmp" => "image/bmp",
        "svg" => "image/svg+xml",
        "pdf" => "application/pdf",
        "txt" => "text/plain",
        "html" | "htm" => "text/html",
        "csv" => "text/csv",
        "json" => "application/json",
        "zip" => "application/zip",
        "doc" => "application/msword",
        "docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "xls" => "application/vnd.ms-excel",
        "xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "ppt" => "application/vnd.ms-powerpoint",
        "pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "eml" => "message/rfc822",
        "msg" => "application/vnd.ms-outlook",
        _ => return None,
    };
    Some(mime.to_string())
}

/// Whether an attachment is inline (`cid:`-referenced) content rather than a
/// standalone attachment. `content_id` is already cleaned of `<>` (see
/// `clean_content_id`).
///
/// Inline when it has a Content-ID and either Outlook marked it hidden or the
/// HTML body references it as `cid:<content_id>` (case-insensitive, whole id).
/// A Content-ID the HTML never references is usually a regular attachment some
/// senders label with a CID anyway; no Content-ID is never inline.
pub fn is_inline(content_id: Option<&str>, hidden: bool, html_body: &str) -> bool {
    let Some(id) = content_id.filter(|id| !id.is_empty()) else {
        return false;
    };
    if hidden {
        return true;
    }
    let needle = format!("cid:{}", id.to_lowercase());
    let haystack = html_body.to_lowercase();
    // The reference must end where the id ends, so `cid:img10@x` does not
    // count as a reference to `img1@x`.
    haystack.match_indices(&needle).any(|(start, _)| {
        haystack[start + needle.len()..]
            .chars()
            .next()
            .is_none_or(|ch| ch.is_whitespace() || matches!(ch, '"' | '\'' | '(' | ')' | '<' | '>'))
    })
}

/// Decodes a COM wide string (BSTR contents, UTF-16) into a Rust `String`.
/// Well-formed UTF-16 (Hebrew, other BMP scripts, emoji as surrogate pairs)
/// converts losslessly. Only an unpaired surrogate, which no `String` can
/// hold, becomes U+FFFD; that matches `BSTR`'s own `Display`.
pub fn string_from_utf16(wide: &[u16]) -> String {
    String::from_utf16_lossy(wide)
}

/// Maps a Windows `LOCALE_IFIRSTDAYOFWEEK` value (`"0"` = Monday through
/// `"6"` = Sunday) to a weekday. Surrounding whitespace and NULs are ignored;
/// anything else is `None`.
pub fn first_day_of_week_from_locale_value(value: &str) -> Option<chrono::Weekday> {
    use chrono::Weekday::*;
    match value.trim_matches(|c: char| c == '\0' || c.is_whitespace()) {
        "0" => Some(Mon),
        "1" => Some(Tue),
        "2" => Some(Wed),
        "3" => Some(Thu),
        "4" => Some(Fri),
        "5" => Some(Sat),
        "6" => Some(Sun),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn first_day_of_week_maps_every_locale_value() {
        use chrono::Weekday::*;
        let expected = [Mon, Tue, Wed, Thu, Fri, Sat, Sun];
        for (i, day) in expected.into_iter().enumerate() {
            assert_eq!(first_day_of_week_from_locale_value(&i.to_string()), Some(day));
        }
        assert_eq!(first_day_of_week_from_locale_value("6\0"), Some(Sun));
        assert_eq!(first_day_of_week_from_locale_value(" 5 "), Some(Sat));
        for bad in ["", "7", "-1", "06", "Monday", "\0"] {
            assert_eq!(first_day_of_week_from_locale_value(bad), None, "{bad:?}");
        }
    }

    #[test]
    fn make_and_parse_item_id_round_trip() {
        let id = make_item_id("entry-1", "store-1");
        assert_eq!(id, "entry-1|store-1");
        assert_eq!(parse_item_id(&id).unwrap(), ("entry-1".to_string(), "store-1".to_string()));
    }

    #[test]
    fn parse_item_id_rejects_malformed_input() {
        assert!(parse_item_id("no-separator").is_err());
        assert!(parse_item_id("|missing-entry").is_err());
        assert!(parse_item_id("missing-store|").is_err());
    }

    #[test]
    fn jet_datetime_formats_us_style_no_seconds() {
        use chrono::NaiveDate;
        let dt = NaiveDate::from_ymd_opt(2026, 6, 10).unwrap().and_hms_opt(14, 30, 0).unwrap();
        assert_eq!(jet_datetime(&dt), "06/10/2026 02:30 PM");
    }

    #[test]
    fn safe_filename_strips_unsafe_characters() {
        assert_eq!(safe_filename("report:final*.pdf"), "report_final_.pdf");
        assert_eq!(safe_filename("   "), "attachment");
        assert_eq!(safe_filename(""), "attachment");
    }

    #[test]
    fn categories_round_trip_and_trim() {
        assert_eq!(parse_categories("Work, Receipts"), vec!["Work", "Receipts"]);
        assert_eq!(parse_categories("  Work ,  Personal "), vec!["Work", "Personal"]);
        assert_eq!(parse_categories(""), Vec::<String>::new());
        assert_eq!(join_categories(&["Work".into(), "Personal".into()]), "Work, Personal");
        assert_eq!(join_categories(&[]), "");
    }

    #[test]
    fn clean_content_id_strips_whitespace_and_angle_brackets() {
        assert_eq!(clean_content_id(Some("<image001.png@01D9>")), Some("image001.png@01D9".to_string()));
        assert_eq!(clean_content_id(Some("  < abc@x >  ")), Some("abc@x".to_string()));
        assert_eq!(clean_content_id(Some("plain-id")), Some("plain-id".to_string()));
        // Only a matched pair is stripped.
        assert_eq!(clean_content_id(Some("<half")), Some("<half".to_string()));
        assert_eq!(clean_content_id(Some("<>")), None);
        assert_eq!(clean_content_id(Some("   ")), None);
        assert_eq!(clean_content_id(Some("")), None);
        assert_eq!(clean_content_id(None), None);
    }

    #[test]
    fn normalize_cid_request_strips_cid_prefix_and_brackets() {
        assert_eq!(normalize_cid_request("image001.png@01D9"), Some("image001.png@01D9".to_string()));
        assert_eq!(normalize_cid_request("cid:image001.png@01D9"), Some("image001.png@01D9".to_string()));
        assert_eq!(normalize_cid_request("  CID:<abc@x>  "), Some("abc@x".to_string()));
        assert_eq!(normalize_cid_request("Cid: <abc@x>"), Some("abc@x".to_string()));
        assert_eq!(normalize_cid_request("<abc@x>"), Some("abc@x".to_string()));
        // Only a leading prefix is stripped, and case is preserved.
        assert_eq!(normalize_cid_request("Logo-cid:A@B"), Some("Logo-cid:A@B".to_string()));
        assert_eq!(normalize_cid_request("cid:"), None);
        assert_eq!(normalize_cid_request("cid:<>"), None);
        assert_eq!(normalize_cid_request("  "), None);
        assert_eq!(normalize_cid_request(""), None);
        // Multi-byte input shorter than / straddling the prefix length is fine.
        assert_eq!(normalize_cid_request("é@x"), Some("é@x".to_string()));
        assert_eq!(normalize_cid_request("ciéd"), Some("ciéd".to_string()));
    }

    #[test]
    fn guess_mime_prefers_the_tag_lowercased() {
        assert_eq!(guess_mime(Some(" Image/PNG "), "photo.jpg"), Some("image/png".to_string()));
        assert_eq!(guess_mime(Some("  "), "photo.jpg"), Some("image/jpeg".to_string()));
        assert_eq!(guess_mime(None, "photo.jpg"), Some("image/jpeg".to_string()));
    }

    #[test]
    fn guess_mime_maps_common_extensions_case_insensitively() {
        assert_eq!(guess_mime(None, "a.PNG"), Some("image/png".to_string()));
        assert_eq!(guess_mime(None, "a.jpeg"), Some("image/jpeg".to_string()));
        assert_eq!(guess_mime(None, "a.svg"), Some("image/svg+xml".to_string()));
        assert_eq!(guess_mime(None, "report.final.pdf"), Some("application/pdf".to_string()));
        assert_eq!(guess_mime(None, "notes.txt"), Some("text/plain".to_string()));
        assert_eq!(
            guess_mime(None, "deck.pptx"),
            Some("application/vnd.openxmlformats-officedocument.presentationml.presentation".to_string())
        );
        assert_eq!(guess_mime(None, "fwd.eml"), Some("message/rfc822".to_string()));
        assert_eq!(guess_mime(None, "fwd.msg"), Some("application/vnd.ms-outlook".to_string()));
    }

    #[test]
    fn guess_mime_returns_none_when_unknown() {
        assert_eq!(guess_mime(None, "archive.xyz"), None);
        assert_eq!(guess_mime(None, "no-extension"), None);
        assert_eq!(guess_mime(None, ""), None);
    }

    #[test]
    fn is_inline_needs_a_content_id() {
        assert!(!is_inline(None, false, r#"<img src="cid:logo@x">"#));
        assert!(!is_inline(Some(""), false, r#"<img src="cid:logo@x">"#));
    }

    #[test]
    fn is_inline_when_referenced_from_html() {
        assert!(is_inline(Some("logo@x"), false, r#"<img src="cid:logo@x">"#));
        assert!(is_inline(Some("logo@x"), false, "cid:logo@x"));
        assert!(is_inline(Some("logo@x"), false, "url(cid:logo@x)"));
        assert!(is_inline(Some("logo@x"), false, "<img src='cid:logo@x' >"));
    }

    #[test]
    fn is_inline_reference_is_case_insensitive() {
        assert!(is_inline(Some("Image001.PNG@01D9"), false, r#"<img src="CID:image001.png@01d9">"#));
    }

    #[test]
    fn is_inline_cid_not_referenced_is_regular() {
        assert!(!is_inline(Some("logo@x"), false, "<p>no images here</p>"));
        assert!(!is_inline(Some("logo@x"), false, ""));
    }

    #[test]
    fn is_inline_requires_the_whole_id() {
        // `cid:img10@x` is not a reference to `img1@x`, nor `cid:img1@xyz`.
        assert!(!is_inline(Some("img1@x"), false, r#"<img src="cid:img10@x">"#));
        assert!(!is_inline(Some("img1@x"), false, r#"<img src="cid:img1@xyz">"#));
        // ...but a later whole reference still counts.
        assert!(is_inline(Some("img1@x"), false, r#"<img src="cid:img1@xyz"><img src="cid:img1@x">"#));
    }

    #[test]
    fn is_inline_hidden() {
        assert!(!is_inline(None, true, ""));
        assert!(is_inline(Some("logo@x"), true, ""));
    }

    #[test]
    fn string_from_utf16_is_lossless_for_hebrew_and_emoji() {
        for text in ["מייל שיקוף", "סיכום עשייה", "Re: סיכום (Q3) 📧", ""] {
            let wide: Vec<u16> = text.encode_utf16().collect();
            assert_eq!(string_from_utf16(&wide), text);
        }
    }

    #[test]
    fn string_from_utf16_replaces_only_unpaired_surrogates() {
        // "מ", a lone high surrogate, "x": only the surrogate is replaced.
        assert_eq!(string_from_utf16(&[0x05DE, 0xD83D, 0x0078]), "מ\u{FFFD}x");
        // A lone low surrogate at the end.
        assert_eq!(string_from_utf16(&[0x0061, 0xDCE7]), "a\u{FFFD}");
    }

    #[test]
    fn variant_to_string_round_trips_hebrew_bstr() {
        let text = "מייל שיקוף — סיכום עשייה";
        assert_eq!(variant_to_string(&variant_from_str(text)), text);
    }

    #[test]
    fn variant_to_iso_string_round_trips_a_vt_date_variant() {
        use chrono::NaiveDate;
        let dt = NaiveDate::from_ymd_opt(2026, 6, 10).unwrap().and_hms_opt(14, 30, 0).unwrap();
        let v = variant_from_datetime(&dt).expect("variant_from_datetime should succeed");
        assert_eq!(variant_to_iso_string(&v), Some("2026-06-10T14:30:00".to_string()));
    }

    // Guards the non-VT_DATE fallback path: the fix branches on `vt`, so this
    // pins that a VARIANT the fallback can't decode still yields `None` (not
    // garbage). A non-numeric string is used because the fallback is
    // `f64::try_from` == Win32 `VariantToDouble`, which *coerces* numeric
    // VARIANTs (e.g. VT_I4 42 -> 42.0 -> a valid-but-bogus date) but rejects a
    // non-numeric string with a type mismatch. That coercion is pre-existing
    // Win32 behavior, unchanged by this fix (the fallback branch is a verbatim
    // `f64::try_from`).
    #[test]
    fn variant_to_iso_string_returns_none_for_non_date_variant() {
        let v = variant_from_str("not a date");
        assert_eq!(variant_to_iso_string(&v), None);
    }
}

// ---------------------------------------------------------------------------
// Real Win32 COM interop below. Verified directly against the installed
// `windows` 0.62.2 / `windows-result` 0.4.1 crate source (not just docs),
// since this is the highest-risk, least-forgiving part of the port.
// ---------------------------------------------------------------------------

use windows::core::{Error as WinError, Result as WinResult, BSTR, GUID, PCWSTR};
use windows::Win32::Globalization::{GetLocaleInfoEx, GetUserDefaultLCID, LOCALE_IFIRSTDAYOFWEEK};
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoUninitialize, CLSIDFromProgID, IDispatch,
    CLSCTX_LOCAL_SERVER, COINIT_APARTMENTTHREADED, DISPATCH_METHOD, DISPATCH_PROPERTYGET,
    DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO,
};
use windows::Win32::System::Variant::{
    SystemTimeToVariantTime, VariantTimeToSystemTime, VARIANT, VT_DATE,
};
use windows::Win32::Foundation::SYSTEMTIME;

/// The first day of the week from the current Windows user's regional
/// settings (`GetLocaleInfoEx(LOCALE_NAME_USER_DEFAULT, LOCALE_IFIRSTDAYOFWEEK)`),
/// which is what `start_of_week` / `end_of_week` in the date grammar use.
/// Looked up on every call (it is cheap, and the user can change the setting
/// while the server runs). Falls back to Monday (ISO 8601) only if the API
/// fails or returns an unexpected value.
pub fn user_first_day_of_week() -> chrono::Weekday {
    let mut buf = [0u16; 8];
    // A null locale name is LOCALE_NAME_USER_DEFAULT. The return value is the
    // number of UTF-16 units written, including the terminating NUL; 0 = error.
    let written = unsafe { GetLocaleInfoEx(PCWSTR::null(), LOCALE_IFIRSTDAYOFWEEK, Some(&mut buf)) };
    let len = usize::try_from(written).unwrap_or(0).min(buf.len());
    first_day_of_week_from_locale_value(&string_from_utf16(&buf[..len])).unwrap_or(chrono::Weekday::Mon)
}

/// One per COM call (mirrors `pythoncom.CoInitialize()` inside `client.py`'s
/// `@_com` decorator): initializes this OS thread for apartment-threaded COM
/// on construction, uninitializes on drop. Must be created on the same
/// thread `spawn_blocking`'s closure runs on (see `run_blocking` in
/// `src/server.rs`), and must outlive every COM object used within that call.
pub struct ComGuard;

impl ComGuard {
    pub fn new() -> WinResult<Self> {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) }.ok()?;
        Ok(ComGuard)
    }
}

impl Drop for ComGuard {
    fn drop(&mut self) {
        unsafe { CoUninitialize() };
    }
}

pub fn create_com_object(prog_id: &str) -> WinResult<IDispatch> {
    let wide: Vec<u16> = prog_id.encode_utf16().chain(std::iter::once(0)).collect();
    let clsid: GUID = unsafe { CLSIDFromProgID(PCWSTR(wide.as_ptr()))? };
    unsafe { CoCreateInstance(&clsid, None, CLSCTX_LOCAL_SERVER) }
}

fn name_to_dispid(disp: &IDispatch, name: &str) -> WinResult<i32> {
    let wide: Vec<u16> = name.encode_utf16().chain(std::iter::once(0)).collect();
    let names = [PCWSTR(wide.as_ptr())];
    let mut dispid = 0i32;
    unsafe {
        disp.GetIDsOfNames(&GUID::zeroed(), names.as_ptr(), 1, GetUserDefaultLCID(), &mut dispid)?;
    }
    Ok(dispid)
}

const DISP_E_EXCEPTION: i32 = -2147352567; // 0x80020009, from winerror.h

fn enrich_error(err: WinError, excepinfo: &EXCEPINFO) -> WinError {
    if err.code().0 == DISP_E_EXCEPTION && !excepinfo.bstrDescription.is_empty() {
        return WinError::new(err.code(), excepinfo.bstrDescription.to_string());
    }
    err
}

/// `EXCEPINFO`'s three `BSTR` fields are allocated by the callee (Outlook)
/// whenever `Invoke` reports `DISP_E_EXCEPTION`; nothing else frees them.
/// Safe to call unconditionally — `SysFreeString` on a null/empty `BSTR` is
/// a documented no-op, so this is harmless on the (common) success path
/// where they were never populated.
fn free_excepinfo(excepinfo: &mut EXCEPINFO) {
    unsafe {
        let _ = std::mem::ManuallyDrop::take(&mut excepinfo.bstrSource);
        let _ = std::mem::ManuallyDrop::take(&mut excepinfo.bstrDescription);
        let _ = std::mem::ManuallyDrop::take(&mut excepinfo.bstrHelpFile);
    }
}

fn invoke(
    disp: &IDispatch,
    name: &str,
    flags: windows::Win32::System::Com::DISPATCH_FLAGS,
    args: &mut [VARIANT],
) -> WinResult<VARIANT> {
    let dispid = name_to_dispid(disp, name)?;
    // COM wants arguments in reverse order. This mutates the caller's slice
    // in place; harmless today since every call site builds a fresh,
    // single-use temporary array, but worth knowing if that ever changes.
    args.reverse();
    let is_put = flags == DISPATCH_PROPERTYPUT;
    let mut put_dispid: i32 = -3; // DISPID_PROPERTYPUT
    let params = DISPPARAMS {
        rgvarg: args.as_mut_ptr(),
        rgdispidNamedArgs: if is_put { &mut put_dispid } else { std::ptr::null_mut() },
        cArgs: args.len() as u32,
        cNamedArgs: if is_put { 1 } else { 0 },
    };
    let mut result = VARIANT::default();
    let mut excepinfo = EXCEPINFO::default();
    let mut arg_err = 0u32;
    let outcome = unsafe {
        disp.Invoke(
            dispid, &GUID::zeroed(), GetUserDefaultLCID(), flags,
            &params, Some(&mut result), Some(&mut excepinfo), Some(&mut arg_err),
        )
    }
    .map_err(|e| enrich_error(e, &excepinfo));
    free_excepinfo(&mut excepinfo);
    outcome?;
    Ok(result)
}

pub fn get_property(disp: &IDispatch, name: &str) -> WinResult<VARIANT> {
    invoke(disp, name, DISPATCH_PROPERTYGET, &mut [])
}

pub fn put_property(disp: &IDispatch, name: &str, value: VARIANT) -> WinResult<()> {
    invoke(disp, name, DISPATCH_PROPERTYPUT, &mut [value])?;
    Ok(())
}

pub fn call_method(disp: &IDispatch, name: &str, args: &mut [VARIANT]) -> WinResult<VARIANT> {
    invoke(disp, name, DISPATCH_METHOD, args)
}

/// Read a MAPI property by DASL schema name via
/// `disp.PropertyAccessor.GetProperty(schema)`. Returns `None` on any error:
/// Outlook raises a COM error for a property that isn't set, and some objects
/// have no `PropertyAccessor` at all.
pub fn get_mapi_prop(disp: &IDispatch, schema: &str) -> Option<VARIANT> {
    let accessor = IDispatch::try_from(&get_property(disp, "PropertyAccessor").ok()?).ok()?;
    call_method(&accessor, "GetProperty", &mut [variant_from_str(schema)]).ok()
}

/// Read an item's color categories (empty vec if the property is missing or blank).
pub fn get_item_categories(disp: &IDispatch) -> Vec<String> {
    let raw = get_property(disp, "Categories")
        .map(|v| variant_to_string(&v))
        .unwrap_or_default();
    parse_categories(&raw)
}

/// Overwrite an item's categories with the given list.
pub fn set_item_categories(disp: &IDispatch, cats: &[String]) -> WinResult<()> {
    put_property(disp, "Categories", variant_from_str(&join_categories(cats)))
}

/// Mirrors Python's `hasattr(obj, name)` for a COM dispatch member: resolves
/// the name via `GetIDsOfNames`, returning `false` when the member doesn't
/// exist. Used by `respond_to_meeting` to distinguish a `MeetingItem` (which
/// exposes `GetAssociatedAppointment`) from an `AppointmentItem` (which does
/// not) without invoking the method for its side effect.
pub fn has_member(disp: &IDispatch, name: &str) -> bool {
    name_to_dispid(disp, name).is_ok()
}

/// Translation of `outlook_mcp/errors.py::format_com_error`. Errors enriched
/// by `enrich_error` above already carry the COM exception's own description
/// text in `message()` (equivalent to Python's `excepinfo[2]`).
pub fn format_com_error(err: &WinError) -> String {
    format!("Outlook error: {} (HRESULT {:#010x})", err.message(), err.code().0)
}

// ---- VARIANT conversions ---------------------------------------------
//
// `VARIANT::from(i32)`/`from(bool)` are ungated — generated by the crate's
// internal `variant_from_value!(i32, VT_I4, lVal, ...)` /
// `variant_from_value!(bool, VT_BOOL, boolVal, ...)` macro invocations,
// invisible to a plain-text `impl From<` search but present unconditionally.
// The `Win32_System_Com_StructuredStorage` feature (enabled in Cargo.toml)
// is only needed for `TryFrom<&VARIANT> for BSTR` and `Display for VARIANT`,
// both used by `variant_to_string` below.

pub fn variant_from_str(value: &str) -> VARIANT {
    VARIANT::from(value)
}

pub fn variant_from_i32(value: i32) -> VARIANT {
    VARIANT::from(value)
}

pub fn variant_from_bool(value: bool) -> VARIANT {
    VARIANT::from(value)
}

/// Builds a `VT_DATE` VARIANT (OLE Automation date) from a `NaiveDateTime`,
/// so Outlook properties typed as `Date` (e.g. `AppointmentItem.Start`/`End`)
/// receive the exact type pywin32 marshals a Python `datetime` to. There is no
/// `From<…> for VARIANT` that yields `VT_DATE` (a plain `f64` becomes `VT_R8`),
/// so this converts via `SystemTimeToVariantTime` and writes the union by hand.
pub fn variant_from_datetime(dt: &chrono::NaiveDateTime) -> WinResult<VARIANT> {
    use chrono::{Datelike, Timelike};
    let st = SYSTEMTIME {
        wYear: dt.year() as u16,
        wMonth: dt.month() as u16,
        wDayOfWeek: 0,
        wDay: dt.day() as u16,
        wHour: dt.hour() as u16,
        wMinute: dt.minute() as u16,
        wSecond: dt.second() as u16,
        wMilliseconds: 0,
    };
    let mut date: f64 = 0.0;
    // Returns nonzero on success; 0 signals an out-of-range date.
    if unsafe { SystemTimeToVariantTime(&st, &mut date) } == 0 {
        return Err(WinError::new(
            windows::core::HRESULT(-2147024809), // E_INVALIDARG (0x80070057)
            format!("Could not convert {dt} to an OLE Automation date"),
        ));
    }
    let mut variant = VARIANT::default();
    unsafe {
        let inner = &mut *variant.Anonymous.Anonymous;
        inner.vt = VT_DATE;
        inner.Anonymous.date = date;
    }
    Ok(variant)
}

/// For VT_BSTR-typed properties (Outlook `Subject`/`Name`/etc.). Returns an
/// empty string if the VARIANT isn't a string — use `variant_to_i32`/
/// `variant_to_bool`/`variant_to_iso_string` for other VT kinds instead of
/// relying on this as a general-purpose fallback. The text stays UTF-16 until
/// `string_from_utf16`, so no code page is involved (see its doc comment).
pub fn variant_to_string(value: &VARIANT) -> String {
    BSTR::try_from(value).map(|b| string_from_utf16(&b)).unwrap_or_default()
}

pub fn variant_to_i32(value: &VARIANT) -> Option<i32> {
    i32::try_from(value).ok()
}

pub fn variant_to_bool(value: &VARIANT) -> Option<bool> {
    bool::try_from(value).ok()
}

/// Converts a VT_DATE-typed VARIANT (OLE Automation date: an f64 count of
/// days since 1899-12-30) to an ISO-8601 string, mirroring `_to_iso` in
/// `outlook_mcp/outlook/client.py`. Returns `None` if the VARIANT isn't a
/// date the Win32 `VariantTimeToSystemTime` call can decode.
pub fn variant_to_iso_string(value: &VARIANT) -> Option<String> {
    // Outlook returns every Date-typed property as VT_DATE, but the crate's
    // `f64::try_from(&VARIANT)` only accepts VT_R8 — it rejects VT_DATE with a
    // type mismatch. Read the OLE Automation date out of the union directly for
    // VT_DATE (mirroring how `variant_from_datetime` writes it), and keep the
    // type-checked `f64::try_from` path for any VT_R8 caller.
    let date = unsafe {
        let inner = &*value.Anonymous.Anonymous;
        if inner.vt == VT_DATE {
            inner.Anonymous.date
        } else {
            f64::try_from(value).ok()?
        }
    };
    let mut sys_time = SYSTEMTIME::default();
    unsafe {
        if VariantTimeToSystemTime(date, &mut sys_time) == 0 {
            return None;
        }
    }
    chrono::NaiveDate::from_ymd_opt(sys_time.wYear as i32, sys_time.wMonth as u32, sys_time.wDay as u32)
        .and_then(|d| {
            d.and_hms_milli_opt(
                sys_time.wHour as u32,
                sys_time.wMinute as u32,
                sys_time.wSecond as u32,
                sys_time.wMilliseconds as u32,
            )
        })
        .map(|dt| dt.format("%Y-%m-%dT%H:%M:%S").to_string())
}
