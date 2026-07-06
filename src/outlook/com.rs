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

#[cfg(test)]
mod tests {
    use super::*;

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
}

// ---------------------------------------------------------------------------
// Real Win32 COM interop below. Verified directly against the installed
// `windows` 0.62.2 / `windows-result` 0.4.1 crate source (not just docs),
// since this is the highest-risk, least-forgiving part of the port.
// ---------------------------------------------------------------------------

use windows::core::{Error as WinError, Result as WinResult, BSTR, GUID, PCWSTR};
use windows::Win32::Globalization::GetUserDefaultLCID;
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoUninitialize, CLSIDFromProgID, IDispatch,
    CLSCTX_LOCAL_SERVER, COINIT_APARTMENTTHREADED, DISPATCH_METHOD, DISPATCH_PROPERTYGET,
    DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO,
};
use windows::Win32::System::Variant::{VARIANT, VariantTimeToSystemTime};
use windows::Win32::Foundation::SYSTEMTIME;

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

fn invoke(
    disp: &IDispatch,
    name: &str,
    flags: windows::Win32::System::Com::DISPATCH_FLAGS,
    args: &mut [VARIANT],
) -> WinResult<VARIANT> {
    let dispid = name_to_dispid(disp, name)?;
    args.reverse(); // COM wants arguments in reverse order
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
    unsafe {
        disp.Invoke(
            dispid, &GUID::zeroed(), GetUserDefaultLCID(), flags,
            &params, Some(&mut result), Some(&mut excepinfo), Some(&mut arg_err),
        )
    }
    .map_err(|e| enrich_error(e, &excepinfo))?;
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

/// Translation of `outlook_mcp/errors.py::format_com_error`. Errors enriched
/// by `enrich_error` above already carry the COM exception's own description
/// text in `message()` (equivalent to Python's `excepinfo[2]`).
pub fn format_com_error(err: &WinError) -> String {
    format!("Outlook error: {} (HRESULT {:#010x})", err.message(), err.code().0)
}

// ---- VARIANT conversions ---------------------------------------------
//
// `windows` needs the `Win32_System_Com_StructuredStorage` feature enabled
// for VARIANT to get `From<i32>`/`From<bool>` (both confirmed present via
// the crate's internal `variant_from_value!` macro — `variant_from_value!
// (i32, VT_I4, lVal, ...)`, `variant_from_value!(bool, VT_BOOL, boolVal,
// ...)`), `TryFrom<&VARIANT>` for the primitive types, and `Display`. That
// feature is enabled in Cargo.toml specifically for this.

pub fn variant_from_str(value: &str) -> VARIANT {
    VARIANT::from(value)
}

pub fn variant_from_i32(value: i32) -> VARIANT {
    VARIANT::from(value)
}

pub fn variant_from_bool(value: bool) -> VARIANT {
    VARIANT::from(value)
}

/// Best-effort string conversion: tries a real string value first (the
/// common case for Outlook `Subject`/`Name`/etc. properties), and falls back
/// to VARIANT's own `Display` impl (which handles numeric/date/other VT
/// kinds reasonably) if the value isn't VT_BSTR.
pub fn variant_to_string(value: &VARIANT) -> String {
    BSTR::try_from(value).map(|b| b.to_string()).unwrap_or_else(|_| value.to_string())
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
    let date = f64::try_from(value).ok()?;
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
