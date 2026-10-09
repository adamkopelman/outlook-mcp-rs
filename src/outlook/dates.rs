//! The one date grammar every date parameter on every tool accepts (issue #30).
//!
//! ```text
//! date    := iso | anchor offset* | offset+
//! iso     := YYYY-MM-DD | YYYY-MM-DD(T| )HH:MM[:SS[.fff]]       (local time)
//! anchor  := now | today | yesterday | tomorrow
//!          | start_of_week | end_of_week | start_of_month | end_of_month
//!          | start_of_year | end_of_year
//! offset  := (+|-) digits unit
//! unit    := m (minutes) | h (hours) | d (days) | w (weeks) | mo (months) | y (years)
//! ```
//!
//! - Keywords and units are case-insensitive; whitespace between the parts of
//!   a keyword expression is ignored (`"today - 1d"` = `"today-1d"`).
//! - An offset with no anchor is relative to `now` (the current local time,
//!   not midnight): `-14d` is exactly 14 days ago.
//! - `today`/`yesterday`/`tomorrow` and every `start_of_*` are midnight
//!   (00:00:00); every `end_of_*` is the last second of that period
//!   (23:59:59 on its last day).
//! - Weeks start on the first day of the week from the Windows user's
//!   regional settings (`LOCALE_IFIRSTDAYOFWEEK`: Monday in most of Europe,
//!   Sunday in the US, Saturday in some Middle-East locales). The parser
//!   takes it as an input; the COM client reads it per call with
//!   `com::user_first_day_of_week` (Monday only if that lookup fails), and
//!   the in-memory fake always uses Monday so tests are deterministic.
//! - `mo`/`y` move by calendar months/years, clamping the day to the target
//!   month's length (`2026-03-31` + `1mo` = `2026-04-30`).
//! - Timezone suffixes (`Z`, `+02:00`) are not accepted: every value is local
//!   time, as Outlook shows it.
//!
//! Everything here is pure (the caller passes `now` and the week start) so
//! it can be unit tested; [`parse_date_param`] is the wrapper the COM client uses.
//!
//! The result is a `NaiveDateTime`; turning it into a `Restrict` filter
//! string is `com::jet_datetime`'s job (and its day-first-locale bug is
//! issue #1, not this module's).

use chrono::{Datelike, Duration, Months, NaiveDate, NaiveDateTime, NaiveTime, Weekday};

use crate::error::ToolError;

/// One parsed date value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DateValue {
    pub at: NaiveDateTime,
    /// True only for a bare ISO date (`2026-06-10`): it names a whole day,
    /// which matters for an inclusive upper bound (see [`DateValue::upper_bound`]).
    pub date_only: bool,
}

impl DateValue {
    /// The value as the upper bound of a `*_before` range filter: a bare ISO
    /// date includes that whole day (`received_before: "2026-06-30"` keeps
    /// mail from June 30th); anything else is used as-is (`"today"` is
    /// today's midnight, so `received_before: "today"` means before today).
    pub fn upper_bound(self) -> NaiveDateTime {
        if self.date_only {
            self.at.date().and_hms_micro_opt(23, 59, 59, 999_999).unwrap()
        } else {
            self.at
        }
    }
}

/// Summary of the grammar for error messages and tool descriptions.
pub const DATE_GRAMMAR_HELP: &str = "use an ISO local date/datetime like '2026-06-10' or \
'2026-06-10T14:30', a keyword (now, today, yesterday, tomorrow, start_of_week, end_of_week, \
start_of_month, end_of_month, start_of_year, end_of_year) optionally followed by offsets like \
'today-1d', or an offset from now like '-14d', '+3h', '-2w' (units: m=minutes, h=hours, \
d=days, w=weeks, mo=months, y=years)";

const KEYWORDS: [&str; 10] = [
    "start_of_month", "end_of_month", "start_of_week", "end_of_week", "start_of_year",
    "end_of_year", "yesterday", "tomorrow", "today", "now",
];

fn midnight(d: NaiveDate) -> NaiveDateTime {
    d.and_time(NaiveTime::MIN)
}

/// Last second of day `d`.
fn last_second(d: NaiveDate) -> NaiveDateTime {
    d.and_hms_opt(23, 59, 59).unwrap()
}

fn parse_iso(s: &str) -> Option<DateValue> {
    // A single space separator is accepted like `T`.
    let normalized = s.replacen(' ', "T", 1);
    for fmt in ["%Y-%m-%dT%H:%M:%S%.f", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M"] {
        if let Ok(at) = NaiveDateTime::parse_from_str(&normalized, fmt) {
            return Some(DateValue { at, date_only: false });
        }
    }
    NaiveDate::parse_from_str(s, "%Y-%m-%d")
        .ok()
        .map(|d| DateValue { at: midnight(d), date_only: true })
}

/// The first day (`week_start`) of the week containing `day`.
fn start_of_week(day: NaiveDate, week_start: Weekday) -> NaiveDate {
    day - Duration::days(day.weekday().days_since(week_start) as i64)
}

fn anchor(keyword: &str, now: NaiveDateTime, first_day: Weekday) -> NaiveDateTime {
    let today = now.date();
    let week_start = start_of_week(today, first_day);
    let month_start = today.with_day(1).unwrap();
    let year_start = NaiveDate::from_ymd_opt(today.year(), 1, 1).unwrap();
    match keyword {
        "now" => now,
        "today" => midnight(today),
        "yesterday" => midnight(today - Duration::days(1)),
        "tomorrow" => midnight(today + Duration::days(1)),
        "start_of_week" => midnight(week_start),
        "end_of_week" => last_second(week_start + Duration::days(6)),
        "start_of_month" => midnight(month_start),
        "end_of_month" => {
            last_second(month_start.checked_add_months(Months::new(1)).unwrap() - Duration::days(1))
        }
        "start_of_year" => midnight(year_start),
        "end_of_year" => last_second(NaiveDate::from_ymd_opt(today.year(), 12, 31).unwrap()),
        _ => unreachable!("anchor() is only called with a KEYWORDS entry"),
    }
}

/// Applies one signed offset to `at`; `None` when the result is out of range.
fn apply_offset(at: NaiveDateTime, negative: bool, amount: u32, unit: &str) -> Option<NaiveDateTime> {
    let n = amount as i64;
    let delta = match unit {
        "m" => Duration::try_minutes(n)?,
        "h" => Duration::try_hours(n)?,
        "d" => Duration::try_days(n)?,
        "w" => Duration::try_weeks(n)?,
        "mo" | "y" => {
            let months = if unit == "y" { amount.checked_mul(12)? } else { amount };
            return if negative {
                at.checked_sub_months(Months::new(months))
            } else {
                at.checked_add_months(Months::new(months))
            };
        }
        _ => return None,
    };
    if negative { at.checked_sub_signed(delta) } else { at.checked_add_signed(delta) }
}

/// Parses one date expression (see the module docs) relative to `now`, with
/// weeks starting on `week_start` (for `start_of_week` / `end_of_week`).
/// The error is a short reason, without the field name or the grammar help.
pub fn parse_date_value(input: &str, now: NaiveDateTime, week_start: Weekday) -> Result<DateValue, String> {
    let trimmed = input.trim();
    if trimmed.is_empty() {
        return Err("empty value".to_string());
    }
    if let Some(v) = parse_iso(trimmed) {
        return Ok(v);
    }
    let s: String = trimmed.chars().filter(|c| !c.is_whitespace()).collect::<String>().to_lowercase();
    let (mut at, mut rest) = if s.starts_with(['+', '-']) {
        (now, s.as_str())
    } else {
        match KEYWORDS.iter().find(|k| s.starts_with(*k)) {
            Some(k) => (anchor(k, now, week_start), &s[k.len()..]),
            None => return Err("not a recognized date".to_string()),
        }
    };
    while !rest.is_empty() {
        let negative = match rest.as_bytes()[0] {
            b'+' => false,
            b'-' => true,
            _ => return Err(format!("unexpected {rest:?} (an offset must start with + or -)")),
        };
        rest = &rest[1..];
        let digits = rest.bytes().take_while(u8::is_ascii_digit).count();
        if digits == 0 {
            return Err("an offset needs a number after its sign, like '-14d'".to_string());
        }
        let amount: u32 = rest[..digits].parse().map_err(|_| "offset is too large".to_string())?;
        rest = &rest[digits..];
        let unit_len = rest.bytes().take_while(u8::is_ascii_alphabetic).count();
        let unit = &rest[..unit_len];
        if !matches!(unit, "m" | "h" | "d" | "w" | "mo" | "y") {
            return Err(format!(
                "unknown offset unit {unit:?} (use m, h, d, w, mo or y)"
            ));
        }
        rest = &rest[unit_len..];
        at = apply_offset(at, negative, amount, unit).ok_or_else(|| "offset is too large".to_string())?;
    }
    Ok(DateValue { at, date_only: false })
}

/// [`parse_date_value`] relative to the current local time, with a
/// [`ToolError`] naming the parameter on failure.
pub fn parse_date_param(value: &str, field: &str, week_start: Weekday) -> Result<DateValue, ToolError> {
    parse_date_param_at(value, field, chrono::Local::now().naive_local(), week_start)
}

/// [`parse_date_value`] with a [`ToolError`] naming the parameter on failure.
pub fn parse_date_param_at(value: &str, field: &str, now: NaiveDateTime, week_start: Weekday)
    -> Result<DateValue, ToolError> {
    parse_date_value(value, now, week_start)
        .map_err(|reason| ToolError::new(format!("Invalid {field} {value:?}: {reason}; {DATE_GRAMMAR_HELP}")))
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Thursday 2026-10-08 14:30:15.
    fn now() -> NaiveDateTime {
        NaiveDate::from_ymd_opt(2026, 10, 8).unwrap().and_hms_opt(14, 30, 15).unwrap()
    }

    fn dt(s: &str) -> NaiveDateTime {
        NaiveDateTime::parse_from_str(s, "%Y-%m-%dT%H:%M:%S").unwrap()
    }

    fn at(s: &str) -> NaiveDateTime {
        parse_date_value(s, now(), Weekday::Mon).unwrap_or_else(|e| panic!("{s:?}: {e}")).at
    }

    fn err(s: &str) -> String {
        parse_date_value(s, now(), Weekday::Mon).unwrap_err()
    }

    #[test]
    fn iso_dates_and_datetimes() {
        let v = parse_date_value("2026-06-10", now(), Weekday::Mon).unwrap();
        assert_eq!(v, DateValue { at: dt("2026-06-10T00:00:00"), date_only: true });
        assert_eq!(at("2026-06-10T14:30"), dt("2026-06-10T14:30:00"));
        assert_eq!(at("2026-06-10 14:30"), dt("2026-06-10T14:30:00"));
        assert_eq!(at("2026-06-10T14:30:45"), dt("2026-06-10T14:30:45"));
        assert_eq!(at("  2026-06-10T14:30:45.250 "), dt("2026-06-10T14:30:45") + Duration::milliseconds(250));
        assert!(!parse_date_value("2026-06-10T00:00", now(), Weekday::Mon).unwrap().date_only);
    }

    #[test]
    fn iso_is_year_month_day_never_day_first() {
        // 2026-11-09 is 9 November, not 11 September (issue #1's symptom).
        assert_eq!(at("2026-11-09"), dt("2026-11-09T00:00:00"));
        assert!(parse_date_value("2026-13-01", now(), Weekday::Mon).is_err());
        assert!(parse_date_value("2026-02-30", now(), Weekday::Mon).is_err());
    }

    #[test]
    fn day_keywords_are_midnight() {
        assert_eq!(at("today"), dt("2026-10-08T00:00:00"));
        assert_eq!(at("yesterday"), dt("2026-10-07T00:00:00"));
        assert_eq!(at("tomorrow"), dt("2026-10-09T00:00:00"));
        assert_eq!(at("now"), now());
        assert_eq!(at("TODAY"), dt("2026-10-08T00:00:00"));
    }

    /// `(start_of_week, end_of_week)` for `now` = `day` at 13:00.
    fn week(day: &str, first: Weekday) -> (NaiveDateTime, NaiveDateTime) {
        let now = dt(&format!("{day}T13:00:00"));
        let get = |s: &str| parse_date_value(s, now, first).unwrap_or_else(|e| panic!("{s:?}: {e}")).at;
        (get("start_of_week"), get("end_of_week"))
    }

    #[test]
    fn week_starting_monday() {
        // 2026-10-08 is a Thursday.
        assert_eq!(at("start_of_week"), dt("2026-10-05T00:00:00"));
        assert_eq!(at("end_of_week"), dt("2026-10-11T23:59:59"));
        let this_week = (dt("2026-10-05T00:00:00"), dt("2026-10-11T23:59:59"));
        assert_eq!(week("2026-10-05", Weekday::Mon), this_week); // on the start day
        assert_eq!(week("2026-10-11", Weekday::Mon), this_week); // Sunday, the last day
        // The Sunday before the Monday start belongs to the previous week.
        assert_eq!(week("2026-10-04", Weekday::Mon), (dt("2026-09-28T00:00:00"), dt("2026-10-04T23:59:59")));
    }

    #[test]
    fn week_starting_sunday() {
        let this_week = (dt("2026-10-04T00:00:00"), dt("2026-10-10T23:59:59"));
        assert_eq!(week("2026-10-08", Weekday::Sun), this_week); // Thursday
        assert_eq!(week("2026-10-04", Weekday::Sun), this_week); // on the start day
        assert_eq!(week("2026-10-10", Weekday::Sun), this_week); // Saturday, the day before the next start
        assert_eq!(week("2026-10-11", Weekday::Sun), (dt("2026-10-11T00:00:00"), dt("2026-10-17T23:59:59")));
    }

    #[test]
    fn week_starting_saturday() {
        let this_week = (dt("2026-10-03T00:00:00"), dt("2026-10-09T23:59:59"));
        assert_eq!(week("2026-10-08", Weekday::Sat), this_week); // Thursday
        assert_eq!(week("2026-10-03", Weekday::Sat), this_week); // on the start day
        assert_eq!(week("2026-10-09", Weekday::Sat), this_week); // Friday, the day before the next start
        assert_eq!(week("2026-10-10", Weekday::Sat), (dt("2026-10-10T00:00:00"), dt("2026-10-16T23:59:59")));
        // A week across a month and year boundary.
        assert_eq!(week("2027-01-01", Weekday::Sat), (dt("2026-12-26T00:00:00"), dt("2027-01-01T23:59:59")));
    }

    #[test]
    fn week_start_reaches_combined_forms() {
        let now = now(); // Thursday 2026-10-08
        let get = |s: &str, first| parse_date_value(s, now, first).unwrap().at;
        assert_eq!(get("start_of_week-1w", Weekday::Mon), dt("2026-09-28T00:00:00"));
        assert_eq!(get("start_of_week-1w", Weekday::Sun), dt("2026-09-27T00:00:00"));
        assert_eq!(get("start_of_week-1w", Weekday::Sat), dt("2026-09-26T00:00:00"));
        assert_eq!(get("end_of_week+1w", Weekday::Sun), dt("2026-10-17T23:59:59"));
        assert_eq!(get("Start_Of_Week + 1d", Weekday::Sat), dt("2026-10-04T00:00:00"));
        // Nothing else depends on the week start.
        for s in ["today", "-14d", "start_of_month", "end_of_year", "2026-06-10"] {
            assert_eq!(get(s, Weekday::Sun), get(s, Weekday::Mon), "{s}");
        }
    }

    #[test]
    fn month_and_year_keywords() {
        assert_eq!(at("start_of_month"), dt("2026-10-01T00:00:00"));
        assert_eq!(at("end_of_month"), dt("2026-10-31T23:59:59"));
        assert_eq!(at("start_of_year"), dt("2026-01-01T00:00:00"));
        assert_eq!(at("end_of_year"), dt("2026-12-31T23:59:59"));
        let feb = NaiveDate::from_ymd_opt(2028, 2, 10).unwrap().and_hms_opt(0, 0, 0).unwrap();
        assert_eq!(parse_date_value("end_of_month", feb, Weekday::Mon).unwrap().at, dt("2028-02-29T23:59:59"));
        let dec = NaiveDate::from_ymd_opt(2026, 12, 31).unwrap().and_hms_opt(0, 0, 0).unwrap();
        assert_eq!(parse_date_value("end_of_month", dec, Weekday::Mon).unwrap().at, dt("2026-12-31T23:59:59"));
    }

    #[test]
    fn bare_offsets_are_relative_to_now() {
        assert_eq!(at("-14d"), dt("2026-09-24T14:30:15"));
        assert_eq!(at("+3h"), dt("2026-10-08T17:30:15"));
        assert_eq!(at("-2w"), dt("2026-09-24T14:30:15"));
        assert_eq!(at("-90m"), dt("2026-10-08T13:00:15"));
        assert_eq!(at("+1mo"), dt("2026-11-08T14:30:15"));
        assert_eq!(at("-1y"), dt("2025-10-08T14:30:15"));
        assert_eq!(at("-0d"), now());
    }

    #[test]
    fn anchors_combine_with_several_offsets() {
        assert_eq!(at("today-1d"), dt("2026-10-07T00:00:00"));
        assert_eq!(at("start_of_week-1w"), dt("2026-09-28T00:00:00"));
        assert_eq!(at("start_of_month-1mo"), dt("2026-09-01T00:00:00"));
        assert_eq!(at("tomorrow+9h+30m"), dt("2026-10-09T09:30:00"));
        assert_eq!(at("today +1d -2h"), dt("2026-10-08T22:00:00"));
        assert_eq!(at("-1d+2h"), dt("2026-10-07T16:30:15"));
        assert_eq!(at("Start_Of_Week + 1W"), dt("2026-10-12T00:00:00"));
    }

    #[test]
    fn month_offsets_clamp_the_day() {
        let mar31 = NaiveDate::from_ymd_opt(2026, 3, 31).unwrap().and_hms_opt(8, 0, 0).unwrap();
        assert_eq!(parse_date_value("+1mo", mar31, Weekday::Mon).unwrap().at, dt("2026-04-30T08:00:00"));
        assert_eq!(parse_date_value("-1mo", mar31, Weekday::Mon).unwrap().at, dt("2026-02-28T08:00:00"));
    }

    #[test]
    fn keyword_values_are_not_date_only() {
        assert!(!parse_date_value("today", now(), Weekday::Mon).unwrap().date_only);
        assert!(!parse_date_value("-1d", now(), Weekday::Mon).unwrap().date_only);
    }

    #[test]
    fn upper_bound_includes_a_whole_bare_date() {
        let v = parse_date_value("2026-06-30", now(), Weekday::Mon).unwrap();
        assert_eq!(v.upper_bound(), dt("2026-06-30T23:59:59") + Duration::microseconds(999_999));
        let v = parse_date_value("2026-06-30T12:00", now(), Weekday::Mon).unwrap();
        assert_eq!(v.upper_bound(), dt("2026-06-30T12:00:00"));
        let v = parse_date_value("today", now(), Weekday::Mon).unwrap();
        assert_eq!(v.upper_bound(), dt("2026-10-08T00:00:00"));
    }

    #[test]
    fn rejects_garbage_with_a_reason() {
        assert_eq!(err(""), "empty value");
        assert_eq!(err("   "), "empty value");
        assert_eq!(err("next tuesday"), "not a recognized date");
        assert_eq!(err("06/10/2026"), "not a recognized date");
        assert!(err("14d").contains("not a recognized date"));
        assert!(err("-d").contains("needs a number"));
        assert!(err("-14").contains("unknown offset unit"));
        assert!(err("-14x").contains("unknown offset unit"));
        assert!(err("-14days").contains("unknown offset unit"));
        assert!(err("today1d").contains("must start with + or -"));
        assert!(err("todayish").contains("must start with + or -"));
        assert!(err("2026-06-10Z").contains("not a recognized date"));
        assert!(err("2026-06-10T14:30+02:00").contains("not a recognized date"));
    }

    #[test]
    fn huge_offsets_error_instead_of_panicking() {
        assert_eq!(err("+99999999999d"), "offset is too large");
        assert_eq!(err("+4000000000y"), "offset is too large");
        assert_eq!(err("+999999999w"), "offset is too large");
        assert_eq!(err("-300000y"), "offset is too large");
    }

    #[test]
    fn non_ascii_input_does_not_panic() {
        assert!(parse_date_value("היום", now(), Weekday::Mon).is_err());
        assert!(parse_date_value("-1דקה", now(), Weekday::Mon).is_err());
        assert!(parse_date_value("today-1é", now(), Weekday::Mon).is_err());
    }
}
