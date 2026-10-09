//! Date filters for JET `Items.Restrict` that cannot be silently misread.
//!
//! `Restrict` parses the date strings in a filter with the *current user's*
//! regional format, so the old hard-coded US `MM/DD/YYYY` string meant
//! 2026-09-11 was read as 9 November on a day-first machine (issue #1). Rather
//! than trust any one guess of the parser's day/month order, the filter is
//! built so that order cannot matter:
//!
//! 1. **Unambiguous dates only.** Each bound is snapped outwards to the nearest
//!    date whose day is 13 or more, or equals its month, written with a 4-digit
//!    year. Under every way of reading three numbers as year/month/day, such a
//!    string names exactly one valid date — the intended one — so a parser that
//!    disagrees about the order either reads the right date or rejects it.
//!    The snap widens the window by at most 12 days on each side.
//! 2. **Date-only literals.** No time, so no seconds/AM-PM/time-separator issues.
//! 3. **Rendering order:** the user's short-date order and separator (what
//!    Outlook uses to parse) with a 4-digit year, then ISO `yyyy-mm-dd` if that
//!    is rejected; if both are rejected the tool returns an error instead of
//!    unfiltered results.
//! 4. **Exact bounds client-side.** Each item that passes `Restrict` is compared
//!    against the exact requested bounds ([`DateFilter::classify`]), which also
//!    removes JET's minute granularity. An item well outside the window that was
//!    sent to `Restrict` means the filter was misread, and is reported as an
//!    error rather than silently dropped.
//!
//! Everything here is pure (chrono only) so it is unit-tested on any platform;
//! the Win32 locale lookup lives in `com.rs` and the COM calls in `client.rs`.

use chrono::{Datelike, NaiveDate, NaiveDateTime};

/// One of the three numeric fields of a short date.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DatePart {
    Year,
    Month,
    Day,
}

/// The order of year/month/day in the user's short-date pattern, plus the
/// separator to write between them.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DateOrder {
    pub parts: [DatePart; 3],
    pub separator: char,
}

/// True when `d`, written numerically with a 4-digit year, names one date under
/// any assignment of its three numbers to year/month/day: the year is the only
/// number above 31, and of the two left either only the month is ≤ 12
/// (day ≥ 13) or both readings agree (day == month).
pub fn is_unambiguous(d: NaiveDate) -> bool {
    d.day() >= 13 || d.day() == d.month()
}

/// The latest unambiguous date on or before `d` (at most 12 days earlier).
pub fn unambiguous_floor(d: NaiveDate) -> NaiveDate {
    let mut cur = d;
    while !is_unambiguous(cur) {
        match cur.pred_opt() {
            Some(p) => cur = p,
            None => return cur,
        }
    }
    cur
}

/// The earliest unambiguous date on or after `d` (at most 12 days later).
pub fn unambiguous_ceil(d: NaiveDate) -> NaiveDate {
    let mut cur = d;
    while !is_unambiguous(cur) {
        match cur.succ_opt() {
            Some(n) => cur = n,
            None => return cur,
        }
    }
    cur
}

/// Reads the year/month/day order and separator out of a Windows short-date
/// pattern (`LOCALE_SSHORTDATE`, e.g. `M/d/yyyy`, `dd.MM.yyyy`, `yyyy-MM-dd`).
/// `d`/`dd` is the day (`ddd`+ is a weekday name and is skipped), any run of
/// `M` the month, any run of `y` the year; `'quoted'` text is literal. The
/// separator is the text between the first two fields when, trimmed, it is one
/// of `/ . -`; otherwise `/`. Returns `None` unless each field appears once.
pub fn date_order_from_pattern(pattern: &str) -> Option<DateOrder> {
    let chars: Vec<char> = pattern.chars().collect();
    let mut parts: Vec<DatePart> = Vec::new();
    let mut between: Vec<String> = Vec::new();
    let mut literal = String::new();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c == '\'' {
            // Quoted literal; `''` inside quotes is an escaped quote.
            i += 1;
            while i < chars.len() {
                if chars[i] == '\'' {
                    if chars.get(i + 1) == Some(&'\'') {
                        literal.push('\'');
                        i += 2;
                        continue;
                    }
                    break;
                }
                literal.push(chars[i]);
                i += 1;
            }
            i += 1;
            continue;
        }
        let run = chars[i..].iter().take_while(|&&x| x == c).count();
        let part = match c {
            'd' if run <= 2 => Some(DatePart::Day),
            'M' => Some(DatePart::Month),
            'y' => Some(DatePart::Year),
            _ => None,
        };
        match part {
            Some(p) => {
                if parts.contains(&p) {
                    return None;
                }
                if !parts.is_empty() {
                    between.push(std::mem::take(&mut literal));
                }
                literal.clear();
                parts.push(p);
            }
            // A weekday name (`ddd`), era (`g`) or other letters: not a field,
            // and not a simple separator either.
            None if c.is_alphabetic() => {
                for _ in 0..run {
                    literal.push('\u{0}');
                }
            }
            None => {
                for _ in 0..run {
                    literal.push(c);
                }
            }
        }
        i += run;
    }
    let parts: [DatePart; 3] = parts.try_into().ok()?;
    let separator = match between.first().map(|s| s.trim()) {
        Some(s) if s.chars().count() == 1 && "/.-".contains(s) => s.chars().next().unwrap(),
        _ => '/',
    };
    Some(DateOrder { parts, separator })
}

/// The 4-digit year range a date bound can be written in. Outside it a bound
/// is left out of `Restrict` (the client-side check still applies it).
fn renderable(d: NaiveDate) -> bool {
    (1000..=9999).contains(&d.year())
}

/// `d` written numerically in `order` with zero-padded day/month and a 4-digit
/// year, e.g. `11.09.2026` for `dd.MM.yyyy`. `None` if the year isn't 4 digits.
pub fn render_date(d: NaiveDate, order: &DateOrder) -> Option<String> {
    if !renderable(d) {
        return None;
    }
    let field = |p: DatePart| match p {
        DatePart::Year => format!("{:04}", d.year()),
        DatePart::Month => format!("{:02}", d.month()),
        DatePart::Day => format!("{:02}", d.day()),
    };
    let sep = order.separator.to_string();
    Some(order.parts.map(field).join(&sep))
}

/// `d` as ISO `yyyy-mm-dd`. `None` if the year isn't 4 digits.
pub fn render_iso(d: NaiveDate) -> Option<String> {
    renderable(d).then(|| d.format("%Y-%m-%d").to_string())
}

/// Where an item's date falls relative to a [`DateFilter`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Placement {
    /// Before the requested `after` bound (but inside the widened window).
    Before,
    /// Within the exact requested bounds (inclusive on both ends).
    InRange,
    /// After the requested `before` bound (but inside the widened window).
    After,
    /// Well outside the window sent to `Restrict`: Outlook misread the filter.
    OutsideWindow,
}

/// The slack around the `Restrict` window before an item counts as
/// [`Placement::OutsideWindow`], so time-zone or all-day edge cases at the
/// window's own edges never trip the misread check.
const WINDOW_SLACK_DAYS: i64 = 2;

/// An exact, inclusive date range for a list tool, and the wider date-only
/// window sent to `Restrict` to pre-filter it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DateFilter {
    /// Keep items at or after this instant.
    pub after: Option<NaiveDateTime>,
    /// Keep items at or before this instant.
    pub before: Option<NaiveDateTime>,
    /// `Restrict` lower bound (`>=`): unambiguous, on or before `after`'s date.
    pub window_start: Option<NaiveDate>,
    /// `Restrict` upper bound (`<`): unambiguous, after `before`'s date.
    pub window_end: Option<NaiveDate>,
}

impl DateFilter {
    /// `None` when neither bound is given (no date filter at all).
    pub fn new(after: Option<NaiveDateTime>, before: Option<NaiveDateTime>) -> Option<Self> {
        if after.is_none() && before.is_none() {
            return None;
        }
        let window_start =
            after.map(|a| unambiguous_floor(a.date())).filter(|d| renderable(*d));
        let window_end = before
            .and_then(|b| b.date().succ_opt())
            .map(unambiguous_ceil)
            .filter(|d| renderable(*d));
        Some(Self { after, before, window_start, window_end })
    }

    /// The `Restrict` filters to try, in order: the user's short-date order
    /// (when known), then ISO. Each is the same window; duplicates are dropped.
    /// Empty when neither window bound is renderable.
    pub fn restrict_filters(&self, field: &str, order: Option<&DateOrder>) -> Vec<String> {
        let mut out: Vec<String> = Vec::new();
        let renderers: [&dyn Fn(NaiveDate) -> Option<String>; 2] =
            [&|d| order.and_then(|o| render_date(d, o)), &render_iso];
        for render in renderers {
            let lower = self.window_start.map(render);
            let upper = self.window_end.map(render);
            // A renderer that can't write a bound we have is skipped entirely.
            if matches!(lower, Some(None)) || matches!(upper, Some(None)) {
                continue;
            }
            let clauses: Vec<String> = [
                lower.flatten().map(|s| format!("[{field}] >= '{s}'")),
                upper.flatten().map(|s| format!("[{field}] < '{s}'")),
            ]
            .into_iter()
            .flatten()
            .collect();
            if clauses.is_empty() {
                continue;
            }
            let f = clauses.join(" AND ");
            if !out.contains(&f) {
                out.push(f);
            }
        }
        out
    }

    /// Where `t` falls: inside the exact bounds, before/after them, or so far
    /// outside the `Restrict` window that the filter must have been misread.
    pub fn classify(&self, t: NaiveDateTime) -> Placement {
        let slack = chrono::Duration::days(WINDOW_SLACK_DAYS);
        let midnight = |d: NaiveDate| d.and_hms_opt(0, 0, 0).unwrap();
        if self.window_start.is_some_and(|ws| t < midnight(ws) - slack)
            || self.window_end.is_some_and(|we| t >= midnight(we) + slack)
        {
            return Placement::OutsideWindow;
        }
        if self.after.is_some_and(|a| t < a) {
            Placement::Before
        } else if self.before.is_some_and(|b| t > b) {
            Placement::After
        } else {
            Placement::InRange
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use DatePart::{Day, Month, Year};

    fn ymd(y: i32, m: u32, d: u32) -> NaiveDate {
        NaiveDate::from_ymd_opt(y, m, d).unwrap()
    }

    fn dt(s: &str) -> NaiveDateTime {
        NaiveDateTime::parse_from_str(s, "%Y-%m-%dT%H:%M:%S").unwrap()
    }

    /// Every date a parser could read out of `s` (three numbers separated by
    /// any non-digit), under all six assignments of the numbers to Y/M/D.
    fn all_readings(s: &str) -> Vec<NaiveDate> {
        let nums: Vec<u32> = s
            .split(|c: char| !c.is_ascii_digit())
            .filter(|p| !p.is_empty())
            .map(|p| p.parse().unwrap())
            .collect();
        assert_eq!(nums.len(), 3, "{s}");
        let perms = [[0, 1, 2], [0, 2, 1], [1, 0, 2], [1, 2, 0], [2, 0, 1], [2, 1, 0]];
        let mut out: Vec<NaiveDate> = perms
            .iter()
            .filter_map(|[y, m, d]| NaiveDate::from_ymd_opt(nums[*y] as i32, nums[*m], nums[*d]))
            .collect();
        out.sort();
        out.dedup();
        out
    }

    fn every_day(from: i32, to: i32) -> impl Iterator<Item = NaiveDate> {
        ymd(from, 1, 1).iter_days().take_while(move |d| d.year() <= to)
    }

    const ORDERS: [[DatePart; 3]; 4] =
        [[Month, Day, Year], [Day, Month, Year], [Year, Month, Day], [Year, Day, Month]];

    #[test]
    fn snapping_brackets_every_day_within_twelve_days_onto_unambiguous_dates() {
        for d in every_day(2023, 2029) {
            let lo = unambiguous_floor(d);
            let hi = unambiguous_ceil(d);
            assert!(lo <= d && d <= hi, "{d}: {lo}..{hi}");
            assert!((d - lo).num_days() <= 12, "{d} floor {lo}");
            assert!((hi - d).num_days() <= 12, "{d} ceil {hi}");
            assert!(is_unambiguous(lo) && is_unambiguous(hi), "{d}");
            if is_unambiguous(d) {
                assert_eq!((lo, hi), (d, d));
            }
        }
    }

    #[test]
    fn snapping_examples() {
        assert_eq!(unambiguous_floor(ymd(2026, 9, 11)), ymd(2026, 9, 9));
        assert_eq!(unambiguous_ceil(ymd(2026, 9, 11)), ymd(2026, 9, 13));
        assert_eq!(unambiguous_floor(ymd(2026, 12, 11)), ymd(2026, 11, 30));
        assert_eq!(unambiguous_ceil(ymd(2026, 1, 2)), ymd(2026, 1, 13));
        assert_eq!(unambiguous_floor(ymd(2026, 1, 12)), ymd(2026, 1, 1));
    }

    #[test]
    fn every_rendering_of_every_snapped_date_has_exactly_one_reading() {
        for d in every_day(2024, 2028) {
            for snapped in [unambiguous_floor(d), unambiguous_ceil(d)] {
                for parts in ORDERS {
                    for separator in ['/', '.', '-'] {
                        let s = render_date(snapped, &DateOrder { parts, separator }).unwrap();
                        assert_eq!(all_readings(&s), vec![snapped], "{s}");
                    }
                }
                assert_eq!(all_readings(&render_iso(snapped).unwrap()), vec![snapped]);
            }
        }
    }

    #[test]
    fn the_oracle_does_catch_the_issue_1_ambiguity() {
        // Sanity check of the oracle itself: the old US string for 2026-09-11
        // reads as two different dates.
        assert_eq!(all_readings("09/11/2026"), vec![ymd(2026, 9, 11), ymd(2026, 11, 9)]);
    }

    #[test]
    fn pattern_order_parsing() {
        let cases: [(&str, [DatePart; 3], char); 11] = [
            ("M/d/yyyy", [Month, Day, Year], '/'),
            ("dd/MM/yyyy", [Day, Month, Year], '/'),
            ("dd.MM.yyyy", [Day, Month, Year], '.'),
            ("yyyy-MM-dd", [Year, Month, Day], '-'),
            ("yyyy/M/d", [Year, Month, Day], '/'),
            ("d-MMM-yy", [Day, Month, Year], '-'),
            ("dd/MM/yy", [Day, Month, Year], '/'),
            ("d. M. yyyy", [Day, Month, Year], '.'),
            ("yyyy. MM. dd.", [Year, Month, Day], '.'),
            ("ddd dd/MM/yyyy", [Day, Month, Year], '/'),
            ("yyyy'年'M'月'd'日'", [Year, Month, Day], '/'),
        ];
        for (pattern, parts, separator) in cases {
            assert_eq!(
                date_order_from_pattern(pattern),
                Some(DateOrder { parts, separator }),
                "{pattern}"
            );
        }
    }

    #[test]
    fn pattern_without_all_three_fields_is_rejected() {
        for pattern in ["", "MM/yyyy", "dd/MM", "dd/MM/dd/yyyy", "dddd, MMMM"] {
            assert_eq!(date_order_from_pattern(pattern), None, "{pattern}");
        }
    }

    #[test]
    fn rendering_follows_the_locale_order_with_a_four_digit_year() {
        let d = ymd(2026, 9, 13);
        let render = |pattern: &str| render_date(d, &date_order_from_pattern(pattern).unwrap());
        assert_eq!(render("M/d/yyyy").unwrap(), "09/13/2026");
        assert_eq!(render("dd/MM/yy").unwrap(), "13/09/2026");
        assert_eq!(render("dd.MM.yyyy").unwrap(), "13.09.2026");
        assert_eq!(render("yyyy/M/d").unwrap(), "2026/09/13");
        assert_eq!(render("d-MMM-yy").unwrap(), "13-09-2026");
        assert_eq!(render_iso(d).unwrap(), "2026-09-13");
        let us = date_order_from_pattern("M/d/yyyy").unwrap();
        assert_eq!(render_date(ymd(999, 9, 13), &us), None);
        assert_eq!(render_iso(ymd(10000, 9, 13)), None);
    }

    #[test]
    fn issue_1_window_is_sent_unambiguously_and_wider_than_requested() {
        // list_events 2026-09-11..2026-09-18 (end day inclusive).
        let f = DateFilter::new(Some(dt("2026-09-11T00:00:00")), Some(dt("2026-09-18T23:59:59")))
            .unwrap();
        assert_eq!(f.window_start, Some(ymd(2026, 9, 9)));
        assert_eq!(f.window_end, Some(ymd(2026, 9, 19)));
        let il = date_order_from_pattern("dd/MM/yyyy").unwrap();
        assert_eq!(
            f.restrict_filters("Start", Some(&il)),
            vec![
                "[Start] >= '09/09/2026' AND [Start] < '19/09/2026'".to_string(),
                "[Start] >= '2026-09-09' AND [Start] < '2026-09-19'".to_string(),
            ]
        );
        let us = date_order_from_pattern("M/d/yyyy").unwrap();
        assert_eq!(
            f.restrict_filters("Start", Some(&us))[0],
            "[Start] >= '09/09/2026' AND [Start] < '09/19/2026'"
        );
    }

    #[test]
    fn restrict_filters_fall_back_to_iso_only_and_dedupe() {
        let f = DateFilter::new(Some(dt("2026-03-05T10:00:00")), None).unwrap();
        assert_eq!(f.restrict_filters("ReceivedTime", None), vec!["[ReceivedTime] >= '2026-03-03'"]);
        let iso = date_order_from_pattern("yyyy-MM-dd").unwrap();
        assert_eq!(f.restrict_filters("ReceivedTime", Some(&iso)).len(), 1);
        let f = DateFilter::new(None, Some(dt("2026-03-05T10:00:00"))).unwrap();
        assert_eq!(f.restrict_filters("ReceivedTime", None), vec!["[ReceivedTime] < '2026-03-13'"]);
    }

    #[test]
    fn unrenderable_bounds_are_left_to_the_client_side_check() {
        let f = DateFilter::new(Some(dt("0500-01-01T00:00:00")), None).unwrap();
        assert!(f.restrict_filters("Start", None).is_empty());
        assert_eq!(f.classify(dt("2026-01-01T00:00:00")), Placement::InRange);
        assert_eq!(DateFilter::new(None, None), None);
    }

    #[test]
    fn every_window_covers_the_exact_range() {
        for d in every_day(2025, 2027) {
            let after = d.and_hms_opt(17, 45, 30).unwrap();
            let before = after + chrono::Duration::days(3);
            let f = DateFilter::new(Some(after), Some(before)).unwrap();
            assert!(f.window_start.unwrap().and_hms_opt(0, 0, 0).unwrap() <= after);
            assert!(f.window_end.unwrap().and_hms_opt(0, 0, 0).unwrap() > before);
            assert_eq!(f.classify(after), Placement::InRange);
            assert_eq!(f.classify(before), Placement::InRange);
        }
    }

    #[test]
    fn classify_applies_exact_bounds_to_the_second() {
        let f = DateFilter::new(Some(dt("2026-09-11T09:30:00")), Some(dt("2026-09-18T17:00:00")))
            .unwrap();
        assert_eq!(f.classify(dt("2026-09-11T09:29:59")), Placement::Before);
        assert_eq!(f.classify(dt("2026-09-11T09:30:00")), Placement::InRange);
        assert_eq!(f.classify(dt("2026-09-18T17:00:00")), Placement::InRange);
        assert_eq!(f.classify(dt("2026-09-18T17:00:01")), Placement::After);
        // Inside the widened Restrict window (2026-09-09 .. 2026-09-19).
        assert_eq!(f.classify(dt("2026-09-09T00:00:00")), Placement::Before);
        assert_eq!(f.classify(dt("2026-09-20T12:00:00")), Placement::After);
    }

    #[test]
    fn classify_flags_items_far_outside_the_restrict_window_as_a_misread() {
        let f = DateFilter::new(Some(dt("2026-09-11T00:00:00")), Some(dt("2026-09-18T23:59:59")))
            .unwrap();
        // What the swapped issue-1 reading would have returned.
        assert_eq!(f.classify(dt("2026-11-09T10:00:00")), Placement::OutsideWindow);
        assert_eq!(f.classify(dt("2026-09-06T23:59:59")), Placement::OutsideWindow);
        assert_eq!(f.classify(dt("2026-09-07T00:00:00")), Placement::Before);
        assert_eq!(f.classify(dt("2026-09-20T23:59:59")), Placement::After);
        assert_eq!(f.classify(dt("2026-09-21T00:00:00")), Placement::OutsideWindow);
    }
}
