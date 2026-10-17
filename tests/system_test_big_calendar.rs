//! Live system test of a big calendar (2026-10-17). Plan:
//! `SYSTEM_TEST_PLAN_2026-10-17-BIG-CALENDAR.md`.
//!
//! Seeds ~490 events in 2098 plus two long recurring series (400 daily and
//! 104 weekly occurrences), then checks `list_events`' caps and paging
//! against exact ground truth: the 250-row page cap, `offset` paging across
//! more than 250 rows, recurrence expansion, the date-window slack from #42
//! (bait events just outside the range must not eat the page), a dense
//! #43-style week, output controls, filters before paging, Hebrew subjects /
//! locations / categories / queries, and 100-item batch writes with `undo`.
//!
//! Everything goes through the real MCP tool layer (see `common`). Nothing is
//! sent: every event is a personal appointment, or a meeting saved with
//! `notify: false` to `example.invalid` attendees. Never panics before
//! cleanup.
//!
//!     cargo test --test system_test_big_calendar -- --ignored --nocapture

mod common;

use std::collections::{BTreeMap, BTreeSet};
use std::time::Instant;

use chrono::{Datelike, Duration, NaiveDate, NaiveDateTime, Weekday};
use common::{ids_of, s, temp_dir, Log, Wire};
use serde_json::{json, Value};

const TAG: &str = "[outlook-mcp-rs systest bigcal]";
const SELF: &str = "adamkopelman@outlook.com";
const HE_SUBJ: &str = "פגישת צוות";
const HE_LOC: &str = "חדר ישיבות";
const HE_CAT: &str = "צוות";
const HE_CAT2: &str = "קטגוריה נוספת";
const HE_SERIES: &str = "סדרה שבועית";
const SHOW: [&str; 5] = ["busy", "free", "tentative", "out_of_office", "working_elsewhere"];

/// One expected `list_events` row, from the seed data.
#[derive(Clone, Debug)]
struct Exp {
    subject: String,
    start: NaiveDateTime,
    show_as: String,
    categories: Vec<String>,
    all_day: bool,
    meeting: bool,
}

impl Exp {
    fn key(&self) -> String {
        format!("{}|{}", self.subject, iso(self.start))
    }
}

fn iso(t: NaiveDateTime) -> String {
    t.format("%Y-%m-%dT%H:%M:%S").to_string()
}

fn at(y: i32, m: u32, d: u32, h: u32, min: u32) -> NaiveDateTime {
    NaiveDate::from_ymd_opt(y, m, d).unwrap().and_hms_opt(h, min, 0).unwrap()
}

/// `subject|start` of a returned row.
fn key(r: &Value) -> String {
    format!("{}|{}", s(r, "subject"), s(r, "start"))
}

/// A row's identity: the occurrence id for an occurrence, else the id.
fn rid(r: &Value) -> String {
    let o = s(r, "occurrence_id");
    if o.is_empty() { s(r, "id") } else { o }
}

fn rows(v: &Result<Value, String>) -> Vec<Value> {
    v.as_ref().ok().and_then(|v| v.as_array().cloned()).unwrap_or_default()
}

fn sorted_by_start(rows: &[Value]) -> bool {
    rows.windows(2).all(|p| s(&p[0], "start") <= s(&p[1], "start"))
}

/// Every page of `args` at `page` rows per call, until a short page.
async fn all_pages(w: &mut Wire, args: &Value, page: i64) -> (Vec<Value>, Vec<usize>, Option<String>) {
    let mut all = Vec::new();
    let mut sizes = Vec::new();
    let mut offset = 0;
    loop {
        let mut a = args.clone();
        a["count"] = json!(page);
        a["offset"] = json!(offset);
        match w.call("list_events", a).await {
            Ok(v) => {
                let r = v.as_array().cloned().unwrap_or_default();
                sizes.push(r.len());
                let n = r.len() as i64;
                all.extend(r);
                if n < page || sizes.len() > 100 {
                    return (all, sizes, None);
                }
                offset += page;
            }
            Err(e) => return (all, sizes, Some(e)),
        }
    }
}

fn set_of(rows: &[Value]) -> BTreeSet<String> {
    rows.iter().map(key).collect()
}

fn exp_set<'a>(e: impl Iterator<Item = &'a Exp>) -> BTreeSet<String> {
    e.map(Exp::key).collect()
}

/// A short diff of two key sets for failure messages.
fn diff(got: &BTreeSet<String>, want: &BTreeSet<String>) -> String {
    let missing: Vec<_> = want.difference(got).take(3).collect();
    let extra: Vec<_> = got.difference(want).take(3).collect();
    format!("got {} want {}; missing {} e.g. {missing:?}; extra {} e.g. {extra:?}",
        got.len(), want.len(), want.difference(got).count(), got.difference(want).count())
}

#[tokio::test(flavor = "multi_thread")]
#[ignore]
async fn system_test_big_calendar() {
    let run = format!("r{}", chrono::Local::now().timestamp());
    let mut w = common::connect("systest-bigcal").await;
    let mut log = Log::default();
    let mut created: Vec<String> = Vec::new();
    println!("run token {run}");

    body(&mut w, &mut log, &mut created, &run).await;

    println!("\n---- cleanup");
    cleanup(&mut w, &mut log, &created, &run).await;
    println!("\n---- summary");
    let ok = log.summary();
    assert!(ok, "one or more steps failed");
}

async fn body(w: &mut Wire, log: &mut Log, created: &mut Vec<String>, run: &str) {
    let t = |label: &str| format!("{TAG} {run} {label}");
    let q = format!("\"{run}\"");
    let mut exp: Vec<Exp> = Vec::new(); // every single (non-recurring) seed
    let mut failures = Vec::new();

    // ---- phase 0: seed ------------------------------------------------------
    let seeded = Instant::now();
    let mut seeds: Vec<(Value, Exp)> = Vec::new();
    let single = |subject: String, start: NaiveDateTime, minutes: i64, show_as: &str, cats: Vec<String>,
                  location: &str, all_day: bool, meeting: bool| -> (Value, Exp) {
        let end = start + Duration::minutes(minutes);
        let mut a = json!({
            "subject": subject, "start": if all_day { start.format("%Y-%m-%d").to_string() } else { start.format("%Y-%m-%dT%H:%M").to_string() },
            "end": if all_day { end.format("%Y-%m-%d").to_string() } else { end.format("%Y-%m-%dT%H:%M").to_string() },
            "show_as": show_as, "notify": false, "all_day": all_day,
        });
        if !cats.is_empty() { a["categories"] = json!(cats); }
        if !location.is_empty() { a["location"] = json!(location); }
        if meeting {
            a["required_attendees"] = json!(["bigcal-req1@example.invalid", "bigcal-req2@example.invalid"]);
            a["optional_attendees"] = json!(["bigcal-opt@example.invalid"]);
        }
        (a, Exp { subject, start, show_as: show_as.into(), categories: cats, all_day, meeting })
    };
    // March 2098: 10 per day 08:00..12:30. Slot 9 is Hebrew.
    for d in 1..=31u32 {
        for slot in 0..10u32 {
            let start = at(2098, 3, d, 8 + slot / 2, (slot % 2) * 30);
            if slot == 9 {
                seeds.push(single(t(&format!("{HE_SUBJ} {d:02}")), start, 30, "busy", vec![HE_CAT.into()], HE_LOC, false, false));
            } else {
                let cat = match slot % 4 { 0 => vec!["Red Category".to_string()], 1 => vec!["Blue Category".into()], 2 => vec!["Green Category".into()], _ => vec![] };
                let loc = if slot % 3 == 0 { format!("Room {slot}") } else { String::new() };
                seeds.push(single(t(&format!("m{d:02}-{slot:02}")), start, 30, SHOW[((d + slot) % 5) as usize], cat, &loc, false, false));
            }
        }
    }
    // The dense week (#43): 6 more per day, 14:00..16:30, as unsent meetings.
    for d in 9..=15u32 {
        for k in 0..6u32 {
            seeds.push(single(t(&format!("w{d:02}-{k}")), at(2098, 3, d, 14 + k / 2, (k % 2) * 30), 30, "busy", vec![], "", false, true));
        }
    }
    for d in [1u32, 6, 11, 16, 21, 26, 31] {
        seeds.push(single(t(&format!("allday {d:02}")), at(2098, 3, d, 0, 0), 24 * 60, "free", vec![], "", true, false));
    }
    // Bait in the #42 slack window on both sides, and the exact edges.
    for d in 17..=28u32 {
        for k in 0..5u32 { seeds.push(single(t(&format!("bait feb{d:02}-{k}")), at(2098, 2, d, 9 + k, 0), 30, "busy", vec![], "", false, false)); }
    }
    for d in 1..=12u32 {
        for k in 0..5u32 { seeds.push(single(t(&format!("bait apr{d:02}-{k}")), at(2098, 4, d, 9 + k, 0), 30, "busy", vec![], "", false, false)); }
    }
    seeds.push(single(t("edge feb28-2330"), at(2098, 2, 28, 23, 30), 20, "busy", vec![], "", false, false));
    seeds.push(single(t("edge mar01-0000"), at(2098, 3, 1, 0, 0), 20, "busy", vec![], "", false, false));
    seeds.push(single(t("edge mar31-2359"), at(2098, 3, 31, 23, 59), 20, "busy", vec![], "", false, false));
    seeds.push(single(t("edge apr01-0000"), at(2098, 4, 1, 0, 0), 20, "busy", vec![], "", false, false));
    seeds.push(single(t("span feb27-mar03"), at(2098, 2, 27, 10, 0), 4 * 24 * 60, "busy", vec![], "", false, false));

    let mut by_subject: BTreeMap<String, String> = BTreeMap::new();
    for (args, e) in seeds {
        match w.call("create_event", args).await {
            Ok(v) if !s(&v, "id").is_empty() => {
                created.push(s(&v, "id"));
                by_subject.insert(e.subject.clone(), s(&v, "id"));
                exp.push(e);
            }
            other => failures.push(format!("{}: {other:?}", e.subject)),
        }
    }
    // Two long series: 400 daily occurrences from 2097-12-01, and a Hebrew
    // weekly one on Mondays (2098-01-06 is a Monday), 104 occurrences.
    let daily = t("daily");
    let weekly = t(HE_SERIES);
    for (subject, start, end, rec) in [
        (&daily, "2097-12-01T07:00", "2097-12-01T07:15", json!({"pattern": "daily", "occurrences": 400})),
        (&weekly, "2098-01-06T07:30", "2098-01-06T07:45", json!({"pattern": "weekly", "days_of_week": ["monday"], "occurrences": 104})),
    ] {
        match w.call("create_event", json!({"subject": subject, "start": start, "end": end, "notify": false, "recurrence": rec})).await {
            Ok(v) if !s(&v, "id").is_empty() => { created.push(s(&v, "id")); by_subject.insert(subject.clone(), s(&v, "id")); }
            other => failures.push(format!("{subject}: {other:?}")),
        }
    }
    let seed_secs = seeded.elapsed().as_secs_f64();
    if !log.check("SEED", failures.is_empty(), format!("{} events + 2 series in {seed_secs:.0}s; failures {failures:?}", exp.len())) {
        return;
    }

    // Expected occurrences of the two series, by date.
    let daily_occ: Vec<NaiveDateTime> = (0..400).map(|i| at(2097, 12, 1, 7, 0) + Duration::days(i)).collect();
    let weekly_occ: Vec<NaiveDateTime> = (0..104).map(|i| at(2098, 1, 6, 7, 30) + Duration::weeks(i)).collect();
    assert_eq!(at(2098, 1, 6, 0, 0).weekday(), Weekday::Mon);

    // Move the daily series' 2098-03-15 occurrence to 19:00 (an exception).
    let moved_to = at(2098, 3, 15, 19, 0);
    let occ15 = rows(&w.call("list_events", json!({"query": format!("{q} daily"), "start_after": "2098-03-15", "start_before": "2098-03-15"})).await);
    let moved = match occ15.first() {
        Some(r) => w.call("update_event", json!({"event_id": s(r, "occurrence_id"), "start": "2098-03-15T19:00", "end": "2098-03-15T19:15", "notify": false})).await,
        None => Err("no 2098-03-15 occurrence".into()),
    };
    log.check("EXC", moved.is_ok(), format!("moved the 03-15 occurrence: {:?}", moved.as_ref().err()));
    let daily_march: Vec<NaiveDateTime> = daily_occ.iter().copied()
        .filter(|d| d.year() == 2098 && d.month() == 3)
        .map(|d| if d.day() == 15 { moved_to } else { d })
        .collect();
    let series_exp = |d: NaiveDateTime, subject: &str| Exp {
        subject: subject.to_string(), start: d, show_as: "busy".into(), categories: vec![], all_day: false, meeting: false,
    };
    let in_range = |t: NaiveDateTime, a: NaiveDateTime, b: NaiveDateTime| t >= a && t <= b;
    let (m0, m1) = (at(2098, 3, 1, 0, 0), at(2098, 3, 31, 23, 59) + Duration::seconds(59));
    let mut march: Vec<Exp> = exp.iter().filter(|e| in_range(e.start, m0, m1)).cloned().collect();
    march.extend(daily_march.iter().map(|d| series_exp(*d, &daily)));
    march.extend(weekly_occ.iter().filter(|d| in_range(**d, m0, m1)).map(|d| series_exp(*d, &weekly)));
    let march_keys = exp_set(march.iter());
    println!("expected March rows: {}", march.len());
    let march_args = json!({"query": q, "start_after": "2098-03-01", "start_before": "2098-03-31"});

    // ---- caps and paging (BC) -------------------------------------------------
    let started = Instant::now();
    let r = w.call("list_events", json!({"query": q, "start_after": "2098-03-01", "start_before": "2098-03-31", "count": 1000})).await;
    let first_page_secs = started.elapsed().as_secs_f64();
    let got = rows(&r);
    log.check("BC1", got.len() == 250 && sorted_by_start(&got) && got.iter().all(|x| march_keys.contains(&key(x))),
        format!("count 1000 -> {} rows (cap 250), sorted {}, all in March {}, {first_page_secs:.1}s; {:?}",
            got.len(), sorted_by_start(&got), got.iter().all(|x| march_keys.contains(&key(x))), r.as_ref().err()));
    let r = w.call("list_events", march_args.clone()).await;
    log.check("BC2", rows(&r).len() == 250, format!("no count -> {} rows (default = cap 250)", rows(&r).len()));

    let (all250, sizes250, err) = all_pages(w, &march_args, 250).await;
    let got_keys = set_of(&all250);
    let ids: BTreeSet<String> = all250.iter().map(rid).collect();
    log.check("BC3", err.is_none() && got_keys == march_keys && ids.len() == all250.len() && sizes250 == vec![250, march.len() - 250] && sorted_by_start(&all250),
        format!("pages of 250: sizes {sizes250:?}; distinct ids {}; {}", ids.len(), diff(&got_keys, &march_keys)));
    let (all37, sizes37, err) = all_pages(w, &march_args, 37).await;
    let k37 = set_of(&all37);
    let want_pages = march.len().div_ceil(37);
    log.check("BC4", err.is_none() && k37 == march_keys && all37.len() == march.len() && sizes37.len() == want_pages + usize::from(march.len() % 37 == 0)
        && sorted_by_start(&all37),
        format!("pages of 37: {} pages, last {:?}; {}", sizes37.len(), sizes37.last(), diff(&k37, &march_keys)));
    let past = w.call("list_events", json!({"query": q, "start_after": "2098-03-01", "start_before": "2098-03-31", "offset": march.len() + 5})).await;
    let neg = w.call("list_events", json!({"query": q, "start_after": "2098-03-01", "start_before": "2098-03-31", "count": 5, "offset": -7})).await;
    let zero = w.call("list_events", json!({"query": q, "start_after": "2098-03-01", "start_before": "2098-03-31", "count": 5})).await;
    log.check("BC5", rows(&past).is_empty() && rows(&neg).iter().map(key).collect::<Vec<_>>() == rows(&zero).iter().map(key).collect::<Vec<_>>(),
        format!("offset past the end -> {} rows; offset -7 == offset 0: {}", rows(&past).len(),
            rows(&neg).iter().map(key).collect::<Vec<_>>() == rows(&zero).iter().map(key).collect::<Vec<_>>()));
    let first = all250.first().map(|r| s(r, "start")).unwrap_or_default();
    let last = all250.last().map(|r| s(r, "start")).unwrap_or_default();
    let outside = all250.iter().filter(|r| { let st = s(r, "start"); st.as_str() < "2098-03-01" || st.as_str() > "2098-03-31T23:59:59" }).count();
    log.check("BC6", outside == 0 && first == "2098-03-01T00:00:00" && last == "2098-03-31T23:59:00",
        format!("slack bait: {outside} rows outside March; first {first}, last {last}"));
    let has = |sub: &str| all250.iter().any(|r| s(r, "subject") == t(sub));
    log.check("BC7", has("edge mar01-0000") && has("edge mar31-2359") && !has("edge feb28-2330") && !has("edge apr01-0000") && !has("span feb27-mar03"),
        "edges: 03-01 00:00 and 03-31 23:59 in; 02-28 23:30, 04-01 00:00 and the 02-27..03-03 span (starts before) out".to_string());

    // ---- recurrence expansion (R) ----------------------------------------------
    let r = w.call("list_events", json!({"query": format!("{q} daily"), "start_after": "2098-03-01", "start_before": "2098-03-31"})).await;
    let got = rows(&r);
    let occ_ids: BTreeSet<String> = got.iter().map(|x| s(x, "occurrence_id")).collect();
    let starts: Vec<String> = got.iter().map(|x| s(x, "start")).collect();
    let want: Vec<String> = { let mut v: Vec<String> = daily_march.iter().map(|d| iso(*d)).collect(); v.sort(); v };
    let mut sorted_starts = starts.clone();
    sorted_starts.sort();
    log.check("R1", got.len() == 31 && occ_ids.len() == 31 && !occ_ids.contains("") && sorted_starts == want && starts.contains(&iso(moved_to)),
        format!("March of the daily series: {} rows, {} distinct occurrence ids, moved one at 19:00 {}", got.len(), occ_ids.len(), starts.contains(&iso(moved_to))));
    let long = json!({"query": format!("{q} daily"), "start_after": "2097-12-01", "start_before": "2099-01-31"});
    let started = Instant::now();
    let (occ_all, occ_sizes, err) = all_pages(w, &long, 250).await;
    let occ_secs = started.elapsed().as_secs_f64();
    let distinct: BTreeSet<String> = occ_all.iter().map(rid).collect();
    log.check("R2", err.is_none() && occ_all.len() == 400 && distinct.len() == 400 && occ_sizes == vec![250, 150]
        && occ_all.last().map(|r| s(r, "start")) == Some(iso(*daily_occ.last().unwrap())),
        format!("all 400 occurrences: pages {occ_sizes:?}, {} distinct, last {:?}, {occ_secs:.1}s", distinct.len(), occ_all.last().map(|r| s(r, "start"))));
    let mut deep = long.clone();
    deep["count"] = json!(10);
    deep["offset"] = json!(390);
    let r = w.call("list_events", deep.clone()).await;
    let want_tail: Vec<String> = daily_occ[390..].iter().map(|d| iso(*d)).collect();
    deep["offset"] = json!(400);
    let empty = w.call("list_events", deep).await;
    log.check("R3", rows(&r).iter().map(|x| s(x, "start")).collect::<Vec<_>>() == want_tail && rows(&empty).is_empty(),
        format!("offset 390 count 10 -> the last 10 occurrences; offset 400 -> {} rows", rows(&empty).len()));

    // ---- the dense week (W) -----------------------------------------------------
    let (w0, w1) = (at(2098, 3, 9, 0, 0), at(2098, 3, 15, 23, 59) + Duration::seconds(59));
    let week: Vec<&Exp> = march.iter().filter(|e| in_range(e.start, w0, w1)).collect();
    let week_args = json!({"query": q, "start_after": "2098-03-09", "start_before": "2098-03-15"});
    let full = w.call("list_events", week_args.clone()).await;
    let full_rows = rows(&full);
    let full_bytes = serde_json::to_string(&full_rows).unwrap_or_default().len();
    log.check("W1", set_of(&full_rows) == exp_set(week.iter().copied()) && full_rows.len() > 105,
        format!("dense week: {} rows (#43 saw ~105); {}", full_rows.len(), diff(&set_of(&full_rows), &exp_set(week.iter().copied()))));
    let mut a = week_args.clone();
    a["compact"] = json!(true);
    let compact = rows(&w.call("list_events", a).await);
    let compact_bytes = serde_json::to_string(&compact).unwrap_or_default().len();
    let mut a = week_args.clone();
    a["fields"] = json!(["id", "start"]);
    let tiny = rows(&w.call("list_events", a).await);
    let tiny_bytes = serde_json::to_string(&tiny).unwrap_or_default().len();
    // Each id is the item's EntryID plus its store's full StoreID (~460
    // characters), and occurrences repeat it in occurrence_id, so ids are
    // most of a row; compact keeps the id, so it can only save the rest.
    let id_bytes: usize = full_rows.iter().map(|r| s(r, "id").len() + s(r, "occurrence_id").len()).sum();
    log.check("W2", compact.len() == full_rows.len() && tiny.len() == full_rows.len() && compact_bytes < full_bytes && tiny_bytes < compact_bytes,
        format!("JSON size: full {full_bytes} B ({} B/row, of which ids {} B/row), compact {compact_bytes} B ({}%), fields [id,start] {tiny_bytes} B ({}%)",
            full_bytes / full_rows.len().max(1), id_bytes / full_rows.len().max(1),
            compact_bytes * 100 / full_bytes.max(1), tiny_bytes * 100 / full_bytes.max(1)));
    let dir = temp_dir("bigcal-rows");
    let mut a = week_args.clone();
    a["output_dir"] = json!(dir);
    match w.call("list_events", a).await {
        Ok(v) => {
            let text = std::fs::read_to_string(s(&v, "rows_file")).unwrap_or_default();
            let lines: Vec<Value> = text.lines().filter_map(|l| serde_json::from_str(l).ok()).collect();
            log.check("W3", v["count"].as_u64() == Some(full_rows.len() as u64) && set_of(&lines) == set_of(&full_rows) && text.contains(HE_SUBJ),
                format!("output_dir: count {}, {} JSON lines, same rows {}, Hebrew in the file as UTF-8 {}", v["count"], lines.len(),
                    set_of(&lines) == set_of(&full_rows), text.contains(HE_SUBJ)));
        }
        Err(e) => log.fail("W3", e),
    }
    let meetings: Vec<&Value> = full_rows.iter().filter(|r| s(r, "subject").contains(" w")).collect();
    let roles_ok = meetings.len() == 42 && meetings.iter().all(|r| r["is_meeting"] == true && r["required_count"] == 2
        && r["optional_count"] == 1 && r["user_is_required"] == true);
    let personal_ok = full_rows.iter().filter(|r| r["is_meeting"] == false).all(|r| r["user_is_required"] == true && r["required_count"] == 0);
    log.check("W4", roles_ok && personal_ok, format!("{} meetings: 2 required + 1 optional each, you required (organizer) {roles_ok}; personal rows required {personal_ok}", meetings.len()));

    // ---- filters before paging (F) ------------------------------------------------
    let filt = |extra: Value| { let mut a = march_args.clone(); a.as_object_mut().unwrap().extend(extra.as_object().unwrap().clone()); a };
    for (id, extra, want) in [
        ("F1", json!({"category": "Red Category"}), exp_set(march.iter().filter(|e| e.categories.iter().any(|c| c == "Red Category")))),
        ("F2", json!({"show_as": "free"}), exp_set(march.iter().filter(|e| e.show_as == "free"))),
        ("F3", json!({"show_as": ["out_of_office", "working_elsewhere"]}), exp_set(march.iter().filter(|e| e.show_as == "out_of_office" || e.show_as == "working_elsewhere"))),
        ("F4", json!({"all_day": true}), exp_set(march.iter().filter(|e| e.all_day))),
        ("F5", json!({"meetings_only": true}), exp_set(march.iter().filter(|e| e.meeting))),
        ("F6", json!({"query": format!("{q} m05-")}), exp_set(march.iter().filter(|e| e.subject.contains(" m05-")))),
    ] {
        let (got, sizes, err) = all_pages(w, &filt(extra.clone()), 25).await;
        let k = set_of(&got);
        log.check(id, err.is_none() && k == want && got.len() == want.len(), format!("{extra} in pages of 25 ({} pages): {}", sizes.len(), diff(&k, &want)));
    }

    // ---- Hebrew on the big calendar (H) -------------------------------------------
    let he_rows: Vec<&Exp> = march.iter().filter(|e| e.subject.contains(HE_SUBJ)).collect();
    let he_keys = exp_set(he_rows.iter().copied());
    let (got, _, _) = all_pages(w, &filt(json!({"query": format!("{q} {HE_SUBJ}")})), 250).await;
    let exact = got.iter().all(|r| he_rows.iter().any(|e| e.subject == s(r, "subject")) && s(r, "location") == HE_LOC
        && r["categories"].as_array().is_some_and(|c| c.iter().any(|x| x == HE_CAT)));
    log.check("H1", set_of(&got) == he_keys && exact, format!("query \"{HE_SUBJ}\": {}; subject/location/category byte-exact {exact}", diff(&set_of(&got), &he_keys)));
    let (got, _, _) = all_pages(w, &filt(json!({"query": format!("{q} location:ישיבות")})), 250).await;
    log.check("H2", set_of(&got) == he_keys, format!("location:ישיבות -> {}", diff(&set_of(&got), &he_keys)));
    let (got, _, _) = all_pages(w, &filt(json!({"query": format!("{q} subject:פגישת*צוות")})), 250).await;
    let (rev, _, _) = all_pages(w, &filt(json!({"query": format!("{q} subject:צוות*פגישת")})), 250).await;
    log.check("H3", set_of(&got) == he_keys && rev.is_empty(), format!("subject:פגישת*צוות -> {}; reversed order -> {} rows", diff(&set_of(&got), &he_keys), rev.len()));
    let (got, _, _) = all_pages(w, &filt(json!({"category": HE_CAT})), 25).await;
    log.check("H4", set_of(&got) == he_keys, format!("category \"{HE_CAT}\" in pages of 25: {}", diff(&set_of(&got), &he_keys)));
    let year = json!({"query": format!("{q} {HE_SERIES}"), "start_after": "2098-01-01", "start_before": "2098-12-31"});
    let (got, sizes, err) = all_pages(w, &year, 20).await;
    let want: BTreeSet<String> = weekly_occ.iter().filter(|d| d.year() == 2098).map(|d| format!("{weekly}|{}", iso(*d))).collect();
    let distinct: BTreeSet<String> = got.iter().map(rid).collect();
    log.check("H5", err.is_none() && set_of(&got) == want && distinct.len() == got.len() && got.iter().all(|r| s(r, "subject") == weekly),
        format!("Hebrew weekly series over 2098 in pages of 20 ({} pages): {}", sizes.len(), diff(&set_of(&got), &want)));
    let he_ids: Vec<String> = he_rows.iter().filter_map(|e| by_subject.get(&e.subject).cloned()).collect();
    match w.call("update_event", json!({"event_id": he_ids, "add_categories": [HE_CAT2], "notify": false})).await {
        Ok(v) => {
            let undos: Vec<Value> = v.as_array().cloned().unwrap_or_default().iter().map(|x| x["undo"].clone()).collect();
            let (tagged, _, _) = all_pages(w, &filt(json!({"category": HE_CAT2})), 250).await;
            let back = w.call("update_event", json!({"updates": undos})).await;
            let (after, _, _) = all_pages(w, &filt(json!({"category": HE_CAT2})), 250).await;
            log.check("H6", tagged.len() == 31 && back.is_ok() && after.is_empty(),
                format!("add Hebrew category \"{HE_CAT2}\" to 31 events in one call: {} tagged, undo {:?}, after undo {}", tagged.len(), back.err(), after.len()));
        }
        Err(e) => log.fail("H6", e),
    }

    // ---- batch writes at the limit (B) ---------------------------------------------
    let hundred: Vec<String> = march.iter().filter(|e| e.subject.contains(" m") && !e.subject.contains(HE_SUBJ))
        .filter_map(|e| by_subject.get(&e.subject).cloned()).take(100).collect();
    let started = Instant::now();
    match w.call("update_event", json!({"event_id": hundred, "add_categories": ["Orange Category"], "notify": false})).await {
        Ok(v) => {
            let res = v.as_array().cloned().unwrap_or_default();
            let secs = started.elapsed().as_secs_f64();
            let undos: Vec<Value> = res.iter().map(|x| x["undo"].clone()).collect();
            let (orange, _, _) = all_pages(w, &filt(json!({"category": "Orange Category"})), 250).await;
            let back = w.call("update_event", json!({"updates": undos})).await;
            let back_ok = back.as_ref().is_ok_and(|b| b.as_array().is_some_and(|a| a.len() == 100 && a.iter().all(|x| x.get("error").is_none())));
            let (after, _, _) = all_pages(w, &filt(json!({"category": "Orange Category"})), 250).await;
            log.check("B1", res.len() == 100 && res.iter().all(|x| x.get("error").is_none()) && orange.len() == 100 && back_ok && after.is_empty(),
                format!("100 events in one update ({secs:.1}s): {} ok, {} orange; 100 undos in one call ok {back_ok}; orange after {}",
                    res.iter().filter(|x| x.get("error").is_none()).count(), orange.len(), after.len()));
        }
        Err(e) => log.fail("B1", e),
    }
    let mut over = hundred.clone();
    over.push("bogus|id".into());
    let r = w.call("update_event", json!({"event_id": over, "add_categories": ["Orange Category"], "notify": false})).await;
    let (orange, _, _) = all_pages(w, &filt(json!({"category": "Orange Category"})), 250).await;
    log.check("B2", r.as_ref().is_err_and(|e| e.contains("100")) && orange.is_empty(), format!("101 ids -> {:?}; nothing changed: {}", r.err(), orange.is_empty()));
    let bait: Vec<String> = exp.iter().filter(|e| e.subject.contains(" bait ")).filter_map(|e| by_subject.get(&e.subject).cloned()).take(100).collect();
    let started = Instant::now();
    let r = w.call("delete_event", json!({"event_id": bait, "permanent": true, "notify": false})).await;
    let secs = started.elapsed().as_secs_f64();
    let res = rows(&r);
    let left = rows(&w.call("list_events", json!({"query": format!("{q} bait"), "start_after": "2098-02-01", "start_before": "2098-04-30"})).await);
    let in_deleted = ids_of(&w.call("list_emails", json!({"folder": "deleted", "query": format!("{q} bait"), "count": 200})).await.unwrap_or_default());
    let in_drafts = ids_of(&w.call("list_emails", json!({"folder": "drafts", "query": format!("{q} bait"), "count": 200})).await.unwrap_or_default());
    log.check("B3", res.len() == 100 && res.iter().all(|x| x["permanent"] == true) && left.len() == 20 && in_deleted.is_empty() && in_drafts.is_empty(),
        format!("100 permanent deletes in one call ({secs:.1}s): {} ok; bait left {} (want 20); in Deleted Items {}; junk drafts {}",
            res.iter().filter(|x| x["permanent"] == true).count(), left.len(), in_deleted.len(), in_drafts.len()));
    created.retain(|id| !bait.contains(id));

    // ---- availability and speed (A) ---------------------------------------------------
    // Every 30-minute slot of 2098-03-10 08:00-18:00 holds at most one seeded
    // event, so the exact status of all 20 slots is known. Free/busy is
    // published asynchronously: poll until it matches, for up to 2 minutes.
    let day = at(2098, 3, 10, 8, 0);
    let want: Vec<String> = (0..20).map(|i| {
        let slot = day + Duration::minutes(30 * i);
        exp.iter().find(|e| !e.all_day && e.start == slot).map(|e| e.show_as.clone()).unwrap_or_else(|| "free".into())
    }).collect();
    let started = Instant::now();
    let mut got: Vec<String> = Vec::new();
    let mut resolved = false;
    while started.elapsed().as_secs() < 120 {
        if let Ok(v) = w.call("check_availability", json!({"people": [SELF], "start": "2098-03-10T08:00", "end": "2098-03-10T18:00", "interval_minutes": 30})).await {
            resolved = v["people"][0]["resolved"] == true;
            got = v["people"][0]["slots"].as_array().cloned().unwrap_or_default().iter().map(|x| s(x, "status")).collect();
            if got == want { break; }
        }
        tokio::time::sleep(std::time::Duration::from_secs(10)).await;
    }
    log.check("A1", resolved && got == want,
        format!("2098-03-10 08:00-18:00 for yourself, all 20 slots exact after {:.0}s: {}; want {want:?}, got {got:?}",
            started.elapsed().as_secs_f64(), got == want));
    let started = Instant::now();
    let with_roles = w.call("list_events", march_args.clone()).await;
    let roles_secs = started.elapsed().as_secs_f64();
    let mut a = march_args.clone();
    a["fields"] = json!(["id", "subject", "start"]);
    let started = Instant::now();
    let without = w.call("list_events", a).await;
    let plain_secs = started.elapsed().as_secs_f64();
    let started = Instant::now();
    let year = w.call("list_events", json!({"query": q, "start_after": "2098-01-01", "start_before": "2098-12-31"})).await;
    let year_secs = started.elapsed().as_secs_f64();
    log.check("A2", rows(&with_roles).len() == 250 && rows(&without).len() == 250 && rows(&year).len() == 250 && sorted_by_start(&rows(&year))
        && roles_secs < 60.0 && plain_secs < 60.0 && year_secs < 60.0,
        format!("250-row page: with user roles {roles_secs:.1}s, fields without roles {plain_secs:.1}s; whole year 2098 {year_secs:.1}s"));
}

/// Deletes every created event (in batches of 100, permanently), then sweeps
/// the run token over 2097-11..2100-01 page by page until nothing is left:
/// the calendar holds more than one page, so a single capped query would not
/// see everything.
async fn cleanup(w: &mut Wire, log: &mut Log, created: &[String], run: &str) {
    let q = format!("\"{run}\"");
    let mut errors = Vec::new();
    for chunk in created.chunks(100) {
        match w.call("delete_event", json!({"event_id": chunk, "permanent": true, "notify": false})).await {
            Ok(v) => errors.extend(v.as_array().cloned().unwrap_or_default().into_iter().filter(|x| x.get("error").is_some()).map(|x| s(&x, "error"))),
            Err(e) => errors.push(e),
        }
    }
    let mut swept = 0;
    for pass in 0..20 {
        let left = rows(&w.call("list_events", json!({"query": q, "start_after": "2097-11-01", "start_before": "2100-01-31"})).await);
        let series: BTreeSet<String> = left.iter().map(|r| s(r, "id")).collect();
        println!("sweep pass {pass}: {} rows, {} items", left.len(), series.len());
        if series.is_empty() {
            break;
        }
        swept += series.len();
        let ids: Vec<String> = series.into_iter().collect();
        for chunk in ids.chunks(100) {
            if let Err(e) = w.call("delete_event", json!({"event_id": chunk, "permanent": true, "notify": false})).await {
                errors.push(e);
            }
        }
    }
    let mut stray = Vec::new();
    for folder in ["deleted", "drafts", "sent", "outbox"] {
        let found = ids_of(&w.call("list_emails", json!({"folder": folder, "query": q, "count": 200})).await.unwrap_or_default());
        for id in &found {
            let _ = w.call("delete_email", json!({"email_id": id, "permanent": true})).await;
        }
        if !found.is_empty() {
            stray.push(format!("{folder}: {}", found.len()));
        }
    }
    log.check("CLEAN", errors.is_empty() && stray.is_empty(),
        format!("{} deleted by id, {swept} more swept; stray mail items {stray:?}; errors {:?}", created.len(), errors.iter().take(3).collect::<Vec<_>>()));
}
