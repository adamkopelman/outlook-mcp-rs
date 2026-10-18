//! Live system test of edge cases (2026-10-18). Plan:
//! `SYSTEM_TEST_PLAN_2026-10-18-EDGE.md`.
//!
//! Each test names a hypothesis from a code review (colliding attachment
//! names, commas in categories, DASL wildcards, all-day and DST boundaries,
//! odd availability intervals, recurrence edges, permanent deletes, paging
//! under mutation, non-default folders, attendee no-ops, size limits, flags,
//! stale ids). A confirmed hypothesis is recorded as `[FAIL]` with what was
//! observed; this suite does not fix anything.
//!
//! Everything goes through the real MCP tool layer (see `common`), except
//! what the tools don't expose: creating and removing the test folders,
//! attaching an item to a draft, and a few read-only probes (reminders, raw
//! category strings, the final tag count) use raw Outlook COM from
//! PowerShell. Traffic: one self-loop email to the mailbox itself and its
//! reply-all (also only to itself); everything else is a draft to
//! `nobody@example.invalid`, a personal appointment, or an unsent meeting
//! saved with `notify: false` to `example.invalid` attendees. Never panics
//! before cleanup.
//!
//!     cargo test --test system_test_edge -- --ignored --nocapture

mod common;

use std::collections::BTreeSet;
use std::path::{Path, PathBuf};
use std::time::{Duration, Instant};

use base64::Engine;
use chrono::{Datelike, NaiveDate, NaiveDateTime, Weekday};
use common::{ids_of, s, subjects_of, Log, Wire};
use serde_json::{json, Value};

const TAG: &str = "[outlook-mcp-rs systest edge]";
const SELF: &str = "adamkopelman@outlook.com";
const NOBODY: &str = "nobody@example.invalid";

/// What the run created, by tool family (latest known ids), plus names.
struct Ctx {
    run: String,
    root: PathBuf,
    sub: String,
    nested: String,
    emails: Vec<String>,
    events: Vec<String>,
    tasks: Vec<String>,
    /// Inbound copies to wait for before the Inbox sweep (self-loop, reply-all).
    inbound: usize,
}

impl Ctx {
    fn t(&self, label: &str) -> String {
        format!("{TAG} {} {label}", self.run)
    }
    /// A query scoped to this run: the run token plus `rest`.
    fn q(&self, rest: &str) -> String {
        format!("{} {rest}", self.run)
    }
    fn sub_path(&self) -> String {
        format!("Inbox/{}", self.sub)
    }
}

fn rows(v: &Result<Value, String>) -> Vec<Value> {
    v.as_ref().ok().and_then(|v| v.as_array().cloned()).unwrap_or_default()
}

fn starts(rows: &[Value]) -> Vec<String> {
    rows.iter().map(|r| s(r, "start")).collect()
}

fn at(y: i32, m: u32, d: u32, h: u32, min: u32) -> NaiveDateTime {
    NaiveDate::from_ymd_opt(y, m, d).unwrap().and_hms_opt(h, min, 0).unwrap()
}

fn iso(t: NaiveDateTime) -> String {
    t.format("%Y-%m-%dT%H:%M:%S").to_string()
}

fn first_weekday(y: i32, m: u32, wd: Weekday) -> NaiveDate {
    (1..=7).map(|d| NaiveDate::from_ymd_opt(y, m, d).unwrap()).find(|d| d.weekday() == wd).unwrap()
}

fn cats(v: &Value) -> Vec<String> {
    let mut c: Vec<String> = v["categories"].as_array().cloned().unwrap_or_default()
        .iter().filter_map(|x| x.as_str().map(str::to_string)).collect();
    c.sort();
    c
}

/// `(EntryID, StoreID)` of an item id.
fn eid_sid(id: &str) -> (String, String) {
    let mut p = id.split('|');
    (p.next().unwrap_or("").to_string(), p.next().unwrap_or("").to_string())
}

/// A PowerShell single-quoted string literal.
fn psq(x: &str) -> String {
    format!("'{}'", x.replace('\'', "''"))
}

/// Runs a PowerShell script against Outlook's COM object model (`$ns` is
/// the MAPI namespace) and returns its trimmed stdout. Passed as
/// -EncodedCommand (UTF-16), so Hebrew and quotes survive; killed after
/// `timeout_s` so a hung COM call can't hang the suite.
fn ps_blocking(script: &str, timeout_s: u64) -> Result<String, String> {
    use std::io::Read;
    use std::process::{Command, Stdio};
    let full = format!(
        "$ErrorActionPreference='Stop'; $ProgressPreference='SilentlyContinue'; \
         [Console]::OutputEncoding=[System.Text.Encoding]::UTF8; \
         try {{ $ns=(New-Object -ComObject Outlook.Application).GetNamespace('MAPI'); {script} }} \
         catch {{ Write-Output ('ERROR ' + $_.Exception.Message); exit 1 }}"
    );
    let bytes: Vec<u8> = full.encode_utf16().flat_map(u16::to_le_bytes).collect();
    let enc = base64::engine::general_purpose::STANDARD.encode(bytes);
    let mut child = Command::new("powershell.exe")
        .args(["-NoProfile", "-NonInteractive", "-EncodedCommand", &enc])
        .stdin(Stdio::null()).stdout(Stdio::piped()).stderr(Stdio::piped())
        .spawn().map_err(|e| e.to_string())?;
    let mut out = child.stdout.take().unwrap();
    let mut err = child.stderr.take().unwrap();
    let to = std::thread::spawn(move || { let mut b = Vec::new(); let _ = out.read_to_end(&mut b); b });
    let te = std::thread::spawn(move || { let mut b = Vec::new(); let _ = err.read_to_end(&mut b); b });
    let started = Instant::now();
    let status = loop {
        match child.try_wait() {
            Ok(Some(st)) => break Some(st),
            Ok(None) => {}
            Err(e) => return Err(e.to_string()),
        }
        if started.elapsed().as_secs() > timeout_s {
            let _ = child.kill();
            break None;
        }
        std::thread::sleep(Duration::from_millis(100));
    };
    let clean = |b: Vec<u8>| String::from_utf8_lossy(&b).trim().trim_start_matches('\u{feff}').trim().to_string();
    let o = clean(to.join().unwrap_or_default());
    let e = clean(te.join().unwrap_or_default());
    match status {
        Some(st) if st.success() => Ok(o),
        Some(_) => Err(format!("{o} {e}").trim().chars().take(400).collect()),
        None => Err(format!("timed out after {timeout_s}s; {o} {e}")),
    }
}

async fn ps(script: String) -> Result<String, String> {
    tokio::task::spawn_blocking(move || ps_blocking(&script, 120)).await.unwrap_or_else(|e| Err(e.to_string()))
}

/// Expected free/busy slots over [start, end) on the `interval`-minute grid
/// from 00:00 of start's date (how `check_availability` slices it): each
/// slot overlapping the window, with the status of the first busy interval
/// it overlaps, else "free".
fn expect_grid(start: NaiveDateTime, end: NaiveDateTime, interval: i64,
    busy: &[(NaiveDateTime, NaiveDateTime, &str)]) -> Vec<(String, String)> {
    let midnight = start.date().and_hms_opt(0, 0, 0).unwrap();
    let step = chrono::Duration::minutes(interval);
    let mut out = Vec::new();
    let mut k = 0;
    loop {
        let a = midnight + step * k;
        let b = a + step;
        k += 1;
        if a >= end { break; }
        if b <= start { continue; }
        let st = busy.iter().find(|(x, y, _)| *x < b && *y > a).map(|(_, _, s)| s.to_string()).unwrap_or_else(|| "free".into());
        out.push((iso(a), st));
    }
    out
}

fn got_slots(v: &Value, person: usize) -> Vec<(String, String)> {
    v["people"][person]["slots"].as_array().cloned().unwrap_or_default().iter()
        .map(|x| (s(x, "start"), s(x, "status"))).collect()
}

fn slot_diff(got: &[(String, String)], want: &[(String, String)]) -> String {
    let bad: Vec<String> = want.iter().zip(got.iter()).filter(|(w, g)| w != g)
        .take(4).map(|(w, g)| format!("want {}={} got {}={}", w.0, w.1, g.0, g.1)).collect();
    format!("{} slots (want {}); first mismatches {bad:?}", got.len(), want.len())
}

#[tokio::test(flavor = "multi_thread")]
#[ignore]
async fn system_test_edge() {
    let run = format!("r{}", chrono::Local::now().timestamp());
    let root = std::env::temp_dir().join(format!("outlook-mcp-rs-systest-edge-{run}"));
    let _ = std::fs::remove_dir_all(&root);
    std::fs::create_dir_all(&root).unwrap();
    let mut c = Ctx {
        sub: format!("{TAG} {run} systest edge תיקייה"),
        nested: format!("{TAG} {run} nested"),
        run, root, emails: Vec::new(), events: Vec::new(), tasks: Vec::new(), inbound: 0,
    };
    let mut w = common::connect("systest-edge").await;
    let mut log = Log::default();
    println!("run token {}", c.run);

    body(&mut w, &mut log, &mut c).await;

    println!("\n---- cleanup");
    cleanup(&mut w, &mut log, &c).await;

    println!("\n---- summary");
    let failed = log.0.iter().filter(|(_, ok, _)| !ok).count();
    for (id, ok, detail) in &log.0 {
        println!("{:<8} {} {}", id, if *ok { "PASS" } else { "FAIL" }, if *ok { "" } else { detail });
    }
    println!("{}/{} passed", log.0.len() - failed, log.0.len());
    assert!(failed == 0, "{failed} step(s) failed");
}

async fn draft(w: &mut Wire, c: &mut Ctx, label: &str, extra: Value) -> Result<String, String> {
    let mut a = json!({"to": [NOBODY], "subject": c.t(label), "body": "x"});
    a.as_object_mut().unwrap().extend(extra.as_object().cloned().unwrap_or_default());
    let r = w.call("create_draft", a).await?;
    let id = s(&r, "id");
    if id.is_empty() { return Err(format!("no id: {r}")); }
    c.emails.push(id.clone());
    Ok(id)
}

async fn event(w: &mut Wire, c: &mut Ctx, label: &str, start: &str, end: &str, extra: Value) -> Result<String, String> {
    let mut a = json!({"subject": c.t(label), "start": start, "end": end, "notify": false});
    a.as_object_mut().unwrap().extend(extra.as_object().cloned().unwrap_or_default());
    let r = w.call("create_event", a).await?;
    let id = s(&r, "id");
    if id.is_empty() { return Err(format!("no id: {r}")); }
    c.events.push(id.clone());
    Ok(id)
}

/// `list_events` for this run's `label` between two dates.
async fn evs(w: &mut Wire, c: &Ctx, label: &str, after: &str, before: &str, extra: Value) -> Result<Vec<Value>, String> {
    let mut a = json!({"query": c.q(label), "start_after": after, "start_before": before});
    a.as_object_mut().unwrap().extend(extra.as_object().cloned().unwrap_or_default());
    w.call("list_events", a).await.map(|v| v.as_array().cloned().unwrap_or_default())
}

/// `list_emails` ids in `folder` for this run's query `rest`.
async fn mails(w: &mut Wire, c: &Ctx, folder: &str, rest: &str, extra: Value) -> Result<Vec<String>, String> {
    let mut a = json!({"folder": folder, "query": c.q(rest), "count": 200});
    a.as_object_mut().unwrap().extend(extra.as_object().cloned().unwrap_or_default());
    w.call("list_emails", a).await.map(|v| ids_of(&v))
}

/// `(start, end, all_day)` of an event, from get_event.
async fn times(w: &mut Wire, id: &str) -> (String, String, bool) {
    match w.call("get_event", json!({"event_id": id, "include": []})).await {
        Ok(v) => (s(&v, "start"), s(&v, "end"), v["all_day"] == true),
        Err(e) => (format!("ERR {e}"), String::new(), false),
    }
}

async fn body(w: &mut Wire, log: &mut Log, c: &mut Ctx) {
    // ---- setup: folders, Deleted Items count -----------------------------------
    let r = ps(format!(
        "$i=$ns.GetDefaultFolder(6); $f=$i.Folders.Add({}); $n=$f.Folders.Add({}); 'ok ' + $f.Name",
        psq(&c.sub), psq(&c.nested))).await;
    let folder_ok = r.as_ref().is_ok_and(|o| o.starts_with("ok ") && o.contains("תיקייה"));
    println!("folder setup: {r:?}");
    let deleted_classes = ps("$d=$ns.GetDefaultFolder(3); $g=$d.Items | Group-Object MessageClass | Sort-Object Count -Descending; \
        ($g | ForEach-Object { $_.Name + '=' + $_.Count }) -join ', '".into()).await;
    println!("Deleted Items by class at start: {deleted_classes:?}");

    // ---- SEED --------------------------------------------------------------------
    let mut seed_fail: Vec<String> = Vec::new();
    if !folder_ok { seed_fail.push(format!("subfolder: {r:?}")); }
    // The one self-loop email; it arrives while the rest runs.
    let selfloop = c.t("SELFLOOP-mail");
    let r = w.call("send_email", json!({"to": [SELF], "subject": selfloop, "body": "edge self-loop"})).await;
    if r.is_err() { seed_fail.push(format!("self-loop: {r:?}")); } else { c.inbound += 1; }
    let sent_at = Instant::now();
    let q_labels = ["50% off", "50 percent off", "a_b", "axb", "[draft] x", "draft x", "O'Brien",
        "say \"hi\"", "back\\slash", "star*lit", "star-lit"];
    let mut qd: Vec<(String, String)> = Vec::new();
    for l in q_labels {
        match draft(w, c, l, json!({})).await {
            Ok(id) => qd.push((l.to_string(), id)),
            Err(e) => seed_fail.push(format!("draft {l}: {e}")),
        }
    }
    // Attachments: two a.txt with different bytes, a_b.txt, and .a_b.txt (the
    // leading dot is trimmed by safe_filename, standing in for a:b.txt, which
    // no Windows file can be named). CON.txt can't be attached at all (see AT3).
    let att = c.root.join("att");
    let att_src: Vec<(PathBuf, Vec<u8>)> = vec![
        (att.join("d1").join("a.txt"), b"first a.txt".to_vec()),
        (att.join("d2").join("a.txt"), b"second a.txt, different bytes".to_vec()),
        (att.join("a_b.txt"), b"underscore a_b".to_vec()),
        (att.join(".a_b.txt"), b"leading-dot .a_b".to_vec()),
    ];
    for (p, b) in &att_src {
        let _ = std::fs::create_dir_all(p.parent().unwrap());
        if let Err(e) = std::fs::write(p, b) { seed_fail.push(format!("{}: {e}", p.display())); }
    }
    let paths: Vec<String> = att_src.iter().map(|(p, _)| p.to_string_lossy().into_owned()).collect();
    let at_draft = draft(w, c, "AT-attachments", json!({"attachments": paths})).await;
    if let Err(e) = &at_draft { seed_fail.push(format!("attachments draft: {e}")); }
    // Free/busy and DST events, seeded early so free/busy is published by the time it's read.
    let busy_seeds = [
        ("AV1-a", "2099-08-04T09:00", "2099-08-04T10:00", "busy"),
        ("AV1-b", "2099-08-05T09:00", "2099-08-05T10:00", "busy"),
        ("AV2-a", "2099-08-10T23:30", "2099-08-11T00:00", "busy"),
        ("AV2-b", "2099-08-11T00:00", "2099-08-11T00:30", "out_of_office"),
        ("AV3-far", "2099-12-01T10:00", "2099-12-01T11:00", "busy"),
        ("DST3-spring", "2099-03-27T10:00", "2099-03-27T11:00", "busy"),
        ("DST3-fall", "2099-10-25T10:00", "2099-10-25T11:00", "busy"),
    ];
    for (l, a, b, st) in busy_seeds {
        if let Err(e) = event(w, c, l, a, b, json!({"show_as": st})).await { seed_fail.push(format!("{l}: {e}")); }
    }
    log.check("SEED", seed_fail.is_empty(), format!(
        "subfolder + nested, self-loop sent, {} query drafts, attachments draft, {} free/busy events; failures {seed_fail:?}",
        qd.len(), busy_seeds.len()));

    // ---- AT: attachments with clashing names ----------------------------------------
    match &at_draft {
        Ok(id) => {
            let list = w.call("list_attachments", json!({"email_id": id})).await;
            let names: Vec<String> = rows(&list).iter().map(|a| s(a, "filename")).collect();
            println!("attachment names: {names:?}");
            let save = c.root.join("save");
            let r = w.call("save_attachments", json!({"email_id": id, "save_dir": save})).await;
            let res = rows(&r);
            // index (1-based) -> (filename, status, saved_to, bytes on disk == source)
            let info: Vec<(String, String, String, bool)> = res.iter().map(|x| {
                let i = x["index"].as_u64().unwrap_or(0) as usize;
                let to = s(x, "saved_to");
                let ok = i >= 1 && i <= att_src.len() && std::fs::read(&to).ok().as_deref() == Some(&att_src[i - 1].1[..]);
                (s(x, "filename"), s(x, "status"), to, ok)
            }).collect();
            println!("save_attachments: {info:?}");
            let pick = |want: &[usize]| -> Vec<&(String, String, String, bool)> {
                res.iter().zip(info.iter()).filter(|(x, _)| want.contains(&(x["index"].as_u64().unwrap_or(0) as usize))).map(|(_, i)| i).collect()
            };
            let a = pick(&[1, 2]);
            let distinct: BTreeSet<&String> = a.iter().map(|x| &x.2).collect();
            log.check("AT1", r.is_ok() && a.len() == 2 && a.iter().all(|x| x.1 == "saved" && x.3) && distinct.len() == 2,
                format!("two attachments named a.txt: statuses {:?}, saved_to {:?}, bytes intact {:?}; {:?}",
                    a.iter().map(|x| &x.1).collect::<Vec<_>>(), a.iter().map(|x| short_path(&x.2)).collect::<Vec<_>>(),
                    a.iter().map(|x| x.3).collect::<Vec<_>>(), r.as_ref().err()));
            let b = pick(&[3, 4]);
            let distinct: BTreeSet<&String> = b.iter().map(|x| &x.2).collect();
            log.check("AT2", b.len() == 2 && b.iter().all(|x| x.1 == "saved" && x.3) && distinct.len() == 2,
                format!("a_b.txt and .a_b.txt (stand-in for a:b.txt: safe_filename trims the leading dot): names {:?}, saved_to {:?}, bytes intact {:?}",
                    b.iter().map(|x| &x.0).collect::<Vec<_>>(), b.iter().map(|x| short_path(&x.2)).collect::<Vec<_>>(), b.iter().map(|x| x.3).collect::<Vec<_>>()));
            let all_on_disk = info.iter().filter(|x| x.1 == "saved").all(|x| Path::new(&x.2).is_file());
            log.check("AT3", all_on_disk,
                format!("every reported save is a file on disk: {all_on_disk}. Not reachable here: CON.txt (Outlook refuses to attach it: \
                    \"Value does not fall within the expected range\"), \"...\" and \"a:b.txt\" (no Windows file has those names; \
                    MAPI refuses to change an attachment's FileName; .eml import hangs), so no attachment can carry them"));
        }
        Err(e) => { for id in ["AT1", "AT2", "AT3"] { log.fail(id, format!("no attachments draft: {e}")); } }
    }

    // ---- CT: categories containing a comma ---------------------------------------------
    // Outlook can't store a comma in a category name (it separates categories
    // with commas), so every tool that writes categories refuses one.
    let comma = |r: &Result<Value, String>| r.as_ref().is_err_and(|e| e.contains("comma"));
    let bad_ev = w.call("create_event", json!({"subject": c.t("CT1-commas"), "start": "2099-02-11T10:00",
        "end": "2099-02-11T10:30", "categories": ["Smith, John", "Red Category"], "notify": false})).await;
    let bad_tk = w.call("create_task", json!({"subject": c.t("CT1-commas task"), "categories": ["Smith, John", "Red Category"]})).await;
    if let Ok(v) = &bad_ev { c.events.push(s(v, "id")); }
    if let Ok(v) = &bad_tk { c.tasks.push(s(v, "id")); }
    let left = evs(w, c, "CT1-commas", "2099-02-11", "2099-02-11", json!({})).await.map(|r| r.len());
    log.check("CT1", comma(&bad_ev) && comma(&bad_tk) && left == Ok(0),
        format!("create_event -> {:?}; create_task -> {:?}; events created: {left:?}", bad_ev.as_ref().err(), bad_tk.as_ref().err()));
    let ev = event(w, c, "CT2-cats", "2099-02-11T11:00", "2099-02-11T11:30", json!({"categories": ["Red Category"]})).await;
    let tk = w.call("create_task", json!({"subject": c.t("CT2-cats task"), "categories": ["Red Category"]})).await
        .map(|v| s(&v, "id"));
    if let Ok(id) = &tk { c.tasks.push(id.clone()); }
    match (&ev, &tk) {
        (Ok(e), Ok(t)) => {
            let red = vec!["Red Category".to_string()];
            let ue = w.call("update_event", json!({"event_id": e, "add_categories": ["Smith, John"], "notify": false})).await;
            let ut = w.call("update_task", json!({"task_id": t, "add_categories": ["Smith, John"]})).await;
            let ge = w.call("get_event", json!({"event_id": e, "include": []})).await.unwrap_or_default();
            let gt = w.call("get_task", json!({"task_id": t, "include": []})).await.unwrap_or_default();
            log.check("CT2", comma(&ue) && comma(&ut) && cats(&ge) == red && cats(&gt) == red,
                format!("add \"Smith, John\": event {:?} -> {:?}; task {:?} -> {:?}", ue.as_ref().err(), cats(&ge), ut.as_ref().err(), cats(&gt)));
            let ue = w.call("update_event", json!({"event_id": e, "remove_categories": ["Red Category"], "notify": false})).await;
            let ut = w.call("update_task", json!({"task_id": t, "remove_categories": ["Red Category"]})).await;
            let ge = w.call("get_event", json!({"event_id": e, "include": []})).await.unwrap_or_default();
            let gt = w.call("get_task", json!({"task_id": t, "include": []})).await.unwrap_or_default();
            log.check("CT3", ue.is_ok() && ut.is_ok() && cats(&ge).is_empty() && cats(&gt).is_empty(),
                format!("remove \"Red Category\": event {:?}, task {:?}; {:?} {:?}", cats(&ge), cats(&gt), ue.err(), ut.err()));
        }
        _ => { for id in ["CT2", "CT3"] { log.fail(id, format!("seed failed: {:?} {:?}", ev, tk)); } }
    }

    // ---- Q: wildcard and quote characters in queries ------------------------------------
    let label_of = |id: &String| qd.iter().find(|(_, i)| i == id).map(|(l, _)| l.clone()).unwrap_or_else(|| "other".into());
    let q_one = |got: &Result<Vec<String>, String>| -> Vec<String> {
        got.as_ref().map(|v| v.iter().map(&label_of).collect()).unwrap_or_else(|e| vec![format!("ERR {e}")])
    };
    let r = mails(w, c, "drafts", "\"50% off\"", json!({})).await;
    let got = q_one(&r);
    log.check("Q1", got == vec!["50% off".to_string()], format!("query \"50% off\" -> {got:?} (want only the 50% off draft, not \"50 percent off\")"));
    let r = mails(w, c, "drafts", "a_b", json!({})).await;
    let got = q_one(&r);
    log.check("Q2", got == vec!["a_b".to_string()], format!("query a_b -> {got:?} (want only a_b, not axb)"));
    let mut q3 = Vec::new();
    let mut q3_ok = true;
    for (query, want) in [("\"[draft] x\"", "[draft] x"), ("O'Brien", "O'Brien"), ("say \"hi\"", "say \"hi\""), ("back\\slash", "back\\slash")] {
        let r = mails(w, c, "drafts", query, json!({})).await;
        let got = q_one(&r);
        q3_ok &= got == vec![want.to_string()];
        q3.push(format!("{query} -> {got:?}"));
    }
    log.check("Q3", q3_ok, q3.join("; "));
    let r = mails(w, c, "drafts", "\"star*lit\"", json!({})).await;
    let got = q_one(&r);
    let mut sorted = got.clone();
    sorted.sort();
    log.check("Q4", got == vec!["star*lit".to_string()] || sorted == vec!["star*lit".to_string(), "star-lit".to_string()],
        format!("query \"star*lit\" -> {got:?} (the tool documents * as a wildcard, quotes or not, so star-lit matching too is documented)"));
    let all_ours = mails(w, c, "drafts", "", json!({})).await.map(|v| v.len()).unwrap_or(0);
    let r1 = mails(w, c, "drafts", "subject:", json!({})).await;
    let r2 = mails(w, c, "drafts", "*", json!({})).await;
    log.check("Q5", r1.is_ok() && r2.is_ok(),
        format!("\"subject:\" -> {:?}; bare * -> {:?} (this run has {all_ours} drafts: both act as no text filter)",
            r1.as_ref().map(Vec::len), r2.as_ref().map(Vec::len)));
    let mut q6 = Vec::new();
    let mut q6_ok = true;
    for (k, v) in [("from", "o'brien"), ("to", "50%"), ("from", "%"), ("to", "%"), ("to", "_")] {
        let r = mails(w, c, "drafts", "", json!({k: v})).await;
        q6_ok &= r.as_ref().is_ok_and(Vec::is_empty);
        q6.push(format!("{k}: {v:?} -> {:?}", r.map(|x| x.len())));
    }
    log.check("Q6", q6_ok, format!("{} (want no error and no match: no draft is from/to those strings)", q6.join("; ")));
    let mut q7_ids = Vec::new();
    for l in ["Q7 50% off", "Q7 50 percent off", "Q7 axb"] {
        if let Ok(id) = event(w, c, l, "2099-02-10T10:00", "2099-02-10T10:30", json!({})).await { q7_ids.push(id); }
    }
    let r = evs(w, c, "\"50% off\"", "2099-02-10", "2099-02-10", json!({})).await;
    let got = r.as_ref().map(|v| subjects_of(&json!(v))).unwrap_or_default();
    log.check("Q7", q7_ids.len() == 3 && got == vec![c.t("Q7 50% off")], format!("list_events \"50% off\" -> {got:?}; {:?}", r.err()));

    // ---- AD: all-day and multi-day boundaries ------------------------------------------
    match event(w, c, "AD1-allday", "2099-07-01T00:00", "2099-07-01T00:00", json!({"all_day": true})).await {
        Ok(id) => {
            let (st, en, ad) = times(w, &id).await;
            let on1 = evs(w, c, "AD1-allday", "2099-07-01", "2099-07-01", json!({})).await.map(|r| r.len());
            let on2 = evs(w, c, "AD1-allday", "2099-07-02", "2099-07-02", json!({})).await.map(|r| r.len());
            log.check("AD1", ad && st == "2099-07-01T00:00:00" && en == "2099-07-02T00:00:00" && on1 == Ok(1) && on2 == Ok(0),
                format!("start = end = 07-01T00:00, all_day -> all_day {ad}, {st} .. {en}; listed on 07-01 {on1:?}, on 07-02 {on2:?}"));
        }
        Err(e) => log.fail("AD1", e),
    }
    match event(w, c, "AD2-timed-allday", "2099-07-02T10:00", "2099-07-02T11:00", json!({"all_day": true})).await {
        Ok(id) => {
            let (st, en, ad) = times(w, &id).await;
            log.check("AD2", ad && st == "2099-07-02T00:00:00" && en == "2099-07-03T00:00:00",
                format!("all_day with 10:00-11:00 -> all_day {ad}, {st} .. {en} (want the whole day 07-02)"));
        }
        Err(e) => { log.check("AD2", true, format!("refused with an error: {e}")); }
    }
    match event(w, c, "AD3-threeday", "2099-07-10", "2099-07-13", json!({"all_day": true})).await {
        Ok(id) => {
            let (st, en, ad) = times(w, &id).await;
            let mut on = Vec::new();
            for d in 9..=13 {
                let day = format!("2099-07-{d:02}");
                on.push(evs(w, c, "AD3-threeday", &day, &day, json!({})).await.map(|r| r.len()).unwrap_or(99));
            }
            log.check("AD3a", ad && st == "2099-07-10T00:00:00" && en == "2099-07-13T00:00:00" && on[0] == 0 && on[1] == 1 && on[4] == 0,
                format!("07-10..07-12 all-day: {st} .. {en} all_day {ad}; rows on 07-09..07-13 = {on:?}"));
            log.check("AD3b", on[2] == 0 && on[3] == 0,
                format!("on 07-11 and 07-12: {} and {} rows; list_events documents that start_after/start_before bound the event START, \
                    so a multi-day event is listed on its first day only (big-calendar BC7 asserts the same)", on[2], on[3]));
        }
        Err(e) => { log.fail("AD3a", e.clone()); log.fail("AD3b", e); }
    }
    let mut ad4 = Vec::new();
    let mut ad4_ok = true;
    for (label, a, b, all_day, flip) in [("AD4-was-allday", "2099-07-20", "2099-07-21", true, false),
                                         ("AD4-was-timed", "2099-07-22T10:00", "2099-07-22T11:00", false, true)] {
        match event(w, c, label, a, b, json!({"all_day": all_day})).await {
            Ok(id) => {
                let before = times(w, &id).await;
                let u = w.call("update_event", json!({"event_id": id, "all_day": flip, "notify": false})).await;
                let mid = times(w, &id).await;
                let undo = u.as_ref().map(|v| v["undo"].clone()).unwrap_or_default();
                let back = w.call("update_event", undo.clone()).await;
                let after = times(w, &id).await;
                ad4_ok &= u.is_ok() && back.is_ok() && after == before && mid.2 == flip;
                ad4.push(format!("{label}: before {before:?}, after all_day={flip} {mid:?}, after undo {after:?} (undo {undo})"));
            }
            Err(e) => { ad4_ok = false; ad4.push(format!("{label}: {e}")); }
        }
    }
    log.check("AD4", ad4_ok, ad4.join("; "));
    match event(w, c, "AD5-allday-daily", "2099-07-25", "2099-07-26", json!({"all_day": true, "recurrence": {"pattern": "daily", "occurrences": 5}})).await {
        Ok(_) => {
            let r = evs(w, c, "AD5-allday-daily", "2099-07-24", "2099-07-30", json!({})).await;
            let rs = r.clone().unwrap_or_default();
            let want: Vec<String> = (25..=29).map(|d| format!("2099-07-{d}T00:00:00")).collect();
            log.check("AD5", starts(&rs) == want && rs.iter().all(|x| x["all_day"] == true),
                format!("all-day daily x5: starts {:?}, all all_day {}; {:?}", starts(&rs), rs.iter().all(|x| x["all_day"] == true), r.err()));
        }
        Err(e) => log.fail("AD5", e),
    }

    // ---- DST: daylight saving time (Israel: 2099-03-27 02:00->03:00, 2099-10-25 02:00->01:00) ----
    let tz = ps("$z=[TimeZoneInfo]::Local; $z.Id + '|' + $z.SupportsDaylightSavingTime + '|' + $z.IsInvalidTime([datetime]'2099-03-27T02:30') + '|' + $z.IsAmbiguousTime([datetime]'2099-10-25T01:30')".into()).await;
    println!("time zone: {tz:?}");
    let zone_ok = tz.as_ref().is_ok_and(|z| z.contains("|True|True|True"));
    if !zone_ok {
        for id in ["DST1", "DST2", "DST3", "DST4"] {
            log.check(id, true, format!("skipped: the zone's 2099 transitions are not 03-27 02:00 / 10-25 02:00 ({tz:?})"));
        }
    } else {
        let mut dst2 = Vec::new();
        let mut dst2_ok = true;
        for (label, a, b, d0) in [("DST2-spring", "2099-03-25T09:00", "2099-03-25T09:15", at(2099, 3, 25, 9, 0)),
                                  ("DST2-fall", "2099-10-23T09:00", "2099-10-23T09:15", at(2099, 10, 23, 9, 0))] {
            match event(w, c, label, a, b, json!({"recurrence": {"pattern": "daily", "occurrences": 5}})).await {
                Ok(_) => {
                    let after = (d0 - chrono::Duration::days(1)).format("%Y-%m-%d").to_string();
                    let before = (d0 + chrono::Duration::days(6)).format("%Y-%m-%d").to_string();
                    let rs = evs(w, c, label, &after, &before, json!({})).await.unwrap_or_default();
                    let want: Vec<String> = (0..5).map(|i| iso(d0 + chrono::Duration::days(i))).collect();
                    let occ: BTreeSet<String> = rs.iter().map(|x| s(x, "occurrence_id")).collect();
                    let ok = starts(&rs) == want && occ.len() == 5 && !occ.contains("");
                    dst2_ok &= ok;
                    dst2.push(format!("{label}: starts {:?}, {} distinct occurrence ids", starts(&rs), occ.len()));
                }
                Err(e) => { dst2_ok = false; dst2.push(format!("{label}: {e}")); }
            }
        }
        log.check("DST2", dst2_ok, dst2.join("; "));
        let mut dst4 = Vec::new();
        let mut dst4_ok = true;
        for (day, wants) in [("2099-03-27", ["DST2-spring", "DST3-spring"]), ("2099-10-25", ["DST2-fall", "DST3-fall"])] {
            let r = evs(w, c, "DST", day, day, json!({})).await;
            let rs = r.clone().unwrap_or_default();
            let keys: Vec<String> = rs.iter().map(|x| format!("{}@{}", s(x, "subject").rsplit(' ').next().unwrap_or(""), s(x, "start"))).collect();
            let distinct: BTreeSet<&String> = keys.iter().collect();
            let present = wants.iter().all(|l| rs.iter().filter(|x| s(x, "subject") == c.t(l)).count() == 1);
            dst4_ok &= r.is_ok() && present && distinct.len() == keys.len();
            dst4.push(format!("{day}: {keys:?}{}", r.err().map(|e| format!(" ERROR {e}")).unwrap_or_default()));
        }
        log.check("DST4", dst4_ok, format!("{} (each seeded event once)", dst4.join("; ")));
        // DST1 last: the skipped-hour create is the risky one.
        for (id, label, a, b) in [("DST1a", "DST1-skipped", "2099-03-27T02:30", "2099-03-27T03:00"),
                                  ("DST1b", "DST1-skipped-long", "2099-03-27T02:30", "2099-03-27T04:00"),
                                  ("DST1c", "DST1-repeated", "2099-10-25T01:30", "2099-10-25T02:00")] {
            // A start in the skipped hour doesn't exist on the local clock: refused.
            let skipped = id != "DST1c";
            match event(w, c, label, a, b, json!({})).await {
                Ok(eid) => {
                    let (st, en, _) = times(w, &eid).await;
                    log.check(id, !skipped && !st.starts_with("ERR"), format!("{a}..{b} -> stored {st} .. {en}"));
                }
                Err(e) => { log.check(id, skipped && e.contains("doesn't exist"), format!("{a}..{b} -> refused: {e}")); }
            }
        }
        // After a failed skipped-hour create: does the day still list, and is anything left behind?
        let r = evs(w, c, "DST", "2099-03-27", "2099-03-27", json!({})).await;
        let stray = ps(format!(
            "$tag={}; $hits=@(); foreach ($f in @($ns.GetDefaultFolder(9), $ns.GetDefaultFolder(6).Parent.Folders.Item('Sync Issues').Folders.Item('Local Failures'))) \
             {{ foreach ($it in $f.Items) {{ if (([string]$it.Subject) -eq $tag) {{ $hits += $f.Name }} }} }}; 'found in: ' + ($hits -join ', ')",
            psq(&c.t("DST1-skipped")))).await;
        let listed = match &r {
            Ok(v) => format!("{:?}", v.iter().map(|x| format!("{}@{}", s(x, "subject").rsplit(' ').next().unwrap_or(""), s(x, "start"))).collect::<Vec<_>>()),
            Err(e) => format!("ERROR {e}"),
        };
        log.check("DST1d", r.is_ok() && stray.as_ref().is_ok_and(|x| x.trim_end() == "found in:"),
            format!("list_events over 2099-03-27 after DST1a: {listed}; items titled DST1-skipped (raw COM, Calendar + Local Failures): {stray:?}"));
    }

    // ---- AV: check_availability intervals and windows (and DST3) --------------------------
    // Free/busy is published asynchronously: wait until the AV1 event shows.
    let started = Instant::now();
    let mut published = false;
    while started.elapsed().as_secs() < 180 {
        // Both AV1 days: one day can publish well before the other (seen live).
        if let Ok(v) = w.call("check_availability", json!({"people": [SELF], "start": "2099-08-04T08:00", "end": "2099-08-05T12:00", "interval_minutes": 30})).await {
            let g = got_slots(&v, 0);
            let busy = |at: &str| g.iter().any(|(st, x)| st == at && x == "busy");
            if busy("2099-08-04T09:00:00") && busy("2099-08-05T09:00:00") { published = true; break; }
        }
        tokio::time::sleep(Duration::from_secs(10)).await;
    }
    println!("free/busy published: {published} after {:.0}s", started.elapsed().as_secs_f64());
    let av1_busy = [(at(2099, 8, 4, 9, 0), at(2099, 8, 4, 10, 0), "busy"), (at(2099, 8, 5, 9, 0), at(2099, 8, 5, 10, 0), "busy")];
    let r = avail(w, &[SELF], "2099-08-04T08:00".into(), "2099-08-04T12:00".into(), 0).await;
    let g = r.as_ref().map(|v| got_slots(v, 0)).unwrap_or_default();
    log.check("AV1a", r.is_err(), format!("interval_minutes 0 -> {}: {} slots of {} (the tool doesn't document 0: rejected or a documented clamp expected)",
        if r.is_ok() { "accepted" } else { "error" }, g.len(),
        g.first().map(|x| format!("{}..", x.0)).unwrap_or_default()));
    if let Err(e) = &r { println!("AV1a error: {e}"); }
    let mut av1 = Vec::new();
    let mut av1_ok = true;
    for (id, iv, a, b) in [("AV1b", 7, at(2099, 8, 4, 8, 0), at(2099, 8, 5, 12, 0)),
                           ("AV1c", 45, at(2099, 8, 4, 8, 0), at(2099, 8, 5, 12, 0)),
                           ("AV1d", 1440, at(2099, 8, 4, 0, 0), at(2099, 8, 7, 0, 0))] {
        let r = avail(w, &[SELF], iso(a), iso(b), iv).await;
        let g = r.as_ref().map(|v| got_slots(v, 0)).unwrap_or_default();
        let want = expect_grid(a, b, iv, &av1_busy);
        let ok = r.is_ok() && g == want;
        let busy: Vec<&String> = g.iter().filter(|x| x.1 != "free").map(|x| &x.0).collect();
        log.check(id, ok, format!("interval {iv} over {}..{}: {}; busy slots {busy:?}; {:?}", iso(a), iso(b), slot_diff(&g, &want), r.as_ref().err()));
        av1_ok &= ok;
        av1.push(id);
    }
    let _ = (av1, av1_ok);
    let av2_busy = [(at(2099, 8, 10, 23, 30), at(2099, 8, 11, 0, 0), "busy"), (at(2099, 8, 11, 0, 0), at(2099, 8, 11, 0, 30), "out_of_office")];
    let mut av2 = Vec::new();
    let mut av2_ok = true;
    for (a, b) in [(at(2099, 8, 10, 23, 30), at(2099, 8, 11, 0, 30)), (at(2099, 8, 10, 23, 30), at(2099, 8, 13, 23, 30))] {
        let r = avail(w, &[SELF], iso(a), iso(b), 30).await;
        let g = r.as_ref().map(|v| got_slots(v, 0)).unwrap_or_default();
        let want = expect_grid(a, b, 30, &av2_busy);
        av2_ok &= r.is_ok() && g == want;
        av2.push(format!("{}..{}: {}; first {:?}", iso(a), iso(b), slot_diff(&g, &want), g.iter().take(3).collect::<Vec<_>>()));
    }
    log.check("AV2", av2_ok, av2.join("; "));
    let (a, b) = (at(2099, 12, 1, 9, 0), at(2099, 12, 1, 12, 0));
    let r = avail(w, &[SELF], iso(a), iso(b), 30).await;
    let g = r.as_ref().map(|v| got_slots(v, 0)).unwrap_or_default();
    let want = expect_grid(a, b, 30, &[(at(2099, 12, 1, 10, 0), at(2099, 12, 1, 11, 0), "busy")]);
    log.check("AV3", r.is_ok() && g == want && r.as_ref().is_ok_and(|v| v["people"][0]["resolved"] == true),
        format!("2099-12-01 (over 40 days out): {}; {:?}", slot_diff(&g, &want), r.as_ref().err()));
    let (a, b) = (at(2099, 8, 4, 8, 0), at(2099, 8, 4, 12, 0));
    let r = avail(w, &[SELF, NOBODY], iso(a), iso(b), 30).await;
    let g = r.as_ref().map(|v| got_slots(v, 0)).unwrap_or_default();
    let want = expect_grid(a, b, 30, &av1_busy);
    let nob = r.as_ref().map(|v| (v["people"][1]["resolved"].clone(), v["people"][1]["slots"].as_array().map_or(0, Vec::len))).ok();
    log.check("AV4", r.is_ok() && g == want && nob == Some((json!(false), 0)),
        format!("you: {}; nobody@example.invalid: (resolved, slots) {nob:?}; {:?}", slot_diff(&g, &want), r.as_ref().err()));
    if zone_ok {
        for (id, day, next) in [("DST3a", at(2099, 3, 27, 0, 0), at(2099, 3, 28, 0, 0)), ("DST3b", at(2099, 10, 25, 0, 0), at(2099, 10, 26, 0, 0))] {
            // Each day's free/busy can publish later than AV1's (seen once:
            // August and December published, both DST days still all free).
            // Wait until the day shows anything busy.
            let waited = Instant::now();
            let mut r = avail(w, &[SELF], iso(day), iso(next), 30).await;
            while waited.elapsed().as_secs() < 180
                && r.as_ref().map(|v| got_slots(v, 0)).unwrap_or_default().iter().all(|x| x.1 == "free")
            {
                tokio::time::sleep(Duration::from_secs(10)).await;
                r = avail(w, &[SELF], iso(day), iso(next), 30).await;
            }
            println!("{id}: free/busy for {} after {:.0}s", iso(day), waited.elapsed().as_secs_f64());
            let g = r.as_ref().map(|v| got_slots(v, 0)).unwrap_or_default();
            let status = |h: u32, m: u32| g.iter().find(|x| x.0 == iso(day.date().and_hms_opt(h, m, 0).unwrap())).map(|x| x.1.clone()).unwrap_or_default();
            let real_hours = if day.month() == 3 { 23 } else { 25 };
            let busy: Vec<String> = g.iter().filter(|x| x.1 != "free").map(|x| x.0[11..16].to_string()).collect();
            let ok = r.is_ok() && status(10, 0) == "busy" && status(10, 30) == "busy" && status(11, 0) == "free" && status(9, 30) == "free";
            log.check(id, ok, format!("{}: {} slots (the day really has {real_hours} h = {} half-hours); busy at {busy:?} \
                (want 10:00 and 10:30 busy for the 10:00-11:00 event, 09:00 for DST2's 09:00-09:15 occurrence, plus DST1's event); {:?}",
                day.format("%Y-%m-%d"), g.len(), real_hours * 2, r.as_ref().err()));
        }
    }

    // ---- RC: recurrence edge cases ------------------------------------------------------
    let mon = first_weekday(2099, 9, Weekday::Mon);
    let d = |n: i64| iso(mon.and_hms_opt(8, 0, 0).unwrap() + chrono::Duration::days(n));
    match event(w, c, "RC1-biweekly", &format!("{}T08:00", mon), &format!("{}T08:30", mon),
        json!({"recurrence": {"pattern": "weekly", "days_of_week": ["monday", "wednesday"], "interval": 2, "occurrences": 6}})).await {
        Ok(_) => {
            let rs = evs(w, c, "RC1-biweekly", "2099-09-01", "2099-10-31", json!({})).await.unwrap_or_default();
            let want = vec![d(0), d(2), d(14), d(16), d(28), d(30)];
            log.check("RC1", starts(&rs) == want, format!("Mon+Wed every 2 weeks x6 from {mon}: {:?} (want {want:?})", starts(&rs)));
        }
        Err(e) => log.fail("RC1", e),
    }
    match event(w, c, "RC2-day31", "2099-01-31T07:00", "2099-01-31T07:15",
        json!({"recurrence": {"pattern": "monthly", "day_of_month": 31, "occurrences": 4}})).await {
        Ok(_) => {
            let rs = evs(w, c, "RC2-day31", "2099-01-01", "2099-06-30", json!({})).await.unwrap_or_default();
            let st = starts(&rs);
            let months: Vec<String> = st.iter().map(|x| x.get(..7).unwrap_or("").to_string()).collect();
            log.check("RC2", months == vec!["2099-01", "2099-02", "2099-03", "2099-04"],
                format!("monthly on day 31 x4 from 2099-01-31: {st:?} (one per month, none skipped or doubled)"));
        }
        Err(e) => log.fail("RC2", e),
    }
    match event(w, c, "RC3-leap", "2096-02-29T07:00", "2096-02-29T07:15", json!({"recurrence": {"pattern": "yearly", "occurrences": 3}})).await {
        Ok(_) => {
            let rs = evs(w, c, "RC3-leap", "2096-01-01", "2099-12-31", json!({})).await.unwrap_or_default();
            let st = starts(&rs);
            let years: Vec<String> = st.iter().map(|x| x.get(..4).unwrap_or("").to_string()).collect();
            log.check("RC3", years == vec!["2096", "2097", "2098"], format!("yearly x3 from 2096-02-29: {st:?}"));
        }
        Err(e) => log.fail("RC3", e),
    }
    let mut rc4 = Vec::new();
    let mut rc4_ok = true;
    for (label, rec) in [("RC4-until-before", json!({"pattern": "daily", "until": "2099-06-01"})),
                         ("RC4-zero", json!({"pattern": "daily", "occurrences": 0})),
                         ("RC4-negative", json!({"pattern": "daily", "occurrences": -1}))] {
        let r = event(w, c, label, "2099-06-10T07:00", "2099-06-10T07:15", json!({"recurrence": rec})).await;
        let leaked = evs(w, c, label, "2099-05-01", "2099-07-31", json!({})).await.map(|v| v.len()).unwrap_or(0);
        let detail = match &r {
            Ok(id) => {
                let g = w.call("get_event", json!({"event_id": id, "include": []})).await.unwrap_or_default();
                format!("created: recurrence {} ", g["recurrence"])
            }
            Err(e) => format!("error {:?} ", e.chars().take(160).collect::<String>()),
        };
        rc4_ok &= r.is_err() && leaked == 0;
        rc4.push(format!("{label}: {detail}rows {leaked}"));
    }
    log.check("RC4", rc4_ok, rc4.join("; "));
    let jmon = first_weekday(2099, 6, Weekday::Mon);
    match event(w, c, "RC5-switch", &format!("{jmon}T07:00"), &format!("{jmon}T07:15"),
        json!({"recurrence": {"pattern": "weekly", "days_of_week": ["monday"], "occurrences": 3}})).await {
        Ok(id) => {
            let u = w.call("update_event", json!({"event_id": id, "notify": false,
                "recurrence": {"pattern": "monthly", "day_of_month": jmon.day(), "occurrences": 3}})).await;
            let g = w.call("get_event", json!({"event_id": id, "include": []})).await.unwrap_or_default();
            let rs = evs(w, c, "RC5-switch", "2099-06-01", "2099-09-30", json!({})).await.unwrap_or_default();
            let want: Vec<String> = [6, 7, 8].iter().map(|m| iso(at(2099, *m, jmon.day(), 7, 0))).collect();
            log.check("RC5", u.is_ok() && g["recurrence"]["pattern"] == "monthly" && starts(&rs) == want,
                format!("weekly -> monthly day {}: pattern {}, starts {:?}; {:?}", jmon.day(), g["recurrence"], starts(&rs), u.err()));
        }
        Err(e) => log.fail("RC5", e),
    }
    match event(w, c, "RC6-end", "2099-09-20T12:00", "2099-09-20T12:15", json!({"recurrence": {"pattern": "daily", "occurrences": 5}})).await {
        Ok(_) => {
            let rs = evs(w, c, "RC6-end", "2099-09-18", "2099-09-30", json!({})).await.unwrap_or_default();
            let want: Vec<String> = (20..=24).map(|dd| iso(at(2099, 9, dd, 12, 0))).collect();
            log.check("RC6", starts(&rs) == want, format!("daily x5 listed over 09-18..09-30: {:?}", starts(&rs)));
        }
        Err(e) => log.fail("RC6", e),
    }

    // ---- PD: permanent delete edge cases ---------------------------------------------------
    match draft(w, c, "PD1-twice", json!({})).await {
        Ok(id) => {
            let soft = w.call("delete_email", json!({"email_id": id})).await;
            let in_del = mails(w, c, "deleted", "PD1-twice", json!({})).await.unwrap_or_default();
            let hard = match in_del.first() {
                Some(d) => w.call("delete_email", json!({"email_id": d, "permanent": true})).await,
                None => Err("not found in Deleted Items".into()),
            };
            let after = mails(w, c, "deleted", "PD1-twice", json!({})).await;
            log.check("PD1", soft.is_ok() && in_del.len() == 1 && hard.as_ref().is_ok_and(|v| v["permanent"] == true) && after.as_ref().is_ok_and(Vec::is_empty),
                format!("soft delete {:?}; in Deleted Items {}; permanent from there {:?}; left there {:?}",
                    soft.err(), in_del.len(), hard.as_ref().map(|v| v["permanent"].clone()), after.map(|v| v.len())));
        }
        Err(e) => log.fail("PD1", e),
    }
    let mut stale_mail = String::new();
    match draft(w, c, "PD2-again", json!({})).await {
        Ok(id) => {
            let first = w.call("delete_email", json!({"email_id": id, "permanent": true})).await;
            let second = w.call("delete_email", json!({"email_id": id, "permanent": true})).await;
            log.check("PD2", first.is_ok() && second.as_ref().is_err_and(|e| e.to_lowercase().contains("not found")),
                format!("first {:?}; second {:?}", first.as_ref().map(|v| v["status"].clone()), second));
            stale_mail = id;
        }
        Err(e) => log.fail("PD2", e),
    }
    let pd3d = draft(w, c, "PD3-O'Neil draft", json!({})).await;
    let pd3e = event(w, c, "PD3-O'Neil event", "2099-02-16T10:00", "2099-02-16T10:30", json!({})).await;
    match (&pd3d, &pd3e) {
        (Ok(d), Ok(e)) => {
            let a = w.call("delete_email", json!({"email_id": d, "permanent": true})).await;
            let b = w.call("delete_event", json!({"event_id": e, "permanent": true, "notify": false})).await;
            let left_del = mails(w, c, "deleted", "PD3-O'Neil", json!({})).await;
            let left_dr = mails(w, c, "drafts", "PD3-O'Neil", json!({})).await;
            let left_cal = evs(w, c, "PD3-O'Neil", "2099-02-16", "2099-02-16", json!({})).await.map(|v| v.len());
            log.check("PD3", a.is_ok() && b.is_ok() && left_del.as_ref().is_ok_and(Vec::is_empty) && left_dr.as_ref().is_ok_and(Vec::is_empty) && left_cal == Ok(0),
                format!("draft {:?}, event {:?}; left in Deleted Items {:?}, Drafts {:?}, Calendar {left_cal:?}",
                    a.err(), b.err(), left_del.map(|v| v.len()), left_dr.map(|v| v.len())));
        }
        _ => log.fail("PD3", format!("{pd3d:?} {pd3e:?}")),
    }
    match event(w, c, "PD4-occurrence", "2099-11-23T09:00", "2099-11-23T09:30", json!({"recurrence": {"pattern": "daily", "occurrences": 3}})).await {
        Ok(_) => {
            let rs = evs(w, c, "PD4-occurrence", "2099-11-22", "2099-11-27", json!({})).await.unwrap_or_default();
            match rs.get(1) {
                Some(o) => {
                    let r = w.call("delete_event", json!({"event_id": s(o, "occurrence_id"), "permanent": true, "notify": false})).await;
                    let after = evs(w, c, "PD4-occurrence", "2099-11-22", "2099-11-27", json!({})).await.unwrap_or_default();
                    let in_del = mails(w, c, "deleted", "PD4-occurrence", json!({})).await.map(|v| v.len());
                    log.check("PD4", r.as_ref().is_ok_and(|v| v["permanent"] != true) && starts(&after) == vec![s(&rs[0], "start"), s(&rs[2], "start")] && in_del == Ok(0),
                        format!("occurrence delete with permanent: result {}; left {:?}; in Deleted Items {in_del:?}",
                            r.as_ref().map(|v| v.to_string()).unwrap_or_else(|e| e.clone()), starts(&after)));
                }
                None => log.fail("PD4", format!("series has {} occurrences", rs.len())),
            }
        }
        Err(e) => log.fail("PD4", e),
    }
    let mut pd5_id = String::new();
    match draft(w, c, "PD5-subfolder", json!({})).await {
        Ok(id) => {
            let mv = w.call("update_email", json!({"email_id": id, "move_to": c.sub_path()})).await;
            let nid = mv.as_ref().map(|v| s(v, "id")).unwrap_or_default();
            if !nid.is_empty() { c.emails.push(nid.clone()); }
            let del = w.call("delete_email", json!({"email_id": nid, "permanent": true})).await;
            let sub = mails(w, c, &c.sub_path(), "PD5-subfolder", json!({})).await;
            let dl = mails(w, c, "deleted", "PD5-subfolder", json!({})).await;
            log.check("PD5", mv.is_ok() && del.is_ok() && sub.as_ref().is_ok_and(Vec::is_empty) && dl.as_ref().is_ok_and(Vec::is_empty),
                format!("move {:?}; permanent delete {:?}; left in subfolder {:?}, Deleted Items {:?}",
                    mv.as_ref().err(), del.err(), sub.map(|v| v.len()), dl.map(|v| v.len())));
            pd5_id = nid;
        }
        Err(e) => log.fail("PD5", e),
    }
    let _ = pd5_id;
    let mut stale_event = String::new();
    match event(w, c, "PD6-stale", "2099-02-17T10:00", "2099-02-17T10:30", json!({})).await {
        Ok(id) => {
            let a = w.call("delete_event", json!({"event_id": id, "permanent": true, "notify": false})).await;
            let b = w.call("delete_event", json!({"event_id": id, "permanent": true, "notify": false})).await;
            log.check("PD6", a.is_ok() && b.is_err(), format!("first {:?}; again with the stale id {:?}", a.as_ref().err(), b));
            stale_event = id;
        }
        Err(e) => log.fail("PD6", e),
    }
    let max_class = deleted_classes.as_ref().ok()
        .and_then(|x| x.split(", ").next().and_then(|p| p.rsplit('=').next()).and_then(|n| n.parse::<usize>().ok()))
        .unwrap_or(0);
    log.check("PD7", deleted_classes.is_ok() && max_class <= 500,
        format!("not reachable: the largest same-class group in Deleted Items has {max_class} items (soft_delete_then_purge scans 500); \
            not filled to 500 by design ({deleted_classes:?})"));

    // ---- OD: deleting a meeting you organize -------------------------------------------------
    match event(w, c, "OD1-meeting", "2099-11-02T10:00", "2099-11-02T10:30", json!({"required_attendees": ["od1@example.invalid"]})).await {
        Ok(id) => {
            let r = w.call("delete_event", json!({"event_id": id, "notify": false, "permanent": false})).await;
            let sent = mails(w, c, "sent", "OD1-meeting", json!({})).await;
            let outbox = mails(w, c, "outbox", "OD1-meeting", json!({})).await;
            let del = mails(w, c, "deleted", "OD1-meeting", json!({})).await;
            let canc = evs(w, c, "OD1-meeting", "2099-11-02", "2099-11-02", json!({"is_canceled": true})).await.map(|v| v.len());
            let any = evs(w, c, "OD1-meeting", "2099-11-02", "2099-11-02", json!({})).await.map(|v| v.len());
            log.check("OD1", r.is_ok() && sent.as_ref().is_ok_and(Vec::is_empty) && outbox.as_ref().is_ok_and(Vec::is_empty)
                && del.as_ref().is_ok_and(|v| v.len() == 1) && canc == Ok(0) && any == Ok(0),
                format!("delete quietly: {:?}; Sent Items {:?}, Outbox {:?}, Deleted Items {:?}; calendar is_canceled {canc:?}, all {any:?}",
                    r.as_ref().map(|v| s(v, "note")), sent.map(|v| v.len()), outbox.map(|v| v.len()), del.map(|v| v.len())));
        }
        Err(e) => log.fail("OD1", e),
    }
    match event(w, c, "OD2-series", "2099-11-09T10:00", "2099-11-09T10:30",
        json!({"required_attendees": ["od2@example.invalid"], "recurrence": {"pattern": "daily", "occurrences": 3}})).await {
        Ok(id) => {
            let rs = evs(w, c, "OD2-series", "2099-11-08", "2099-11-13", json!({})).await.unwrap_or_default();
            match rs.get(1) {
                Some(o) => {
                    let r = w.call("delete_event", json!({"event_id": s(o, "occurrence_id"), "notify": false})).await;
                    let after = evs(w, c, "OD2-series", "2099-11-08", "2099-11-13", json!({})).await.unwrap_or_default();
                    let g = w.call("get_event", json!({"event_id": id, "include": []})).await;
                    let canceled: Vec<bool> = after.iter().map(|x| x["is_canceled"] == true).collect();
                    let sent = mails(w, c, "sent", "OD2-series", json!({})).await.map(|v| v.len());
                    log.check("OD2", r.is_ok() && after.len() == 2 && canceled == vec![false, false]
                        && g.as_ref().is_ok_and(|v| v["is_canceled"] == false) && sent == Ok(0),
                        format!("one occurrence deleted quietly: {:?}; left {:?}, is_canceled {canceled:?}; series is_canceled {:?}; Sent Items {sent:?}",
                            r.as_ref().err(), starts(&after), g.map(|v| v["is_canceled"].clone())));
                }
                None => log.fail("OD2", format!("series has {} occurrences", rs.len())),
            }
        }
        Err(e) => log.fail("OD2", e),
    }

    // ---- PG: paging under mutation -----------------------------------------------------------
    let mut pg = Vec::new();
    for i in 0..60 {
        if let Ok(id) = draft(w, c, &format!("PGmail {i:02}"), json!({})).await { pg.push(id); }
    }
    let full = page(w, c, 0, 60).await.unwrap_or_default();
    let p1 = page(w, c, 0, 25).await.unwrap_or_default();
    let p1b = page(w, c, 0, 25).await.unwrap_or_default();
    let gone: Vec<String> = p1.iter().take(5).cloned().collect();
    let del = w.call("delete_email", json!({"email_id": gone, "permanent": true})).await;
    let p2 = page(w, c, 25, 25).await;
    let want = full.get(30..55).map(|x| x.to_vec()).unwrap_or_default();
    let p2v = p2.clone().unwrap_or_default();
    let shift = if p2v == want { "5 rows skipped (offset semantics)".to_string() }
        else if Some(&p2v[..]) == full.get(25..50) { "no shift".to_string() }
        else { format!("other: first {:?}", p2v.first().and_then(|x| full.iter().position(|y| y == x))) };
    log.check("PG1", pg.len() == 60 && full.len() == 60 && Some(&p1[..]) == full.get(..25) && del.is_ok() && p2.is_ok() && p2v == want,
        format!("60 drafts; page 1 == full[0..25] {}; deleted 5 of page 1 ({:?}); page 2: {} rows, {shift}",
            Some(&p1[..]) == full.get(..25), del.as_ref().err(), p2v.len()));
    log.check("PG2", !p1.is_empty() && p1 == p1b, format!("two identical calls back to back: identical {}", p1 == p1b));
    let remaining: Vec<String> = pg.iter().filter(|x| !gone.contains(x)).cloned().collect();
    let mut pge = Vec::new();
    for i in 0..60u32 {
        let (day, slot) = (3 + i / 20, i % 20);
        let st = at(2099, 5, day, 8 + slot / 2, (slot % 2) * 30);
        if let Ok(id) = event(w, c, &format!("PGevent {i:02}"), &st.format("%Y-%m-%dT%H:%M").to_string(),
            &(st + chrono::Duration::minutes(15)).format("%Y-%m-%dT%H:%M").to_string(), json!({})).await { pge.push(id); }
    }
    let efull = epage(w, c, 0, 60).await.unwrap_or_default();
    let ep1 = epage(w, c, 0, 25).await.unwrap_or_default();
    let egone: Vec<String> = ep1.iter().take(5).cloned().collect();
    let edel = w.call("delete_event", json!({"event_id": egone, "permanent": true, "notify": false})).await;
    let ep2 = epage(w, c, 25, 25).await.unwrap_or_default();
    let ewant = efull.get(30..55).map(|x| x.to_vec()).unwrap_or_default();
    log.check("PG3", pge.len() == 60 && efull.len() == 60 && edel.is_ok() && ep2 == ewant,
        format!("60 events; deleted 5 of page 1 ({:?}); page 2: {} rows, == full[30..55] {} (== full[25..50] {})",
            edel.as_ref().err(), ep2.len(), ep2 == ewant, Some(&ep2[..]) == efull.get(25..50)));

    // ---- FO: non-default folders and calendar_of ------------------------------------------
    match draft(w, c, "FO1-mover", json!({})).await {
        Ok(id) => {
            let mv = w.call("update_email", json!({"email_id": id, "move_to": c.sub_path()})).await;
            let nid = mv.as_ref().map(|v| s(v, "id")).unwrap_or_default();
            if !nid.is_empty() { c.emails.push(nid.clone()); }
            let there = mails(w, c, &c.sub_path(), "FO1-mover", json!({})).await;
            let opened = w.call("get_email", json!({"email_id": nid, "include": []})).await;
            log.check("FO1", mv.is_ok() && there.as_ref().is_ok_and(|v| v == &vec![nid.clone()]) && opened.as_ref().is_ok_and(|v| s(v, "subject") == c.t("FO1-mover")),
                format!("moved to the subfolder: {:?}; listed there under the new id {:?}; get_email {:?}",
                    mv.as_ref().err(), there.as_ref().map(|v| v.contains(&nid)), opened.as_ref().map(|v| s(v, "subject")).err()));
            let back = w.call("update_email", json!({"email_id": nid, "move_to": "Drafts"})).await;
            let bid = back.as_ref().map(|v| s(v, "id")).unwrap_or_default();
            if !bid.is_empty() { c.emails.push(bid.clone()); }
            let in_drafts = mails(w, c, "drafts", "FO1-mover", json!({})).await;
            let in_sub = mails(w, c, &c.sub_path(), "FO1-mover", json!({})).await;
            log.check("FO2", back.is_ok() && in_drafts.as_ref().is_ok_and(|v| v == &vec![bid.clone()]) && in_sub.as_ref().is_ok_and(Vec::is_empty),
                format!("moved back to Drafts: {:?}; in Drafts {:?}; in subfolder {:?}", back.as_ref().err(),
                    in_drafts.map(|v| v.len()), in_sub.map(|v| v.len())));
        }
        Err(e) => { log.fail("FO1", e.clone()); log.fail("FO2", e); }
    }
    let lf = w.call("list_folders", json!({})).await;
    let folders = rows(&lf);
    let fsub = folders.iter().find(|f| s(f, "name") == c.sub);
    log.check("FO3", fsub.is_some_and(|f| s(f, "path") == c.sub_path()),
        format!("subfolder listed: {:?}; Hebrew name exact {}", fsub.map(|f| s(f, "path")), fsub.is_some_and(|f| s(f, "name") == c.sub)));
    let nested_path = format!("{}/{}", c.sub_path(), c.nested);
    let fnested = folders.iter().find(|f| s(f, "name") == c.nested);
    log.check("G2", fnested.is_some_and(|f| s(f, "path") == nested_path),
        format!("nested folder: {:?} (want {nested_path:?}); {:?}", fnested.map(|f| s(f, "path")), lf.as_ref().err()));
    match event(w, c, "FO4-own", "2099-11-20T10:00", "2099-11-20T10:30", json!({})).await {
        Ok(id) => {
            let r = evs(w, c, "FO4-own", "2099-11-20", "2099-11-20", json!({"calendar_of": SELF})).await;
            log.check("FO4", r.as_ref().is_ok_and(|v| ids_of(&json!(v)).contains(&id) || subjects_of(&json!(v)).contains(&c.t("FO4-own"))),
                format!("calendar_of yourself -> {:?}", r.as_ref().map(|v| subjects_of(&json!(v)))));
        }
        Err(e) => log.fail("FO4", e),
    }
    let r = evs(w, c, "FO4-own", "2099-11-20", "2099-11-20", json!({"calendar_of": NOBODY})).await;
    log.check("FO5", r.as_ref().is_err_and(|e| e.contains("nobody@example.invalid")), format!("calendar_of nobody@example.invalid -> {:?}", r.map(|v| v.len())));

    // ---- AE: attendee edge cases ---------------------------------------------------------
    let ae = event(w, c, "AE-meeting", "2099-11-16T10:00", "2099-11-16T10:30", json!({"required_attendees": ["ae-req@example.invalid"]})).await;
    match &ae {
        Ok(id) => {
            let count_x = |v: &Value| -> usize {
                format!("{};{}", s(v, "required_attendees"), s(v, "optional_attendees")).split(';')
                    .filter(|p| p.trim().eq_ignore_ascii_case("x@example.invalid")).count()
            };
            let u1 = w.call("update_event", json!({"event_id": id, "add_required_attendees": ["x@example.invalid", "x@example.invalid"], "notify": false})).await;
            let g1 = w.call("get_event", json!({"event_id": id, "include": []})).await.unwrap_or_default();
            log.check("AE1", u1.is_ok() && count_x(&g1) == 1, format!("add x twice in one call: x listed {} time(s); required {:?}; {:?}",
                count_x(&g1), s(&g1, "required_attendees"), u1.err()));
            let u2 = w.call("update_event", json!({"event_id": id, "add_required_attendees": ["X@Example.Invalid"], "notify": false})).await;
            let g2 = w.call("get_event", json!({"event_id": id, "include": []})).await.unwrap_or_default();
            log.check("AE2", u2.is_ok() && count_x(&g2) == count_x(&g1), format!("add X@Example.Invalid again: x listed {} time(s) (was {}); required {:?}; {:?}",
                count_x(&g2), count_x(&g1), s(&g2, "required_attendees"), u2.err()));
            let u3 = w.call("update_event", json!({"event_id": id, "remove_attendees": ["not-there@example.invalid"], "notify": false})).await;
            let g3 = w.call("get_event", json!({"event_id": id, "include": []})).await.unwrap_or_default();
            let same = s(&g3, "required_attendees") == s(&g2, "required_attendees") && s(&g3, "optional_attendees") == s(&g2, "optional_attendees");
            log.check("AE3", u3.is_ok() && same, format!("remove someone not on the meeting: {} ; attendees unchanged {same}",
                u3.as_ref().map(|v| format!("status {} changed {}", v["status"], v["changed"])).unwrap_or_else(|e| e.clone())));
        }
        Err(e) => { for id in ["AE1", "AE2", "AE3"] { log.fail(id, e.clone()); } }
    }

    // ---- L: size limits ----------------------------------------------------------------
    let big_body = "0123456789".repeat(100_000);
    let body_path = c.root.join("l1-body.txt");
    let _ = std::fs::write(&body_path, &big_body);
    match draft(w, c, "L1-bigbody", json!({"body": Value::Null, "body_file": body_path})).await {
        Ok(id) => {
            let mut l1 = Vec::new();
            let mut l1_ok = true;
            for (n, want_len, want_trunc) in [(999u32, 1000usize, true), (1000, 1000, true), (5_000_001, 0, false)] {
                let r = w.call("get_email", json!({"email_id": id, "include": ["body"], "max_body_chars": n})).await;
                let b = r.as_ref().map(|v| s(v, "body")).unwrap_or_default();
                let len = b.chars().count();
                let trunc = r.as_ref().map(|v| v["body_truncated"] == true).unwrap_or(false);
                let full = r.as_ref().map(|v| v["body_length"].as_u64().unwrap_or(0)).unwrap_or(0);
                // A cut body ends with the marker "\n\n[... truncated at N characters]" (read.rs truncate).
                let marker = format!("\n\n[... truncated at {want_len} characters]");
                let head = b.strip_suffix(&marker).map(|h| h.chars().count());
                let ok = r.is_ok() && trunc == want_trunc && if want_trunc {
                    head == Some(want_len) && big_body.starts_with(b.strip_suffix(&marker).unwrap_or("-"))
                } else { len as u64 == full && b.contains(&big_body) };
                l1_ok &= ok;
                l1.push(format!("max {n}: {len} chars ({head:?} before the truncation marker), truncated {trunc}, body_length {full}{}",
                    r.as_ref().err().map(|e| format!(" ERR {e}")).unwrap_or_default()));
            }
            log.check("L1", l1_ok, format!("1 MB body: {}", l1.join("; ")));
        }
        Err(e) => log.fail("L1", e),
    }
    let base = c.t("L2-long ");
    let long_subj: String = format!("{base}{}", "x".repeat(300 - base.chars().count()));
    match w.call("create_draft", json!({"to": [NOBODY], "subject": long_subj, "body": "x"})).await {
        Ok(v) => {
            let id = s(&v, "id");
            c.emails.push(id.clone());
            let g = w.call("get_email", json!({"email_id": id, "include": []})).await.unwrap_or_default();
            // Outlook keeps the first 255 characters of a subject (its own limit).
            let want: String = long_subj.chars().take(255).collect();
            log.check("L2", s(&g, "subject") == want, format!("300-char subject reads back as its first 255 characters (Outlook's limit): {} ({} chars)", s(&g, "subject") == want, s(&g, "subject").chars().count()));
        }
        Err(e) => log.fail("L2", e),
    }
    let big = c.root.join("big12.png");
    let mut img = vec![0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A];
    img.extend((0..12 * 1024 * 1024).map(|i: u32| (i % 251) as u8));
    let _ = std::fs::write(&big, &img);
    match draft(w, c, "L3-bigimage", json!({"body": Value::Null, "html_body": "<p>big</p><img src=\"cid:big12\">",
        "inline_images": [{"content_id": "big12", "path": big}]})).await {
        Ok(id) => {
            let r = w.call("get_inline_image", json!({"email_id": id, "content_id": "big12"})).await;
            log.check("L3", r.as_ref().is_err_and(|e| e.contains("10 MB")),
                format!("12 MB inline image -> {}", match &r { Ok(v) => format!("returned {} bytes", v["size"]), Err(e) => format!("error {e:?}") }));
        }
        Err(e) => log.fail("L3", e),
    }
    let _ = std::fs::remove_file(&big);
    if remaining.is_empty() {
        log.fail("L4", "no PG drafts left to read");
    } else {
        let ids200: Vec<String> = remaining.iter().cycle().take(200).cloned().collect();
        let mut ids201 = ids200.clone();
        ids201.push(remaining[0].clone());
        let r200 = w.call("get_email", json!({"email_id": ids200, "include": []})).await;
        let r201 = w.call("get_email", json!({"email_id": ids201, "include": []})).await;
        let ok200 = rows(&r200).iter().filter(|x| x.get("error").is_none()).count();
        log.check("L4", ok200 == 200 && rows(&r200).len() == 200,
            format!("200 ids ({} distinct drafts, repeated) -> {ok200} results; 201 ids -> {} (not applicable as planned: get_email documents no batch limit, \
                unlike the write tools' 100)", remaining.len(), match &r201 { Ok(v) => format!("{} results", v.as_array().map_or(0, Vec::len)), Err(e) => format!("error {e:?}") }));
    }

    // ---- self-loop: FL, G5, G3 -------------------------------------------------------------
    let mut received: Option<Value> = None;
    while sent_at.elapsed().as_secs() < 300 {
        let r = w.call("list_emails", json!({"query": c.q("SELFLOOP-mail"), "count": 5})).await;
        if let Some(x) = rows(&r).into_iter().find(|x| s(x, "subject") == selfloop) { received = Some(x); break; }
        tokio::time::sleep(Duration::from_secs(5)).await;
    }
    println!("self-loop arrived: {} after {:.0}s", received.is_some(), sent_at.elapsed().as_secs_f64());
    match received.clone() {
        Some(row) => {
            let mut rid = s(&row, "id");
            let mut fl1 = Vec::new();
            let mut fl1_ok = true;
            for (f, want) in [("complete", [false, true, false]), ("clear", [false, false, true]), ("follow_up", [true, false, false])] {
                let u = w.call("update_email", json!({"email_id": rid, "flag": f})).await;
                if let Ok(v) = &u && !s(v, "id").is_empty() { rid = s(v, "id"); }
                let st = flag_state(w, c, &rid).await;
                fl1_ok &= u.is_ok() && st == want;
                fl1.push(format!("{f}: found by (follow_up, complete, clear) {st:?}{}", u.err().map(|e| format!(" ERR {e}")).unwrap_or_default()));
            }
            log.check("FL1", fl1_ok, fl1.join("; "));
            let u = w.call("update_email", json!({"email_id": rid, "mark_read": false})).await;
            let a = rows(&w.call("list_emails", json!({"query": c.q("SELFLOOP-mail"), "unread_only": true})).await).iter().any(|x| s(x, "id") == rid);
            let u2 = w.call("update_email", json!({"email_id": rid, "mark_read": true})).await;
            let b = rows(&w.call("list_emails", json!({"query": c.q("SELFLOOP-mail"), "unread_only": true})).await).iter().any(|x| s(x, "id") == rid);
            log.check("FL2", u.is_ok() && u2.is_ok() && a && !b, format!("unread_only finds it after mark unread {a}, after mark read {b}"));
            let recv = NaiveDateTime::parse_from_str(&s(&row, "received"), "%Y-%m-%dT%H:%M:%S").ok();
            match recv {
                Some(t0) => {
                    let fmt = |t: NaiveDateTime| t.format("%Y-%m-%dT%H:%M").to_string();
                    let plus = fmt(t0 + chrono::Duration::minutes(1));
                    let minus = fmt(t0 - chrono::Duration::minutes(1));
                    let inc = rows(&w.call("list_emails", json!({"query": c.q("SELFLOOP-mail"), "received_before": plus})).await).iter().any(|x| s(x, "id") == rid);
                    let exc = rows(&w.call("list_emails", json!({"query": c.q("SELFLOOP-mail"), "received_before": minus})).await).iter().any(|x| s(x, "id") == rid);
                    log.check("G5", inc && !exc, format!("received {}: received_before {plus} includes it {inc}; {minus} includes it {exc}", iso(t0)));
                }
                None => log.fail("G5", format!("unparsable received {:?}", s(&row, "received"))),
            }
            // G3: reply-all, saved first so its recipients are checked before anything is sent.
            let me_name = s(&row, "sender");
            let rd = w.call("reply_email", json!({"email_id": rid, "reply_all": true, "send": false, "body": "edge reply-all"})).await;
            let rdid = rd.as_ref().map(|v| s(v, "id")).unwrap_or_default();
            if !rdid.is_empty() { c.emails.push(rdid.clone()); }
            let g = w.call("get_email", json!({"email_id": rdid, "include": []})).await.unwrap_or_default();
            let recips: Vec<String> = format!("{};{};{}", s(&g, "to"), s(&g, "cc"), s(&g, "bcc")).split(';')
                .map(|x| x.trim().to_string()).filter(|x| !x.is_empty()).collect();
            let is_me = |x: &str| { let x = x.trim().trim_matches(|ch| ch == '\'' || ch == '"'); x.eq_ignore_ascii_case(SELF) || x == me_name };
            let only_self = !recips.is_empty() && recips.iter().all(|x| is_me(x));
            let mut g3 = format!("reply-all draft recipients {recips:?} (you are {me_name:?} <{SELF}>)");
            let mut g3_ok = false;
            if only_self {
                let sent = w.call("update_draft", json!({"email_id": rdid, "send": true})).await;
                if sent.is_ok() { c.inbound += 1; }
                let started = Instant::now();
                let mut in_sent = Vec::new();
                while started.elapsed().as_secs() < 120 {
                    let r = w.call("list_emails", json!({"folder": "sent", "query": c.q("SELFLOOP-mail"), "count": 10})).await;
                    in_sent = rows(&r).into_iter().filter(|x| s(x, "subject").to_lowercase().starts_with("re:")).collect();
                    if !in_sent.is_empty() { break; }
                    tokio::time::sleep(Duration::from_secs(5)).await;
                }
                let sent_to: Vec<String> = in_sent.iter().map(|x| s(x, "to")).collect();
                g3_ok = sent.is_ok() && in_sent.len() == 1 && sent_to.iter().all(|t| t.split(';').all(|x| is_me(x)));
                g3 = format!("{g3}; sent {:?}; in Sent Items {} reply, to {sent_to:?}", sent.as_ref().map(|v| s(v, "status")), in_sent.len());
            } else {
                g3 = format!("{g3}; NOT sent: recipients are not only you");
            }
            log.check("G3", g3_ok, g3);
            // FL3: the reply (newer, unflagged, matching the query) must not use up count 1.
            let started = Instant::now();
            let mut newest = String::new();
            while c.inbound >= 2 && started.elapsed().as_secs() < 180 {
                let r = w.call("list_emails", json!({"query": c.q("SELFLOOP-mail"), "count": 1})).await;
                newest = rows(&r).first().map(|x| s(x, "subject")).unwrap_or_default();
                if newest.to_lowercase().starts_with("re:") { break; }
                tokio::time::sleep(Duration::from_secs(5)).await;
            }
            let r = w.call("list_emails", json!({"query": c.q("SELFLOOP-mail"), "flag": "follow_up", "count": 1})).await;
            let got = ids_of(&r.clone().unwrap_or_default());
            log.check("FL3", newest.to_lowercase().starts_with("re:") && got == vec![rid.clone()],
                format!("newest match is the unflagged reply ({newest:?}); flag follow_up, count 1 -> {} row(s), the flagged original {}",
                    got.len(), got == vec![rid.clone()]));
        }
        None => { for id in ["FL1", "FL2", "G5", "G3", "FL3"] { log.fail(id, "the self-loop email never arrived in the Inbox"); } }
    }

    // ---- G: gaps in earlier suites -----------------------------------------------------------
    let plain = draft(w, c, "G1-plain", json!({})).await;
    match (&ae, &plain) {
        (Ok(m), Ok(p)) => {
            let a = w.call("get_email", json!({"email_id": m, "include": ["meeting"]})).await;
            let b = w.call("get_email", json!({"email_id": p, "include": ["meeting"]})).await;
            let am = a.as_ref().map(|v| (v["item_type"].clone(), v["is_meeting"].clone(), v.get("meeting").cloned())).ok();
            let bm = b.as_ref().map(|v| (v["is_meeting"].clone(), v.get("meeting").is_none_or(Value::is_null))).ok();
            log.check("G1a", a.is_ok(), format!("not applicable as planned: an unsent meeting has no meeting-request message (only sending creates one, \
                which this suite never does), and get_email's is_meeting/meeting describe a meeting-request message. On the meeting's calendar \
                item it returns (item_type, is_meeting, meeting) {am:?}; {:?}", a.as_ref().err()));
            log.check("G1b", bm == Some((json!(false), true)), format!("plain draft with include [\"meeting\"]: (is_meeting, no meeting) {bm:?}"));
        }
        _ => { log.fail("G1a", format!("{ae:?} {plain:?}")); log.fail("G1b", "no draft"); }
    }
    match event(w, c, "G4-reminder", "2099-02-12T10:00", "2099-02-12T10:30", json!({"reminder_minutes": 15})).await {
        Ok(id) => {
            let (eid, sid) = eid_sid(&id);
            let probe = format!("$i=$ns.GetItemFromID({}, {}); [string]$i.ReminderSet + '|' + $i.ReminderMinutesBeforeStart", psq(&eid), psq(&sid));
            let a = ps(probe.clone()).await;
            let g = w.call("get_event", json!({"event_id": id, "include": []})).await.unwrap_or_default();
            let has_field = g.as_object().is_some_and(|o| o.keys().any(|k| k.contains("reminder")));
            let u = w.call("update_event", json!({"event_id": id, "reminder_minutes": -1, "notify": false})).await;
            let b = ps(probe).await;
            log.check("G4a", a.as_ref().is_ok_and(|x| x == "True|15"),
                format!("reminder_minutes 15 -> raw COM (ReminderSet|minutes) {a:?}; get_event has a reminder field: {has_field} (it documents none)"));
            log.check("G4b", u.is_ok() && b.as_ref().is_ok_and(|x| x.starts_with("False")),
                format!("update reminder_minutes -1 -> raw COM {b:?} (want no reminder); {:?}", u.err()));
        }
        Err(e) => { log.fail("G4a", e.clone()); log.fail("G4b", e); }
    }
    log.check("G6", true, "not applicable: list_events has no importance filter and its docs don't claim one (only list_emails and list_tasks filter by importance)");
    match event(w, c, "G7-roles", "2099-02-14T10:00", "2099-02-14T10:30",
        json!({"required_attendees": ["g7-req@example.invalid"], "optional_attendees": ["g7-opt@example.invalid"]})).await {
        Ok(id) => {
            let mut g7 = Vec::new();
            let mut g7_ok = true;
            for (who, role, want) in [("g7-opt@example.invalid", "optional", true), ("g7-req@example.invalid", "optional", false),
                                      ("g7-req@example.invalid", "required", true), ("g7-opt@example.invalid", "required", false)] {
                let r = evs(w, c, "G7-roles", "2099-02-14", "2099-02-14", json!({"attendees": who, "attendee_role": role})).await;
                let hit = r.as_ref().is_ok_and(|v| v.iter().any(|x| s(x, "id") == id));
                g7_ok &= hit == want && r.is_ok();
                g7.push(format!("{who} as {role}: {hit}"));
            }
            log.check("G7", g7_ok, g7.join("; "));
        }
        Err(e) => log.fail("G7", e),
    }
    match draft(w, c, "G8-archive", json!({})).await {
        Ok(id) => {
            let mv = w.call("update_email", json!({"email_id": id, "move_to": "Archive"})).await;
            let nid = mv.as_ref().map(|v| s(v, "id")).unwrap_or_default();
            if !nid.is_empty() { c.emails.push(nid.clone()); }
            let there = mails(w, c, "Archive", "G8-archive", json!({})).await;
            log.check("G8", mv.is_ok() && there.as_ref().is_ok_and(|v| v == &vec![nid.clone()]),
                format!("move_to Archive {:?}; list_emails(folder Archive) -> {:?}", mv.as_ref().err(), there.map(|v| v.len())));
        }
        Err(e) => log.fail("G8", e),
    }
    let outer = draft(w, c, "G9-outer", json!({})).await;
    let inner = draft(w, c, "G9-inner", json!({})).await;
    match (&outer, &inner) {
        (Ok(o), Ok(i)) => {
            let (oe, os) = eid_sid(o);
            let (ie, is) = eid_sid(i);
            let r = ps(format!("$o2=$ns.GetItemFromID({}, {}); $in=$ns.GetItemFromID({}, {}); $null=$o2.Attachments.Add($in, 5); $o2.Save(); 'ok ' + $o2.Attachments.Count",
                psq(&oe), psq(&os), psq(&ie), psq(&is))).await;
            let la = w.call("list_attachments", json!({"email_id": o})).await;
            let ge = w.call("get_email", json!({"email_id": o, "include": ["attachments"]})).await;
            let sv = w.call("save_attachments", json!({"email_id": o, "save_dir": c.root.join("save-g9")})).await;
            let item = rows(&la).first().cloned().unwrap_or_default();
            let sres = rows(&sv).first().cloned().unwrap_or_default();
            let saved_ok = (s(&sres, "status") == "saved" && Path::new(&s(&sres, "saved_to")).is_file())
                || (s(&sres, "status") == "failed" && !s(&sres, "error").is_empty());
            log.check("G9", r.is_ok() && s(&item, "type") == "item" && ge.as_ref().is_ok_and(|v| v["attachments"].as_array().is_some_and(|a| a.len() == 1)) && saved_ok,
                format!("embedded via COM {r:?}; list_attachments type {:?} name {:?}; get_email attachments {:?}; save_attachments status {:?} {:?}",
                    s(&item, "type"), s(&item, "filename"), ge.as_ref().map(|v| v["attachments"].clone()).ok(), s(&sres, "status"),
                    if s(&sres, "status") == "saved" { short_path(&s(&sres, "saved_to")) } else { s(&sres, "error") }));
        }
        _ => log.fail("G9", format!("{outer:?} {inner:?}")),
    }

    // ---- ID: stale and malformed ids --------------------------------------------------------
    let a = w.call("get_email", json!({"email_id": ""})).await;
    let b = w.call("update_email", json!({"email_id": "", "mark_read": true})).await;
    let d = w.call("delete_email", json!({"email_id": ""})).await;
    log.check("ID1", a.is_err() && b.is_err() && d.is_err(), format!("empty id: get {:?}; update {:?}; delete {:?}", a.err(), b.err(), d.err()));
    if stale_mail.is_empty() {
        log.fail("ID2", "no permanently deleted draft id (PD2 failed)");
    } else {
        let a = w.call("get_email", json!({"email_id": stale_mail})).await;
        let b = w.call("update_email", json!({"email_id": stale_mail, "mark_read": true})).await;
        let d = w.call("delete_email", json!({"email_id": stale_mail})).await;
        let e = w.call("get_event", json!({"event_id": stale_event})).await;
        log.check("ID2", a.is_err() && b.is_err() && d.is_err() && e.is_err(),
            format!("hard-deleted draft: get {}; update {}; delete {}; hard-deleted event via get_event {}",
                a.is_err(), b.is_err(), d.is_err(), e.is_err()));
    }
    match (&tk, &ev) {
        (Ok(t), Ok(e)) => {
            let a = w.call("get_event", json!({"event_id": t, "include": []})).await;
            let b = w.call("get_task", json!({"task_id": e, "include": []})).await;
            log.check("ID3", a.is_err() && b.is_err(), format!("task id to get_event -> {}; event id to get_task -> {}",
                match &a { Ok(v) => format!("returned {:?} (start {:?})", s(v, "subject"), v["start"]), Err(e) => format!("error {e:?}") },
                match &b { Ok(v) => format!("returned {:?} (status {:?})", s(v, "subject"), v["status"]), Err(e) => format!("error {e:?}") }));
        }
        _ => log.fail("ID3", "CT2's task or event is missing"),
    }
    match event(w, c, "ID4-series", "2099-11-26T09:00", "2099-11-26T09:30", json!({"recurrence": {"pattern": "daily", "occurrences": 2}})).await {
        Ok(id) => {
            let a = w.call("get_event", json!({"event_id": format!("{id}|2099-11-29T09:00:00")})).await;
            let b = w.call("get_event", json!({"event_id": format!("{id}|2099-11-26T09:30:00")})).await;
            log.check("ID4", a.is_err() && b.is_err(), format!("a date after the series ends -> {:?}; a time off the pattern -> {:?}",
                a.as_ref().err().map(|e| e.chars().take(120).collect::<String>()), b.as_ref().err().map(|e| e.chars().take(120).collect::<String>())));
        }
        Err(e) => log.fail("ID4", e),
    }
    match &plain {
        Ok(p) => {
            let r = w.call("respond_to_meeting", json!({"event_id": p, "response": "accept", "notify": false})).await;
            let clear = r.as_ref().is_err_and(|e| { let e = e.to_lowercase(); e.contains("meeting") || e.contains("appointment") || e.contains("not a") });
            log.check("ID5", clear, format!("a draft's id -> {r:?} (want an error that says the id is not a meeting)"));
        }
        Err(e) => log.fail("ID5", e.clone()),
    }
}

async fn avail(w: &mut Wire, people: &[&str], a: String, b: String, iv: i64) -> Result<Value, String> {
    w.call("check_availability", json!({"people": people, "start": a, "end": b, "interval_minutes": iv})).await
}

async fn page(w: &mut Wire, c: &Ctx, off: i64, n: i64) -> Result<Vec<String>, String> {
    w.call("list_emails", json!({"folder": "drafts", "query": c.q("PGmail"), "count": n, "offset": off})).await.map(|v| ids_of(&v))
}

async fn epage(w: &mut Wire, c: &Ctx, off: i64, n: i64) -> Result<Vec<String>, String> {
    w.call("list_events", json!({"query": c.q("PGevent"), "start_after": "2099-05-01", "start_before": "2099-05-31",
        "count": n, "offset": off})).await.map(|v| ids_of(&v))
}

/// Whether the self-loop email `id` is found by flag follow_up, complete, clear.
async fn flag_state(w: &mut Wire, c: &Ctx, id: &str) -> Vec<bool> {
    let mut out = Vec::new();
    for f in ["follow_up", "complete", "clear"] {
        let r = w.call("list_emails", json!({"query": c.q("SELFLOOP-mail"), "flag": f})).await;
        out.push(rows(&r).iter().any(|x| s(x, "id") == id));
    }
    out
}

fn short_path(p: &str) -> String {
    Path::new(p).file_name().map(|f| f.to_string_lossy().into_owned()).unwrap_or_else(|| p.to_string())
}

/// Deletes everything the run created: tracked ids first (permanently), then
/// a tag sweep of every touched folder until a pass finds nothing, the test
/// folders, and the local directory. Ground truth is a raw COM count by the
/// run's tag at the end.
async fn cleanup(w: &mut Wire, log: &mut Log, c: &Ctx) {
    let mut errors = Vec::new();
    // Wait (up to 2 min) for the inbound copies of what was sent, so the Inbox sweep sees them.
    let started = Instant::now();
    loop {
        let r = w.call("list_emails", json!({"query": c.q("SELFLOOP-mail"), "count": 10})).await;
        let n = rows(&r).len();
        if n >= c.inbound || started.elapsed().as_secs() > 120 {
            println!("inbound copies in the Inbox: {n} of {} after {:.0}s", c.inbound, started.elapsed().as_secs_f64());
            break;
        }
        tokio::time::sleep(Duration::from_secs(5)).await;
    }
    for chunk in c.emails.chunks(100) {
        let _ = w.call("delete_email", json!({"email_id": chunk, "permanent": true})).await;
    }
    for chunk in c.events.chunks(100) {
        let _ = w.call("delete_event", json!({"event_id": chunk, "permanent": true, "notify": false})).await;
    }
    for id in &c.tasks {
        let _ = w.call("delete_task", json!({"task_id": id, "permanent": true})).await;
    }
    let mut folders: Vec<String> = ["drafts", "inbox", "sent", "deleted", "outbox", "Archive"].iter().map(|f| f.to_string()).collect();
    folders.push(c.sub_path());
    folders.push(format!("{}/{}", c.sub_path(), c.nested));
    // Only rows whose subject carries this run's exact tag are touched.
    let tag = format!("{TAG} {}", c.run);
    let ours = |rows: &[Value]| -> Vec<String> {
        rows.iter().filter(|x| s(x, "subject").contains(&tag)).map(|x| s(x, "id")).collect()
    };
    let mut swept = 0;
    for pass in 0..8 {
        let mut found = 0;
        for f in &folders {
            let ids = ours(&rows(&w.call("list_emails", json!({"folder": f, "query": c.run, "count": 200})).await));
            for chunk in ids.chunks(100) {
                found += chunk.len();
                if let Err(e) = w.call("delete_email", json!({"email_id": chunk, "permanent": true})).await { errors.push(format!("{f}: {e}")); }
            }
        }
        let mut series = BTreeSet::new();
        for (a, b) in [("2096-01-01", "2097-12-31"), ("2098-01-01", "2099-06-30"), ("2099-07-01", "2100-12-31")] {
            let mut off = 0;
            loop {
                let r = rows(&w.call("list_events", json!({"query": c.run, "start_after": a, "start_before": b, "count": 250, "offset": off, "fields": ["id", "subject"]})).await);
                series.extend(ours(&r));
                if r.len() < 250 { break; }
                off += 250;
            }
        }
        let series: Vec<String> = series.into_iter().collect();
        for chunk in series.chunks(100) {
            found += chunk.len();
            if let Err(e) = w.call("delete_event", json!({"event_id": chunk, "permanent": true, "notify": false})).await { errors.push(e); }
        }
        let tasks = ours(&rows(&w.call("list_tasks", json!({"query": c.run, "include_completed": true})).await));
        let notes = ours(&rows(&w.call("list_notes", json!({"query": c.run})).await));
        for id in &tasks { found += 1; if let Err(e) = w.call("delete_task", json!({"task_id": id, "permanent": true})).await { errors.push(e); } }
        for id in &notes { found += 1; if let Err(e) = w.call("delete_note", json!({"note_id": id, "permanent": true})).await { errors.push(e); } }
        println!("cleanup sweep pass {pass}: {found} found");
        swept += found;
        if found == 0 { break; }
    }
    // Raw COM fallback for anything tagged the tools couldn't reach or delete (e.g. a failed
    // create stranded in Sync Issues/Local Failures, whose Delete() fails): Delete(), else
    // Items.Remove(index); then purge Deleted Items. Exact tag + run token only.
    let tag = format!("{TAG} {}", c.run);
    let raw_sweep = ps(format!(
        "$tag={}; $del=$ns.GetDefaultFolder(3); $out=@(); \
         function Sweep($f) {{ if ($f.EntryID -ne $del.EntryID) {{ $items=$f.Items; for ($i=$items.Count; $i -ge 1; $i--) {{ $it=$items.Item($i); \
           if (([string]$it.Subject).Contains($tag)) {{ try {{ $it.Delete(); $script:out += ('deleted in ' + $f.Name) }} \
           catch {{ try {{ $items.Remove($i); $script:out += ('removed in ' + $f.Name) }} catch {{ $script:out += ('STUCK in ' + $f.Name) }} }} }} }} }}; \
           foreach ($s in $f.Folders) {{ Sweep $s }} }}; \
         foreach ($f in $ns.GetDefaultFolder(6).Parent.Folders) {{ Sweep $f }}; \
         for ($k=0; $k -lt 5; $k++) {{ $items=$del.Items; $n=0; for ($i=$items.Count; $i -ge 1; $i--) {{ $it=$items.Item($i); \
           if (([string]$it.Subject).Contains($tag)) {{ $it.Delete(); $n++ }} }}; if ($n -eq 0) {{ break }}; $script:out += ('purged ' + $n) }}; \
         'raw sweep: ' + ($out -join ', ')",
        psq(&tag))).await;
    println!("{raw_sweep:?}");
    if raw_sweep.as_ref().is_ok_and(|x| x.contains("STUCK")) || raw_sweep.is_err() { errors.push(format!("raw sweep {raw_sweep:?}")); }
    // The folders: Folder.Delete() moves a folder to Deleted Items; delete it there too.
    let rm = ps(format!(
        "$tag={}; $n=0; for ($k=0; $k -lt 6; $k++) {{ $hit=$null; \
           foreach ($p in @($ns.GetDefaultFolder(6), $ns.GetDefaultFolder(3))) {{ foreach ($f in $p.Folders) {{ if ($f.Name.Contains($tag)) {{ $hit=$f }} }} }}; \
           if (-not $hit) {{ break }}; $hit.Delete(); $n++ }}; 'removed ' + $n",
        psq(&tag))).await;
    println!("folder removal: {rm:?}");
    let _ = std::fs::remove_dir_all(&c.root);
    let local_gone = !c.root.exists();
    // Ground truth: raw COM count of anything carrying this run's tag, in every folder.
    tokio::time::sleep(Duration::from_secs(3)).await;
    let raw = ps(format!(
        "$tag={}; $script:n=0; $script:where=@(); \
         function Walk($f, $p) {{ try {{ foreach ($it in $f.Items) {{ $s=''; try {{ $s=[string]$it.Subject }} catch {{}}; $b=''; \
           if ($it.Class -eq 44) {{ try {{ $b=[string]$it.Body }} catch {{}} }}; \
           if ($s.Contains($tag) -or $b.Contains($tag)) {{ $script:n++; $script:where += $p }} }} }} catch {{}}; \
           if ($f.Name.Contains($tag)) {{ $script:n++; $script:where += ('folder ' + $p) }}; \
           foreach ($sub in $f.Folders) {{ Walk $sub ($p + '/' + $sub.Name) }} }}; \
         foreach ($f in $ns.GetDefaultFolder(6).Parent.Folders) {{ Walk $f $f.Name }}; [string]$script:n + ' ' + (($script:where | Select-Object -Unique) -join ', ')",
        psq(&tag))).await;
    let raw_zero = raw.as_ref().is_ok_and(|x| x.trim() == "0");
    log.check("CLEAN", errors.is_empty() && raw_zero && local_gone && rm.is_ok(),
        format!("{} emails, {} events, {} tasks deleted by id; {swept} more swept; folders {rm:?}; local dir removed {local_gone}; \
            raw sweep {raw_sweep:?}; raw COM count of the run's tag in every folder: {raw:?}; errors {:?}",
            c.emails.len(), c.events.len(), c.tasks.len(), errors.iter().take(3).collect::<Vec<_>>()));
}
