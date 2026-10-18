//! Live system test with real external mail and meeting invites (2026-10-17).
//! Plan: `SYSTEM_TEST_PLAN_2026-10-17-INBOUND.md`.
//!
//! Needs a connector step first: an agent (or a person) sends three emails
//! from the external Gmail account and creates eight Google Calendar invites
//! to this mailbox, all tagged with one run token, then runs:
//!
//!     SYSTEST_INBOUND_RUN=r<token> cargo test --test system_test_inbound -- --ignored --nocapture
//!
//! The test waits for everything to arrive, checks reading, Hebrew, inline
//! images, Hebrew attachment names, batch flags/categories with undo, real
//! meeting responses (notify true and false, one occurrence, a batch, via the
//! meeting request), roles and require_no_attendees on real invites, sends
//! three Hebrew emails back to Gmail, and prints what the connector step must
//! verify on the Gmail / Google Calendar side ("VERIFY ..." lines). Then it
//! deletes everything it touched on the Outlook side. Never panics before
//! cleanup.

mod common;

use std::time::{Duration, Instant};

use base64::Engine;
use common::{ids_of, s, subjects_of, temp_dir, Log, Wire};
use serde_json::{json, Value};

const TAG: &str = "[outlook-mcp-rs systest inbound]";
const GMAIL: &str = "adamkopelman2@gmail.com";
const SELF: &str = "adamkopelman@outlook.com";
const PNG_B64: &str = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";
/// A Hebrew word that the connector step puts only in G1's body.
const HE_BODY_WORD: &str = "תפוזכתום";
const HE_ATTACHMENT: &str = "מסמך.txt";
const HE_ATTACHMENT_TEXT: &str = "שלום מקובץ מצורף - בדיקה";
const HE_CATEGORY: &str = "דחוף";

fn rows(v: &Result<Value, String>) -> Vec<Value> {
    v.as_ref().ok().and_then(|v| v.as_array().cloned()).unwrap_or_default()
}

#[tokio::test(flavor = "multi_thread")]
#[ignore]
async fn system_test_inbound() {
    let mut log = Log::default();
    let Ok(run) = std::env::var("SYSTEST_INBOUND_RUN") else {
        log.fail("SETUP", "SYSTEST_INBOUND_RUN is not set: run the connector step first (see the plan)");
        assert!(log.summary());
        return;
    };
    let mut w = common::connect("systest-inbound").await;
    println!("run token {run}");
    body(&mut w, &mut log, &run).await;
    println!("\n---- cleanup");
    cleanup(&mut w, &mut log, &run).await;
    println!("\n---- summary");
    let ok = log.summary();
    assert!(ok, "one or more steps failed");
}

/// The inbox email whose subject contains `label` (and the run token).
fn find<'a>(mail: &'a [Value], label: &str) -> Option<&'a Value> {
    mail.iter().find(|e| s(e, "subject").contains(label))
}

async fn body(w: &mut Wire, log: &mut Log, run: &str) {
    let t = |label: &str| format!("{TAG} {run} {label}");
    let q = format!("\"{run}\"");
    let june = json!({"query": q, "start_after": "2099-06-01", "start_before": "2099-07-31"});

    // ---- phase 0: wait for the connector step's mail and invites --------------
    let started = Instant::now();
    let (mut mail, mut events) = (Vec::new(), Vec::new());
    while started.elapsed() < Duration::from_secs(300) {
        mail = rows(&w.call("list_emails", json!({"query": q, "count": 50, "from": GMAIL})).await);
        events = rows(&w.call("list_events", june.clone()).await);
        let have = |l: &str| find(&mail, l).is_some();
        let have_ev = |l: &str| events.iter().any(|e| s(e, "subject").contains(l));
        if ["G1", "G2", "G3"].iter().all(|l| have(l)) && ["M1", "M2", "M3", "M4a", "M4b", "M5", "M6", "M7"].iter().all(|l| have_ev(l)) {
            break;
        }
        tokio::time::sleep(Duration::from_secs(10)).await;
    }
    let ev = |label: &str, events: &[Value]| events.iter().find(|e| s(e, "subject").contains(label)).cloned();
    let labels: Vec<String> = events.iter().map(|e| s(e, "subject").replace(&format!("{TAG} {run} "), "")).collect();
    if !log.check("ARRIVE", find(&mail, "G1").is_some() && find(&mail, "G2").is_some() && find(&mail, "G3").is_some() && events.len() >= 13,
        format!("after {:.0}s: {} emails from Gmail, {} invite rows {labels:?}", started.elapsed().as_secs_f64(), mail.len(), events.len())) {
        return;
    }
    let g1 = find(&mail, "G1").cloned().unwrap_or_default();
    let g2 = find(&mail, "G2").cloned().unwrap_or_default();
    let g3 = find(&mail, "G3").cloned().unwrap_or_default();

    // ---- reading real external mail, in Hebrew (G) ---------------------------------
    let d1 = w.call("get_email", json!({"email_id": s(&g1, "id")})).await.unwrap_or_default();
    log.check("G1", s(&g1, "subject") == t("G1 שלום מג'ימייל") && s(&g1, "sender_email").eq_ignore_ascii_case(GMAIL)
        && s(&d1, "item_type") == "email" && s(&d1, "body").contains(HE_BODY_WORD) && s(&d1, "body").contains("גוף ההודעה בעברית"),
        format!("subject {:?}, sender {:?}, item_type {:?}, Hebrew body {:?}", s(&g1, "subject"), s(&g1, "sender_email"), s(&d1, "item_type"),
            s(&d1, "body").chars().take(70).collect::<String>()));
    let r = w.call("list_emails", json!({"query": format!("{q} {HE_BODY_WORD}"), "from": GMAIL})).await;
    log.check("G2", ids_of(&r.clone().unwrap_or_default()) == vec![s(&g1, "id")],
        format!("Hebrew word only in the body -> {:?}", r.as_ref().map(subjects_of)));
    // Gmail gives the image its own Content-ID ("ii_...") and strips the
    // <img src="cid:pixel.png"> that doesn't match it from the HTML it sends
    // (its sent copy's HTML is just the paragraph). So the image is correctly
    // not inline, nothing resolves or stays unresolved, and the bytes still
    // come back by Content-ID.
    let atts = rows(&w.call("list_attachments", json!({"email_id": s(&g2, "id")})).await);
    match atts.iter().find(|a| !s(a, "content_id").is_empty()).cloned() {
        Some(a) => {
            let got = w.call("get_inline_image", json!({"email_id": s(&g2, "id"), "content_id": s(&a, "content_id")})).await.unwrap_or_default();
            let bytes = s(&got, "data_uri").split_once(',').and_then(|(_, b)| base64::engine::general_purpose::STANDARD.decode(b).ok());
            let want = base64::engine::general_purpose::STANDARD.decode(PNG_B64).unwrap();
            let resolved = w.call("get_email", json!({"email_id": s(&g2, "id"), "resolve_inline_images": true})).await.unwrap_or_default();
            let unresolved = resolved["inline_images_unresolved"].clone();
            let html = s(&resolved, "html_body");
            log.check("G3", bytes.as_deref() == Some(&want[..]) && a["is_inline"] == false && !html.contains("cid:")
                && resolved["inline_images_resolved"].as_u64().unwrap_or(0) == 0 && unresolved.as_array().is_none_or(|u| u.is_empty()),
                format!("Gmail image cid {:?}: bytes equal {}; is_inline {}; HTML has a cid: reference {}; resolved {}, unresolved {unresolved}",
                    s(&a, "content_id"), bytes.as_deref() == Some(&want[..]), a["is_inline"], html.contains("cid:"), resolved["inline_images_resolved"]));
        }
        None => log.fail("G3", format!("no attachment with a Content-ID among {atts:?}")),
    }
    let atts = rows(&w.call("list_attachments", json!({"email_id": s(&g3, "id")})).await);
    let names: Vec<String> = atts.iter().map(|a| s(a, "filename")).collect();
    let dir = temp_dir("inbound-attachments");
    let saved = w.call("save_attachments", json!({"email_id": s(&g3, "id"), "save_dir": dir})).await;
    let bytes_of = |name: &str| std::fs::read(dir.join(name)).unwrap_or_default();
    let exact = |name: &str| bytes_of(name) == HE_ATTACHMENT_TEXT.as_bytes();
    // The connector step sends the same Hebrew text three ways; the names
    // must survive, and the octet-stream copy proves the save is byte-exact.
    log.check("G4", names.iter().any(|n| n == HE_ATTACHMENT) && names.iter().any(|n| n == "notes.txt") && saved.is_ok()
        && exact("מסמך-utf8.txt") && exact("מסמך.bin") && bytes_of("notes.txt") == b"plain notes",
        format!("attachment names {names:?}; saved {:?}; bytes exact: text/plain {}, text/plain;charset=UTF-8 {}, octet-stream {} (text/plain bytes {:?})",
            saved.as_ref().err(), exact(HE_ATTACHMENT), exact("מסמך-utf8.txt"), exact("מסמך.bin"),
            String::from_utf8_lossy(&bytes_of(HE_ATTACHMENT))));
    let ids = vec![s(&g1, "id"), s(&g2, "id"), s(&g3, "id")];
    match w.call("update_email", json!({"email_id": ids, "mark_read": true, "flag": "follow_up", "add_categories": [HE_CATEGORY]})).await {
        Ok(v) => {
            let res = v.as_array().cloned().unwrap_or_default();
            let undos: Vec<Value> = res.iter().map(|x| x["undo"].clone()).collect();
            let tagged = rows(&w.call("list_emails", json!({"query": q, "from": GMAIL, "category": HE_CATEGORY, "flag": "follow_up"})).await).len();
            let back = w.call("update_email", json!({"updates": undos})).await;
            let after = rows(&w.call("list_emails", json!({"query": q, "from": GMAIL, "category": HE_CATEGORY})).await).len();
            let flagged = rows(&w.call("list_emails", json!({"query": q, "from": GMAIL, "flag": "follow_up"})).await).len();
            log.check("G6", res.len() == 3 && tagged == 3 && back.is_ok() && after == 0 && flagged == 0,
                format!("3 received emails: read + flag + Hebrew category \"{HE_CATEGORY}\" in one call -> {tagged} match; undo -> category {after}, flagged {flagged}; {:?}", back.err()));
        }
        Err(e) => log.fail("G6", e),
    }
    let recent = ids_of(&w.call("list_emails", json!({"query": q, "from": GMAIL, "received_after": "-6h", "to": SELF, "item_type": "email"})).await.unwrap_or_default());
    let old = ids_of(&w.call("list_emails", json!({"query": q, "from": GMAIL, "received_before": "-6h"})).await.unwrap_or_default());
    log.check("G7", ids.iter().all(|i| recent.contains(i)) && old.is_empty(),
        format!("received_after -6h + to + item_type email: {} of 3; received_before -6h: {}", ids.iter().filter(|i| recent.contains(i)).count(), old.len()));

    // A referenced inline image over real transport: a self-loop whose HTML
    // has a data: image (#37 turns it into a cid: attachment on send).
    let g8 = t("G8 תמונה בלולאה");
    let html8 = format!("<p>לולאה</p><img src=\"data:image/png;base64,{PNG_B64}\">");
    let sent = w.call("send_email", json!({"to": [SELF], "subject": g8, "html_body": html8})).await;
    let started = Instant::now();
    let mut got = None;
    while sent.is_ok() && started.elapsed() < Duration::from_secs(120) {
        if let Some(e) = rows(&w.call("list_emails", json!({"query": format!("{q} G8"), "count": 5})).await).into_iter().find(|e| s(e, "subject") == g8) {
            got = Some(e);
            break;
        }
        tokio::time::sleep(Duration::from_secs(5)).await;
    }
    match got {
        Some(e) => {
            let atts = rows(&w.call("list_attachments", json!({"email_id": s(&e, "id")})).await);
            let resolved = w.call("get_email", json!({"email_id": s(&e, "id"), "resolve_inline_images": true})).await.unwrap_or_default();
            log.check("G8", atts.len() == 1 && atts[0]["is_inline"] == true && resolved["inline_images_resolved"] == 1
                && s(&resolved, "html_body").contains("data:image/png;base64,") && s(&resolved, "html_body").contains("לולאה"),
                format!("self-loop after {:.0}s: attachments {:?}; resolved {}", started.elapsed().as_secs_f64(),
                    atts.iter().map(|a| (s(a, "content_id"), a["is_inline"].clone())).collect::<Vec<_>>(), resolved["inline_images_resolved"]));
            let dir2 = temp_dir("inbound-inline-filter");
            let only_regular = w.call("save_attachments", json!({"email_id": s(&e, "id"), "save_dir": dir2, "inline": false})).await;
            let n_regular = std::fs::read_dir(&dir2).map(|d| d.count()).unwrap_or(99);
            let dir3 = temp_dir("inbound-inline-only");
            let only_inline = w.call("save_attachments", json!({"email_id": s(&e, "id"), "save_dir": dir3, "inline": true})).await;
            let n_inline = std::fs::read_dir(&dir3).map(|d| d.count()).unwrap_or(99);
            // An email with only an inline image has nothing for inline:false:
            // save_attachments reports that as an error by design (#38).
            log.check("G5", only_regular.as_ref().is_err_and(|e| e.contains("No attachments matched")) && n_regular == 0
                && only_inline.is_ok() && n_inline == 1,
                format!("save_attachments inline:false -> {n_regular} files, {:?}; inline:true -> {n_inline} file", only_regular.err()));
        }
        None => log.fail("G8", format!("self-loop did not arrive: {sent:?}")),
    }

    // ---- sending Hebrew back to Gmail (O) -------------------------------------------
    let he_html = format!("<p>תשובה בעברית {run}</p><img src=\"data:image/png;base64,{PNG_B64}\">");
    let r = w.call("reply_email", json!({"email_id": s(&g1, "id"), "html_body": he_html, "send": true})).await;
    log.check("O1", r.as_ref().is_ok_and(|v| s(v, "status") == "sent"), format!("reply_email (Hebrew HTML + data: image) -> {r:?}"));
    println!("VERIFY O1: Gmail thread of G1 gets a reply containing \"תשובה בעברית {run}\" with one inline image");
    let att_dir = temp_dir("inbound-outbound");
    let att = att_dir.join("דוח.txt");
    std::fs::write(&att, "תוכן הקובץ המצורף").unwrap();
    let o2 = t("O2 הודעה יוצאת");
    let r = w.call("send_email", json!({"to": [GMAIL], "subject": o2, "body": format!("גוף בעברית {run}"), "attachments": [att], "importance": "high"})).await;
    log.check("O2", r.as_ref().is_ok_and(|v| s(v, "status") == "sent"), format!("send_email (Hebrew subject, body, attachment דוח.txt) -> {r:?}"));
    println!("VERIFY O2: Gmail receives {o2:?} with body \"גוף בעברית {run}\" and attachment דוח.txt = \"תוכן הקובץ המצורף\"");
    let o3 = t("O3 טיוטה שנשלחה");
    let r = match w.call("create_draft", json!({"to": [GMAIL], "subject": t("O3 draft"), "body": "טיוטה"})).await {
        Ok(v) => w.call("update_draft", json!({"email_id": s(&v, "id"), "subject": o3, "body": format!("נשלח מטיוטה {run}"), "send": true})).await,
        Err(e) => Err(e),
    };
    log.check("O3", r.as_ref().is_ok_and(|v| s(v, "status") == "sent"), format!("create_draft + update_draft(send) -> {r:?}"));
    println!("VERIFY O3: Gmail receives {o3:?} with body \"נשלח מטיוטה {run}\"");

    // ---- real meeting invites (M) -----------------------------------------------------
    let organizer_ok = ["M1", "M2", "M4a", "M5", "M7"].iter().all(|l| ev(l, &events).is_some_and(|e| e["is_meeting"] == true));
    log.check("M0", organizer_ok, format!("invites from Google are meetings: {organizer_ok}; organizers {:?}",
        events.iter().map(|e| s(e, "organizer")).collect::<std::collections::BTreeSet<_>>()));
    let m1 = ev("M1", &events).unwrap_or_default();
    let r = w.call("respond_to_meeting", json!({"event_id": s(&m1, "id"), "response": "accept", "notify": false})).await;
    let new_id = r.as_ref().map(|v| s(v, "id")).unwrap_or_default();
    let g = w.call("get_event", json!({"event_id": new_id, "include": []})).await;
    log.check("M1", r.as_ref().is_ok_and(|v| s(v, "status") == "accept_saved") && g.as_ref().is_ok_and(|g| s(g, "my_response") == "accepted"),
        format!("accept, notify:false -> {:?}; returned id opens, my_response {:?}", r.as_ref().map(|v| s(v, "status")), g.as_ref().map(|g| s(g, "my_response"))));
    println!("VERIFY M1: Google Calendar: {SELF} is still needsAction on {:?}; Gmail has no \"Accepted\" reply for it", s(&m1, "subject"));
    let m2 = ev("M2", &events).unwrap_or_default();
    let r = w.call("respond_to_meeting", json!({"event_id": s(&m2, "id"), "response": "tentative", "comment": format!("אולי אגיע {run}"), "notify": true})).await;
    log.check("M2", r.as_ref().is_ok_and(|v| s(v, "status") == "tentative_sent"), format!("tentative with a Hebrew comment, notify:true -> {r:?}"));
    println!("VERIFY M2: Google Calendar: {SELF} is tentative on {:?}; Gmail gets the response with \"אולי אגיע {run}\"", s(&m2, "subject"));
    let series: Vec<Value> = events.iter().filter(|e| s(e, "subject").contains("M3")).cloned().collect();
    let occ: Vec<String> = series.iter().map(|e| s(e, "occurrence_id")).collect();
    if series.len() == 6 && occ.iter().all(|o| !o.is_empty()) {
        let r = w.call("respond_to_meeting", json!({"event_id": occ[2], "response": "decline", "notify": false})).await;
        let left = rows(&w.call("list_events", json!({"query": format!("{q} M3"), "start_after": "2099-06-01", "start_before": "2099-07-31"})).await);
        let declined_gone = !left.iter().any(|e| s(e, "start") == s(&series[2], "start"));
        log.check("M3", r.is_ok() && left.len() == 5 && declined_gone,
            format!("decline occurrence 3 of 6, notify:false -> {:?}; {} occurrences left, the declined one gone {declined_gone}", r.as_ref().err(), left.len()));
        println!("VERIFY M3: Google Calendar: no response from {SELF} on {:?} (any occurrence)", s(&series[0], "subject"));
    } else {
        log.fail("M3", format!("expected 6 occurrences with occurrence ids, got {} {occ:?}", series.len()));
    }
    let batch: Vec<String> = ["M4a", "M4b"].iter().filter_map(|l| ev(l, &events)).map(|e| s(&e, "id")).collect();
    let r = w.call("respond_to_meeting", json!({"event_id": batch, "response": "accept", "notify": false})).await;
    let res = rows(&r);
    let mut seen = Vec::new();
    let mut opened = 0;
    for x in &res {
        let g = w.call("get_event", json!({"event_id": s(x, "id"), "include": []})).await;
        if g.as_ref().is_ok_and(|g| s(g, "subject").contains("M4") && s(g, "my_response") == "accepted") { opened += 1; }
        seen.push(match &g {
            Ok(g) => format!("{} -> {:?} {:?}", s(x, "status"), s(g, "subject").replace(&format!("{TAG} {run} "), ""), s(g, "my_response")),
            Err(e) => format!("{} {:?} -> {e}", s(x, "status"), s(x, "error")),
        });
    }
    log.check("M4", res.len() == 2 && opened == 2, format!("batch accept of 2, notify:false -> {seen:?}; {:?}", r.as_ref().err()));
    let m5 = ev("M5", &events).unwrap_or_default();
    let m1r = ev("M1", &events).unwrap_or_default();
    log.check("M5", m5["user_is_optional"] == true && m5["user_is_required"] == false && m1r["user_is_required"] == true && m1r["user_is_optional"] == false,
        format!("optional invite: user_is_optional {} user_is_required {}; required invite: {} / {}",
            m5["user_is_optional"], m5["user_is_required"], m1r["user_is_required"], m1r["user_is_optional"]));
    let requests = rows(&w.call("list_emails", json!({"query": format!("{q} M6"), "item_type": "meeting"})).await);
    match requests.first() {
        Some(req) => {
            let r = w.call("respond_to_meeting", json!({"event_id": s(req, "id"), "response": "accept", "notify": false})).await;
            let g = w.call("get_event", json!({"event_id": r.as_ref().map(|v| s(v, "id")).unwrap_or_default(), "include": []})).await;
            log.check("M6", r.is_ok() && g.as_ref().is_ok_and(|g| s(g, "subject").contains("M6") && s(g, "my_response") == "accepted"),
                format!("respond via the Inbox meeting request ({:?}) -> {:?}; the returned id opens {:?}", s(req, "subject"),
                    r.as_ref().map(|v| s(v, "status")), g.as_ref().map(|g| (s(g, "subject"), s(g, "my_response"), s(g, "start")))));
        }
        None => log.fail("M6", "no meeting request for M6 in the Inbox"),
    }
    let m7 = ev("M7", &events).unwrap_or_default();
    let d7 = w.call("get_event", json!({"event_id": s(&m7, "id")})).await.unwrap_or_default();
    // Google appends the Meet link to the location ("...; https://meet.google.com/...").
    log.check("M7", s(&m7, "subject") == t("M7 ישיבת תכנון") && s(&m7, "location").starts_with("תל אביב, חדר 3") && s(&d7, "body").contains("סדר יום: תכנון רבעוני"),
        format!("Hebrew invite: subject {:?}, location {:?}, body has the Hebrew agenda {}", s(&m7, "subject"), s(&m7, "location"), s(&d7, "body").contains("סדר יום")));
    let u = w.call("update_event", json!({"event_id": s(&m7, "id"), "location": "x", "require_no_attendees": true})).await;
    let d = w.call("delete_event", json!({"event_id": s(&m7, "id"), "require_no_attendees": true})).await;
    let still = w.call("get_event", json!({"event_id": s(&m7, "id"), "include": []})).await;
    log.check("M8", u.is_err() && d.is_err() && still.as_ref().is_ok_and(|g| s(g, "location") == s(&m7, "location")),
        format!("require_no_attendees on a real invite: update refused {}, delete refused {}, unchanged {}", u.is_err(), d.is_err(), still.is_ok()));
}

/// Deletes the run's calendar items quietly (no decline reaches Google), then
/// sweeps Inbox / Sent Items / Deleted Items / Drafts until nothing of the
/// run has turned up for 30 s (responses and sends reach Sent Items late).
async fn cleanup(w: &mut Wire, log: &mut Log, run: &str) {
    let q = format!("\"{run}\"");
    let mut errors = Vec::new();
    for _ in 0..3 {
        let left = rows(&w.call("list_events", json!({"query": q, "start_after": "2099-06-01", "start_before": "2099-07-31"})).await);
        let ids: std::collections::BTreeSet<String> = left.iter().map(|e| s(e, "id")).collect();
        if ids.is_empty() { break; }
        let ids: Vec<String> = ids.into_iter().collect();
        if let Err(e) = w.call("delete_event", json!({"event_id": ids, "notify": false, "permanent": true})).await { errors.push(e); }
    }
    let started = Instant::now();
    let mut quiet = Instant::now();
    let mut swept = 0;
    loop {
        let mut busy = !ids_of(&w.call("list_emails", json!({"folder": "outbox", "query": q})).await.unwrap_or_default()).is_empty();
        for folder in ["inbox", "sent", "drafts", "deleted"] {
            for id in ids_of(&w.call("list_emails", json!({"folder": folder, "query": q, "count": 200})).await.unwrap_or_default()) {
                busy = true;
                match w.call("delete_email", json!({"email_id": id, "permanent": true})).await {
                    Ok(_) => swept += 1,
                    Err(e) => errors.push(e),
                }
            }
        }
        if busy { quiet = Instant::now(); }
        if quiet.elapsed() >= Duration::from_secs(30) || started.elapsed() >= Duration::from_secs(240) { break; }
        tokio::time::sleep(Duration::from_secs(5)).await;
    }
    // Throttled deletes are retried by later passes; only what is still there counts.
    let mut left = Vec::new();
    for folder in ["inbox", "sent", "drafts", "deleted", "outbox"] {
        let n = ids_of(&w.call("list_emails", json!({"folder": folder, "query": q, "count": 200})).await.unwrap_or_default()).len();
        if n > 0 { left.push(format!("{folder}: {n}")); }
    }
    let events_left = rows(&w.call("list_events", json!({"query": q, "start_after": "2099-06-01", "start_before": "2099-07-31"})).await).len();
    log.check("CLEAN", left.is_empty() && events_left == 0,
        format!("swept {swept} mail items in {:.0}s; left {left:?}; calendar rows left {events_left}; transient errors {}", started.elapsed().as_secs_f64(), errors.len()));
}
