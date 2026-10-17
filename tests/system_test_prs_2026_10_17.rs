//! Live system test of open PRs #37–#56 merged together (2026-10-17).
//! Plan: `SYSTEM_TEST_PLAN_2026-10-17-PRS.md`.
//!
//! Unlike the earlier system tests, every call goes through the real MCP
//! tool layer: `OutlookMcpServer` over a real `WindowsOutlookClient`, served
//! in-process and called with JSON-RPC `tools/call` over rmcp's stdio codec.
//! So aliases, batching, `undo`, `require_no_attendees` and the output
//! controls all run exactly as a client sees them.
//!
//! Never panics before cleanup: each step is recorded in a results log, and
//! the only `assert!` is the last line.
//!
//!     cargo test --test system_test_prs_2026_10_17 -- --ignored --nocapture

use std::sync::Arc;
use std::time::{Duration, Instant};

use base64::Engine;
use outlook_mcp_rs::outlook::client::WindowsOutlookClient;
use outlook_mcp_rs::server::OutlookMcpServer;
use serde_json::{json, Value};

const TAG: &str = "[outlook-mcp-rs systest 1017]";
const SELF: &str = "adamkopelman@outlook.com";
const NOBODY: &str = "nobody@example.invalid";
const HE: &str = "שלום עולם בדיקה";
const PNG_B64: &str = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

// ---- JSON-RPC over the stdio codec ------------------------------------------

struct Wire {
    reader: tokio::io::BufReader<tokio::io::ReadHalf<tokio::io::DuplexStream>>,
    writer: tokio::io::WriteHalf<tokio::io::DuplexStream>,
    next_id: u64,
}

impl Wire {
    async fn start(server: OutlookMcpServer) -> Self {
        use rmcp::ServiceExt;
        let (client_io, server_io) = tokio::io::duplex(16 * 1024 * 1024);
        tokio::spawn(async move {
            if let Ok(running) = server.serve(server_io).await {
                let _ = running.waiting().await;
            }
        });
        let (read_half, writer) = tokio::io::split(client_io);
        let mut w = Wire { reader: tokio::io::BufReader::new(read_half), writer, next_id: 1 };
        w.request("initialize", json!({
            "protocolVersion": "2025-06-18", "capabilities": {},
            "clientInfo": {"name": "systest-1017", "version": "0"}}))
            .await
            .expect("initialize");
        w.send(&json!({"jsonrpc": "2.0", "method": "notifications/initialized"})).await;
        w
    }

    async fn send(&mut self, msg: &Value) {
        use tokio::io::AsyncWriteExt;
        let mut bytes = serde_json::to_vec(msg).unwrap();
        bytes.push(b'\n');
        self.writer.write_all(&bytes).await.unwrap();
    }

    /// One request; Err if the wire text isn't UTF-8 or the server closed.
    async fn request(&mut self, method: &str, params: Value) -> Result<Value, String> {
        use tokio::io::AsyncBufReadExt;
        let id = self.next_id;
        self.next_id += 1;
        self.send(&json!({"jsonrpc": "2.0", "id": id, "method": method, "params": params})).await;
        loop {
            let mut line = Vec::new();
            let n = self.reader.read_until(b'\n', &mut line).await.map_err(|e| e.to_string())?;
            if n == 0 {
                return Err("server closed".into());
            }
            let wire = String::from_utf8(line).map_err(|_| "response is not valid UTF-8".to_string())?;
            let msg: Value = serde_json::from_str(&wire).map_err(|e| e.to_string())?;
            if msg["id"] == id {
                return Ok(msg);
            }
        }
    }

    /// `tools/call`: Ok(parsed result) or Err(error text), whether the tool
    /// failed with a JSON-RPC error or an `isError` result.
    async fn call(&mut self, tool: &str, args: Value) -> Result<Value, String> {
        let msg = self.request("tools/call", json!({"name": tool, "arguments": args})).await?;
        if let Some(e) = msg.get("error") {
            return Err(e["message"].as_str().unwrap_or("").to_string());
        }
        let r = &msg["result"];
        let text = r["content"][0]["text"].as_str().unwrap_or("").to_string();
        if r["isError"] == true {
            return Err(text);
        }
        Ok(serde_json::from_str(&text).unwrap_or(Value::String(text)))
    }
}

// ---- results log ------------------------------------------------------------

#[derive(Default)]
struct Log(Vec<(String, bool, String)>);

impl Log {
    fn check(&mut self, id: &str, ok: bool, detail: impl Into<String>) -> bool {
        let detail = detail.into();
        println!("[{}] {id}: {detail}", if ok { "PASS" } else { "FAIL" });
        self.0.push((id.to_string(), ok, detail));
        ok
    }
    fn fail(&mut self, id: &str, detail: impl Into<String>) {
        self.check(id, false, detail);
    }
}

fn s(v: &Value, k: &str) -> String {
    v[k].as_str().unwrap_or("").to_string()
}

fn ids_of(rows: &Value) -> Vec<String> {
    rows.as_array().map(|a| a.iter().map(|r| s(r, "id")).collect()).unwrap_or_default()
}

fn subjects_of(rows: &Value) -> Vec<String> {
    rows.as_array().map(|a| a.iter().map(|r| s(r, "subject")).collect()).unwrap_or_default()
}

fn temp_dir(name: &str) -> std::path::PathBuf {
    let dir = std::env::temp_dir().join(format!("outlook-mcp-rs-systest-1017-{}-{name}", std::process::id()));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();
    dir
}

/// Items the test created, by tool family, with their latest known id.
#[derive(Default)]
struct Created {
    emails: Vec<String>,
    events: Vec<String>,
    tasks: Vec<String>,
    notes: Vec<String>,
}

#[tokio::test(flavor = "multi_thread")]
#[ignore]
async fn system_test_prs_2026_10_17() {
    let run = format!("r{}", chrono::Local::now().timestamp());
    let t = |label: &str| format!("{TAG} {run} {label}");
    let server = OutlookMcpServer::new(Arc::new(WindowsOutlookClient::new()));
    let mut w = Wire::start(server).await;
    let mut log = Log::default();
    let mut made = Created::default();
    println!("run token {run}");

    body(&mut w, &mut log, &mut made, &run, &t).await;

    // ---- cleanup (always) ----------------------------------------------------
    println!("\n---- cleanup");
    cleanup(&mut w, &mut log, &made, &run).await;

    println!("\n---- summary");
    let failed: Vec<_> = log.0.iter().filter(|(_, ok, _)| !ok).collect();
    for (id, ok, detail) in &log.0 {
        println!("{:<6} {} {}", id, if *ok { "PASS" } else { "FAIL" }, if *ok { "" } else { detail });
    }
    println!("{} / {} passed", log.0.len() - failed.len(), log.0.len());
    assert!(failed.is_empty(), "{} step(s) failed", failed.len());
}

async fn body(w: &mut Wire, log: &mut Log, made: &mut Created, run: &str, t: &dyn Fn(&str) -> String) {
    let dir = temp_dir("files");

    // ---- phase 0: seed ------------------------------------------------------
    let seed = |res: Result<Value, String>, label: &str, log: &mut Log| -> Option<String> {
        match res {
            Ok(v) if !s(&v, "id").is_empty() => Some(s(&v, "id")),
            other => {
                log.fail("SEED", format!("{label}: {other:?}"));
                None
            }
        }
    };
    let he_body = format!("{HE} גוף המשימה");
    let r = w.call("create_task", json!({"subject": t("task A"), "due_date": "2099-03-10",
        "importance": "high", "categories": ["Red Category"], "body": he_body})).await;
    let ta = seed(r, "TA", log);
    let r = w.call("create_task", json!({"subject": t("task B"), "due_date": "2099-03-20T09:00",
        "categories": ["Blue Category"]})).await;
    let tb = seed(r, "TB", log);
    let r = w.call("create_task", json!({"subject": t("task C"), "importance": "low"})).await;
    let tc = seed(r, "TC", log);
    let r = w.call("create_note", json!({"body": t("alpha note")})).await;
    let na = seed(r, "NA", log);
    let nb_text = format!("{} {HE}", t("hebrew note"));
    let r = w.call("create_note", json!({"body": nb_text})).await;
    let nb = seed(r, "NB", log);
    let ev = |subj: String, start: &str, end: &str, extra: Value| {
        let mut a = json!({"subject": subj, "start": start, "end": end, "notify": false});
        a.as_object_mut().unwrap().extend(extra.as_object().unwrap().clone());
        a
    };
    let r = w.call("create_event", ev(t("event A"), "2099-04-13T10:00", "2099-04-13T10:30",
        json!({"show_as": "busy", "categories": ["Red Category"]}))).await;
    let mut ea = seed(r, "EA", log);
    let r = w.call("create_event", ev(t("event B"), "2099-04-14T11:00", "2099-04-14T11:30",
        json!({"show_as": "free", "categories": ["Blue Category"]}))).await;
    let mut eb = seed(r, "EB", log);
    let r = w.call("create_event", ev(t("event X"), "2099-03-04T09:00", "2099-03-04T09:30", json!({}))).await;
    let ex = seed(r, "EX", log);
    let r = w.call("create_event", ev(t("event Y"), "2099-04-03T09:00", "2099-04-03T09:30", json!({}))).await;
    let ey = seed(r, "EY", log);
    let r = w.call("create_event", ev(t("event R"), "2099-04-20T08:00", "2099-04-20T08:15",
        json!({"recurrence": {"pattern": "daily", "occurrences": 5}}))).await;
    let er = seed(r, "ER", log);
    let r = w.call("create_draft", json!({"to": [NOBODY], "subject": t("draft alpha beta"), "body": "x"})).await;
    let da = seed(r, "DA", log);
    let r = w.call("create_draft", json!({"to": [NOBODY], "subject": t("draft beta alpha"), "body": "x"})).await;
    let db = seed(r, "DB", log);
    for id in [&ta, &tb, &tc].into_iter().flatten() { made.tasks.push(id.clone()); }
    for id in [&na, &nb].into_iter().flatten() { made.notes.push(id.clone()); }
    for id in [&ea, &eb, &ex, &ey, &er].into_iter().flatten() { made.events.push(id.clone()); }
    for id in [&da, &db].into_iter().flatten() { made.emails.push(id.clone()); }
    let (Some(ta), Some(tb), Some(tc), Some(na), Some(nb), Some(_), Some(_), Some(ex), Some(ey), Some(er), Some(da), Some(db)) =
        (ta, tb, tc, na, nb, ea.clone(), eb.clone(), ex, ey, er, da, db)
    else {
        log.fail("SEED", "seeding failed; skipping the tests");
        return;
    };
    log.check("SEED", true, "12 seed items created");

    // ---- #37 content-writing tools (W) ---------------------------------------
    let html_path = dir.join("body.html");
    let html = format!("<p>{HE}</p><img src=\"data:image/png;base64,{PNG_B64}\">");
    std::fs::write(&html_path, &html).unwrap();
    match w.call("create_draft", json!({"to": [NOBODY], "subject": t("W1 file draft"),
        "html_body_file": html_path, "importance": "high", "categories": ["Green Category"]})).await {
        Ok(v) => {
            let id = s(&v, "id");
            made.emails.push(id.clone());
            let d = w.call("get_email", json!({"email_id": id, "body_format": "html"})).await.unwrap_or_default();
            let h = s(&d, "html_body");
            log.check("W1a", h.contains("cid:img-") && !h.contains("data:image") && h.contains(HE),
                format!("html cid/no data:/hebrew: {}", h.chars().take(160).collect::<String>()));
            // Email results have no `importance` field, so read it back through
            // #39's importance filter.
            let high = w.call("list_emails", json!({"folder": "drafts", "query": format!("\"{run} W1\""), "importance": "high"})).await;
            log.check("W1b", high.as_ref().is_ok_and(|v| ids_of(v) == vec![id.clone()])
                && d["categories"].as_array().is_some_and(|c| c.iter().any(|x| x == "Green Category")),
                format!("importance filter {:?} categories {}", high.as_ref().map(subjects_of), d["categories"]));
            let atts = w.call("list_attachments", json!({"email_id": id})).await.unwrap_or_default();
            let img = atts.as_array().and_then(|a| a.iter().find(|x| s(x, "content_id").starts_with("img-")).cloned());
            match img {
                Some(a) => {
                    let cid = s(&a, "content_id");
                    let got = w.call("get_inline_image", json!({"email_id": id, "content_id": cid})).await.unwrap_or_default();
                    let bytes = s(&got, "data_uri").split_once(',')
                        .and_then(|(_, b)| base64::engine::general_purpose::STANDARD.decode(b).ok());
                    let want = base64::engine::general_purpose::STANDARD.decode(PNG_B64).unwrap();
                    log.check("W1c", a["hidden"] == true && bytes.as_deref() == Some(&want[..]),
                        format!("hidden {} bytes equal {}", a["hidden"], bytes.as_deref() == Some(&want[..])));
                }
                None => log.fail("W1c", format!("no img-* attachment: {atts}")),
            }
        }
        Err(e) => log.fail("W1a", e),
    }

    let txt_path = dir.join("body.txt");
    let file_text = format!("{HE} מקובץ\nsecond line");
    std::fs::write(&txt_path, &file_text).unwrap();
    for (tool, getter, idk, args) in [
        ("create_task", "get_task", "task_id", json!({"subject": t("W2 task"), "body_file": txt_path})),
        ("create_note", "get_note", "note_id", json!({"body_file": txt_path})),
        ("create_event", "get_event", "event_id", json!({"subject": t("W2 event"), "start": "2099-05-01T09:00",
            "end": "2099-05-01T09:30", "notify": false, "body_file": txt_path})),
    ] {
        match w.call(tool, args).await {
            Ok(v) => {
                let id = s(&v, "id");
                match tool {
                    "create_task" => made.tasks.push(id.clone()),
                    "create_note" => made.notes.push(id.clone()),
                    _ => made.events.push(id.clone()),
                }
                let d = w.call(getter, json!({idk: id})).await.unwrap_or_default();
                // Outlook re-renders an appointment's body (it is stored as RTF), so
                // its line breaks come back as " \r\n" or "\r\n\r\n" (confirmed with a
                // raw COM probe): compare the words, not the whitespace.
                let words = |x: &str| x.split_whitespace().collect::<Vec<_>>().join(" ");
                let b = s(&d, "body");
                log.check(&format!("W2-{}", &tool[7..]), words(&b).contains(&words(&file_text)),
                    format!("body {:?}", b.chars().take(60).collect::<String>()));
            }
            Err(e) => log.fail(&format!("W2-{}", &tool[7..]), e),
        }
    }
    // A note's subject is its first line: the W2 note is titled with Hebrew, not the tag.

    let w3_subject = t("W3 must not exist");
    let r = w.call("create_task", json!({"subject": w3_subject, "body": "a", "body_file": txt_path})).await;
    let leaked = w.call("list_tasks", json!({"query": format!("\"{w3_subject}\"")})).await.unwrap_or_default();
    log.check("W3", r.is_err() && ids_of(&leaked).is_empty(), format!("result {:?}; matching tasks {}", r.err(), ids_of(&leaked).len()));

    // W5: deprecated html: true + body, and update_draft's draft_id alias.
    match w.call("create_draft", json!({"to": [NOBODY], "subject": t("W5 legacy html"), "html": true, "body": "<b>bold</b>"})).await {
        Ok(v) => {
            let id = s(&v, "id");
            made.emails.push(id.clone());
            let d = w.call("get_email", json!({"email_id": id, "body_format": "html"})).await.unwrap_or_default();
            log.check("W5a", s(&d, "html_body").contains("<b>bold</b>"), "html: true + body stores HTML");
            let u = w.call("update_draft", json!({"draft_id": id, "subject": t("W5 legacy renamed")})).await;
            let new_id = u.as_ref().map(|v| s(v, "id")).unwrap_or_default();
            if !new_id.is_empty() { made.emails.push(new_id.clone()); }
            let d = w.call("get_email", json!({"email_id": if new_id.is_empty() { id } else { new_id }})).await.unwrap_or_default();
            log.check("W5b", u.is_ok() && s(&d, "subject") == t("W5 legacy renamed"), format!("{u:?}"));
        }
        Err(e) => log.fail("W5a", e),
    }

    // W4: the one real send, to the mailbox itself.
    let w4_subject = t("W4 self-loop via update_draft send");
    let mut received: Option<String> = None;
    match w.call("create_draft", json!({"to": [SELF], "subject": w4_subject, "body": "self-loop"})).await {
        Ok(v) => {
            let id = s(&v, "id");
            made.emails.push(id.clone());
            let r = w.call("update_draft", json!({"email_id": id, "send": true})).await;
            let sent = r.as_ref().is_ok_and(|v| s(v, "status") == "sent");
            log.check("W4a", sent, format!("{r:?}"));
            if sent {
                let started = Instant::now();
                let mut arrived = false;
                while started.elapsed() < Duration::from_secs(120) {
                    let rows = w.call("list_emails", json!({"query": format!("\"{run} W4\""), "count": 5})).await.unwrap_or_default();
                    if let Some(r) = rows.as_array().and_then(|a| a.iter().find(|r| s(r, "subject") == w4_subject)) {
                        received = Some(s(r, "id"));
                        arrived = true;
                        break;
                    }
                    tokio::time::sleep(Duration::from_secs(5)).await;
                }
                log.check("W4b", arrived, format!("arrived in Inbox after {:?}", started.elapsed()));
            }
        }
        Err(e) => log.fail("W4a", e),
    }

    // ---- #38 read tools (R) ---------------------------------------------------
    match w.call("get_task", json!({"task_id": [ta, "bogus|id"], "include": []})).await {
        Ok(v) => {
            let a = v.as_array().cloned().unwrap_or_default();
            log.check("R1", a.len() == 2 && s(&a[0], "id") == ta && a[0].get("body").is_none_or(Value::is_null)
                && s(&a[1], "id") == "bogus|id" && !s(&a[1], "error").is_empty(), format!("{v}"));
        }
        Err(e) => log.fail("R1", e),
    }
    let ea_id = ea.clone().unwrap_or_default();
    let eb_id = eb.clone().unwrap_or_default();
    match w.call("get_event", json!({"event_id": [ea_id, eb_id], "include": []})).await {
        Ok(v) => { log.check("R2a", v.as_array().is_some_and(|a| a.len() == 2), format!("{} results", v.as_array().map_or(0, Vec::len))); }
        Err(e) => log.fail("R2a", e),
    };
    // max_body_chars is clamped to 1,000..5,000,000, so the body must be longer than 1,000.
    let long_event = w.call("create_event", json!({"subject": t("R2 long body"), "start": "2099-05-05T09:00",
        "end": "2099-05-05T09:30", "notify": false, "body": "0123456789".repeat(300)})).await.map(|v| s(&v, "id")).ok();
    if let Some(id) = &long_event {
        made.events.push(id.clone());
        match w.call("get_event", json!({"event_id": id, "max_body_chars": 1000})).await {
            Ok(v) => { log.check("R2b", v["body_truncated"] == true && v["body_length"].as_u64().unwrap_or(0) >= 3000,
                format!("truncated {} length {}", v["body_truncated"], v["body_length"])); }
            Err(e) => log.fail("R2b", e),
        };
    }
    let out = temp_dir("note-out");
    match w.call("get_note", json!({"note_id": nb, "output_dir": out})).await {
        Ok(v) => {
            let f = s(&v, "body_file");
            let content = std::fs::read_to_string(&f).unwrap_or_default();
            log.check("R3", !f.is_empty() && content.replace("\r\n", "\n").contains(&nb_text), format!("body_file {f}"));
        }
        Err(e) => log.fail("R3", e),
    };
    let r = w.call("get_email", json!({"email_id": da, "body_format": "text", "prefer_html": true})).await;
    log.check("R4", r.as_ref().is_err_and(|e| e.contains("prefer_html") && e.contains("body_format")), format!("{r:?}"));

    // ---- #39 list conventions (L) ---------------------------------------------
    let tq = format!("\"{run}\"");
    let tasks = |v: &Value| -> Vec<String> { ids_of(v) };
    let r = w.call("list_tasks", json!({"query": tq, "due_after": "2099-03-01", "due_before": "2099-03-15"})).await;
    log.check("L1a", r.as_ref().is_ok_and(|v| tasks(v) == vec![ta.clone()]), format!("{:?}", r.as_ref().map(subjects_of)));
    let r = w.call("list_tasks", json!({"query": tq, "due_after": "2099-03-01", "due_before": "2099-03-20"})).await;
    let got = r.as_ref().map(tasks).unwrap_or_default();
    log.check("L1b", got.len() == 2 && got.contains(&ta) && got.contains(&tb), format!("{:?}", r.as_ref().map(subjects_of)));
    let r = w.call("list_tasks", json!({"query": tq, "importance": ["high", "low"]})).await;
    let got = r.as_ref().map(tasks).unwrap_or_default();
    log.check("L2a", got.len() == 2 && got.contains(&ta) && got.contains(&tc), format!("{:?}", r.as_ref().map(subjects_of)));
    let r = w.call("list_tasks", json!({"query": tq, "category": ["Red Category", "Blue Category"]})).await;
    let got = r.as_ref().map(tasks).unwrap_or_default();
    log.check("L2b", got.len() == 2 && got.contains(&ta) && got.contains(&tb), format!("{:?}", r.as_ref().map(subjects_of)));
    let april = json!({"query": tq, "start_after": "2099-04-01", "start_before": "2099-04-30"});
    let with = |extra: Value| { let mut a = april.clone(); a.as_object_mut().unwrap().extend(extra.as_object().unwrap().clone()); a };
    let r = w.call("list_events", with(json!({"show_as": "free"}))).await;
    log.check("L3a", r.as_ref().is_ok_and(|v| subjects_of(v) == vec![t("event B")]), format!("{:?}", r.as_ref().map(subjects_of)));
    let r = w.call("list_events", with(json!({"category": ["Red Category"]}))).await;
    log.check("L3b", r.as_ref().is_ok_and(|v| subjects_of(v) == vec![t("event A")]), format!("{:?}", r.as_ref().map(subjects_of)));
    let r = w.call("list_notes", json!({"query": format!("{run} {HE}")})).await;
    log.check("L4", r.as_ref().is_ok_and(|v| ids_of(v) == vec![nb.clone()]), format!("{:?}", r.as_ref().map(subjects_of)));
    let r = w.call("list_notes", json!({"query": run, "created_after": "-1h"})).await;
    let got = r.as_ref().map(ids_of).unwrap_or_default();
    log.check("L5", got.contains(&na) && got.contains(&nb), format!("{:?}", r.as_ref().map(subjects_of)));
    let errs = [
        ("list_events", json!({"start_after": "2099-04-10", "start_before": "2099-04-01"})),
        ("list_tasks", json!({"importance": "urgent"})),
        ("list_emails", json!({"item_type": "fax"})),
        ("list_emails", json!({"since_days": 3, "received_after": "-10d"})),
    ];
    for (i, (tool, args)) in errs.into_iter().enumerate() {
        let r = w.call(tool, args.clone()).await;
        log.check(&format!("L6{}", (b'a' + i as u8) as char), r.is_err(), format!("{tool} {args} -> {:?}", r.err()));
    }
    let r = w.call("list_events", json!({"query": tq, "start_date": "2099-04-13", "end_date": "2099-04-14"})).await;
    let got = r.as_ref().map(subjects_of).unwrap_or_default();
    log.check("L7a", got.contains(&t("event A")) && got.contains(&t("event B")), format!("start_date/end_date aliases: {got:?}"));
    let r = w.call("list_emails", json!({"since_days": 1, "flagged": false, "count": 3})).await;
    log.check("L7b", r.is_ok(), format!("since_days + flagged: {:?}", r.err()));
    let r = w.call("list_emails", json!({"folder": "drafts", "query": format!("{run} subject:alpha*beta")})).await;
    let got = r.as_ref().map(ids_of).unwrap_or_default();
    log.check("L8", got == vec![da.clone()], format!("alpha*beta -> {:?}", r.as_ref().map(subjects_of)));
    let _ = &db;
    let mut pages = Vec::new();
    for off in 0..3 {
        let r = w.call("list_tasks", json!({"query": format!("{run} task"), "count": 1, "offset": off})).await;
        pages.push(r.as_ref().map(ids_of).unwrap_or_default());
    }
    let mut union: Vec<String> = pages.iter().flatten().cloned().collect();
    union.sort();
    let mut want = vec![ta.clone(), tb.clone(), tc.clone()];
    want.sort();
    log.check("L9", pages.iter().all(|p| p.len() == 1) && union == want, format!("pages {pages:?}"));

    // ---- #40 UTF-8 (U) -----------------------------------------------------------
    let d = w.call("get_task", json!({"task_id": ta})).await.unwrap_or_default();
    let n = w.call("get_note", json!({"note_id": nb})).await.unwrap_or_default();
    log.check("U1", s(&d, "body").contains(&he_body) && s(&n, "body").contains(&nb_text),
        format!("task body {:?} / note body {:?}", s(&d, "body"), s(&n, "body")));

    // ---- #42 locale-proof dates (Q) ---------------------------------------------
    let ids_in = |v: &Result<Value, String>| v.as_ref().map(ids_of).unwrap_or_default();
    let r = w.call("list_events", json!({"query": tq, "start_after": "2099-03-01", "start_before": "2099-03-10"})).await;
    log.check("Q1", ids_in(&r).contains(&ex) && !ids_in(&r).contains(&ey), format!("{:?}", r.as_ref().map(subjects_of)));
    let r = w.call("list_events", json!({"query": tq, "start_after": "2099-04-01", "start_before": "2099-04-05"})).await;
    log.check("Q2", ids_in(&r).contains(&ey) && !ids_in(&r).contains(&ex), format!("{:?}", r.as_ref().map(subjects_of)));
    let r = w.call("list_events", json!({"query": tq, "start_after": "2099-04-13T10:30", "start_before": "2099-04-14"})).await;
    let got = r.as_ref().map(subjects_of).unwrap_or_default();
    log.check("Q3", got == vec![t("event B")], format!("{got:?}"));

    // ---- #49 output controls (O) -----------------------------------------------
    let r = w.call("list_tasks", json!({"query": tq, "fields": ["subject", "due_date"]})).await;
    let ok = r.as_ref().is_ok_and(|v| v.as_array().is_some_and(|a| !a.is_empty() && a.iter().all(|row| {
        let mut k: Vec<&String> = row.as_object().unwrap().keys().collect();
        k.sort();
        k == vec!["due_date", "subject"]
    })));
    log.check("O1a", ok, format!("{:?}", r.as_ref().map(|v| v.get(0).cloned())));
    let r = w.call("list_tasks", json!({"fields": ["subject", "nonsense"]})).await;
    log.check("O1b", r.as_ref().is_err_and(|e| e.contains("nonsense") && e.contains("due_date")), format!("{:?}", r.err()));
    let r = w.call("list_events", with(json!({"compact": true}))).await;
    let row = r.as_ref().ok().and_then(|v| v.get(0).cloned()).unwrap_or_default();
    log.check("O2a", row.get("id").is_some() && row.get("start").is_some() && row.get("required_attendees").is_none(),
        format!("compact row keys {:?}", row.as_object().map(|o| o.keys().cloned().collect::<Vec<_>>())));
    let rows_dir = temp_dir("rows");
    let r = w.call("list_events", with(json!({"output_dir": rows_dir}))).await;
    match &r {
        Ok(v) => {
            let lines = std::fs::read_to_string(s(v, "rows_file")).unwrap_or_default();
            let parsed = lines.lines().filter(|l| serde_json::from_str::<Value>(l).is_ok_and(|x| x.get("id").is_some())).count();
            log.check("O2b", v["count"].as_u64() == Some(parsed as u64) && parsed >= 7, format!("count {} lines {parsed}", v["count"]));
        }
        Err(e) => log.fail("O2b", e.clone()),
    }
    let r = w.call("list_events", with(json!({"query": format!("{run} \"event A\"")}))).await;
    let row = r.as_ref().ok().and_then(|v| v.get(0).cloned()).unwrap_or_default();
    log.check("O3", row["user_is_required"] == true && row["required_count"] == 0,
        format!("user_is_required {} required_count {}", row["user_is_required"], row["required_count"]));
    let r1 = w.call("list_folders", json!({"compact": true})).await;
    let r2 = w.call("list_attachments", json!({"email_id": da, "fields": ["filename"]})).await;
    log.check("O4", r1.is_ok() && r2.is_ok(), format!("{:?} / {:?}", r1.err(), r2.err()));

    // ---- #50 ids after a write, occurrences (I) ----------------------------------
    let mut i1 = Vec::new();
    for (tool, getter, key, id, extra) in [
        ("update_task", "get_task", "task_id", tc.clone(), json!({"percent_complete": 25})),
        ("update_note", "get_note", "note_id", na.clone(), json!({"add_categories": ["Green Category"]})),
        ("update_event", "get_event", "event_id", ex.clone(), json!({"location": "Room 9", "notify": false})),
    ] {
        let mut args = json!({key: id});
        args.as_object_mut().unwrap().extend(extra.as_object().unwrap().clone());
        let r = w.call(tool, args).await;
        let new_id = r.as_ref().map(|v| s(v, "id")).unwrap_or_default();
        let opened = !new_id.is_empty() && w.call(getter, json!({key: new_id})).await.is_ok();
        i1.push(format!("{tool}: id {} opens {opened}", if new_id.is_empty() { "missing" } else { "returned" }));
        if !opened { i1.push(format!("{r:?}")); }
    }
    log.check("I1", i1.iter().all(|x| x.ends_with("opens true")), i1.join("; "));

    let series = json!({"query": format!("{run} \"event R\""), "start_after": "2099-04-20", "start_before": "2099-04-24"});
    let occ_rows = w.call("list_events", series.clone()).await.unwrap_or_default();
    let occ: Vec<(String, String)> = occ_rows.as_array().cloned().unwrap_or_default().iter()
        .map(|r| (s(r, "occurrence_id"), s(r, "start"))).collect();
    let distinct: std::collections::BTreeSet<_> = occ.iter().map(|(o, _)| o.clone()).collect();
    log.check("I2a", occ.len() == 5 && distinct.len() == 5 && !distinct.contains(""), format!("{occ:?}"));
    if occ.len() == 5 {
        let g = w.call("get_event", json!({"event_id": occ[1].0})).await.unwrap_or_default();
        log.check("I2b", s(&g, "start") == occ[1].1, format!("start {} vs {}", s(&g, "start"), occ[1].1));
        let u = w.call("update_event", json!({"event_id": occ[1].0, "subject": t("event R moved one"), "notify": false})).await;
        let rows = w.call("list_events", json!({"query": run, "start_after": "2099-04-20", "start_before": "2099-04-24"})).await.unwrap_or_default();
        let subs = subjects_of(&rows);
        let changed = subs.iter().filter(|x| **x == t("event R moved one")).count();
        log.check("I3a", u.is_ok() && changed == 1 && subs.iter().filter(|x| **x == t("event R")).count() == 4,
            format!("{u:?} subjects {subs:?}"));
        let r = w.call("update_event", json!({"event_id": occ[2].0, "recurrence": {"pattern": "daily", "occurrences": 2}, "notify": false})).await;
        log.check("I3b", r.is_err(), format!("{:?}", r.err()));
        let d = w.call("delete_event", json!({"event_id": occ[2].0, "notify": false})).await;
        let rows = w.call("list_events", json!({"query": run, "start_after": "2099-04-20", "start_before": "2099-04-24"})).await.unwrap_or_default();
        let left = rows.as_array().map_or(0, |a| a.iter().filter(|r| s(r, "subject").contains("event R")).count());
        log.check("I4", d.is_ok() && left == 4, format!("{d:?}; {left} left"));
    }
    let _ = &er;

    // ---- #51 batch writes and undo (B) -----------------------------------------
    match w.call("update_task", json!({"task_id": [ta, tb], "add_categories": ["Green Category"]})).await {
        Ok(v) => {
            let a = v.as_array().cloned().unwrap_or_default();
            let undos: Vec<Value> = a.iter().map(|x| x["undo"].clone()).collect();
            let rows = w.call("list_tasks", json!({"query": tq, "category": "Green Category"})).await.unwrap_or_default();
            let green = ids_of(&rows).len();
            let back = w.call("update_task", json!({"updates": undos})).await;
            let rows = w.call("list_tasks", json!({"query": tq, "category": "Green Category"})).await.unwrap_or_default();
            log.check("B1", a.len() == 2 && undos.iter().all(Value::is_object) && green == 2 && back.is_ok() && ids_of(&rows).is_empty(),
                format!("{} results, green before undo {green}, after {}; undo {:?}", a.len(), ids_of(&rows).len(), back.err()));
        }
        Err(e) => log.fail("B1", e),
    }
    let r = w.call("update_event", json!({"notify": false, "updates": [
        {"event_id": ea_id, "location": "B2 room", "notify": false},
        {"event_id": eb_id, "show_as": "tentative", "notify": false},
        {"event_id": "bogus|id", "location": "x", "notify": false},
    ]})).await;
    match &r {
        Ok(v) => {
            let a = v.as_array().cloned().unwrap_or_default();
            log.check("B2", a.len() == 3 && a[0].get("error").is_none() && a[1].get("error").is_none() && !s(&a[2], "error").is_empty(),
                format!("{v}"));
            if let Some(x) = a.first() && !s(x, "id").is_empty() { ea = Some(s(x, "id")); }
            if let Some(x) = a.get(1) && !s(x, "id").is_empty() { eb = Some(s(x, "id")); }
        }
        Err(e) => log.fail("B2", e.clone()),
    }
    // Outlook can't flag a draft ("Draft items cannot be marked"), so this uses
    // W4's received copy. Email results have no flag/importance fields: read
    // them back through #39's flag / importance filters.
    match received.clone() {
        Some(rid) => {
            let before = w4_state(w, run, &rid).await;
            match w.call("update_email", json!({"email_id": rid, "flag": "follow_up", "importance": "high"})).await {
                Ok(v) => {
                    let during = w4_state(w, run, &s(&v, "id")).await;
                    let undone = w.call("update_email", v["undo"].clone()).await;
                    let after = w4_state(w, run, &s(&v, "id")).await;
                    log.check("B3", before == (false, false) && during == (true, true) && undone.is_ok() && after == before,
                        format!("(flagged, high) before {before:?} during {during:?} after undo {after:?}; undo {}", v["undo"]));
                }
                Err(e) => log.fail("B3", e),
            }
        }
        None => log.fail("B3", "no received W4 copy to flag"),
    }
    match w.call("update_note", json!({"note_id": na, "color": "pink"})).await {
        Ok(v) => {
            let undo = v["undo"].clone();
            let undone = w.call("update_note", undo.clone()).await;
            log.check("B4", undo.get("color").is_some() && undone.is_ok(), format!("undo {undo} -> {:?}", undone.err()));
        }
        Err(e) => log.fail("B4", e),
    }
    let many: Vec<String> = (0..101).map(|i| format!("bogus|{i}")).collect();
    let r = w.call("update_task", json!({"task_id": many, "percent_complete": 50})).await;
    log.check("B5", r.as_ref().is_err_and(|e| e.contains("100")), format!("{:?}", r.err()));

    // ---- #53 calendar safety (S) --------------------------------------------
    let s1_subject = t("S1 must not exist");
    let r = w.call("create_event", json!({"subject": s1_subject, "start": "2099-05-02T09:00", "end": "2099-05-02T09:30",
        "required_attendees": [NOBODY], "require_no_attendees": true, "notify": false})).await;
    let leaked = w.call("list_events", json!({"query": format!("\"{s1_subject}\""), "start_after": "2099-05-01", "start_before": "2099-05-03"})).await.unwrap_or_default();
    for id in ids_of(&leaked) { made.events.push(id); }
    log.check("S1", r.is_err() && ids_of(&leaked).is_empty(), format!("{:?}; leaked {}", r.err(), ids_of(&leaked).len()));
    let ea_now = ea.clone().unwrap_or_default();
    let r_ok = w.call("update_event", json!({"event_id": ea_now, "location": "S2 room", "require_no_attendees": true, "notify": false})).await;
    if let Ok(v) = &r_ok && !s(v, "id").is_empty() { ea = Some(s(v, "id")); }
    let ea_now = ea.clone().unwrap_or_default();
    let r_no = w.call("update_event", json!({"event_id": ea_now, "add_required_attendees": [NOBODY], "require_no_attendees": true, "notify": false})).await;
    let g = w.call("get_event", json!({"event_id": ea_now, "include": []})).await.unwrap_or_default();
    log.check("S2", r_ok.is_ok() && r_no.is_err() && g["is_meeting"] == false,
        format!("ok {:?}; refused {:?}; is_meeting {}", r_ok.err(), r_no.err(), g["is_meeting"]));
    match w.call("create_event", json!({"subject": t("S3 meeting"), "start": "2099-05-03T09:00", "end": "2099-05-03T09:30",
        "required_attendees": [NOBODY], "notify": false})).await {
        Ok(v) => {
            let id = s(&v, "id");
            made.events.push(id.clone());
            let refused = w.call("delete_event", json!({"event_id": id, "require_no_attendees": true, "notify": false})).await;
            let still = w.call("get_event", json!({"event_id": id, "include": []})).await;
            let gone = w.call("delete_event", json!({"event_id": id, "notify": false, "permanent": true})).await;
            let after = w.call("get_event", json!({"event_id": id, "include": []})).await;
            log.check("S3", refused.is_err() && still.is_ok() && gone.is_ok() && after.is_err(),
                format!("refused {:?}; still {}; deleted {:?}; after {}", refused.err(), still.is_ok(), gone.err(), after.is_ok()));
        }
        Err(e) => log.fail("S3", e),
    }
    match w.request("tools/list", json!({})).await {
        Ok(msg) => {
            let tools = msg["result"]["tools"].as_array().cloned().unwrap_or_default();
            let mut ro: Vec<String> = tools.iter().filter(|t| t["annotations"]["readOnlyHint"] == true).map(|t| s(t, "name")).collect();
            ro.sort();
            let mut want_ro: Vec<String> = tools.iter().map(|t| s(t, "name"))
                .filter(|n| n.starts_with("list_") || n.starts_with("get_") || n == "check_availability").collect();
            want_ro.sort();
            let senders = ["send_email", "reply_email", "update_draft", "create_event", "update_event", "delete_event", "respond_to_meeting"];
            let not_open: Vec<&str> = senders.iter().copied().filter(|n| !tools.iter().any(|t| s(t, "name") == *n && t["annotations"]["openWorldHint"] == true)).collect();
            log.check("S4", tools.len() == 30 && ro == want_ro && not_open.is_empty(),
                format!("{} tools; read-only {ro:?}; senders not open-world {not_open:?}", tools.len()));
        }
        Err(e) => log.fail("S4", e),
    }
    let r = w.call("update_event", json!({"event_id": ea.clone().unwrap_or_default(), "location": "S5 room", "send_update": false})).await;
    if let Ok(v) = &r && !s(v, "id").is_empty() { ea = Some(s(v, "id")); }
    log.check("S5", r.is_ok(), format!("send_update alias: {:?}", r.err()));

    // ---- #41 permanent delete, batched (D) ------------------------------------
    let mut dtasks = Vec::new();
    for n in 1..=2 {
        if let Ok(v) = w.call("create_task", json!({"subject": t(&format!("D1 task {n}"))})).await {
            dtasks.push(s(&v, "id"));
        }
    }
    made.tasks.extend(dtasks.clone());
    let r = w.call("delete_task", json!({"task_id": dtasks, "permanent": true})).await;
    let in_deleted = w.call("list_emails", json!({"folder": "deleted", "query": format!("\"{run} D1\""), "count": 50})).await.unwrap_or_default();
    let in_drafts = w.call("list_emails", json!({"folder": "drafts", "query": format!("\"{run} D1\""), "count": 50})).await.unwrap_or_default();
    log.check("D1", r.as_ref().is_ok_and(|v| v.as_array().is_some_and(|a| a.len() == 2 && a.iter().all(|x| x["permanent"] == true)))
        && ids_of(&in_deleted).is_empty() && ids_of(&in_drafts).is_empty(),
        format!("{r:?}; in Deleted Items {}; in Drafts {}", ids_of(&in_deleted).len(), ids_of(&in_drafts).len()));
    let mut d2 = Vec::new();
    if let Ok(v) = w.call("create_note", json!({"body": t("D2 note")})).await {
        let id = s(&v, "id");
        made.notes.push(id.clone());
        d2.push(w.call("delete_note", json!({"note_id": id, "permanent": true})).await.map(|_| ()));
    }
    if let Ok(v) = w.call("create_event", json!({"subject": t("D2 event"), "start": "2099-05-04T09:00", "end": "2099-05-04T09:30", "notify": false})).await {
        let id = s(&v, "id");
        made.events.push(id.clone());
        d2.push(w.call("delete_event", json!({"event_id": id, "permanent": true, "notify": false})).await.map(|_| ()));
    }
    let in_deleted = w.call("list_emails", json!({"folder": "deleted", "query": format!("\"{run} D2\""), "count": 50})).await.unwrap_or_default();
    let in_drafts = w.call("list_emails", json!({"folder": "drafts", "query": format!("\"{run} D2\""), "count": 50})).await.unwrap_or_default();
    log.check("D2", d2.len() == 2 && d2.iter().all(Result::is_ok) && ids_of(&in_deleted).is_empty() && ids_of(&in_drafts).is_empty(),
        format!("{d2:?}; in Deleted Items {}; in Drafts {}", ids_of(&in_deleted).len(), ids_of(&in_drafts).len()));

    // ---- regression smoke (X) ---------------------------------------------------
    let f = w.call("list_folders", json!({})).await;
    let e = w.call("list_emails", json!({"count": 5})).await;
    let a = w.call("check_availability", json!({"people": [SELF], "start": "2099-04-13T08:00", "end": "2099-04-13T18:00"})).await;
    let folder_ok = f.as_ref().is_ok_and(|v| v.as_array().is_some_and(|x| x.iter().any(|r| s(r, "name") == "Inbox")));
    let emails_ok = e.as_ref().is_ok_and(|v| v.as_array().is_some_and(|x| x.len() <= 5 && x.iter().all(|r| r.get("received").is_some())));
    log.check("X1", folder_ok && emails_ok && a.is_ok(), format!("folders {folder_ok} emails {emails_ok} availability {:?}", a.err()));
    // The latest ids of the events updated above (#50: writes return the current id).
    made.events.extend([ea, eb].into_iter().flatten());
}

/// (follow-up flagged, high importance) of the W4 copy in Inbox, via the
/// `flag` / `importance` filters.
async fn w4_state(w: &mut Wire, run: &str, id: &str) -> (bool, bool) {
    let q = format!("\"{run} W4\"");
    let has = |v: Result<Value, String>| v.is_ok_and(|v| ids_of(&v).iter().any(|x| x == id));
    let flagged = has(w.call("list_emails", json!({"query": q, "flag": "follow_up"})).await);
    let high = has(w.call("list_emails", json!({"query": q, "importance": "high"})).await);
    (flagged, high)
}

/// Deletes everything the run created: the tracked ids first, then a tag
/// sweep over every folder, repeated until a pass finds nothing. Permanent
/// throughout, so nothing is left in Deleted Items.
async fn cleanup(w: &mut Wire, log: &mut Log, made: &Created, run: &str) {
    let mut errors = Vec::new();
    for id in &made.emails {
        let _ = w.call("delete_email", json!({"email_id": id, "permanent": true})).await;
    }
    for id in &made.events {
        let _ = w.call("delete_event", json!({"event_id": id, "permanent": true, "notify": false})).await;
    }
    for id in &made.tasks {
        let _ = w.call("delete_task", json!({"task_id": id, "permanent": true})).await;
    }
    for id in &made.notes {
        let _ = w.call("delete_note", json!({"note_id": id, "permanent": true})).await;
    }
    let mut swept = 0;
    for pass in 0..5 {
        let mut found = 0;
        for folder in ["drafts", "inbox", "sent", "deleted"] {
            let rows = w.call("list_emails", json!({"folder": folder, "query": run, "count": 200})).await.unwrap_or_default();
            for id in ids_of(&rows) {
                found += 1;
                if let Err(e) = w.call("delete_email", json!({"email_id": id, "permanent": true})).await { errors.push(e); }
            }
        }
        let rows = w.call("list_events", json!({"query": run, "start_after": "2099-01-01", "start_before": "2099-12-31"})).await.unwrap_or_default();
        for r in rows.as_array().cloned().unwrap_or_default() {
            found += 1;
            if let Err(e) = w.call("delete_event", json!({"event_id": s(&r, "id"), "permanent": true, "notify": false})).await { errors.push(e); }
        }
        let rows = w.call("list_tasks", json!({"query": run, "include_completed": true})).await.unwrap_or_default();
        for id in ids_of(&rows) {
            found += 1;
            if let Err(e) = w.call("delete_task", json!({"task_id": id, "permanent": true})).await { errors.push(e); }
        }
        // Notes are titled by their first line; W2's starts with Hebrew, so also match the file text.
        let mut notes = ids_of(&w.call("list_notes", json!({"query": run})).await.unwrap_or_default());
        notes.extend(ids_of(&w.call("list_notes", json!({"query": format!("\"{HE} מקובץ\""), "created_after": "-3h"})).await.unwrap_or_default()));
        for id in notes {
            found += 1;
            if let Err(e) = w.call("delete_note", json!({"note_id": id, "permanent": true})).await { errors.push(e); }
        }
        println!("cleanup sweep pass {pass}: {found} found");
        swept += found;
        if found == 0 {
            break;
        }
    }
    log.check("CLEAN", errors.is_empty(), format!("{swept} swept after id deletes; errors {errors:?}"));
}
