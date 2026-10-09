//! System test for open PRs #14-#25 merged together, executed against the
//! real, running Outlook mailbox per `SYSTEM_TEST_PLAN_2026-10-03-PRS.md`.
//! NOT run by plain `cargo test` - `#[ignore]`d. Run explicitly:
//!   cargo test --test system_test_prs -- --ignored --nocapture
//!
//! Every step records PASS/FAIL into a running log instead of panicking, so
//! a single failure never skips cleanup or later steps. Cleanup runs
//! unconditionally at the end; the final line is the only real `assert!`.

use base64::Engine as _;
use outlook_mcp_rs::outlook::client::WindowsOutlookClient;
use outlook_mcp_rs::outlook::types::EmailSummary;
use outlook_mcp_rs::outlook::{DraftUpdate, EmailQuery, InlineImage, MailBody, NewEmail, OutlookClient, ReplyInput};
use std::collections::HashSet;
use std::time::Duration;

const TAG: &str = "[outlook-mcp-rs systest PRS]";
const SELF_ADDR: &str = "adamkopelman@outlook.com";
const NOBODY: &str = "nobody@example.invalid";
const SOMEONE_ELSE: &str = "someone-else@example.invalid";
/// What `get_email` appends to a body cut at `max_body_chars = 1000`.
const TRUNC_MARKER: &str = "\n\n[... truncated at 1000 characters]";
/// A valid 1x1 PNG.
const PNG_B64: &str =
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

fn client() -> WindowsOutlookClient {
    WindowsOutlookClient::new()
}

fn png_bytes() -> Vec<u8> {
    base64::engine::general_purpose::STANDARD.decode(PNG_B64).unwrap()
}

struct Results {
    entries: Vec<(String, bool, String)>,
}
impl Results {
    fn new() -> Self {
        Self { entries: Vec::new() }
    }
    fn record(&mut self, id: &str, pass: bool, note: impl Into<String>) {
        let note = note.into();
        println!("[{id}] {} - {note}", if pass { "PASS" } else { "FAIL" });
        self.entries.push((id.to_string(), pass, note));
    }
    fn fail(&mut self, id: &str, note: impl Into<String>) {
        self.record(id, false, note);
    }
    fn print_summary(&self) {
        println!("\n=== SUMMARY ===");
        for (id, pass, note) in &self.entries {
            println!("{:<6} {:<4} {}", id, if *pass { "PASS" } else { "FAIL" }, note);
        }
        let pass = self.entries.iter().filter(|e| e.1).count();
        println!("\n{pass}/{} passed", self.entries.len());
    }
    fn all_passed(&self) -> bool {
        self.entries.iter().all(|e| e.1)
    }
}

fn q(folder: &str) -> EmailQuery {
    EmailQuery {
        query: None, folder: folder.to_string(), count: 50, offset: 0, unread_only: false,
        from: None, to: None, category: None, received_after: None, received_before: None,
        since_days: None, has_attachments: None, flagged: false, high_importance: false,
    }
}

fn tagged(items: &[EmailSummary]) -> Vec<&EmailSummary> {
    items.iter().filter(|e| e.subject.contains(TAG)).collect()
}

fn ids(items: &[EmailSummary]) -> HashSet<String> {
    tagged(items).into_iter().map(|e| e.id.clone()).collect()
}

/// Polls `folder` for a tagged item whose subject contains `needle` (~90s).
fn wait_for(c: &WindowsOutlookClient, folder: &str, needle: &str) -> Option<EmailSummary> {
    for attempt in 0..30 {
        if attempt > 0 {
            std::thread::sleep(Duration::from_secs(3));
        }
        if let Ok(list) = c.list_emails(EmailQuery { query: Some(needle.to_string()), ..q(folder) })
            && let Some(e) = list.into_iter().find(|e| e.subject.contains(needle))
        {
            return Some(e);
        }
    }
    None
}

fn draft(
    c: &WindowsOutlookClient,
    to: &str,
    subject: &str,
    body: &str,
    html: bool,
    attachments: Option<Vec<String>>,
    inline: Option<Vec<InlineImage>>,
) -> Result<String, String> {
    let v = c
        .create_draft(NewEmail {
            to: vec![to.to_string()], subject: subject.to_string(),
            body: if html { MailBody::Html(body.to_string()) } else { MailBody::Text(body.to_string()) },
            attachments, inline_images: inline, ..Default::default()
        })
        .map_err(|e| e.0)?;
    v["id"].as_str().map(str::to_string).ok_or_else(|| format!("no id in {v}"))
}

fn b64_image(cid: &str) -> InlineImage {
    InlineImage { content_id: cid.into(), data_base64: Some(PNG_B64.into()), ..Default::default() }
}

fn deleted_items_count(c: &WindowsOutlookClient) -> Option<i32> {
    c.list_folders().ok()?.into_iter().find(|f| f.name.eq_ignore_ascii_case("deleted items")).map(|f| f.items)
}

fn decode_data_uri(uri: &str) -> Option<Vec<u8>> {
    let (_, payload) = uri.split_once(',')?;
    base64::engine::general_purpose::STANDARD.decode(payload).ok()
}

#[test]
#[ignore]
fn system_test_open_prs_14_to_25() {
    let c = client();
    let mut r = Results::new();
    let run = format!(
        "r{}",
        std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).unwrap().as_secs()
    );
    let subj = |s: &str| format!("{TAG} {run} {s}");
    let mut cleanup: Vec<(String, String)> = Vec::new(); // (label, id)
    let scratch = std::env::temp_dir().join(format!("outlook-mcp-rs-systest-prs-{run}"));
    let _ = std::fs::create_dir_all(&scratch);
    let png_path = scratch.join("logo.png");
    let txt_path = scratch.join("notes.txt");
    let txt2_path = scratch.join("extra.txt");
    let _ = std::fs::write(&png_path, png_bytes());
    let _ = std::fs::write(&txt_path, b"systest attachment payload\r\n");
    let _ = std::fs::write(&txt2_path, b"second attachment\r\n");
    let png_path_s = png_path.to_string_lossy().to_string();
    let txt_path_s = txt_path.to_string_lossy().to_string();
    let txt2_path_s = txt2_path.to_string_lossy().to_string();
    let deleted_before = deleted_items_count(&c);
    println!("run token {run}; Deleted Items before: {deleted_before:?}");

    // Per-run Hebrew tokens (digits make them unique; the query stays non-ASCII).
    let digits: String = run.chars().filter(|ch| ch.is_ascii_digit()).collect();
    let he_subject_tok = format!("כותרתמבחן{digits}");
    let he_body_tok = format!("גוףמבחן{digits}");
    let he_missing_tok = format!("לאקיים{digits}");

    // ================= Phase 0: seed =================
    println!("\n--- seed ---");
    let mut pg: Vec<String> = Vec::new();
    for n in 1..=5 {
        match draft(&c, NOBODY, &subj(&format!("page {n} pg{run}")), "paging seed", false, None, None) {
            Ok(id) => {
                cleanup.push((format!("PG{n}"), id.clone()));
                pg.push(id);
            }
            Err(e) => r.fail("SEED", format!("PG{n} create_draft: {e}")),
        }
        std::thread::sleep(Duration::from_millis(1100)); // distinct timestamps for ordering
    }
    let he_s = draft(&c, NOBODY, &subj(&format!("{he_subject_tok} שיקוף")), "גוף רגיל", false, None, None);
    let he_b = draft(&c, NOBODY, &subj(&format!("hebrew body {run}")), &format!("שורה ראשונה {he_body_tok} סוף"), false, None, None);
    let to_a = draft(&c, NOBODY, &subj(&format!("to-a {run}")), "to filter seed", false, None, None);
    let to_b = draft(&c, SOMEONE_ELSE, &subj(&format!("to-b {run}")), "to filter seed", false, None, None);
    for (label, res) in [("HE-S", &he_s), ("HE-B", &he_b), ("TO-A", &to_a), ("TO-B", &to_b)] {
        match res {
            Ok(id) => cleanup.push((label.into(), id.clone())),
            Err(e) => r.fail("SEED", format!("{label} create_draft: {e}")),
        }
    }
    let he_s = he_s.ok();
    let he_b = he_b.ok();
    let to_a = to_a.ok();
    let to_b = to_b.ok();

    // ================= PR #14: offset paging / cap =================
    println!("\n--- PR #14 ---");
    let page = |offset: i32, count: i32| -> Result<Vec<String>, String> {
        c.list_emails(EmailQuery { query: Some(format!("pg{run}")), count, offset, ..q("drafts") })
            .map(|l| tagged(&l).into_iter().map(|e| e.id.clone()).collect())
            .map_err(|e| e.0)
    };
    match (page(0, 2), page(2, 2), page(4, 2), page(0, 5)) {
        (Ok(a), Ok(b), Ok(cc), Ok(all)) => {
            let concat: Vec<String> = a.iter().chain(&b).chain(&cc).cloned().collect();
            let set: HashSet<&String> = concat.iter().collect();
            let want: HashSet<&String> = pg.iter().collect();
            let pass = a.len() == 2 && b.len() == 2 && cc.len() == 1 && set.len() == 5 && set == want && concat == all;
            r.record("P14-1", pass, format!("sizes {}/{}/{}, disjoint+complete {}, same order as count=5: {}",
                a.len(), b.len(), cc.len(), set == want && set.len() == 5, concat == all));
        }
        (a, b, cc, all) => r.fail("P14-1", format!("list failed: {:?} {:?} {:?} {:?}", a.err(), b.err(), cc.err(), all.err())),
    }
    match (page(10, 5), page(-3, 2), page(0, 2)) {
        (Ok(past), Ok(neg), Ok(zero)) => r.record("P14-2", past.is_empty() && neg == zero,
            format!("offset 10 -> {} items; offset -3 == offset 0: {}", past.len(), neg == zero)),
        (a, b, cc) => r.fail("P14-2", format!("{:?} {:?} {:?}", a.err(), b.err(), cc.err())),
    }
    match c.list_emails(EmailQuery { count: 500, ..q("deleted") }) {
        Ok(list) => {
            let folder_items = deleted_items_count(&c).unwrap_or(0);
            let pass = list.len() <= 200 && (folder_items <= 55 || list.len() > 50);
            r.record("P14-3", pass, format!("count=500 returned {} (folder holds {folder_items}; >50 required when folder >55)", list.len()));
        }
        Err(e) => r.fail("P14-3", e.0),
    }

    // ================= PR #15: non-ASCII search =================
    println!("\n--- PR #15 ---");
    let search = |query: &str, to: Option<&str>| -> Result<HashSet<String>, String> {
        c.list_emails(EmailQuery { query: Some(query.to_string()), to: to.map(str::to_string), ..q("drafts") })
            .map(|l| ids(&l))
            .map_err(|e| e.0)
    };
    match search(&he_subject_tok, None) {
        Ok(got) => r.record("P15-1", he_s.as_ref().is_some_and(|id| got == HashSet::from([id.clone()])),
            format!("Hebrew subject token -> {} tagged hit(s), HE-S found: {}", got.len(), he_s.as_ref().is_some_and(|id| got.contains(id)))),
        Err(e) => r.fail("P15-1", e),
    }
    match search(&he_body_tok, None) {
        Ok(got) => r.record("P15-2", he_b.as_ref().is_some_and(|id| got == HashSet::from([id.clone()])),
            format!("Hebrew body-only token -> {} tagged hit(s), HE-B found: {}", got.len(), he_b.as_ref().is_some_and(|id| got.contains(id)))),
        Err(e) => r.fail("P15-2", e),
    }
    match search(&he_missing_tok, None) {
        Ok(got) => r.record("P15-3", got.is_empty(), format!("missing Hebrew token -> {} tagged hit(s)", got.len())),
        Err(e) => r.fail("P15-3", e),
    }
    match search(&he_subject_tok, Some(NOBODY)) {
        Ok(got) => r.record("P15-4", he_s.as_ref().is_some_and(|id| got.contains(id)),
            format!("Hebrew query + to=nobody -> {} tagged hit(s)", got.len())),
        Err(e) => r.fail("P15-4", e),
    }

    // ================= PR #17: to filter =================
    println!("\n--- PR #17 ---");
    let to_search = |to: &str| -> Result<HashSet<String>, String> {
        c.list_emails(EmailQuery { query: Some("to-".to_string()), to: Some(to.to_string()), count: 200, ..q("drafts") })
            .map(|l| l.iter().filter(|e| e.subject.contains(&run)).map(|e| e.id.clone()).collect())
            .map_err(|e| e.0)
    };
    match to_search(SOMEONE_ELSE) {
        Ok(got) => r.record("P17-1", to_b.as_ref().is_some_and(|id| got == HashSet::from([id.clone()])),
            format!("to=someone-else -> {} run hit(s)", got.len())),
        Err(e) => r.fail("P17-1", e),
    }
    match to_search(NOBODY) {
        Ok(got) => r.record("P17-2",
            to_a.as_ref().is_some_and(|id| got.contains(id)) && to_b.as_ref().is_some_and(|id| !got.contains(id)),
            format!("to=nobody -> {} run hit(s); TO-A in, TO-B out", got.len())),
        Err(e) => r.fail("P17-2", e),
    }
    match to_search(&NOBODY.to_uppercase()) {
        Ok(got) => r.record("P17-3", to_a.as_ref().is_some_and(|id| got.contains(id)), format!("uppercase -> {} run hit(s)", got.len())),
        Err(e) => r.fail("P17-3", e),
    }
    match to_search(&format!("no-such-recipient-{run}")) {
        Ok(got) => r.record("P17-4", got.is_empty(), format!("unknown recipient -> {} run hit(s)", got.len())),
        Err(e) => r.fail("P17-4", e),
    }

    // ================= PR #18 + #23: attachment metadata / is_inline =================
    println!("\n--- PR #18 / #23 ---");
    let att_draft = draft(&c, NOBODY, &subj("attachments"),
        "<p>Logo below</p><img src=\"cid:logo\">", true,
        Some(vec![txt_path_s.clone()]), Some(vec![b64_image("logo")]));
    match &att_draft {
        Ok(id) => {
            cleanup.push(("ATT".into(), id.clone()));
            match c.list_attachments(id.clone()) {
                Ok(atts) => {
                    let txt = atts.iter().find(|a| a.filename == "notes.txt");
                    let logo = atts.iter().find(|a| a.content_id.as_deref() == Some("logo"));
                    match txt {
                        Some(a) => r.record("P18-1",
                            a.att_type == "file" && a.mime_type.as_deref() == Some("text/plain") && a.content_id.is_none() && !a.hidden && a.size > 0,
                            format!("txt: type={} mime={:?} cid={:?} hidden={} size={}", a.att_type, a.mime_type, a.content_id, a.hidden, a.size)),
                        None => r.fail("P18-1", format!("notes.txt not listed: {atts:?}")),
                    }
                    r.record("P23-1",
                        logo.is_some_and(|a| a.is_inline) && txt.is_some_and(|a| !a.is_inline),
                        format!("logo.is_inline={:?}, notes.txt.is_inline={:?}", logo.map(|a| a.is_inline), txt.map(|a| a.is_inline)));
                }
                Err(e) => {
                    r.fail("P18-1", e.0.clone());
                    r.fail("P23-1", e.0);
                }
            }
            let save_dir = scratch.join("saved");
            let _ = std::fs::create_dir_all(&save_dir);
            match c.save_attachments(id.clone(), save_dir.to_string_lossy().to_string(), None) {
                Ok(entries) => {
                    let txt = entries.iter().find(|v| v["filename"] == "notes.txt");
                    let bytes_ok = txt.and_then(|v| v["saved_to"].as_str()).and_then(|p| std::fs::read(p).ok())
                        == std::fs::read(&txt_path).ok();
                    let meta_ok = txt.is_some_and(|v| v["status"] == "saved" && v["type"] == "file" && v["mime_type"] == "text/plain" && v["hidden"] == false);
                    r.record("P18-2", meta_ok && bytes_ok, format!("txt entry meta ok {meta_ok}, bytes equal {bytes_ok}"));
                    let logo = entries.iter().find(|v| v["content_id"] == "logo");
                    r.record("P23-2", logo.is_some_and(|v| v["is_inline"] == true) && txt.is_some_and(|v| v["is_inline"] == false),
                        format!("save entries is_inline: logo={:?} txt={:?}", logo.map(|v| v["is_inline"].clone()), txt.map(|v| v["is_inline"].clone())));
                }
                Err(e) => {
                    r.fail("P18-2", e.0.clone());
                    r.fail("P23-2", e.0);
                }
            }
        }
        Err(e) => {
            for id in ["P18-1", "P18-2", "P23-1", "P23-2"] {
                r.fail(id, format!("create_draft failed: {e}"));
            }
        }
    }

    // ================= PR #22: inline images on create_draft =================
    println!("\n--- PR #22 ---");
    let img_draft = draft(&c, NOBODY, &subj("inline images"),
        "<p>Line one</p><p>Line two</p><p>Line three</p><img src=\"cid:logo\"><p>after</p><img src=\"cid:chart\">", true, None,
        Some(vec![
            InlineImage { content_id: "logo".into(), path: Some(png_path_s.clone()), ..Default::default() },
            InlineImage { content_id: "<cid:chart>".into(), data_base64: Some(format!("data:image/png;base64,{PNG_B64}")), ..Default::default() },
            b64_image("unreferenced"),
        ]));
    match &img_draft {
        Ok(id) => {
            cleanup.push(("IMG".into(), id.clone()));
            match c.list_attachments(id.clone()) {
                Ok(atts) => {
                    let ok = ["logo", "chart"].iter().all(|cid| atts.iter().any(|a|
                        a.content_id.as_deref() == Some(cid) && a.mime_type.as_deref() == Some("image/png") && a.hidden));
                    r.record("P22-1", ok, format!("{} attachments: {:?}", atts.len(),
                        atts.iter().map(|a| (a.content_id.clone(), a.mime_type.clone(), a.hidden)).collect::<Vec<_>>()));
                }
                Err(e) => r.fail("P22-1", e.0),
            }
        }
        Err(e) => r.fail("P22-1", format!("create_draft failed: {e}")),
    }
    let not_created = |label: &str| -> bool {
        c.list_emails(EmailQuery { query: Some(label.to_string()), ..q("drafts") })
            .map(|l| l.iter().all(|e| !e.subject.contains(label)))
            .unwrap_or(false)
    };
    let bad_plain = subj("bad plain inline");
    let res = draft(&c, NOBODY, &bad_plain, "plain", false, None, Some(vec![b64_image("x")]));
    if let Ok(id) = &res { cleanup.push(("BAD1".into(), id.clone())); }
    r.record("P22-2", res.is_err() && not_created(&bad_plain), format!("html=false with inline_images -> {:?}", res));
    let bad_dup = subj("bad dup inline");
    let res = draft(&c, NOBODY, &bad_dup, "<img src=\"cid:a\">", true, None, Some(vec![b64_image("a"), b64_image("A")]));
    if let Ok(id) = &res { cleanup.push(("BAD2".into(), id.clone())); }
    r.record("P22-3", res.is_err() && not_created(&bad_dup), format!("duplicate content_id -> {:?}", res));

    // ================= PR #24 / #25: get_inline_image + context =================
    println!("\n--- PR #24 / #25 ---");
    if let Ok(id) = &img_draft {
        let want = png_bytes();
        let a = c.get_inline_image(id.clone(), "cid:LOGO".into(), None);
        let b = c.get_inline_image(id.clone(), "<chart>".into(), None);
        match (&a, &b) {
            (Ok(a), Ok(b)) => {
                let ok = a.content_id == "logo" && a.mime_type == "image/png" && decode_data_uri(&a.data_uri) == Some(want.clone())
                    && a.size == want.len() && b.content_id == "chart" && decode_data_uri(&b.data_uri) == Some(want.clone());
                r.record("P24-1", ok, format!("logo: cid={} mime={} size={} bytes equal {}; chart bytes equal {}",
                    a.content_id, a.mime_type, a.size, decode_data_uri(&a.data_uri) == Some(want.clone()), decode_data_uri(&b.data_uri) == Some(want.clone())));
                r.record("P25-2", a.context.is_none(), format!("no context_lines -> context {:?}", a.context));
            }
            _ => {
                r.fail("P24-1", format!("{:?} / {:?}", a.as_ref().err(), b.as_ref().err()));
                r.fail("P25-2", "get_inline_image failed");
            }
        }
        match c.get_inline_image(id.clone(), "missing@nowhere".into(), None) {
            Err(e) => r.record("P24-2", e.0.contains("logo") && e.0.contains("chart"), format!("error: {}", e.0)),
            Ok(_) => r.fail("P24-2", "unknown cid unexpectedly succeeded"),
        }
        match c.get_inline_image(id.clone(), "logo".into(), Some(2)) {
            Ok(img) => r.record("P25-1", img.context.as_deref() == Some("Line two\nLine three"), format!("context {:?}", img.context)),
            Err(e) => r.fail("P25-1", e.0),
        }
        match c.get_inline_image(id.clone(), "unreferenced".into(), Some(3)) {
            Ok(img) => r.record("P25-3", img.context.as_deref() == Some(""), format!("context {:?}", img.context)),
            Err(e) => r.fail("P25-3", e.0),
        }
    } else {
        for id in ["P24-1", "P24-2", "P25-1", "P25-2", "P25-3"] {
            r.fail(id, "no inline-image draft");
        }
    }
    match &to_a {
        Some(id) => match c.get_inline_image(id.clone(), "logo".into(), None) {
            Err(e) => r.record("P24-3", e.0.contains("list_attachments"), format!("error: {}", e.0)),
            Ok(_) => r.fail("P24-3", "draft without attachments unexpectedly returned an image"),
        },
        None => r.fail("P24-3", "no TO-A draft"),
    }

    // ================= PR #20: truncation =================
    println!("\n--- PR #20 ---");
    let long_body: String = "0123456789".repeat(300);
    match draft(&c, NOBODY, &subj("truncation plain"), &long_body, false, None, None) {
        Ok(id) => {
            cleanup.push(("TRUNC1".into(), id.clone()));
            match (c.get_email(id.clone(), false, None), c.get_email(id.clone(), false, Some(1000))) {
                (Ok(full), Ok(cut)) => {
                    // The cut body is the first 1000 chars plus the long-standing
                    // truncation marker.
                    let head: String = full.body.chars().take(1000).collect();
                    let marker_ok = cut.body.strip_prefix(head.as_str()) == Some(TRUNC_MARKER);
                    let pass = !full.body_truncated && full.body_length >= 3000 && cut.body_truncated
                        && marker_ok && cut.body_length == full.body_length;
                    r.record("P20-1", pass, format!("full: truncated={} len={}; cut: truncated={} first-1000+marker {marker_ok} len={}",
                        full.body_truncated, full.body_length, cut.body_truncated, cut.body_length));
                }
                (a, b) => r.fail("P20-1", format!("{:?} / {:?}", a.err(), b.err())),
            }
        }
        Err(e) => r.fail("P20-1", e),
    }
    let long_html = format!("<html><body><p>{}</p></body></html>", "abcdefghij".repeat(300));
    match draft(&c, NOBODY, &subj("truncation html"), &long_html, true, None, None) {
        Ok(id) => {
            cleanup.push(("TRUNC2".into(), id.clone()));
            match c.get_email(id.clone(), true, Some(1000)) {
                Ok(d) => {
                    let marker_ok = d.html_body.as_deref().is_some_and(|h| {
                        h.chars().count() == 1000 + TRUNC_MARKER.chars().count() && h.ends_with(TRUNC_MARKER)
                    });
                    let pass = d.html_truncated == Some(true) && d.html_length.is_some_and(|n| n > 1000) && marker_ok;
                    r.record("P20-2", pass, format!("html_truncated={:?} html_length={:?} first-1000+marker {marker_ok}", d.html_truncated, d.html_length));
                }
                Err(e) => r.fail("P20-2", e.0),
            }
        }
        Err(e) => r.fail("P20-2", e),
    }

    // ================= PR #21: Hebrew round-trip =================
    println!("\n--- PR #21 ---");
    let he_subject = subj("מייל שיקוף — סיכום עשייה 📧");
    let he_line = "סיכום עשייה — שורה ראשונה";
    match draft(&c, NOBODY, &he_subject, &format!("{he_line}\nשורה שנייה"), false, None, None) {
        Ok(id) => {
            cleanup.push(("HE-RT".into(), id.clone()));
            let detail = c.get_email(id.clone(), false, None);
            let listed = c.list_emails(EmailQuery { query: Some(run.clone()), count: 200, ..q("drafts") })
                .ok().and_then(|l| l.into_iter().find(|e| e.id == id));
            match detail {
                Ok(d) => {
                    let pass = d.summary.subject == he_subject && d.body.contains(he_line)
                        && listed.as_ref().is_some_and(|e| e.subject == he_subject);
                    r.record("P21-1", pass, format!("get_email subject exact {}, body has line {}, list subject exact {}",
                        d.summary.subject == he_subject, d.body.contains(he_line), listed.as_ref().is_some_and(|e| e.subject == he_subject)));
                }
                Err(e) => r.fail("P21-1", e.0),
            }
            let new_subject = subj("עדכון כותרת בעברית ✓");
            let res = c.update_draft(DraftUpdate { email_id: id.clone(), subject: Some(new_subject.clone()), ..Default::default() });
            let back = c.get_email(id.clone(), false, None).map(|d| d.summary.subject);
            r.record("P21-2", res.is_ok() && back.as_deref().ok() == Some(new_subject.as_str()),
                format!("update_draft {:?}; read back {:?}", res.map(|v| v["status"].clone()).map_err(|e| e.0), back.map_err(|e| e.0)));
        }
        Err(e) => {
            r.fail("P21-1", e.clone());
            r.fail("P21-2", e);
        }
    }

    // ================= PR #19: update_draft =================
    println!("\n--- PR #19 ---");
    match draft(&c, NOBODY, &subj("update me"), "original", false, Some(vec![txt_path_s.clone()]), None) {
        Ok(id) => {
            cleanup.push(("UPD".into(), id.clone()));
            let new_subject = subj("updated");
            let res = c.update_draft(DraftUpdate {
                email_id: id.clone(),
                subject: Some(new_subject.clone()),
                html_body: Some(format!("<p>marker-{run}</p>")),
                to: Some(vec![SOMEONE_ELSE.into(), NOBODY.into()]),
                cc: Some(vec![NOBODY.into()]),
                ..Default::default()
            });
            match res {
                Ok(v) => {
                    let changed_ok = v["status"] == "draft_updated" && v["changed"] == serde_json::json!(["subject", "html_body", "to", "cc"]);
                    let d = c.get_email(id.clone(), true, None);
                    match d {
                        Ok(d) => {
                            let html_ok = d.html_body.as_deref().is_some_and(|h| h.contains(&format!("marker-{run}")));
                            let to_ok = d.summary.to.contains("someone-else") && d.summary.to.contains("nobody");
                            let pass = changed_ok && d.summary.subject == new_subject && html_ok && to_ok && d.cc.contains("nobody");
                            r.record("P19-1", pass, format!("changed {} ; subject ok {} ; html ok {html_ok} ; to {:?} ; cc {:?}",
                                v["changed"], d.summary.subject == new_subject, d.summary.to, d.cc));
                        }
                        Err(e) => r.fail("P19-1", e.0),
                    }
                }
                Err(e) => r.fail("P19-1", e.0),
            }
            let res = c.update_draft(DraftUpdate { email_id: id.clone(), cc: Some(vec![]), bcc: Some(vec![SOMEONE_ELSE.into()]), ..Default::default() });
            match (res, c.get_email(id.clone(), false, None)) {
                (Ok(_), Ok(d)) => r.record("P19-2", d.cc.trim().is_empty() && d.bcc.contains("someone-else"), format!("cc {:?} bcc {:?}", d.cc, d.bcc)),
                (a, b) => r.fail("P19-2", format!("{:?} / {:?}", a.err().map(|e| e.0), b.err().map(|e| e.0))),
            }
            let res = c.update_draft(DraftUpdate { email_id: id.clone(), attachments: Some(vec![txt2_path_s.clone()]), ..Default::default() });
            match (res, c.list_attachments(id.clone())) {
                (Ok(_), Ok(atts)) => {
                    let names: Vec<&str> = atts.iter().map(|a| a.filename.as_str()).collect();
                    r.record("P19-3", atts.len() == 2 && names.contains(&"notes.txt") && names.contains(&"extra.txt"), format!("attachments {names:?}"));
                }
                (a, b) => r.fail("P19-3", format!("{:?} / {:?}", a.err().map(|e| e.0), b.err().map(|e| e.0))),
            }
            let res = c.update_draft(DraftUpdate { email_id: id.clone(), body: Some("x".into()), html_body: Some("<p>x</p>".into()), subject: Some(subj("SHOULD NOT APPLY")), ..Default::default() });
            let still = c.get_email(id.clone(), false, None).map(|d| d.summary.subject);
            r.record("P19-4", res.is_err() && still.as_deref().ok() == Some(new_subject.as_str()),
                format!("body+html_body -> {:?}; subject now {:?}", res.err().map(|e| e.0), still.map_err(|e| e.0)));
        }
        Err(e) => {
            for id in ["P19-1", "P19-2", "P19-3", "P19-4"] {
                r.fail(id, e.clone());
            }
        }
    }

    // ================= PR #16: permanent delete / empty_deleted_items =================
    println!("\n--- PR #16 ---");
    let in_deleted = |needle: &str| -> Option<EmailSummary> {
        c.list_emails(EmailQuery { query: Some(needle.to_string()), count: 200, ..q("deleted") })
            .ok()?.into_iter().find(|e| e.subject.contains(needle))
    };
    let perm_subject = subj(&format!("perm delete {run}"));
    match draft(&c, NOBODY, &perm_subject, "hard delete me", false, None, None) {
        Ok(id) => match c.delete_email(id.clone(), true) {
            Ok(v) => {
                std::thread::sleep(Duration::from_secs(2));
                let gone = c.get_email(id.clone(), false, None).is_err();
                let not_in_deleted = in_deleted(&perm_subject).is_none();
                r.record("P16-1", v["status"] == "deleted" && v["permanent"] == true && gone && not_in_deleted,
                    format!("result {v}; id resolves no more {gone}; absent from Deleted Items {not_in_deleted}"));
            }
            Err(e) => {
                cleanup.push(("PERM".into(), id));
                r.fail("P16-1", e.0);
            }
        },
        Err(e) => r.fail("P16-1", e),
    }
    let soft_subject = subj(&format!("soft delete {run}"));
    match draft(&c, NOBODY, &soft_subject, "soft delete me", false, None, None) {
        Ok(id) => match c.delete_email(id.clone(), false) {
            Ok(v) => {
                std::thread::sleep(Duration::from_secs(2));
                match in_deleted(&soft_subject) {
                    Some(moved) => {
                        let hard = c.delete_email(moved.id.clone(), true);
                        std::thread::sleep(Duration::from_secs(2));
                        let gone = in_deleted(&soft_subject).is_none();
                        if !gone { cleanup.push(("SOFT".into(), moved.id.clone())); }
                        r.record("P16-2", v["permanent"] == false && hard.as_ref().is_ok_and(|h| h["permanent"] == true) && gone,
                            format!("soft {v}; in Deleted Items yes; hard-delete from there {:?}; gone {gone}", hard.map_err(|e| e.0)));
                    }
                    None => r.fail("P16-2", format!("soft delete {v} but item not found in Deleted Items")),
                }
            }
            Err(e) => {
                cleanup.push(("SOFT".into(), id));
                r.fail("P16-2", e.0);
            }
        },
        Err(e) => r.fail("P16-2", e),
    }
    let before = deleted_items_count(&c);
    match c.empty_deleted_items(false) {
        Err(e) => {
            let after = deleted_items_count(&c);
            r.record("P16-3", e.0.contains("confirm") && before == after && before.is_some(),
                format!("refused: {} ; Deleted Items count {before:?} -> {after:?}", e.0));
        }
        Ok(v) => r.fail("P16-3", format!("empty_deleted_items(false) DID NOT REFUSE: {v}")),
    }

    // ================= S1: self-loop send with inline image (+ P22-4, P24-4, P19-5, R3) =================
    println!("\n--- S1 self-loop send ---");
    let s1_subject = subj("self-loop inline send");
    let mut s1_inbox: Option<EmailSummary> = None;
    match c.send_email(NewEmail {
        to: vec![SELF_ADDR.into()], subject: s1_subject.clone(),
        body: MailBody::Html("<p>Automated system test (PRs #14-#25), sent to self.</p><img src=\"cid:s1img\">".into()),
        inline_images: Some(vec![b64_image("s1img")]), ..Default::default()
    }) {
        Ok(v) => {
            r.record("S1", v["status"] == "sent", format!("send_email -> {v}"));
            s1_inbox = wait_for(&c, "inbox", &s1_subject);
            if let Some(e) = &s1_inbox {
                cleanup.push(("S1-inbox".into(), e.id.clone()));
            } else {
                r.fail("S1", "self-loop copy did not arrive in Inbox within ~90s");
            }
            if let Some(sent) = wait_for(&c, "sent", &s1_subject) {
                cleanup.push(("S1-sent".into(), sent.id));
            }
        }
        Err(e) => r.fail("S1", e.0),
    }
    if let Some(recv) = &s1_inbox {
        match c.list_attachments(recv.id.clone()) {
            Ok(atts) => {
                let a = atts.iter().find(|a| a.content_id.as_deref() == Some("s1img"));
                r.record("P22-4", a.is_some_and(|a| a.is_inline), format!("received attachments {:?}",
                    atts.iter().map(|a| (a.filename.clone(), a.content_id.clone(), a.is_inline)).collect::<Vec<_>>()));
            }
            Err(e) => r.fail("P22-4", e.0),
        }
        match c.get_inline_image(recv.id.clone(), "s1img".into(), Some(1)) {
            Ok(img) => r.record("P24-4", decode_data_uri(&img.data_uri) == Some(png_bytes()),
                format!("received image bytes equal sent: {}; context {:?}", decode_data_uri(&img.data_uri) == Some(png_bytes()), img.context)),
            Err(e) => r.fail("P24-4", e.0),
        }
        let res = c.update_draft(DraftUpdate { email_id: recv.id.clone(), subject: Some(subj("MUST NOT CHANGE")), ..Default::default() });
        let still = c.get_email(recv.id.clone(), false, None).map(|d| d.summary.subject);
        r.record("P19-5", res.as_ref().is_err_and(|e| e.0.contains("only unsent drafts")) && still.as_deref().ok() == Some(recv.subject.as_str()),
            format!("update_draft on received item -> {:?}; subject now {:?}", res.err().map(|e| e.0), still.map_err(|e| e.0)));
        match c.reply_email(ReplyInput {
            email_id: recv.id.clone(), body: MailBody::Text("systest reply draft".into()), ..Default::default()
        }) {
            Ok(v) => {
                if let Some(id) = v["id"].as_str() { cleanup.push(("R3".into(), id.to_string())); }
                r.record("R3", v["status"] == "draft_saved" && v["id"].is_string(), format!("reply_email send=false -> {}", v["status"]));
            }
            Err(e) => r.fail("R3", e.0),
        }
    } else {
        for id in ["P22-4", "P24-4", "P19-5", "R3"] {
            r.fail(id, "no received S1 copy");
        }
    }

    // ================= Regression smoke =================
    println!("\n--- regression smoke ---");
    match c.list_folders() {
        Ok(f) => {
            let ok = ["inbox", "drafts", "deleted items", "sent items"].iter().all(|w| f.iter().any(|x| x.name.eq_ignore_ascii_case(w)));
            r.record("R1", ok, format!("{} folders, core present {ok}", f.len()));
        }
        Err(e) => r.fail("R1", e.0),
    }
    match c.list_emails(EmailQuery { count: 5, ..q("inbox") }) {
        Ok(l) => r.record("R2", l.len() <= 5 && !l.is_empty(), format!("{} items", l.len())),
        Err(e) => r.fail("R2", e.0),
    }

    // ================= Cleanup (unconditional) =================
    println!("\n--- cleanup ---");
    let mut cleanup_failures = Vec::new();
    for (label, id) in &cleanup {
        if let Err(e) = c.delete_email(id.clone(), true) {
            // Already gone (e.g. P16) is fine; the sweep below catches anything else.
            println!("cleanup {label}: {}", e.0);
        }
    }
    for folder in ["drafts", "inbox", "sent", "deleted"] {
        for pass in 0..10 {
            let found: Vec<EmailSummary> = c
                .list_emails(EmailQuery { query: Some(run.clone()), count: 200, ..q(folder) })
                .map(|l| l.into_iter().filter(|e| e.subject.contains(TAG) && e.subject.contains(&run)).collect())
                .unwrap_or_default();
            if found.is_empty() {
                break;
            }
            println!("sweep {folder} pass {pass}: {} item(s)", found.len());
            for e in found {
                if let Err(err) = c.delete_email(e.id.clone(), true) {
                    cleanup_failures.push(format!("{folder}: {} ({})", e.subject, err.0));
                }
            }
            std::thread::sleep(Duration::from_secs(2));
        }
    }
    let _ = std::fs::remove_dir_all(&scratch);
    let deleted_after = deleted_items_count(&c);
    r.record("CLEAN", cleanup_failures.is_empty() && deleted_after == deleted_before,
        format!("failures {cleanup_failures:?}; Deleted Items {deleted_before:?} -> {deleted_after:?}"));

    r.print_summary();
    assert!(r.all_passed(), "system test had failures - see summary above");
}
