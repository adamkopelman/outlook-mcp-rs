"""Builds an HTML review page of every live system test: for each test id,
its input and expected output (from the SYSTEM_TEST_PLAN_*.md docs) and the
actual output and result (from a run log, i.e. the test's --nocapture output).

    python scripts/system_test_report.py <logs-dir> <out.html>

<logs-dir> holds one log per suite, named after the test target
(system_test.txt, system_test_big_calendar.txt, ...), plus offline.txt,
live_outlook.txt and hooks.txt for the summary.
"""
import html
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

SUITES = [
    {
        "target": "system_test_edge",
        "title": "Edge cases (2026-10-18)",
        "plan": "SYSTEM_TEST_PLAN_2026-10-18-EDGE.md",
        "covers": "Code-review hypotheses: clashing attachment names, commas in categories, DASL wildcards, all-day and DST boundaries, "
                  "odd availability intervals, recurrence edges, permanent deletes, paging under mutation, folders, attendees, limits, flags, stale ids.",
        "traffic": "One self-loop email and its reply-all; everything else drafts, personal appointments and unsent meetings.",
    },
    {
        "target": "system_test_big_calendar",
        "title": "Big calendar",
        "plan": "SYSTEM_TEST_PLAN_2026-10-17-BIG-CALENDAR.md",
        "covers": "list_events caps and paging over ~490 events and 500 occurrences, #42's date window, a #43-size week, "
                  "filters before paging, Hebrew queries/categories/series, 100-item batch writes with undo, exact free/busy.",
        "traffic": "Nothing sent (personal appointments, unsent meetings to example.invalid).",
    },
    {
        "target": "system_test_inbound",
        "title": "Inbound mail and meeting invites",
        "plan": "SYSTEM_TEST_PLAN_2026-10-17-INBOUND.md",
        "covers": "Real mail from Gmail and real Google Calendar invites: Hebrew, inline images, attachments, batch flags/categories, "
                  "replies and sends back to Gmail, respond_to_meeting (notify on/off, one occurrence, batch, via the request).",
        "traffic": "Gmail → Outlook (3 emails, 8 invites); Outlook → Gmail (reply, 2 emails, 1 tentative response). Needs the connector step.",
        "run": "SYSTEST_INBOUND_RUN=r<token> cargo test --test system_test_inbound -- --ignored --nocapture",
    },
    {
        "target": "system_test_prs_2026_10_17",
        "title": "PRs #37–#56 (2026-10-17)",
        "plan": "SYSTEM_TEST_PLAN_2026-10-17-PRS.md",
        "covers": "The merged PR stack through the tool layer: *_file inputs, data: images, update_draft send, batch reads, list conventions, "
                  "locale-proof dates, output controls, occurrence ids, batch writes + undo, require_no_attendees, annotations.",
        "traffic": "One self-loop email.",
    },
    {
        "target": "system_test_prs",
        "title": "PRs #14–#25 (2026-10-03)",
        "plan": "SYSTEM_TEST_PLAN_2026-10-03-PRS.md",
        "covers": "list_emails offset/to/non-ASCII search, get_email truncation, update_draft, permanent delete, attachment metadata, "
                  "inline images, get_inline_image, context_lines.",
        "traffic": "One self-loop email.",
    },
    {
        "target": "system_test",
        "title": "Plans 1–9 (email + calendar)",
        "plan": "SYSTEM_TEST_PLAN_2026-07-16.md",
        "covers": "Every email and calendar tool with seeded data: filters, send/reply/draft, update/move/delete, attachments, "
                  "events with filters, updates, attendees and deletes.",
        "traffic": "Self-loop emails, one email and one meeting invite + cancellation to the external Gmail address.",
    },
    {
        "target": "system_test_p10_12",
        "title": "Plans 10–12 (availability, tasks, notes)",
        "plan": "SYSTEM_TEST_PLAN_2026-07-16-P10-12.md",
        "covers": "check_availability, list/create/update/delete for tasks and notes.",
        "traffic": "Nothing sent.",
    },
]

BUGS = [
    ("save_attachments (open)", "Two attachments with the same safe file name overwrite each other, and both report \"saved\"",
     "`client.rs:2848` joins `safe_filename(name)` with no de-duplication: two `a.txt`, or `a_b.txt` and `.a_b.txt` "
     "(`safe_filename` trims the leading dot, `com.rs:25`), land on one path; the first file's bytes are lost.", "Edge AT1, AT2"),
    ("categories (open)", "A category name with a comma is silently split into two",
     "Outlook stores categories as one comma-separated string (raw COM: `\"Smith, John, Red Category\"`). `join_categories` "
     "(`com.rs:46`) writes the name as is and `parse_categories` (`com.rs:37`) splits it back, so the item gets `Smith` and `John`; "
     "the category filter and `remove_categories` then miss it. Nothing rejects the comma.", "Edge CT1, CT2, CT3"),
    ("check_availability (open)", "On a DST day every slot after the transition is labelled an hour off",
     "Raw `FreeBusy` counts real elapsed time from local midnight (a probe put the 10:00 event at index 18 on the 23-hour day). "
     "`parse_freebusy_slots` (`mod.rs:766`) labels slot i as midnight + i × interval in naive local time, so after spring-forward "
     "busy times show an hour early (after fall-back an hour late), and the day always gets 48 half-hour slots.", "Edge DST3a, DST3b"),
    ("create_event / list_events (open)", "An event in the skipped DST hour fails to save and leaves a broken item that breaks list_events",
     "Outlook converts 02:30 on the spring-forward day as 03:30, so End 03:00 is before Start and Save fails with \"The end date you "
     "entered occurs before the start date\". The unsaved item stays in the Calendar (later in Sync Issues/Local Failures, where only "
     "`Items.Remove` deletes it), and every `list_events` window covering that day fails with the same error. `create_event` "
     "(`client.rs:2424`) neither rejects a nonexistent local time nor discards the item on failure.", "Edge DST1a, DST1d"),
    ("create_event recurrence (open)", "A yearly series starting on Feb 29 can't be created",
     "Outlook already derives MonthOfYear 2 / DayOfMonth 29 from the start, then rejects any explicit write of them (zero-write probe). "
     "`apply_recurrence` always writes both (`client.rs:353-354`), so it fails with \"The recurrence pattern is not valid\".", "Edge RC3"),
    ("update_event undo (open)", "Undoing an all_day change doesn't restore the times",
     "The undo records only `all_day` (`client.rs:2581`). Outlook rewrites Start/End when AllDayEvent changes (clearing it gives "
     "08:00–08:30), so 10:00–11:00 → all-day → undo ends at 08:00–08:30.", "Edge AD4"),
    ("create_event all_day (open)", "all_day with start = end makes a zero-length all-day event",
     "`create_event` sets Start/End and then AllDayEvent (`client.rs:2424-2427`) and never extends the end; Outlook accepts a "
     "0-minute all-day event (probe: Duration 0), so the event ends at 07-01 00:00, not 07-02 00:00.", "Edge AD1"),
    ("get_event / get_task (open)", "An id of the wrong item type returns that item as the wrong type",
     "`read_event` (`client.rs:1369`) and `read_task` (`client.rs:1400`) open any id without checking the item class: a task id "
     "comes back as an event starting 1899-12-30, an event id as a not_started task.", "Edge ID3"),
    ("minor (open)", "Unclear or missing validation: interval_minutes 0, no way to clear a reminder, respond_to_meeting on a draft",
     "`check_availability` silently uses 1-minute slots for 0 (`client.rs:2744`); `reminder_minutes` always sets ReminderSet "
     "(`client.rs:2592`), so 0 means \"at start\", never \"none\"; `respond_to_meeting` calls Respond on a mail item and returns "
     "\"Unknown name\" (`client.rs:2509`).", "Edge AV1a, G4b, ID5"),
    ("#42 / list_events", "A one-day window dropped recurring occurrences on its first day",
     "Under IncludeRecurrences Outlook applies a date-only Restrict bound up to a day late to recurring occurrences; #42's "
     "window had no margin when the range started on an \"unambiguous\" date (day ≥ 13).", "Big calendar EXC, BC1, BC3, R1"),
    ("check_availability (Plan 10)", "Every slot shifted by the window's start time",
     "Recipient.FreeBusy's string always starts at 00:00 of the start date; it was read as starting at `start`.", "Big calendar A1"),
    ("check_availability", "\"Working elsewhere\" reported as busy",
     "The free/busy string encodes working elsewhere as '5', not OlBusyStatus 4.", "Big calendar A1"),
    ("#51 × #50 / batch writes", "Batch writes returned stale ids",
     "merge_id kept the input id over the result's current id; a batch accept returned ids that no longer open.", "Inbound M4"),
    ("live_outlook.rs", "update_email_applies_state_then_moves had never passed",
     "It read importance (not a field) and [\"summary\"] (flattened) from JSON.", "First full live run"),
    ("system tests' cleanup", "Old system tests left ~75 items per run; B8 left orphan invites in Google Calendar",
     "Soft deletes, no Sent Items sweep, meeting messages arriving after the sweep, a sent invite removed quietly.", "Re-running the old suites"),
]

ID_RE = r"[A-Z]{1,6}\d*[a-z]?(?:-\d+)?"

# Result lines for steps that plans describe in prose, not as a test id.
KNOWN_STEPS = {
    "seed": ("Create the plan's seed data (tagged, far-future dates)", "Every create succeeds"),
    "clean": ("Delete every created item permanently, then sweep every folder by the run's tag until a pass finds nothing",
              "Nothing of the run is left; no errors"),
    "cleanup": ("Delete every created item, then sweep the touched folders by the test's tag",
                "No leftovers; the tag sweeps come back clean"),
}


def md_inline(text):
    """Escape, then render `code` and **bold**."""
    out = html.escape(text)
    out = re.sub(r"`([^`]+)`", r"<code>\1</code>", out)
    out = re.sub(r"\*\*([^*]+)\*\*", r"<b>\1</b>", out)
    return out


def md_block(text):
    """Paragraphs, bullet lines and pipe tables to HTML."""
    lines = text.strip().splitlines()
    out, i = [], 0
    while i < len(lines):
        line = lines[i].rstrip()
        if line.lstrip().startswith("|"):
            rows = []
            while i < len(lines) and lines[i].lstrip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
                    rows.append(cells)
                i += 1
            if rows:
                head = "".join(f"<th>{md_inline(c)}</th>" for c in rows[0])
                body = "".join("<tr>" + "".join(f"<td dir=\"auto\">{md_inline(c)}</td>" for c in r) + "</tr>" for r in rows[1:])
                out.append(f"<table class=\"inner\"><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>")
            continue
        if not line.strip():
            i += 1
            continue
        para = [line.strip()]
        i += 1
        while i < len(lines) and lines[i].strip() and not lines[i].lstrip().startswith("|"):
            para.append(lines[i].strip())
            i += 1
        out.append("<p dir=\"auto\">" + "<br>".join(md_inline(p) for p in para) + "</p>")
    return "".join(out)


def split_io(text):
    """A bullet's text split into (input, expected) at its first arrow."""
    for sep in ("→", "->"):
        if sep in text:
            a, b = text.split(sep, 1)
            return a.strip(" :;,"), b.strip()
    m = re.search(r"\bExpect(?:ed)?:?\s", text)
    if m:
        return text[:m.start()].strip(), text[m.start():].strip()
    return text.strip(), ""


def expand_ids(label):
    """'S1–S8' -> S1..S8; 'A1' -> A1; 'P14-1' -> P14-1."""
    m = re.fullmatch(r"([A-Z]+)(\d+)[–-]\1?(\d+)", label)
    if m:
        return [f"{m.group(1)}{n}" for n in range(int(m.group(2)), int(m.group(3)) + 1)]
    return [label]


def parse_plan(path):
    """[(group, id_label, ids, title, input_html, expected_html)] in plan order."""
    text = (ROOT / path).read_text(encoding="utf-8")
    tests = text.split("\n## Tests", 1)[-1] if "\n## Tests" in text else text
    entries = []
    group = ""
    # Section style: "### A1. title" + body (07-16 plans).
    sections = re.split(r"\n(?=#{2,3} )", "\n" + tests)
    for sec in sections:
        sec = sec.strip("\n")
        if not sec:
            continue
        head, _, body = sec.partition("\n")
        level = len(head) - len(head.lstrip("#"))
        title = head.lstrip("#").strip()
        m = re.match(rf"({ID_RE}(?:[–-][A-Z]*\d+)?)[.:]\s*(.*)", title)
        bullets = re.findall(rf"^- \*\*({ID_RE})\*\*\s+(.*?)(?=^- \*\*|\Z)", body, re.M | re.S)
        if bullets:
            group = title
            for bid, btext in bullets:
                inp, exp = split_io(" ".join(l.strip() for l in btext.strip().splitlines()))
                entries.append((group, bid, [bid], "", md_block(inp), md_block(exp)))
        elif m and level == 3:
            label, name = m.group(1), m.group(2).strip()
            act = re.search(r"\*\*Action:\*\*(.*?)(?=\*\*Expected:\*\*|\Z)", body, re.S)
            exp = re.search(r"\*\*Expected:\*\*(.*)", body, re.S)
            if act or exp:
                inp_html = md_block(act.group(1)) if act else ""
                exp_html = md_block(exp.group(1)) if exp else ""
            else:
                inp, ex = split_io(body)
                inp_html, exp_html = md_block(inp), md_block(ex)
            entries.append((group, label, expand_ids(label), name, inp_html, exp_html))
        elif level == 2 or (level == 3 and not m):
            group = title
    return entries


LOG_PATTERNS = [
    re.compile(r"^\[(PASS|FAIL)\] (\S+?): (.*)$"),          # new: [PASS] ID: detail
    re.compile(r"^\[(\S+)\] (PASS|FAIL) - (.*)$"),          # old: [ID] PASS - detail
]


def short(detail):
    detail = re.sub(r"[0-9A-F]{40,}(?:\|[0-9A-F]{40,})?(?:\|\d{4}-\d\d-\d\dT[\d:]+)?", "‹id›", detail)
    return detail if len(detail) <= 700 else detail[:700] + " …"


def parse_log(path):
    results = []
    if not path.exists():
        return results, None
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        for i, pat in enumerate(LOG_PATTERNS):
            m = pat.match(line)
            if m:
                if i == 0:
                    status, rid, detail = m.groups()
                else:
                    rid, status, detail = m.groups()
                results.append((rid, status == "PASS", short(detail)))
                break
    summary = None
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        m = re.match(r"^(\d+) ?/ ?(\d+) passed", line.strip())
        if m:
            summary = (int(m.group(1)), int(m.group(2)))
    return results, summary


def owner(rid, entries):
    """The plan entry a log id belongs to: exact, else the longest id prefix."""
    best = None
    for idx, e in enumerate(entries):
        for pid in e[2]:
            if rid == pid:
                return idx
            if rid.startswith(pid) and (len(rid) == len(pid) or not rid[len(pid)].isdigit()):
                if best is None or len(pid) > best[1]:
                    best = (idx, len(pid))
    return best[0] if best else None


def build(logs_dir, out_path):
    logs_dir = Path(logs_dir)
    head = (logs_dir / "HEAD.txt").read_text(encoding="utf-8").strip() if (logs_dir / "HEAD.txt").exists() else ""
    suites_html, nav, overview = [], [], []
    total_pass = total = 0
    for s in SUITES:
        entries = parse_plan(s["plan"])
        results, summary = parse_log(logs_dir / f"{s['target']}.txt")
        by_entry = {i: [] for i in range(len(entries))}
        orphans = []
        for rid, ok, detail in results:
            o = owner(rid, entries)
            (by_entry[o] if o is not None else orphans).append((rid, ok, detail))
        n_ok = sum(1 for _, ok, _ in results if ok)
        total_pass += n_ok
        total += len(results)
        rows, last_group = [], None
        for i, (group, label, _ids, name, inp, exp) in enumerate(entries):
            if group != last_group:
                rows.append(f"<tr class=\"group\"><td colspan=\"5\" dir=\"auto\">{md_inline(group)}</td></tr>")
                last_group = group
            rs = by_entry[i]
            if rs:
                state = "fail" if any(not ok for _, ok, _ in rs) else "pass"
                actual = "".join(
                    f"<div class=\"step {'ok' if ok else 'bad'}\"><span class=\"sid\">{html.escape(rid)}</span> "
                    f"<span dir=\"auto\">{html.escape(d)}</span></div>" for rid, ok, d in rs)
                badge = f"<span class=\"badge {state}\">{'PASS' if state == 'pass' else 'FAIL'}</span>"
                if len(rs) > 1:
                    badge += f"<div class=\"sub\">{sum(1 for _, ok, _ in rs if ok)}/{len(rs)} steps</div>"
            else:
                state, actual = "none", "<span class=\"muted\">(no result line: seed/setup or covered by sub-steps)</span>"
                badge = "<span class=\"badge none\">—</span>"
            name_html = f"<div class=\"tname\" dir=\"auto\">{md_inline(name)}</div>" if name else ""
            rows.append(
                f"<tr class=\"t {state}\"><td class=\"id\">{html.escape(label)}</td>"
                f"<td class=\"io\">{name_html}{inp or '<span class=muted>—</span>'}</td>"
                f"<td class=\"io\">{exp or '<span class=muted>—</span>'}</td>"
                f"<td class=\"actual\">{actual}</td><td class=\"res\">{badge}</td></tr>")
        if orphans:
            rows.append("<tr class=\"group\"><td colspan=\"5\">Setup and cleanup (the plan's Seed data / Cleanup sections)</td></tr>")
            for rid, ok, d in orphans:
                state = "pass" if ok else "fail"
                inp, exp = KNOWN_STEPS.get(rid.lower(), ("—", "—"))
                rows.append(
                    f"<tr class=\"t {state}\"><td class=\"id\">{html.escape(rid)}</td><td class=\"io\">{md_inline(inp)}</td>"
                    f"<td class=\"io\">{md_inline(exp)}</td><td class=\"actual\"><span dir=\"auto\">{html.escape(d)}</span></td>"
                    f"<td class=\"res\"><span class=\"badge {state}\">{'PASS' if ok else 'FAIL'}</span></td></tr>")
        anchor = s["target"]
        run = s.get("run", f"cargo test --test {s['target']} -- --ignored --nocapture")
        score = f"{n_ok}/{len(results)}" if results else "not run"
        cls = "pass" if results and n_ok == len(results) else ("fail" if results else "none")
        nav.append(f"<a href=\"#{anchor}\" class=\"chip {cls}\">{html.escape(s['title'])} <b>{score}</b></a>")
        overview.append(
            f"<tr><td><a href=\"#{anchor}\">{html.escape(s['title'])}</a><div class=\"muted\"><code>tests/{anchor}.rs</code></div></td>"
            f"<td>{html.escape(s['covers'])}</td><td>{html.escape(s['traffic'])}</td>"
            f"<td class=\"res\"><span class=\"badge {cls}\">{score}</span></td></tr>")
        suites_html.append(
            f"<section id=\"{anchor}\"><h2>{html.escape(s['title'])} <span class=\"badge {cls}\">{score}</span></h2>"
            f"<p class=\"meta\">{html.escape(s['covers'])}</p>"
            f"<p class=\"meta\">Plan: <code>{html.escape(s['plan'])}</code> · Run: <code>{html.escape(run)}</code></p>"
            f"<div class=\"scroll\"><table class=\"tests\"><thead><tr><th>ID</th><th>Input</th><th>Expected output</th>"
            f"<th>Actual output (from the run log)</th><th>Result</th></tr></thead><tbody>{''.join(rows)}</tbody></table></div></section>")

    def first(path, pattern):
        p = logs_dir / path
        if not p.exists():
            return "not run"
        m = re.findall(pattern, p.read_text(encoding="utf-8", errors="replace"))
        return m[-1] if m else "?"
    live = first("live_outlook.txt", r"test result: \w+\. (\d+ passed; \d+ failed)")
    hooks = first("hooks.txt", r"(Ran \d+ tests?)") + ", " + first("hooks.txt", r"^(OK|FAILED.*)$")
    offline = "; ".join(re.findall(r"test result: ok\. (\d+ passed)", (logs_dir / "offline.txt").read_text(encoding="utf-8", errors="replace"))) if (logs_dir / "offline.txt").exists() else "not run"
    bugs = "".join(
        f"<tr><td>{md_inline(a)}</td><td dir=\"auto\"><b>{md_inline(b)}</b><div class=\"muted\">{md_inline(c)}</div></td><td>{md_inline(d)}</td></tr>"
        for a, b, c, d in BUGS)

    page = TEMPLATE.format(
        head=html.escape(head), total=f"{total_pass}/{total}", nav="".join(nav), overview="".join(overview),
        live=html.escape(live), hooks=html.escape(hooks), offline=html.escape(offline), bugs=bugs, suites="".join(suites_html))
    Path(out_path).write_text(page, encoding="utf-8")
    print(f"wrote {out_path}: {total_pass}/{total} result lines across {len(SUITES)} suites")


TEMPLATE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>System Test Review</title>
<style>
:root {{ --bg:#f7f7f5; --panel:#ffffff; --ink:#1d1d1b; --muted:#6b6b66; --line:#e3e2dc; --accent:#3557c9;
  --pass:#1f7a45; --pass-bg:#e5f4ea; --fail:#b42318; --fail-bg:#fde8e6; --none:#7a7a73; --none-bg:#efefea;
  --group:#f0efe9; --code:#f1f0ea; }}
@media (prefers-color-scheme: dark) {{ :root:not([data-theme="light"]) {{ --bg:#161614; --panel:#1f1f1c; --ink:#ecebe6;
  --muted:#9d9c95; --line:#34332e; --accent:#8aa4ff; --pass:#5fd08d; --pass-bg:#173523; --fail:#ff8b80; --fail-bg:#3d1b18;
  --none:#a3a29b; --none-bg:#2a2a26; --group:#262622; --code:#2a2a26; }} }}
:root[data-theme="dark"] {{ --bg:#161614; --panel:#1f1f1c; --ink:#ecebe6; --muted:#9d9c95; --line:#34332e; --accent:#8aa4ff;
  --pass:#5fd08d; --pass-bg:#173523; --fail:#ff8b80; --fail-bg:#3d1b18; --none:#a3a29b; --none-bg:#2a2a26; --group:#262622; --code:#2a2a26; }}
* {{ box-sizing:border-box; }}
body {{ margin:0; background:var(--bg); color:var(--ink); font:14px/1.5 system-ui,-apple-system,"Segoe UI",Arial,sans-serif; }}
header {{ padding:28px 16px 8px; max-width:1500px; margin:0 auto; }}
h1 {{ font-size:24px; margin:0 0 4px; }}
h2 {{ font-size:19px; margin:32px 0 6px; display:flex; gap:10px; align-items:center; flex-wrap:wrap; }}
main {{ max-width:1500px; margin:0 auto; padding:0 16px 60px; }}
.meta, .muted {{ color:var(--muted); }}
.meta {{ margin:2px 0; }}
code {{ font:12.5px ui-monospace,Consolas,monospace; background:var(--code); padding:1px 4px; border-radius:4px; word-break:break-word; }}
.chips {{ display:flex; flex-wrap:wrap; gap:8px; margin:14px 0; }}
.chip {{ text-decoration:none; color:var(--ink); background:var(--panel); border:1px solid var(--line); border-radius:999px; padding:5px 12px; }}
.chip.pass b {{ color:var(--pass); }} .chip.fail b {{ color:var(--fail); }}
.panel {{ background:var(--panel); border:1px solid var(--line); border-radius:10px; padding:4px 0; margin:12px 0; }}
.scroll {{ overflow-x:auto; background:var(--panel); border:1px solid var(--line); border-radius:10px; }}
table {{ border-collapse:collapse; width:100%; }}
th, td {{ text-align:start; vertical-align:top; padding:8px 10px; border-bottom:1px solid var(--line); }}
th {{ font-size:12px; text-transform:uppercase; letter-spacing:.04em; color:var(--muted); position:sticky; top:0; background:var(--panel); }}
table.tests td.id {{ font:600 12.5px ui-monospace,Consolas,monospace; white-space:nowrap; }}
table.tests td.io {{ min-width:220px; max-width:420px; }}
table.tests td.actual {{ min-width:260px; max-width:520px; font-size:12.5px; }}
td.res {{ white-space:nowrap; }}
td p {{ margin:0 0 6px; }}
tr.group td {{ background:var(--group); font-weight:600; font-size:13px; }}
.tname {{ font-weight:600; margin-bottom:4px; }}
.badge {{ display:inline-block; font:700 11.5px system-ui; padding:2px 8px; border-radius:999px; }}
.badge.pass {{ color:var(--pass); background:var(--pass-bg); }} .badge.fail {{ color:var(--fail); background:var(--fail-bg); }}
.badge.none {{ color:var(--none); background:var(--none-bg); }}
.sub {{ color:var(--muted); font-size:11.5px; margin-top:3px; }}
.step {{ padding:2px 0 2px 8px; border-inline-start:3px solid var(--line); margin:3px 0; word-break:break-word; }}
.step.ok {{ border-color:var(--pass); }} .step.bad {{ border-color:var(--fail); background:var(--fail-bg); }}
.sid {{ font:600 11.5px ui-monospace,Consolas,monospace; color:var(--muted); }}
table.inner {{ margin:4px 0 8px; font-size:12.5px; }} table.inner th, table.inner td {{ padding:3px 6px; position:static; }}
.controls {{ display:flex; gap:12px; flex-wrap:wrap; align-items:center; margin:8px 0 0; }}
input[type=search] {{ flex:1 1 260px; max-width:420px; padding:8px 10px; border:1px solid var(--line); border-radius:8px; background:var(--panel); color:var(--ink); font:inherit; }}
label {{ display:flex; gap:6px; align-items:center; color:var(--muted); }}
body.hide-actual td.actual, body.hide-actual th:nth-child(4) {{ display:none; }}
@media (max-width:700px) {{ table.tests td.io, table.tests td.actual {{ min-width:180px; }} }}
</style>
</head>
<body>
<header>
<h1>outlook-mcp-rs: system tests</h1>
<p class="meta">Every live system test with its input, expected output and the actual output of its latest run against the real Outlook mailbox. Code: <code>{head}</code></p>
<p class="meta">Result lines across all suites: <b>{total}</b> · <code>live_outlook.rs</code>: {live} · offline suite: {offline} · hook tests: {hooks}</p>
<div class="chips">{nav}</div>
<div class="controls"><input type="search" id="q" placeholder="Filter tests (id, tool, Hebrew text…)" aria-label="Filter tests">
<label><input type="checkbox" id="failOnly"> Failures only</label>
<label><input type="checkbox" id="hideActual"> Hide actual output</label></div>
</header>
<main>
<h2>Suites</h2>
<div class="scroll"><table><thead><tr><th>Suite</th><th>What it covers</th><th>Real traffic</th><th>Result</th></tr></thead><tbody>{overview}</tbody></table></div>
<h2>Bugs these tests found</h2>
<p class="meta">Entries marked <b>open</b> come from the 2026-10-18 edge suite and are not fixed yet; the rest (2026-10-17) are fixed.</p>
<div class="scroll"><table><thead><tr><th>Where</th><th>Bug</th><th>Found by</th></tr></thead><tbody>{bugs}</tbody></table></div>
{suites}
</main>
<script>
const q = document.getElementById('q'), failOnly = document.getElementById('failOnly'), hide = document.getElementById('hideActual');
function apply() {{
  const needle = q.value.trim().toLowerCase();
  document.querySelectorAll('table.tests tbody').forEach(tb => {{
    let groupRow = null, groupHas = false;
    const flush = () => {{ if (groupRow) groupRow.style.display = groupHas ? '' : 'none'; }};
    tb.querySelectorAll('tr').forEach(tr => {{
      if (tr.classList.contains('group')) {{ flush(); groupRow = tr; groupHas = false; return; }}
      const show = (!needle || tr.textContent.toLowerCase().includes(needle)) && (!failOnly.checked || tr.classList.contains('fail'));
      tr.style.display = show ? '' : 'none';
      if (show) groupHas = true;
    }});
    flush();
  }});
}}
q.addEventListener('input', apply); failOnly.addEventListener('change', apply);
hide.addEventListener('change', () => document.body.classList.toggle('hide-actual', hide.checked));
</script>
</body>
</html>
"""

if __name__ == "__main__":
    build(sys.argv[1], sys.argv[2])
