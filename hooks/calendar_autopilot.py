#!/usr/bin/env python3
"""Claude Code PreToolUse hook: "calendar autopilot" for outlook-mcp-rs (issue #48).

Lets a calendar-organizing session run without prompts for everything that
only touches the user's own calendar, and keeps a prompt for anything that
reaches other people:

  allow  reads (list_* / get_* / check_availability)
  allow  respond_to_meeting with notify: false (answer silently, no email),
         for one invite or a list
  allow  update_event with require_no_attendees: true and no attendees added,
         including every entry of an `updates` batch (the server refuses to
         touch a meeting, so no attendee is affected)
  ask    everything else: every tool that can reach other people (send_email,
         reply_email, create_draft, update_draft, create_event, delete_event,
         a notifying respond_to_meeting, update_event without the guard),
         every write to your own mailbox, and any tool it doesn't know. It
         fails closed: an unreadable call is "ask", never "allow".

Claude Code permission rules can only match a tool's name, and "accept
edits" mode only covers file edits, so the split needs a hook that reads the
call's arguments. Install it in .claude/settings.json (see the README,
"Calendar autopilot"). Standard library only; reads the hook event as JSON
on stdin and prints a permissionDecision.
"""

import json
import os
import sys

# Every outlook-mcp-rs tool is classified below; a test checks the lists
# against src/server.rs, so a new tool can't slip through unclassified.

# Tools that never change anything and never contact anyone.
READ_TOOLS = {
    "list_folders", "list_emails", "get_email", "list_attachments", "get_inline_image",
    "list_events", "get_event", "check_availability",
    "list_tasks", "get_task", "list_notes", "get_note",
}

# Tools that can reach other people: send mail, a draft, a meeting response,
# an invitation, an update or a cancellation. Always ask (with the reason),
# except the argument-checked cases in `decide`.
REACHES_OTHERS = {
    "send_email": "sends an email",
    "reply_email": "sends a reply (or saves a reply draft, which may be sent later)",
    "create_draft": "creates an email draft addressed to other people",
    "update_draft": "edits an email draft, and sends it if send is true",
    "create_event": "creates an event and can send meeting invitations",
    "update_event": "changes an event and can send meeting updates to attendees",
    "delete_event": "deletes an event and can send a cancellation to attendees",
    "respond_to_meeting": "emails your response to the organizer",
}

# Tools that only change your own mailbox, calendar, tasks, notes or local
# files. Not automatic either: approve them yourself.
OWN_WRITES = {
    "update_email": "changes an email (read state, flag, categories, folder)",
    "delete_email": "deletes an email",
    "empty_deleted_items": "permanently deletes everything in Deleted Items",
    "save_attachments": "writes attachment files to disk",
    "create_task": "creates a task", "update_task": "changes a task", "delete_task": "deletes a task",
    "create_note": "creates a note", "update_note": "changes a note", "delete_note": "deletes a note",
}


def decide(tool, args):
    """(decision, reason) for one outlook-mcp-rs tool call.

    `tool` is the bare tool name (e.g. "respond_to_meeting"); `args` its
    input. The decision is "allow" or "ask"; anything not explicitly safe
    asks, including tools this hook doesn't know.
    """
    if not isinstance(args, dict):
        args = {}
    if tool in READ_TOOLS:
        return "allow", "calendar autopilot: read-only tool"
    if tool == "respond_to_meeting" and notify_flag(args) is False:
        return "allow", "calendar autopilot: response recorded without emailing the organizer"
    if tool == "update_event":
        # A batch's entries are complete calls; each must carry the guard.
        updates = args.get("updates")
        entries = updates if isinstance(updates, list) and updates else [args]
        if all(guarded_update(e) for e in entries):
            return "allow", ("calendar autopilot: require_no_attendees is set, so the server "
                             "only changes personal appointments and notifies nobody")
        return "ask", (f"{tool} {REACHES_OTHERS[tool]} (require_no_attendees is not set on "
                       "every change, so it may change a meeting with attendees)")
    if tool in REACHES_OTHERS:
        return "ask", f"{tool} {REACHES_OTHERS[tool]}: approve it yourself"
    if tool in OWN_WRITES:
        return "ask", f"{tool} {OWN_WRITES[tool]}: approve it yourself"
    return "ask", f"{tool} is not known to the calendar autopilot hook: approve it yourself"


def notify_flag(args):
    """The calendar notify flag: `notify`, or its deprecated alias `send`."""
    for key in ("notify", "send"):
        if key in args:
            return args[key]
    return None


def guarded_update(args):
    """An update_event call (or `updates` entry) the server keeps away from
    meetings: require_no_attendees is true and no attendee is added."""
    if not isinstance(args, dict):
        return False
    adds = args.get("add_required_attendees") or args.get("add_optional_attendees")
    return args.get("require_no_attendees") is True and not adds


def bare_tool_name(tool_name):
    """"mcp__outlook__list_events" -> "list_events"; None if not an MCP tool."""
    if not tool_name.startswith("mcp__"):
        return None
    return tool_name.rsplit("__", 1)[-1]


def autopilot_enabled(env):
    """False only for the plugin with its calendar_autopilot option off.

    The plugin always registers this hook, and Claude Code exports each
    plugin option to hook processes as CLAUDE_PLUGIN_OPTION_<KEY>. A hook
    installed by hand in settings.json (no CLAUDE_PLUGIN_ROOT) is always on.
    """
    if "CLAUDE_PLUGIN_ROOT" not in env:
        return True
    value = env.get("CLAUDE_PLUGIN_OPTION_CALENDAR_AUTOPILOT", "")
    return value.strip().lower() in ("true", "1", "yes")


def main():
    if not autopilot_enabled(os.environ):
        return 0  # Mode off: the normal permission flow decides (it prompts).
    try:
        event = json.load(sys.stdin)
        tool = bare_tool_name(event.get("tool_name", ""))
        if tool is None:
            return 0  # Not an MCP tool: the normal permission flow decides.
        decision, reason = decide(tool, event.get("tool_input"))
    except Exception as e:  # noqa: BLE001 - fail closed: never allow on a surprise
        decision, reason = "ask", f"calendar autopilot hook could not read the call ({e})"
    json.dump({
        "hookSpecificOutput": {
            "hookEventName": "PreToolUse",
            "permissionDecision": decision,
            "permissionDecisionReason": reason,
        }
    }, sys.stdout)
    return 0


if __name__ == "__main__":
    sys.exit(main())
