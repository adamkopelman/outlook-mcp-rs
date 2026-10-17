"""Tests for calendar_autopilot.py. Run: python -m unittest discover hooks"""

import json
import os
import re
import shutil
import subprocess
import sys
import unittest

sys.path.insert(0, os.path.dirname(__file__))
from calendar_autopilot import (  # noqa: E402
    OWN_WRITES, READ_TOOLS, REACHES_OTHERS, autopilot_enabled, bare_tool_name, decide,
)

SERVER_RS = os.path.join(os.path.dirname(__file__), os.pardir, "src", "server.rs")


def server_tool_attrs():
    """{tool name: its #[tool(...)] attribute} for every tool in src/server.rs."""
    with open(SERVER_RS, encoding="utf-8") as f:
        src = f.read()
    return dict((name, attr) for attr, name in
                re.findall(r"(#\[tool\(.*?\)\])\s*pub async fn (\w+)", src, re.S))


def server_tools():
    """Every tool name the server registers."""
    return set(server_tool_attrs())


def hinted(hint):
    """The tools whose MCP annotations set `hint = true`."""
    return {n for n, attr in server_tool_attrs().items() if f"{hint} = true" in attr}


HOOK = os.path.join(os.path.dirname(__file__), "calendar_autopilot.py")


class CoverageTest(unittest.TestCase):
    def test_every_server_tool_is_classified_exactly_once(self):
        tools = server_tools()
        self.assertGreaterEqual(len(tools), 30, "the regex must find every tool")
        lists = [READ_TOOLS, set(REACHES_OTHERS), set(OWN_WRITES)]
        for tool in tools:
            self.assertEqual(sum(tool in l for l in lists), 1, f"classify {tool} in calendar_autopilot.py")
        for l in lists:
            self.assertEqual(l - tools, set(), "a listed tool no longer exists")

    def test_tables_agree_with_the_servers_mcp_hints(self):
        self.assertEqual(hinted("read_only_hint"), READ_TOOLS)
        self.assertLessEqual(hinted("open_world_hint"), set(REACHES_OTHERS),
                             "a tool that reaches the outside world must ask")

    def test_only_reads_are_allowed_without_arguments(self):
        for tool in server_tools():
            self.assertEqual(decide(tool, {})[0], "allow" if tool in READ_TOOLS else "ask", tool)

    def test_tools_that_reach_others_ask_even_with_quiet_flags(self):
        quiet = {"notify": False, "send": False, "send_update": False, "send_cancellation": False}
        for tool in set(REACHES_OTHERS) - {"respond_to_meeting"}:
            decision, reason = decide(tool, quiet)
            self.assertEqual(decision, "ask", tool)
            self.assertIn(tool, reason)


class DecideTest(unittest.TestCase):
    def test_reads_are_allowed(self):
        for tool in READ_TOOLS:
            self.assertEqual(decide(tool, {})[0], "allow", tool)

    def test_respond_only_silently(self):
        self.assertEqual(decide("respond_to_meeting", {"notify": False})[0], "allow")
        self.assertEqual(decide("respond_to_meeting", {"notify": True})[0], "ask")
        self.assertEqual(decide("respond_to_meeting", {"send": False})[0], "allow", "deprecated alias")
        self.assertEqual(decide("respond_to_meeting", {"event_id": ["a", "b"], "send": False})[0], "allow")
        self.assertEqual(decide("respond_to_meeting", {})[0], "ask", "send defaults to true")
        self.assertEqual(decide("respond_to_meeting", {"send": True})[0], "ask")
        self.assertEqual(decide("respond_to_meeting", {"send": "false"})[0], "ask", "only a real false")
        self.assertEqual(decide("respond_to_meeting", {"send": 0})[0], "ask", "only a real false")
        self.assertEqual(decide("respond_to_meeting", None)[0], "ask")

    def test_moves_only_with_the_server_guard(self):
        self.assertEqual(decide("update_event", {"start": "x", "require_no_attendees": True})[0], "allow")
        guarded = {"event_id": "e", "start": "x", "require_no_attendees": True}
        self.assertEqual(decide("update_event", {"updates": [guarded, guarded]})[0], "allow")
        self.assertEqual(decide("update_event", {"updates": [guarded, {"event_id": "m", "start": "x"}]})[0],
                         "ask", "every entry of a batch needs the guard")
        self.assertEqual(decide("update_event", {"updates": [guarded, "junk"]})[0], "ask")
        self.assertEqual(decide("delete_event", {"event_id": "e", "require_no_attendees": True})[0],
                         "ask", "deletes stay manual")
        self.assertEqual(decide("update_event", {"start": "x"})[0], "ask")
        adds = {"require_no_attendees": True, "add_optional_attendees": ["bob@example.com"]}
        self.assertEqual(decide("update_event", adds)[0], "ask")

    def test_everything_else_asks(self):
        for tool in ["send_email", "create_draft", "update_draft", "delete_event", "create_event",
                     "reply_email", "empty_deleted_items", "save_attachments", "some_future_tool"]:
            decision, reason = decide(tool, {})
            self.assertEqual(decision, "ask", tool)
            self.assertIn(tool, reason)


class PluginOptionTest(unittest.TestCase):
    def test_manual_install_is_always_on(self):
        self.assertTrue(autopilot_enabled({}))

    def test_plugin_follows_its_option(self):
        root = {"CLAUDE_PLUGIN_ROOT": "/p"}
        self.assertFalse(autopilot_enabled(root))
        for on in ["true", "True", "1"]:
            self.assertTrue(autopilot_enabled({**root, "CLAUDE_PLUGIN_OPTION_CALENDAR_AUTOPILOT": on}))
        self.assertFalse(autopilot_enabled({**root, "CLAUDE_PLUGIN_OPTION_CALENDAR_AUTOPILOT": "false"}))


class HookProtocolTest(unittest.TestCase):
    def run_hook(self, event, env=None, cmd=None):
        full_env = {k: v for k, v in os.environ.items() if not k.startswith("CLAUDE_PLUGIN")}
        full_env.update(env or {})
        out = subprocess.run(cmd or [sys.executable, HOOK], input=json.dumps(event),
                             capture_output=True, text=True, check=True, env=full_env).stdout
        return json.loads(out) if out else None

    def test_plugin_launcher_runs_the_hook_when_the_mode_is_on(self):
        launcher = ["sh", os.path.join(os.path.dirname(__file__), "run_calendar_autopilot.sh")]
        event = {"tool_name": "mcp__plugin_outlook-mcp-rs_outlook__list_events", "tool_input": {}}
        on = {"CLAUDE_PLUGIN_ROOT": "/p", "CLAUDE_PLUGIN_OPTION_CALENDAR_AUTOPILOT": "true"}
        out = self.run_hook(event, on, launcher)
        self.assertEqual(out["hookSpecificOutput"]["permissionDecision"], "allow")
        self.assertIsNone(self.run_hook(event, {"CLAUDE_PLUGIN_ROOT": "/p"}, launcher), "mode off")

    def test_launcher_asks_when_no_python_is_found(self):
        launcher = os.path.join(os.path.dirname(__file__), "run_calendar_autopilot.sh")
        event = {"tool_name": "mcp__plugin_outlook-mcp-rs_outlook__send_email", "tool_input": {}}
        # Resolve sh before PATH is emptied; on Windows it is Git Bash's sh.exe, not /bin/sh.
        sh = shutil.which("sh") or "/bin/sh"
        out = self.run_hook(event, {"PATH": "/nonexistent"}, [sh, launcher])
        self.assertEqual(out["hookSpecificOutput"]["permissionDecision"], "ask")

    def test_prints_a_pre_tool_use_decision(self):
        out = self.run_hook({"tool_name": "mcp__outlook__respond_to_meeting",
                             "tool_input": {"event_id": "e|s", "response": "accept", "send": False}})
        self.assertEqual(out["hookSpecificOutput"]["hookEventName"], "PreToolUse")
        self.assertEqual(out["hookSpecificOutput"]["permissionDecision"], "allow")

    def test_fails_closed_on_unreadable_input(self):
        out = subprocess.run([sys.executable, HOOK], input="not json", capture_output=True,
                             text=True, check=True).stdout
        self.assertEqual(json.loads(out)["hookSpecificOutput"]["permissionDecision"], "ask")
        out = self.run_hook({"tool_name": "mcp__outlook__send_email", "tool_input": "oops"})
        self.assertEqual(out["hookSpecificOutput"]["permissionDecision"], "ask")

    def test_ignores_non_mcp_tools(self):
        self.assertIsNone(self.run_hook({"tool_name": "Bash", "tool_input": {"command": "ls"}}))

    def test_bare_tool_name(self):
        self.assertEqual(bare_tool_name("mcp__my_outlook__list_events"), "list_events")
        self.assertIsNone(bare_tool_name("Edit"))


if __name__ == "__main__":
    unittest.main()
