#!/bin/sh
# Plugin launcher for calendar_autopilot.py: runs it with the first working
# Python 3 (python3, python, or the Windows py launcher). Without one, every
# Outlook call asks: the hook never allows anything it couldn't check.
script="$(dirname "$0")/calendar_autopilot.py"
for py in python3 python py; do
  if "$py" -c 'import sys; sys.exit(sys.version_info[0] != 3)' >/dev/null 2>&1 </dev/null; then
    exec "$py" "$script"
  fi
done
cat >/dev/null
printf '%s\n' '{"hookSpecificOutput": {"hookEventName": "PreToolUse", "permissionDecision": "ask", "permissionDecisionReason": "calendar autopilot: no Python 3 found, so every Outlook call asks"}}'
