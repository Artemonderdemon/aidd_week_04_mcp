#!/usr/bin/env python3
"""Qwen PreToolUse hook: deny common network utilities in shell commands."""
import json
import re
import sys

data = json.load(sys.stdin)
tool = data.get("tool_name", "")
tool_input = data.get("tool_input", {})
command = tool_input.get("command", "") if isinstance(tool_input, dict) else ""


def decide(decision: str, reason: str) -> None:
    print(json.dumps({
        "hookSpecificOutput": {
            "hookEventName": "PreToolUse",
            "permissionDecision": decision,
            "permissionDecisionReason": reason,
        }
    }, ensure_ascii=False))
    raise SystemExit(0)


if tool == "run_shell_command" and re.search(
    r"\b(curl|wget|nc|ncat)\b", command, flags=re.IGNORECASE
):
    decide("deny", "Сетевые команды запрещены.")

decide("allow", "ok")
