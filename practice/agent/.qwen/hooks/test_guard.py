import json
import subprocess
import sys
import unittest
from pathlib import Path

HOOK = Path(__file__).with_name("guard.py")


def invoke(tool_name, command):
    payload = json.dumps({
        "tool_name": tool_name,
        "tool_input": {"command": command},
    })
    result = subprocess.run(
        [sys.executable, str(HOOK)],
        input=payload,
        text=True,
        capture_output=True,
        check=True,
    )
    return json.loads(result.stdout)["hookSpecificOutput"]["permissionDecision"]


class GuardHookTest(unittest.TestCase):
    def test_denies_curl_case_insensitively(self):
        self.assertEqual(invoke("run_shell_command", "curl http://127.0.0.1"), "deny")
        self.assertEqual(invoke("run_shell_command", "WGET example.invalid"), "deny")

    def test_denies_common_netcat_names(self):
        for command in ("nc -lk 9999", "ncat example.invalid 80"):
            with self.subTest(command=command):
                self.assertEqual(invoke("run_shell_command", command), "deny")

    def test_allows_non_network_command_and_other_tools(self):
        self.assertEqual(invoke("run_shell_command", "python --version"), "allow")
        self.assertEqual(invoke("write_file", "curl text in another tool"), "allow")


if __name__ == "__main__":
    unittest.main()
