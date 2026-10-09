import json
import subprocess
import sys
import unittest
from pathlib import Path


EXAMPLE = Path(__file__).resolve().parents[1] / "examples" / "read_emails.py"


class ExampleTests(unittest.TestCase):
    def run_example(self, *args):
        return subprocess.run(
            [sys.executable, str(EXAMPLE), *args],
            capture_output=True, text=True, timeout=10,
        )

    def test_offline_outputs_json(self):
        result = self.run_example("--offline")
        self.assertEqual(result.returncode, 0, result.stderr)
        output = json.loads(result.stdout)
        self.assertEqual(output[0]["categories"], ["work", "urgent", "follow-up needed"])
        self.assertEqual(output[0]["email"]["attachments"][0]["name"], "report.pdf")

    def test_offline_zero_limit(self):
        result = self.run_example("--offline", "--limit", "0")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout), [])

    def test_negative_limit(self):
        self.assertEqual(self.run_example("--offline", "--limit", "-1").returncode, 2)

    def test_missing_configuration(self):
        result = self.run_example("--config", "/nonexistent/outlook-settings.json")
        self.assertEqual(result.returncode, 1)
        self.assertIn("Cannot read the configuration JSON", result.stderr)
        self.assertNotIn("Traceback", result.stderr)


if __name__ == "__main__":
    unittest.main()
