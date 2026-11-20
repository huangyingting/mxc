from __future__ import annotations

from pathlib import Path
import tempfile
import unittest

from demo.src.policy_demo import Policy, PolicyDenied


class PolicyDemoTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        (self.root / "input").mkdir()
        (self.root / "input" / "request.txt").write_text("hello", encoding="utf-8")
        config = {
            "filesystem": {
                "readable": ["input"],
                "writable": ["output"],
                "denied": ["input/private"],
            },
            "network": {
                "default": "deny",
                "allow": [{"host": "api.github.com", "ports": [443]}],
            },
        }
        self.policy = Policy(config, self.root)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_reads_only_granted_paths(self) -> None:
        self.assertEqual(self.policy.read_text("input/request.txt"), "hello")
        with self.assertRaises(PolicyDenied):
            self.policy.read_text("input/../other.txt")

    def test_deny_list_overrides_a_read_grant(self) -> None:
        private = self.root / "input" / "private"
        private.mkdir()
        (private / "token.txt").write_text("secret", encoding="utf-8")
        with self.assertRaises(PolicyDenied):
            self.policy.read_text("input/private/token.txt")

    def test_writes_only_to_output(self) -> None:
        target = self.policy.write_text("output/report.txt", "ok")
        self.assertEqual(target.read_text(encoding="utf-8"), "ok")
        with self.assertRaises(PolicyDenied):
            self.policy.write_text("input/report.txt", "no")

    def test_network_requires_exact_host_and_port(self) -> None:
        self.policy.require_network("https://api.github.com/repos/microsoft/mxc")
        with self.assertRaises(PolicyDenied):
            self.policy.require_network("https://sub.api.github.com:443")
        with self.assertRaises(PolicyDenied):
            self.policy.require_network("https://api.github.com:80")
        with self.assertRaises(PolicyDenied):
            self.policy.require_network("http://api.github.com:443")


if __name__ == "__main__":
    unittest.main()
