#!/usr/bin/env python3
"""Verify either the starter or reference implementation of the lab."""

from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest


LAB_ROOT = Path(__file__).resolve().parents[1]


def load_module(path: Path):
    spec = importlib.util.spec_from_file_location("lab_policy", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def load_policy() -> dict:
    with (LAB_ROOT / "policy.json").open(encoding="utf-8") as stream:
        return json.load(stream)


parser = argparse.ArgumentParser()
parser.add_argument("implementation", type=Path)
args = parser.parse_args()
implementation = load_module(args.implementation.resolve())
policy = load_policy()


class PolicyLabTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.base = Path(self.temp.name)
        (self.base / "workspace" / "src").mkdir(parents=True)
        (self.base / "workspace" / "output").mkdir()
        (self.base / "workspace" / "secrets").mkdir()
        (self.base / "workspace-other").mkdir()

    def tearDown(self) -> None:
        self.temp.cleanup()

    def file_allowed(self, path: str, operation: str) -> bool:
        return implementation.file_allowed(path, operation, policy, self.base)

    def test_read_inside_workspace(self) -> None:
        self.assertTrue(self.file_allowed("workspace/src/main.py", "read"))

    def test_write_only_inside_output(self) -> None:
        self.assertTrue(self.file_allowed("workspace/output/report.txt", "write"))
        self.assertFalse(self.file_allowed("workspace/src/main.py", "write"))

    def test_similar_prefix_is_not_inside(self) -> None:
        self.assertFalse(self.file_allowed("workspace-other/data.txt", "read"))

    def test_traversal_is_not_allowed(self) -> None:
        self.assertFalse(self.file_allowed("workspace/../outside.txt", "read"))

    def test_denied_rule_wins(self) -> None:
        self.assertFalse(self.file_allowed("workspace/secrets/token.txt", "read"))

    def test_network_exact_match(self) -> None:
        self.assertTrue(
            implementation.network_allowed(
                "https://api.github.com:443/repos/microsoft/mxc", policy
            )
        )

    def test_network_rejects_other_host_port_or_scheme(self) -> None:
        self.assertFalse(
            implementation.network_allowed("https://sub.api.github.com:443", policy)
        )
        self.assertFalse(
            implementation.network_allowed("https://api.github.com:80", policy)
        )
        self.assertFalse(
            implementation.network_allowed("http://api.github.com:443", policy)
        )


if __name__ == "__main__":
    unittest.main(argv=[__file__])
