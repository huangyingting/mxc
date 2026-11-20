#!/usr/bin/env python3
"""A policy engine teaching demo, not an operating-system sandbox."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Union
from urllib.parse import urlsplit


class PolicyDenied(PermissionError):
    """Raised when a requested operation is outside the policy."""


class Policy:
    def __init__(self, config: dict, base_dir: Path):
        self.config = config
        self.base_dir = base_dir.resolve()

    @classmethod
    def load(cls, path: Path) -> "Policy":
        with path.open(encoding="utf-8") as stream:
            config = json.load(stream)
        return cls(config, path.parent.parent)

    def _path(self, value: Union[Path, str]) -> Path:
        candidate = Path(value)
        if not candidate.is_absolute():
            candidate = self.base_dir / candidate
        return candidate.resolve()

    @staticmethod
    def _within(path: Path, root: Path) -> bool:
        try:
            path.relative_to(root)
            return True
        except ValueError:
            return False

    def _file_allowed(self, path: Union[Path, str], operation: str) -> bool:
        target = self._path(path)
        rules = self.config["filesystem"]
        rule_key = {"read": "readable", "write": "writable"}[operation]
        denied = [self._path(item) for item in rules.get("denied", [])]
        allowed = [self._path(item) for item in rules[rule_key]]
        return not any(self._within(target, root) for root in denied) and any(
            self._within(target, root) for root in allowed
        )

    def require_file(self, path: Union[Path, str], operation: str) -> Path:
        if operation not in ("read", "write"):
            raise ValueError("operation must be 'read' or 'write'")
        target = self._path(path)
        if not self._file_allowed(target, operation):
            raise PolicyDenied(f"{operation} denied: {target}")
        return target

    def require_network(self, url: str) -> None:
        parsed = urlsplit(url)
        network = self.config["network"]
        if parsed.scheme != "https" or parsed.hostname is None:
            raise PolicyDenied(f"network denied: {url}")
        port = parsed.port or 443
        for rule in network.get("allow", []):
            if parsed.hostname.lower() == rule["host"].lower() and port in rule["ports"]:
                return
        raise PolicyDenied(f"network denied: {url}")

    def read_text(self, path: Union[Path, str]) -> str:
        return self.require_file(path, "read").read_text(encoding="utf-8")

    def write_text(self, path: Union[Path, str], content: str) -> Path:
        target = self.require_file(path, "write")
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(content, encoding="utf-8")
        return target


def demonstrate(policy: Policy) -> None:
    request = policy.read_text("input/request.txt").strip()
    report = policy.write_text(
        "output/report.txt",
        f"Policy demo\nRequest: {request}\nResult: least-privilege checks passed\n",
    )
    print(f"ALLOW write -> {report}")

    policy.require_network("https://api.github.com:443/repos/microsoft/mxc")
    print("ALLOW network -> api.github.com:443")

    for action, callback in (
        ("read outside the grant", lambda: policy.read_text("../README.md")),
        (
            "connect to an unlisted host",
            lambda: policy.require_network("https://example.com:443"),
        ),
    ):
        try:
            callback()
        except PolicyDenied as error:
            print(f"DENY  {action} -> {error}")
        else:
            raise RuntimeError(f"policy unexpectedly allowed: {action}")


def main() -> None:
    here = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(
        description="Demonstrate least-privilege policy checks (not an OS sandbox)."
    )
    parser.add_argument(
        "--policy",
        type=Path,
        default=here / "config" / "policy.json",
    )
    args = parser.parse_args()
    demonstrate(Policy.load(args.policy.resolve()))


if __name__ == "__main__":
    main()
