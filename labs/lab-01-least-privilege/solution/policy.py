"""Reference implementation for the least-privilege lab."""

from pathlib import Path
from typing import Union
from urllib.parse import urlsplit


def _inside(path: Path, root: Path) -> bool:
    try:
        path.relative_to(root)
        return True
    except ValueError:
        return False


def _resolved(base_dir: Path, value: Union[str, Path]) -> Path:
    candidate = Path(value)
    if not candidate.is_absolute():
        candidate = base_dir / candidate
    return candidate.resolve()


def file_allowed(path: Union[str, Path], operation: str, policy: dict, base_dir: Path) -> bool:
    if operation not in ("read", "write"):
        raise ValueError("operation must be 'read' or 'write'")

    candidate = _resolved(base_dir, path)
    rules = policy["filesystem"]
    denied = [_resolved(base_dir, root) for root in rules.get("denied", [])]
    if any(_inside(candidate, root) for root in denied):
        return False

    grants = [_resolved(base_dir, root) for root in rules.get(operation, [])]
    return any(_inside(candidate, root) for root in grants)


def network_allowed(url: str, policy: dict) -> bool:
    parsed = urlsplit(url)
    if parsed.scheme != "https" or parsed.hostname is None:
        return False

    port = parsed.port or 443
    return any(
        parsed.hostname.lower() == rule["host"].lower() and port in rule["ports"]
        for rule in policy["network"].get("allow", [])
    )
