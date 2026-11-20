"""Starter implementation for the least-privilege lab."""

from pathlib import Path
from typing import Union
from urllib.parse import urlsplit


def _inside(path: Path, root: Path) -> bool:
    # TODO: compare resolved paths without using string prefixes.
    raise NotImplementedError


def file_allowed(path: Union[str, Path], operation: str, policy: dict, base_dir: Path) -> bool:
    # TODO: resolve the candidate and apply denied rules before read/write grants.
    raise NotImplementedError


def network_allowed(url: str, policy: dict) -> bool:
    # TODO: require HTTPS and exact host plus allowed port matching.
    raise NotImplementedError
