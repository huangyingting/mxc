#!/usr/bin/env python3
"""Build the presentation, Pages site, and demo checks."""

from __future__ import annotations

import argparse
import os
from pathlib import Path
import re
import shutil
import subprocess

from presentation.src import build_html, build_pptx


ROOT = Path(__file__).resolve().parent
IMAGES = ROOT / "presentation" / "assets" / "images"
PPTX_OUT = ROOT / "dist" / "pptx" / "Guardrails with Sandboxing - MXC.pptx"
HTML_OUT = ROOT / "dist" / "html" / "index.html"
PAGES_OUT = ROOT / "docs" / "index.html"


def resolve_pages_url(explicit_url: str | None = None) -> str | None:
    raw = (explicit_url or os.environ.get("PAGES_URL") or "").strip()
    if not raw:
        try:
            remote = subprocess.check_output(
                ["git", "config", "--get", "remote.origin.url"],
                cwd=str(ROOT),
                text=True,
                stderr=subprocess.DEVNULL,
            ).strip()
            match = re.search(r"github\.com[:/](?P<owner>[^/]+)/(?P<repo>[^/.]+?)(?:\.git)?$", remote)
            if match:
                owner = match.group("owner")
                repo = match.group("repo")
                raw = f"https://{owner}.github.io/" if repo.lower() == f"{owner.lower()}.github.io" else f"https://{owner}.github.io/{repo}/"
        except Exception:
            raw = ""
    if not raw:
        return None
    if not raw.endswith(".html"):
        raw = raw.rstrip("/") + "/"
    return raw


def build_presentation(pages_url: str | None = None) -> None:
    build_pptx.build(str(PPTX_OUT), str(IMAGES), pages_url=resolve_pages_url(pages_url))


def build_site(pages_url: str | None = None) -> None:
    resolved_url = resolve_pages_url(pages_url)
    build_html.build(str(HTML_OUT), str(IMAGES), 1600, pages_url=resolved_url, presenter_tools=True)
    PAGES_OUT.parent.mkdir(parents=True, exist_ok=True)
    build_html.build(str(PAGES_OUT), str(IMAGES), 1600, pages_url=resolved_url, presenter_tools=False)
    print(f"Saved GitHub Pages site {PAGES_OUT} (public slides only" + (f", QR -> {resolved_url})" if resolved_url else ")"))


def build_demo() -> None:
    demo_dir = ROOT / "demo"
    lock_file = demo_dir / "package-lock.json"
    install = "ci" if lock_file.exists() else "install"
    npm = shutil.which("npm.cmd" if os.name == "nt" else "npm")
    if npm is None:
        raise RuntimeError("npm is required to build the MXC SDK demo")
    subprocess.run(
        [npm, "--prefix", str(demo_dir), install, "--no-audit"],
        check=True,
    )
    subprocess.run(
        [npm, "--prefix", str(demo_dir), "run", "build"],
        check=True,
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--target",
        choices=("all", "presentation", "pptx", "html", "demo"),
        default="all",
    )
    parser.add_argument(
        "--pages-url",
        default=None,
        help="Final GitHub Pages URL encoded as a QR code on Slide 1",
    )
    args = parser.parse_args()

    if args.target in ("all", "presentation", "pptx"):
        build_presentation(args.pages_url)
    if args.target in ("all", "presentation", "html"):
        build_site(args.pages_url)
    if args.target in ("all", "demo"):
        build_demo()


if __name__ == "__main__":
    main()
