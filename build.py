#!/usr/bin/env python3
"""Build the presentation, Pages site, and demo checks."""

from __future__ import annotations

import argparse
import os
from pathlib import Path
import shutil
import subprocess

from presentation.src import build_html, build_pptx


ROOT = Path(__file__).resolve().parent
IMAGES = ROOT / "presentation" / "assets" / "images"
PPTX_OUT = ROOT / "dist" / "pptx" / "Guardrails with Sandboxing - MXC.pptx"
HTML_OUT = ROOT / "dist" / "html" / "index.html"
PAGES_OUT = ROOT / "docs" / "index.html"


def build_presentation() -> None:
    build_pptx.build(str(PPTX_OUT), str(IMAGES))


def build_site() -> None:
    build_html.build(str(HTML_OUT), str(IMAGES), 1600)
    PAGES_OUT.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(HTML_OUT, PAGES_OUT)
    print(f"Saved GitHub Pages site {PAGES_OUT}")


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
    args = parser.parse_args()

    if args.target in ("all", "presentation", "pptx"):
        build_presentation()
    if args.target in ("all", "presentation", "html"):
        build_site()
    if args.target in ("all", "demo"):
        build_demo()


if __name__ == "__main__":
    main()
