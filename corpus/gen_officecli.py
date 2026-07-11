# -*- coding: utf-8 -*-
"""Seed corpus, OfficeCLI generator: decks authored through the OfficeCLI binary
(https://github.com/iOfficeAI/OfficeCLI) so a third-party, non-Python writer is
covered. Ground truth is by construction: the defect fixture relies on OfficeCLI's
blank-document theme shipping empty ea slots while `add shape` writes no run a:ea,
so a Hangul run resolves to a Latin-only face (the E1 class); the clean fixture sets
the East Asian slot explicitly via `--prop font.ea`.

Requires the `officecli` binary on PATH (npm i -g @officecli/officecli). The committed
.pptx files were generated with OfficeCLI 1.0.135; this script regenerates them but is
not run in CI (CI lints the committed binaries via corpus/run_corpus.py).

Usage: python corpus/gen_officecli.py
"""
import os
import shutil
import subprocess
import sys

HERE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "officecli")


def _officecli():
    # Windows npm installs both an extensionless sh shim and officecli.cmd;
    # CreateProcess can only execute the latter, so resolve explicitly.
    return shutil.which("officecli.cmd") or shutil.which("officecli")


def _run(*args):
    r = subprocess.run([_officecli()] + list(args), capture_output=True,
                       text=True, encoding="utf-8", errors="replace")
    if r.returncode != 0:
        raise SystemExit("officecli %s failed: %s" % (" ".join(args[:2]), r.stderr.strip()))


EM_DASH = chr(0x2014)


def _build(path, title, shapes):
    """shapes: list of (text, extra_props) tuples; each gets the standard frame plus
    extras, stacked vertically unless the extras override x/y/width/height."""
    _run("create", path, "--force")
    _run("add", path, "/", "--type", "slide", "--prop", "title=%s" % title)
    _run("set", path, "/slide[1]/shape[1]", "--prop", "size=28")
    for i, (text, extra) in enumerate(shapes):
        merged = {"text": text, "x": "1in", "y": "%.1fin" % (2.2 + 1.4 * i),
                  "width": "8in", "height": "1.2in", "size": "18"}
        merged.update(dict(extra))
        props = []
        for k, v in merged.items():
            props += ["--prop", "%s=%s" % (k, v)]
        _run("add", path, "/slide[1]", "--type", "shape", *props)
    _run("close", path)


def main():
    if shutil.which("officecli") is None:
        raise SystemExit("officecli binary not found on PATH (npm i -g @officecli/officecli)")
    os.makedirs(HERE, exist_ok=True)

    _build(os.path.join(HERE, "e1_default_hangul.pptx"),
           "Font slots resolve at render time",
           [("한글 본문이 기본 테마에서 폴백됩니다", [])])

    _build(os.path.join(HERE, "clean_bilingual.pptx"),
           "Composed pipeline sanity check",
           [("한글 본문에 동아시아 슬롯을 지정했습니다", [("font.ea", "Malgun Gothic")])])

    _build(os.path.join(HERE, "e2_em_dash.pptx"),
           "Dash punctuation fixture",
           [("The gate runs in CI %s no server, no telemetry" % EM_DASH, [])])

    _build(os.path.join(HERE, "e3_tiny_size.pptx"),
           "Effective size fixture",
           [("This attribution line was shrunk to four points", [("size", "4")])])

    _build(os.path.join(HERE, "e4_tracked_hangul.pptx"),
           "Tracking damage fixture",
           [("자간이 벌어진 한글 본문입니다", [("font.ea", "Malgun Gothic"), ("spacing", "3")])])

    _build(os.path.join(HERE, "w15_overlap.pptx"),
           "Frame collision fixture",
           [("The first body frame carries a full sentence of visible text", []),
            ("The second frame sits directly on the first line of the first one",
             [("y", "2.2in")])])

    _build(os.path.join(HERE, "w16_offcanvas.pptx"),
           "Off canvas fixture",
           [("This body frame starts near the right edge and its glyphs run past the canvas boundary",
             [("x", "11in"), ("width", "4in")])])

    print("generated 7 decks in %s" % HERE)


if __name__ == "__main__":
    sys.exit(main())
