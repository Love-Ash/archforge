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


def _run(*args):
    r = subprocess.run(["officecli"] + list(args), capture_output=True, text=True)
    if r.returncode != 0:
        raise SystemExit("officecli %s failed: %s" % (" ".join(args[:2]), r.stderr.strip()))


def _build(path, title, body_text, body_props):
    _run("create", path, "--force")
    _run("add", path, "/", "--type", "slide", "--prop", "title=%s" % title)
    _run("set", path, "/slide[1]/shape[1]", "--prop", "size=28")
    props = ["--prop", "text=%s" % body_text, "--prop", "x=1in", "--prop", "y=2.2in",
             "--prop", "width=8in", "--prop", "height=1.2in", "--prop", "size=18"]
    for k, v in body_props:
        props += ["--prop", "%s=%s" % (k, v)]
    _run("add", path, "/slide[1]", "--type", "shape", *props)
    _run("close", path)


def main():
    if shutil.which("officecli") is None:
        raise SystemExit("officecli binary not found on PATH (npm i -g @officecli/officecli)")
    os.makedirs(HERE, exist_ok=True)

    _build(os.path.join(HERE, "e1_default_hangul.pptx"),
           "Font slots resolve at render time",
           "한글 본문이 기본 테마에서 폴백됩니다", [])

    _build(os.path.join(HERE, "clean_bilingual.pptx"),
           "Composed pipeline sanity check",
           "한글 본문에 동아시아 슬롯을 지정했습니다", [("font.ea", "Malgun Gothic")])

    print("generated 2 decks in %s" % HERE)


if __name__ == "__main__":
    sys.exit(main())
