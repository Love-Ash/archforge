# -*- coding: utf-8 -*-
"""Corpus runner: every committed deck with a sibling .json manifest is linted through
the CLI and checked against the manifest. Exit 1 on any mismatch. This is the publicly
reproducible accuracy record (external review: accuracy claims relied on a private
real-deck corpus).

The runner drives the actual CLI (not the library) so it checks the exact exit contract,
and it verifies incompleteness in BOTH directions: an unexpected W18 fails a deck that
declared itself complete (0.7.1, external review section 6.1).

Manifest keys:
  expected: {"E1": 1, ...}      exact per-code counts, W18 excluded (empty = must be clean)
  profile: core|full|editorial  (default full)
  expect_exit2: true            the CLI must exit 2 (controlled usage error), not just
                                raise some exception
  expected_incomplete: bool     summary.incomplete must equal this (default false)
  expected_exit: int            override the expected CLI exit code (default: 1 if any
                                expected code is an ERROR, else 0)

Usage: python corpus/run_corpus.py
"""
import glob
import io
import json
import os
import subprocess
import sys
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "..", "src")


def _cli(args):
    env = dict(os.environ)
    env["ARCHFORGE_LANG"] = "en"
    env["PYTHONPATH"] = SRC + os.pathsep + env.get("PYTHONPATH", "")
    return subprocess.run([sys.executable, "-m", "archforge"] + args,
                          capture_output=True, text=True, encoding="utf-8", env=env)


def main():
    manifests = sorted(glob.glob(os.path.join(HERE, "*", "*.json")))
    bad = 0
    for mpath in manifests:
        with io.open(mpath, encoding="utf-8") as f:
            m = json.load(f)
        deck = os.path.splitext(mpath)[0] + ".pptx"
        rel = os.path.relpath(deck, HERE)
        profile = m.get("profile", "full")
        if not os.path.exists(deck):
            print("MISSING  %s" % rel)
            bad += 1
            continue
        if m.get("expect_exit2"):
            r = _cli([deck, "--profile", profile])
            if r.returncode == 2:
                print("ok       %s (controlled exit 2)" % rel)
            else:
                print("FAIL     %s (expected exit 2, got %d)" % (rel, r.returncode))
                bad += 1
            continue
        r = _cli([deck, "--profile", profile, "--json"])
        try:
            doc = json.loads(r.stdout)
        except Exception:
            print("FAIL     %s (no JSON; exit %d, stderr: %s)"
                  % (rel, r.returncode, (r.stderr or "")[:120]))
            bad += 1
            continue
        findings = list(doc.get("errors", [])) + list(doc.get("warnings", []))
        got = Counter(f["code"] for f in findings if f["code"] != "W18")
        want = Counter({k: int(v) for k, v in m.get("expected", {}).items()})
        incomplete = bool(doc["summary"].get("incomplete"))
        want_incomplete = bool(m.get("expected_incomplete",
                                     m.get("expect_incomplete", False)))
        exp_exit = m.get("expected_exit")
        if exp_exit is None:
            exp_exit = 1 if any(c.startswith("E") for c in want) else 0
        problems = []
        if got != want:
            problems.append("findings want %s got %s" % (dict(want), dict(got)))
        if incomplete != want_incomplete:
            problems.append("incomplete want %s got %s" % (want_incomplete, incomplete))
        if r.returncode != exp_exit:
            problems.append("exit want %d got %d" % (exp_exit, r.returncode))
        if problems:
            print("FAIL     %s (%s)" % (rel, "; ".join(problems)))
            bad += 1
        else:
            print("ok       %s %s" % (rel, dict(want)))
    print("---")
    print("corpus: %d manifests, %d failures" % (len(manifests), bad))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
