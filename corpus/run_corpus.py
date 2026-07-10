# -*- coding: utf-8 -*-
"""Corpus runner: every committed deck with a sibling .json manifest is linted and the
finding counts are compared exactly against the manifest's expectations. Exit 1 on any
mismatch. This is the publicly reproducible accuracy record (external review: accuracy
claims relied on a private real-deck corpus).

Manifest keys:
  expected: {"E1": 1, ...}      exact per-code counts (empty object = must be clean)
  profile: core|full|editorial  (default full)
  expect_exit2: true            the file is a usage error (corrupt/hostile package)
  expect_incomplete: true       W18/summary.incomplete must be set

Usage: python corpus/run_corpus.py
"""
import glob
import io
import json
import os
import sys
from collections import Counter

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src"))

import archforge.lint as jl   # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))


def main():
    manifests = sorted(glob.glob(os.path.join(HERE, "*", "*.json")))
    bad = 0
    for mpath in manifests:
        with io.open(mpath, encoding="utf-8") as f:
            m = json.load(f)
        deck = os.path.splitext(mpath)[0] + ".pptx"
        rel = os.path.relpath(deck, HERE)
        if not os.path.exists(deck):
            print("MISSING  %s" % rel)
            bad += 1
            continue
        if m.get("expect_exit2"):
            try:
                jl.lint(deck, profile=m.get("profile", "full"))
                print("FAIL     %s (expected a usage error, got a report)" % rel)
                bad += 1
            except Exception:
                print("ok       %s (usage error as expected)" % rel)
            continue
        try:
            errors, warns = jl.lint(deck, profile=m.get("profile", "full"))
        except Exception as e:
            print("FAIL     %s (unexpected %s: %s)" % (rel, type(e).__name__, e))
            bad += 1
            continue
        got = Counter(f.code for f in list(errors) + list(warns) if f.code != "W18")
        want = Counter({k: int(v) for k, v in m.get("expected", {}).items()})
        incomplete = any(f.code == "W18" for f in warns)
        if m.get("expect_incomplete") and not incomplete:
            print("FAIL     %s (expected incomplete=true)" % rel)
            bad += 1
            continue
        if got != want:
            print("FAIL     %s (want %s, got %s)" % (rel, dict(want), dict(got)))
            bad += 1
        else:
            print("ok       %s %s" % (rel, dict(want)))
    print("---")
    print("corpus: %d manifests, %d failures" % (len(manifests), bad))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
