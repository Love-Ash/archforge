# -*- coding: utf-8 -*-
"""Version-consistency gate (0.8.x, external audit: a QA tool whose own docs cite three
different versions reads as a QA tool that cannot QA itself).

One source of truth: pyproject.toml. Everything below must agree with it exactly,
because the policy since the release-cadence fix is "bump the version only at the
release moment", so on main these are always equal:

  - src/archforge/__init__.py  __version__
  - CHANGELOG.md latest heading
  - README.md / README.ko.md  Action example tag and pre-commit rev
  - .pre-commit-hooks.yaml    usage-comment rev

Exit 1 with a diff-style listing on any mismatch. Runs in CI on every push.
"""
import io
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _read(rel):
    with io.open(os.path.join(ROOT, rel), encoding="utf-8") as f:
        return f.read()


def main():
    truth = re.search(r'^version = "([^"]+)"', _read("pyproject.toml"), re.M).group(1)
    checks = []

    m = re.search(r'__version__ = "([^"]+)"', _read("src/archforge/__init__.py"))
    checks.append(("src/archforge/__init__.py __version__", m.group(1) if m else None))

    m = re.search(r"^## ([0-9][^ ]*) \(", _read("CHANGELOG.md"), re.M)
    checks.append(("CHANGELOG.md latest heading", m.group(1) if m else None))

    for rel in ("README.md", "README.ko.md"):
        body = _read(rel)
        for label, pat in (("Action tag", r"Love-Ash/archforge@v([0-9][^\s\"']*)"),
                           ("pre-commit rev", r"rev: v([0-9][^\s\"']*)")):
            for found in re.findall(pat, body):
                checks.append(("%s %s" % (rel, label), found))

    for found in re.findall(r"rev: v([0-9][^\s\"']*)", _read(".pre-commit-hooks.yaml")):
        checks.append((".pre-commit-hooks.yaml rev", found))

    bad = [(where, got) for where, got in checks if got != truth]
    print("version truth (pyproject.toml): %s  |  %d reference(s) checked"
          % (truth, len(checks)))
    for where, got in bad:
        print("MISMATCH  %-40s %s (expected %s)" % (where, got, truth))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
