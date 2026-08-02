# -*- coding: utf-8 -*-
"""Rule-inventory gate (0.8.x): the registry is the single source of truth, and every
public surface that enumerates rules must agree with it.

Adding a code to src/archforge/rules.py touches more places than it looks. A rule that
lands in the registry but nowhere else still passes the whole test suite, because no test
imports the registry to compare it against the docs; the rule simply exists, undocumented,
and `python scripts/make_rule_docs.py` crashes with a bare KeyError the next time anyone
runs it. That is exactly what happened to a draft W19 in 2026-07, which is why this gate
exists.

Checked against sorted(RULES):

  - scripts/make_rule_docs.py MEANING has an entry (so the generator cannot KeyError)
  - docs/rules/<CODE>.md exists, and no page exists for a code that was removed
  - README.md, README.ko.md and both copies of SKILL.md name the code

llms.txt states its inventory as ranges rather than per-code rows, so it is checked
differently: the highest E code and the highest W code must appear literally, which fails
the moment a new top code is added and the range text goes stale.

Exit 1 with a listing on any mismatch. Runs in CI on every push.
"""
import io
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "src"))
sys.path.insert(0, os.path.join(ROOT, "scripts"))

from archforge.rules import RULES            # noqa: E402
from make_rule_docs import MEANING           # noqa: E402

PER_CODE_SURFACES = (
    "README.md",
    "README.ko.md",
    os.path.join("skills", "archforge-pptx-lint", "SKILL.md"),
    os.path.join("src", "archforge", "skills", "archforge-pptx-lint", "SKILL.md"),
)
RANGE_SURFACES = ("llms.txt",)


def _read(rel):
    with io.open(os.path.join(ROOT, rel), encoding="utf-8") as f:
        return f.read()


def _sort_key(code):
    return (code[0] != "E", int(code[1:]))


def main():
    codes = sorted(RULES, key=_sort_key)
    bad = []

    for code in codes:
        if code not in MEANING:
            bad.append(("scripts/make_rule_docs.py MEANING", code))

    pages = set(f[:-3] for f in os.listdir(os.path.join(ROOT, "docs", "rules"))
                if f.endswith(".md"))
    for code in codes:
        if code not in pages:
            bad.append(("docs/rules/%s.md missing" % code, code))
    for orphan in sorted(pages - set(codes), key=_sort_key):
        bad.append(("docs/rules/%s.md has no registry entry" % orphan, orphan))

    for rel in PER_CODE_SURFACES:
        body = _read(rel)
        for code in codes:
            if not re.search(r"\b%s\b" % code, body):
                bad.append((rel, code))

    tops = [max((c for c in codes if c[0] == prefix), key=_sort_key, default=None)
            for prefix in ("E", "W")]
    for rel in RANGE_SURFACES:
        body = _read(rel)
        for code in tops:
            if code and not re.search(r"\b%s\b" % code, body):
                bad.append(("%s (range text is stale)" % rel, code))

    print("registry: %d codes (%s)  |  %d surface(s) checked"
          % (len(codes), ", ".join(codes), len(PER_CODE_SURFACES) + len(RANGE_SURFACES)))
    for where, code in bad:
        print("MISSING   %-52s %s" % (where, code))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
