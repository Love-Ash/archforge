# -*- coding: utf-8 -*-
"""Generates docs/ACCURACY.md: the regression-corpus verification record.

Named carefully (0.8.x, external audit): these are constructed-fixture results, not an
independent accuracy benchmark. The same author wrote the rules and the fixtures, the
per-gate samples are tiny, and a 1/1 success says almost nothing statistically, so the
table publishes exact binomial lower bounds alongside the raw counts and refuses to
call itself "accuracy" without qualification.

Every number is publicly reproducible (anyone can rerun the shipped decks through the
CLI; CI fails on drift), which is weaker than third-party verified and the doc says so.

Usage: python scripts/make_accuracy_table.py   (re-runs the corpus, rewrites the doc)
"""
import glob
import io
import json
import math
import os
import subprocess
import sys
from collections import Counter, defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CORPUS = os.path.join(ROOT, "corpus")
OUT = os.path.join(ROOT, "docs", "ACCURACY.md")


def _cli(args):
    env = dict(os.environ)
    env["ARCHFORGE_LANG"] = "en"
    env["PYTHONPATH"] = os.path.join(ROOT, "src") + os.pathsep + env.get("PYTHONPATH", "")
    return subprocess.run([sys.executable, "-m", "archforge"] + args,
                          capture_output=True, text=True, encoding="utf-8", env=env,
                          stdin=subprocess.DEVNULL)


def _binom_lower(successes, n, alpha=0.05):
    """Exact (Clopper-Pearson) one-sided lower bound for a binomial proportion, via
    bisection on the beta quantile relation. No scipy dependency."""
    if n == 0:
        return float("nan")
    if successes == 0:
        return 0.0
    # lower bound p such that P(X >= successes | n, p) = alpha  (successes==n case:
    # alpha^(1/n))
    if successes == n:
        return alpha ** (1.0 / n)
    lo, hi = 0.0, 1.0
    for _ in range(60):
        mid = (lo + hi) / 2
        # P(X >= successes) under p=mid
        tail = sum(math.comb(n, k) * mid**k * (1 - mid)**(n - k)
                   for k in range(successes, n + 1))
        if tail < alpha:
            lo = mid
        else:
            hi = mid
    return lo


def main():
    tp = Counter()
    fp = Counter()
    fn = Counter()
    tn_decks = defaultdict(int)
    generators = set()
    per_deck = []
    n_negative = 0
    total_slides = 0
    for mpath in sorted(glob.glob(os.path.join(CORPUS, "*", "*.json"))):
        with io.open(mpath, encoding="utf-8") as f:
            m = json.load(f)
        if m.get("expect_exit2"):
            continue
        deck = os.path.splitext(mpath)[0] + ".pptx"
        if not os.path.exists(deck):
            continue
        generators.add(m.get("generator", "?").split(" ")[0])
        want = Counter({k: int(v) for k, v in m.get("expected", {}).items()})
        if not want:
            n_negative += 1
        r = _cli([deck, "--profile", m.get("profile", "full"), "--json"])
        doc = json.loads(r.stdout)
        got = Counter(f["code"] for f in doc.get("errors", []) + doc.get("warnings", [])
                      if f["code"] != "W18")
        pages = {f["page"] for f in doc.get("errors", []) + doc.get("warnings", [])}
        # slide count: read from the pptx (cheap, python-pptx)
        try:
            from pptx import Presentation
            total_slides += len(Presentation(deck).slides.__iter__.__self__._sldIdLst)
        except Exception:
            total_slides += max(pages) if pages else 1
        per_deck.append((want, got))
        for code in set(want) | set(got):
            w, g = want[code], got[code]
            tp[code] += min(w, g)
            fp[code] += max(0, g - w)
            fn[code] += max(0, w - g)
    n_decks = len(per_deck)
    universe = set(list(tp) + list(fp) + list(fn))
    for want, got in per_deck:
        present = set(want) | set(got)
        for code in universe - present:
            tn_decks[code] += 1

    total_fp = sum(fp.values())
    fp_per_10 = (total_fp / total_slides * 10) if total_slides else float("nan")

    codes = sorted(universe, key=lambda c: (c[0] != "E", int(c[1:])))
    lines = [
        "# Regression corpus results",
        "",
        "Constructed-fixture verification on the [public corpus](../corpus/), regenerated",
        "by `python scripts/make_accuracy_table.py`. Every number is **publicly",
        "reproducible** (the decks and manifests ship in this repository and CI fails on",
        "drift), which is deliberately weaker language than third-party verified: no",
        "independent party has validated these rules yet.",
        "",
        "Read this table for what it is: proof that the gates fire on deliberately seeded",
        "defects and stay silent on clean decks. It is NOT an accuracy benchmark, for two",
        "reasons stated plainly. First, the same author wrote the rules and the fixtures,",
        "so this is regression verification, not independent ground truth. Second, the",
        "per-gate samples are tiny; the exact one-sided 95% lower bound below shows how",
        "little a small perfect score proves (1/1 -> 5%, 3/3 -> 37%).",
        "",
        "Corpus: %d decks, %d generator families, %d clean negatives, %d slides, all"
        % (n_decks, len(generators), n_negative, total_slides),
        "synthetic (no field decks yet). Observed false positives across the whole",
        "corpus: %d (%.2f per 10 slides on this corpus)." % (total_fp, fp_per_10),
        "",
        "| Gate | TP | FP | FN | Deck-level TN | Recall (95% LB) |",
        "|:----:|---:|---:|---:|---:|:---|",
    ]
    for c in codes:
        n = tp[c] + fn[c]
        lb = _binom_lower(tp[c], n) if n else float("nan")
        lines.append("| %s | %d | %d | %d | %d | %d/%d (>= %.0f%%) |"
                     % (c, tp[c], fp[c], fn[c], tn_decks[c], tp[c], n, lb * 100))
    lines += [
        "",
        "Gates absent from the table have no positive fixture in the corpus yet (their",
        "behavior is covered by the unit/property suites, not by corpus ground truth).",
        "W18 is excluded by design: it is an abstention signal, not a finding.",
        "",
        "What would upgrade this into an accuracy record: unseen decks from generators",
        "the author did not write fixtures for (Google Slides / Canva / LibreOffice",
        "exports, real corporate templates), independent labeling, and contributed",
        "false-positive reproductions. That corpus growth is the highest-value",
        "contribution this project takes; see [corpus/README.md](../corpus/README.md).",
        "",
    ]
    with io.open(OUT, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines))
    print("wrote", OUT, "(%d gates, %d decks, %d slides)" % (len(codes), n_decks,
                                                             total_slides))
    if total_fp or sum(fn.values()):
        print("WARNING: corpus has FP/FN; the corpus runner should be failing")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
