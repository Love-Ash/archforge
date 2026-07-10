# -*- coding: utf-8 -*-
"""Generates docs/rules/<CODE>.md from the rule registry, so the per-rule pages,
`archforge explain`, and SARIF helpUri all share one source of truth (0.6.1).

Usage: python scripts/make_rule_docs.py
"""
import io
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src"))

from archforge.rules import RULES, TITLES, PROFILES   # noqa: E402
from archforge.messages import MESSAGES, set_lang     # noqa: E402

MEANING = {
    "E1": "The font that will actually render this text has no glyphs for its script, so PowerPoint silently substitutes an OS fallback (Malgun Gothic for Hangul). No error is raised anywhere. Resolution follows the measured model: run a:ea > paragraph pPr/defRPr > lstStyle inheritance chain (shape > layout placeholder > master placeholder > master txStyles > defaultTextStyle) > theme ea (majorFont for title placeholders, minorFont otherwise; a non-empty theme ea beats run a:latin) > run a:latin only when the theme ea slot is empty > OS fallback.",
    "E2": "A dash-family character (em/en/figure dash, horizontal bar, math minus, fullwidth hyphen, box-drawing line) is used as sentence punctuation: the single most reliable tell of AI-generated copy. Numeric ranges (2020-2024, Q1-Q3, 5%-10%) and minus signs before digits pass by default; --e2-no-exemptions blocks everything.",
    "E3": "The effective font size, after autofit scaling and the full placeholder inheritance chain, is below the hard floor (default 5pt): unreadable on any projector.",
    "E4": "Positive letter-spacing (tracking) on a run containing Hangul: Hangul letterforms visibly drift apart. Kana-containing runs are exempt (tracked kana is normal Japanese practice), and Hanja-only runs are exempt since 0.6.1 (tracked hanzi is legitimate Chinese typography); Hanja still counts toward the consecutive requirement when mixed with Hangul.",
    "W1": "A body-class frame (wide, high character count) renders below the recommended floor (default 9pt).",
    "W5": "No font size exists anywhere in the inheritance chain, so size gates cannot measure this text.",
    "W6": "The same layout skeleton (shape bounding-box signature) repeats across 4+ pages: the recycled-grid tell of generated decks. Tunable via --w6-sim/--w6-cluster.",
    "W7": "Text overlapping a picture renders with a contrast ratio below 2.5 (needs --render with p01.png-style page renders).",
    "W8": "Small CJK text (between the E3 floor and 7.5pt) inside a narrow frame: typical of unreadable labels inside device mockups and cards.",
    "W9": "Three or more accent-colored vertical bars used as list markers: building structure out of color instead of rules, whitespace, and type.",
    "W10": "A hand-drawn diagram (decorative texture marks) is cloned nearly identically across pages.",
    "W11": "AI-tell copy: buzzword clusters anywhere, stock opening cliches on the first pages.",
    "W12": "Footer baselines drift from the dominant baseline across pages: an alignment slip.",
    "W13": "Native PowerPoint shadow/glow/3D effects (2+ on a page): reads dated.",
    "W14": "Most titles are nominal phrases rather than claims (a Korean-title heuristic; titles with figures count as claims). Read the --ghost list top to bottom: it should tell a story.",
    "W15": "The effective glyph areas of two text frames overlap beyond 45% of the smaller one: occlusion or collision. Verify on a render; drop caps and echo typography are auto-excluded.",
    "W16": "Text glyphs or picture ink extend past the canvas edge beyond tolerance. Decorative shape bleed is auto-excluded.",
    "W17": "Text straddles a picture's ink edge (25-75% inside): it will look clipped or split. Captions fully on cards are auto-excluded.",
    "W18": "Some spans or deck-level checks could not run (malformed attributes, vertical text, complex scripts, decode budgets). The result may be incomplete; gate CI with --fail-incomplete.",
}


def main():
    set_lang("en")
    root = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "docs", "rules")
    os.makedirs(root, exist_ok=True)
    for code in sorted(RULES, key=lambda c: (c[0] != "E", int(c[1:]))):
        sev, cat, _mid = RULES[code]
        profiles = sorted(p for p, excl in PROFILES.items() if code not in excl)
        fix = MESSAGES["fix_" + code.lower()]["en"]
        body = "\n".join([
            "# %s: %s" % (code, TITLES[code]),
            "",
            "| | |",
            "|---|---|",
            "| Severity | %s |" % sev,
            "| Category | %s |" % cat,
            "| Profiles | %s |" % ", ".join(profiles),
            "",
            "## What it means",
            "",
            MEANING[code],
            "",
            "## How to fix",
            "",
            fix + ".",
            "",
            "## If you disagree",
            "",
            "False-positive reports with a repro deck are the most valuable contribution "
            "this project takes: use the [FP template](https://github.com/Love-Ash/"
            "archforge/issues/new/choose). Calibration rationale lives in "
            "[docs/CALIBRATION.md](../CALIBRATION.md).",
            "",
        ])
        path = os.path.normpath(os.path.join(root, code + ".md"))
        with io.open(path, "w", encoding="utf-8", newline="\n") as f:
            f.write(body)
        print("wrote", path)


if __name__ == "__main__":
    main()
