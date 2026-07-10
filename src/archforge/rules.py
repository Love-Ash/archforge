# -*- coding: utf-8 -*-
"""Rule registry (0.4.0, phase 1 of the third external review's structural overhaul).

Separates rule metadata (severity, category, message id) and profile definitions from the
check implementation. CLI validation (--skip typos), profiles, and SARIF rules[]
generation all consume this single table. Physically decomposing the check implementation
itself (packaging into rules/typography etc.) is the next phase.
"""

# code -> (severity, category, representative message id)
RULES = {
    "E1": ("error",   "typography", "e1_nofont"),
    "E2": ("error",   "style",      "e2"),
    "E3": ("error",   "typography", "e3"),
    "E4": ("error",   "typography", "e4"),
    "W1": ("warning", "typography", "w1"),
    "W5": ("warning", "typography", "w5"),
    "W6": ("warning", "structure",  "w6"),
    "W7": ("warning", "render",     "w7"),
    "W8": ("warning", "typography", "w8"),
    "W9": ("warning", "style",      "w9"),
    "W10": ("warning", "structure", "w10"),
    "W11": ("warning", "style",     "w11_buzz"),
    "W12": ("warning", "structure", "w12"),
    "W13": ("warning", "style",     "w13"),
    "W14": ("warning", "style",     "w14"),
    "W15": ("warning", "geometry",  "w15"),
    "W16": ("warning", "geometry",  "w16"),
    "W17": ("warning", "geometry",  "w17"),
    "W18": ("warning", "meta",      "w18_page"),
}

ALL_CODES = frozenset(RULES)

# Profile = engine execution policy (since 0.3.1, excluded rules simply do not run).
# 0.4.0: the default profile changed to core (a breaking change). Only objective defects
# run by default; AI-tell/house-style rules (E2, W6, W9-W14) are full opt-in.
PROFILES = {
    "full": frozenset(),
    "core": frozenset({"E2", "W6", "W9", "W10", "W11", "W12", "W13", "W14"}),
    "editorial": frozenset({"W6", "W14"}),
}

DEFAULT_PROFILE = "core"


def severity(code: str) -> str:
    return RULES.get(code, ("warning",))[0]


def category(code: str) -> str:
    return RULES.get(code, ("", "unknown"))[1]
