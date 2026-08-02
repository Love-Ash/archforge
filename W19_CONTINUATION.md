# W19 footer-zone rule: what is left before this can merge

Branch-only file. Delete it in the commit that closes the last item below.

This branch holds the W19 rule and the `contrast_check` ink-contamination guard, both
written on 2026-07-19 from a live deck and then left uncommitted on `main` for three
weeks. It was parked here on 2026-08-02 so that `main` stops carrying an unshippable
working tree and a CHANGELOG entry that claimed test coverage which did not exist.

The engine wiring is done and correct. Everything outside the engine is not.

## What is already done

- `rules.py`: `RULES["W19"]` and `TITLES["W19"]`.
- `messages.py`: `w19` message (ko/en) and `fix_w19`.
- `findings.py`: `_DATA_FIELDS["w19"] = ("below_in",)`.
- `lint.py`: `_footer_rules_y`, `footer_zone_check`, the call site in `lint()` guarded by
  `"W19" not in excl`, the `w19` skip-reason entry, the rule docstring block, and
  `_REASON_RULES["glyph_boxes"]` now lists W19 so a page whose glyph boxes failed reports
  W19 as affected instead of reading clean.

## Blocking correctness

1. **Coordinate-system mismatch.** `_footer_rules_y` iterates `iter_shapes` and reads
   `sp.top` / `sp.width` / `sp.height` directly, which are pre-transform local
   coordinates. `footer_zone_check` then compares those values against the glyph boxes
   built by `_text_glyph_boxes`, whose own docstring states they are absolute coordinates
   derived from `iter_shapes_geo`. When either the footer hairline or the body text sits
   inside a group (grouping a footer rule with a logo is ordinary deck construction), the
   two operands have different origins and the comparison is wrong in both directions.
   Fix: rebuild `_footer_rules_y` on `iter_shapes_geo` so both sides are absolute. The
   group-affine bug already fixed once in this codebase (grpSpPr being matched in the
   wrong namespace) is the precedent for how badly this fails silently.

## Blocking evidence (CONTRIBUTING bar)

2. **No fixtures exist.** `grep -rni "w19|footer_zone|below_in" tests/ corpus/ examples/
   benchmarks/` returns zero. The full test suite passes with the rule present, which
   proves only that nothing else broke. Needed, each with a manifest, in the style of the
   existing geometry fixtures:
   - positive reproduction: content-class text pushed below a qualifying footer hairline;
   - negative: the same text above the rule;
   - negative: an underline decoration, to exercise the `has_footer_text` branch that is
     supposed to stop underlines from being read as footer rules;
   - negative: a footer hairline and its text inside a group, which is the regression test
     for item 1 and must be written before item 1 is fixed, not after.
3. **Thresholds are unrecorded.** `0.85 * sh`, `0.05in`, `>= 50%` width, `10.5pt`, and the
   `0.03` tolerance appear only in code. Every other gate in this project records its
   numbers in `docs/CALIBRATION.md` with the render evidence behind them (W17's 25-75%
   band is the model to copy). The original 2026-07-19 deck is the evidence; without it
   these numbers cannot honestly be defended, so recover that deck or re-derive them.

## Blocking the contrast guard specifically

4. The guard changes the behaviour of W7, which shipped in 0.8.1. Its thresholds (0.06
   luma separation, 20-sample floor, `//6` margin, `//32` ring step) are likewise absent
   from `docs/CALIBRATION.md`. Two tests are needed: one reproducing the original bug (a
   display-size glyph contaminating its own background quantile) and one proving the ring
   fallback does not erase a genuine low-contrast positive. The existing W7 test plants a
   flat single-colour PNG, so inside and outside samples are identical and neither branch
   is exercised.

## Blocking documentation surfaces

5. `scripts/make_rule_docs.py` has no `MEANING["W19"]`, and `main()` indexes
   `MEANING[code]` while iterating `sorted(RULES)`, so the generator raises `KeyError` the
   moment W19 is in the registry. Add the entry, then run the generator to produce
   `docs/rules/W19.md`.
6. Rule inventories that stop at W18 and need the new row: `README.md`, `README.ko.md`,
   `llms.txt`, and both copies of the skill (`skills/archforge-pptx-lint/SKILL.md` and
   `src/archforge/skills/archforge-pptx-lint/SKILL.md`, which a test asserts are
   identical). `docs/HOW_IT_WORKS.md` and `docs/CALIBRATION.md` both describe the geometry
   gates as W15-W17 and need the range widened.
7. `docs/ROADMAP.md` line 51 currently reads that adding a W19 does not move any of the
   three real 1.0 blockers. That sentence was written before this rule existed and is
   still true; if W19 merges, reword it so it does not read as a stale reference to this
   specific rule while remaining honest about rules not being the blocker.

## The check that would have caught most of this

`main` now runs a rules/docs drift gate in CI (`scripts/check_rule_docs.py`). It fails
when the `RULES` registry and `docs/rules/*.md` disagree. Rebase this branch onto `main`
before continuing: the gate will fail until item 5 is done, which is the intended
behaviour and the reason items 5 and 6 went unnoticed for three weeks.
