# Roadmap to 1.0

Archforge is a strong beta. 1.0 is not "more rules"; it is the point where the internal
model is unified, the accuracy claims are independently verifiable, and there is
evidence of outside use. This page is the honest sequence and the bar for each step.
It is shaped by iterative adversarial reviews whose findings were reproduced as
fixtures and are documented in the CHANGELOG and ADRs (review-driven hardening; no
independent third party has audited the project yet, and this page does not claim one
has).

Development is currently maintainer-led and fast. Roadmap issue #5 is tracking, not
up-for-grabs work; contributor-scoped tasks are labeled `good first issue`. If you want to
take one, comment first and it will be held for you, and a held claim lapses after 14 days
without visible progress so an issue never sits reserved indefinitely (see
[CONTRIBUTING](../CONTRIBUTING.md)).

## Version sequence

| Release | Theme | Bar to ship |
|---|---|---|
| 0.6.x | Contract hardening (done) | Action/scan/policy/geometry contracts consistent; NaN, trust boundary, incompleteness all closed |
| 0.7.0 | Contracts done (#6); architecture started (#5) | Shipped: structured `Finding.data`, JSON schema 2.0 (`findings[]`, `capabilities`, `abstentions`), baseline v3 identity, and the `scripts.py` parsing-layer extraction. Verdict-preserving (16-deck A/B and corpus identical). Continuing under #5: the physical split of the interleaved OOXML/resolution/detector body into one document model + one resolver, done as its own verified effort rather than a big-bang. |
| 0.8.0 | Verification and structure | Shipped: typed `Finding.data` at detection sites, formal JSON Schemas (schemas/) validated in tests, the regression-corpus record published as docs/ACCURACY.md (per-gate counts with exact binomial lower bounds) behind a CI drift gate, corpus grown to 19 manifests, baseline artifact identity + `baseline inspect`, and two more #5 kernel extractions (fonts.py, dashes.py). Still open for 0.8.x+: third-party generator exports (Google Slides / Canva / LibreOffice), multi-master corporate templates, JP-CN native review, HTML reporter (#4), renderer matrix expansion |
| 0.9.0 | Release candidate | Shipped: the #5 physical decomposition (`lint.py` 3,412 -> 908 lines across six modules plus `cli.py`), the star-import surface declared with an explicit `__all__` instead of inherited from whatever lacked an underscore, [docs/DEPRECATION.md](DEPRECATION.md), and repairs to three checks that had been passing while blind. Schemas were already frozen and validated in 0.8.0. The 2-4 week soak starts here rather than gating it. Open at ship: #11 (abstention `affected_rules` understates the guard's reach; a machine-contract defect, not a verdict one, so not a P0). |
| 1.0.0 | Stable contracts | 3+ external contributors; outside false-positive fixtures; used in 2+ generators' pipelines; docs/code auto-consistency checks; RC soak clean |

## Where 0.6.x already landed vs the 10-point bar

The review's per-area "10-point" conditions, and where each stands now:

| Area | 10-point condition | Now |
|---|---|---|
| Problem definition | mechanical preflight vs editorial policy cleanly split | core/full/editorial + ADR 001; README states the split |
| Accuracy | public corpus + per-gate FP/FN | public `corpus/` with manifests in CI; per-gate precision/recall still to publish (0.8) |
| Engine | one document/resolution model | monolith today; the 0.7 target (#5) |
| Structure | detector/reporter/CLI separated | partial (rules/reporters split); full split is 0.7 |
| Testing | property, fuzz, corpus, renderer | property + deterministic fuzz + corpus landed; renderer tests need COM/LibreOffice CI (0.8) |
| CLI | rule discovery, subcommands, stable JSON | `rules`/`explain`/`lint`/`scan`/`demo`; single `findings[]` JSON is 0.7 schema 2.0 |
| Action | typed validation, outputs, PR summary, changed-only | all shipped in 0.6.1 |
| Security | budgets, timeout, controlled failure | zip preflight + `--timeout` + honest "not a sandbox"; deeper budgets 0.8 |
| Docs | per-rule pages, versioned docs, executable examples | per-rule pages + HOW_IT_WORKS + ADRs; versioned docs site is 0.8 |
| i18n | report locale vs policy locale | report `--lang` split done; policy-locale (E4-HAN variants) is a candidate |
| Baseline | artifact/policy identity, approvals | v2 today; v3 identity is #6 |
| SARIF | related locations, confidence, stable fingerprints | static titles, helpUri, partialFingerprints done; confidence/related next |
| OSS ops | governance, ADR, recognition | GOVERNANCE + 4 ADRs + label taxonomy + credit in changelog |
| Release | trusted publishing, provenance, deprecation | trusted-publishing workflow in place; provenance/deprecation at RC |
| Recognition | tech writeups, public benchmark, adoption | writeups drafted; corpus public; adoption is the open frontier |

## The honest gap

Measured against the 1.0 row on 2026-08-08, at the 0.9.0 ship: external contributors 1 of 3
(the W14 English work, #7/#9); outside false-positive fixtures 0, since the Google Slides deck
in the corpus is a round trip the maintainer ran rather than a report from a user; known
pipeline adoption 0. The docs/code consistency checks are in CI. So three of five are open,
and none of the three is closed by writing code.

1.0 is blocked less by code than by (a) unifying how the codebase reads a pptx (0.7),
(b) independently verifiable accuracy (corpus growth beyond author-written fixtures),
and (c) real outside adoption. Adding a W19 does not move any of these.

## Rule candidates (not 1.0 blockers)

Tracked here so they are not lost. Per the note above, none of these gate a release.

- Footer-zone rule, parked on the `w19-footer-zone` branch (2026-08-02). The engine
  wiring and a contrast ink-contamination guard were written against a live deck and
  then parked rather than shipped: no fixtures, no rule page in any inventory,
  unrecorded thresholds, and a coordinate bug where `_footer_rules_y` reads
  pre-transform local coordinates and compares them against absolute glyph boxes.
  `W19_CONTINUATION.md` on that branch carries the blocking work. It needs a new rule
  ID before it can ship: W19 was reused in 0.10.0 for the solid-fill contrast gate, so
  the parked branch and main now disagree about what W19 names.

W14-EN used to sit in this list and no longer does, because it shipped in 0.9.0. An
English title counts as a claim when it carries a finite verb from a measured allowlist
or a number with a unit; bare noun phrases do not; structural titles are excluded from
the eligibility pool in both languages; and the majority gate and profile behaviour are
unchanged from the Hangul path (#7, #9, the first outside contribution, by
@AshSgDe29071999).
