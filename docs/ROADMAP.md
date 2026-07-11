# Roadmap to 1.0

Archforge is a strong beta. 1.0 is not "more rules"; it is the point where the internal
model is unified, the accuracy claims are third-party reproducible, and there is
evidence of outside use. This page is the honest sequence and the bar for each step.
It is shaped by an external review that scored 0.6.0 at 8.0/10.

Development is currently maintainer-led and fast. Roadmap issues (#5, #6) are tracking,
not up-for-grabs work; contributor-scoped tasks are labeled `good first issue`. If you
want to take one, comment first and it will be held for you.

## Version sequence

| Release | Theme | Bar to ship |
|---|---|---|
| 0.6.x | Contract hardening (done) | Action/scan/policy/geometry contracts consistent; NaN, trust boundary, incompleteness all closed |
| 0.7.0 | Contracts done (#6); architecture started (#5) | Shipped: structured `Finding.data`, JSON schema 2.0 (`findings[]`, `capabilities`, `abstentions`), baseline v3 identity, and the `scripts.py` parsing-layer extraction. Verdict-preserving (16-deck A/B and corpus identical). Continuing under #5: the physical split of the interleaved OOXML/resolution/detector body into one document model + one resolver, done as its own verified effort rather than a big-bang. |
| 0.8.0 | Verification and structure | Shipped: typed `Finding.data` at detection sites, formal JSON Schemas (schemas/) validated in tests, per-gate precision/recall published as docs/ACCURACY.md with a CI drift gate, corpus grown to 19 manifests, baseline artifact identity + `baseline inspect`, and two more #5 kernel extractions (fonts.py, dashes.py). Still open for 0.8.x+: third-party generator exports (Google Slides / Canva / LibreOffice), multi-master corporate templates, JP-CN native review, HTML reporter (#4), renderer matrix expansion |
| 0.9.0 | Release candidate | JSON + baseline schemas frozen; deprecation policy; public API surface pinned; 2-4 week RC soak; zero open P0 |
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

1.0 is blocked less by code than by (a) unifying how the codebase reads a pptx (0.7),
(b) third-party-reproducible accuracy (0.8 corpus growth + published precision/recall),
and (c) real outside adoption. Adding a W19 does not move any of these.

## Rule candidates (not 1.0 blockers)

Tracked here so they are not lost. Per the note above, none of these gate a release.

- W14-EN: extend the W14 nominal-title heuristic (currently Korean-only) to English.
  Field evidence: English-deck users report the same failure shape, fragment titles
  that name a topic without making a claim ("Eight rungs, gates not dates";
  r/ClaudeAI, 2026-07). English needs its own claim/verb detection, so it ships only
  with its own measured precision bar, like the Korean heuristic did.
