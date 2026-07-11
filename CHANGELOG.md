# Changelog

## 0.8.0 (2026-07-11)

Verification and structure: the release makes the accuracy claims third-party
checkable and continues the #5 decomposition. Verdict-preserving throughout (16-deck
A/B and corpus identical).

### Typed Finding.data

Detection sites now emit structured data directly instead of leaving it to positional
derivation: E1 carries `script` / `effective_font` / `font_source` / `fallback_font`,
E2 the offending code points and function, E3 `nominal_pt` + `autofit_scale` +
`size_source`, E4 unit-explicit tracking (`tracking_raw_hundredths_pt` and
`tracking_pt`), W16 the target `kind` (text | picture). The positional derivation
stays as a fallback; explicit values win.

### Formal JSON Schemas

`schemas/` holds machine-validatable JSON Schemas for report-1.0, report-2.0,
scan-2.0, and baseline-3; tests validate real CLI output against each, and the schemas
ship in the sdist. `schema_version` now points at something a consumer can actually
validate with.

### Accuracy record

The public corpus grew to 19 manifests (new: an en-dash range negative for E2's
exemption contract, W1 body floor, W8 narrow-frame CJK, W6 repeated skeleton), and
`scripts/make_accuracy_table.py` generates docs/ACCURACY.md: per-gate TP / FP / FN /
deck-level TN with precision and recall, measured through the CLI on the shipped
corpus. CI regenerates the table and fails on drift, so the published numbers cannot
go stale. Scope is stated honestly: a seed corpus proving constructed ground truth,
not yet a field-scale benchmark.

### Baseline artifact identity and inspect

A baseline records the deck it was written from (file name + content hash);
applying it to a differently-named deck warns. `archforge baseline inspect PATH`
(with `--json`) shows what a baseline suppresses per code and under which recorded
conditions, ending the opaque-blob era.

### Architecture (#5, continued)

Extracted the E1 font kernel (`fonts.py`: blocklist, slot parsing, theme-token
resolution, the measured E1 decision) and the E2 dash kernel (`dashes.py`) as pure
modules with no Finding and no I/O, re-exported from lint for compatibility.
lint.py is down from ~3350 to ~3050 lines with two more pure seams established.

## 0.7.1 (2026-07-11)

A contract-integrity release driven by an external review of 0.7.0 (82/100: right
direction, but the new contracts were not yet fully consistent). No new rules; the point
is that every contract holds across all reporters and modes. Verdict-preserving: the
16-deck A/B and the public corpus stay identical.

### Contract consistency

- `scan --schema 2` now declares a `scan-2.0` root with `kind` and `file_schema_version`,
  instead of a `1.0` root over `2.0` file objects a consumer would misparse.
- JUnit reflects `--fail-incomplete`: W18 becomes a `<failure>`, so the report matches
  the CLI exit code instead of reading green while the run exited 1.
- Baseline v3 stops trusting generator auto-names (`TextBox N`, `Google Shape;...`) as
  identity and uses the bbox grid for them, so a defect that moves to a far-away box is
  no longer silently re-suppressed.
- Every skip-reason key is registered in the capability map, enforced by a test, so a
  structural abstention never lands with no affected rules while `structure` reads
  `complete`.
- SARIF `partialFingerprints` carry `archforgeFinding/v3` (the structural fingerprint the
  baseline uses) so code scanning and the baseline agree on identity; pair findings use
  the SARIF-standard `relatedLocations`.
- schema 2.0 gains an `invocation` block (profile/policy/config/thresholds) and a `rules`
  split (executed / profile_excluded / user_suppressed).

### Robustness and packaging

- `--timeout` rejects `inf`/`nan` (inf passed `>0` and silently disabled the bound). A
  `--sarif`/`--junit`/`--write-baseline` path under a missing directory is a controlled
  exit 2 up front, not a traceback mid-run. `lint()` maps an un-parseable package to a
  ValueError instead of leaking an lxml exception.
- The fuzz test now fails on any non-ValueError (it previously allowed
  KeyError/TypeError/etc.), which caught a real open-path crash now fixed. The corpus
  runner drives the CLI to check the exact exit code and verifies incompleteness in both
  directions.
- The sdist now includes `action.yml`, `action_runner.py`, and the corpus, and excludes
  `node_modules`, so the published source passes its own full suite (a CI job now
  installs the sdist and runs its tests). The three repo-root-asset tests skip on a wheel
  install rather than fail.

### Docs

- Version references in the README, pre-commit config, and Action example moved to
  v0.7.1; the config module docstring and ADR 004 updated to baseline v3; GOVERNANCE
  scopes the "batched releases" policy to 1.0+ so it no longer contradicts fast pre-1.0
  iteration.

### Deferred (0.8, per the review's own sequencing)

- Full typed `Finding.data` with detectors emitting structured data directly (E1/E2
  payloads, E4 units, E3 nominal/scale, W16 kind), the physical module decomposition
  (#5), artifact-level baseline identity, formal JSON Schema files, and corpus growth
  with per-gate precision/recall.

## 0.7.0 (2026-07-11)

The contract half of the 0.7 architecture work (issue #6), plus the first step of the
physical decomposition (#5). The refactor is verdict-preserving by construction: the
16-deck A/B and the public corpus stay byte-identical, and schema 1.0 output is
unchanged.

### Structured data and JSON schema 2.0

- `Finding.data()` derives a structured numeric payload (effective_pt, overflow_in,
  contrast_ratio, tracking, ...) from the args the detectors already pass, via a
  per-message-id field table. No detection site changed; the verdict is untouched.
- `--schema 2` emits schema 2.0: a single `findings[]` array with `severity` and `data`
  on each item, a `capabilities` map (typography / geometry / structure /
  render_contrast: complete | partial | not_requested), and a structured
  `abstentions[]` (reason, page, count, affected_rules) derived from the W18
  machine-key reasons. Schema 1.0 stays the default and byte-identical for existing
  consumers.

### Baseline v3

- Fingerprints gain a page-free structural location bucket (normalized shape name,
  cell, paragraph, or a coarse bbox grid), so a defect that disappears and reappears in
  a genuinely different place is treated as new rather than silently re-suppressed
  (external review). A threshold hash is recorded and checked on load. v1/v2 baselines
  are rejected with a regenerate message: the single migration ADR 004 committed to.

### Architecture (#5, staged)

- Extracted the per-run Unicode script layer into `scripts.py` (pure, dependency-free),
  establishing the parsing-layer boundary and re-exported for compatibility. The
  remaining physical decomposition of the interleaved OOXML / resolution / detector body
  continues under #5 as its own verified effort rather than a single big-bang, since the
  verdict-preserving contract is the acceptance test and a move that large cannot be
  proven safe against the A/B and corpus alone in one pass.

## 0.6.2 (2026-07-11)

Reporters, reproducibility, and resource bounds. Additive; no contract changes.

- JUnit XML reporter: `--junit PATH` in both single-file and scan modes. One
  testsuite per file, one testcase per executed rule (excluded rules map to
  skipped, since JUnit's skipped means "not run"), ERROR findings as failures,
  WARN findings in system-out unless a warn-failing policy is active, and an
  unreadable file as an error testcase. The implementation follows the design
  discussed with @siddhanttiwari19 in issue #2.
- Public seed corpus under `corpus/`: decks from three generators (python-pptx,
  PptxGenJS, PowerPoint native re-save) plus a malformed set, each committed with a
  manifest of expected findings. `corpus/run_corpus.py` enforces the match and runs in
  CI. This is the externally reproducible accuracy record that the private real-deck
  calibration corpus could not be.
- Property-based tests and a deterministic structured fuzzer (`tests/test_properties.py`,
  hypothesis): determinism, translation invariance of which rules fire, E3 size
  monotonicity, and a no-unhandled-crash guarantee under random attribute mutation.
- `--timeout SECONDS`: a wall-clock limit for the whole run, isolated in a child
  process (portable to Windows, unlike signal.alarm), exit 124 on timeout. Keeps a
  hostile or pathological deck from hanging CI.

## 0.6.1 (2026-07-11)

A contract-consistency release driven by an external review of 0.6.0 (overall verdict:
the integration layer is now sound; the remaining risks are policy/messaging mismatches
and the monolith). Every reproducible finding was verified against the code before
fixing.

### Policy and contract fixes

- E4 now requires actual Hangul in the run. Tracking on a Hanja-only run is legitimate
  Chinese typography; flagging it as a universal ERROR contradicted the "other scripts
  are never falsely flagged" scope promise. Hanja still counts toward the consecutive
  requirement when mixed with Hangul (Korean names, legal terms), which keeps the
  Korean-deck coverage the rule was built for.
- `summary.policy` records the active failure policy (fail_on_warning /
  fail_incomplete / e2_no_exemptions): identical counts can pass or fail depending on
  flags, and JSON consumers could not tell why. All integration docs (README, skill
  pack, agent loop) now gate on `summary.pass` with `--fail-incomplete`.
- scan validates global CLI values once up front (exit 2) instead of degrading a bad
  flag into N identical per-file error entries; per-file problems (corrupt deck, that
  deck's config/baseline) still isolate per file.
- scan tracks match counts per input pattern: one input matching nothing exits 2 even
  when another matched (a typo'd glob could hide behind a populated directory).
  `--allow-empty-pattern` opts out; the aggregate JSON records
  `scan.inputs[].{pattern,matches}`.
- The GitHub Action validates boolean inputs exactly (a typo like `ture` fails the job
  instead of silently disabling a safety default) and rejects unknown `source` values.
  The runner also emits step outputs (passed, error-count, warning-count, incomplete,
  checked-files, failed-files), writes a GITHUB_STEP_SUMMARY table, and gains
  `changed-only`/`base-ref` mode for large repositories.
- Config numeric strictness: booleans are rejected as thresholds (float(True) == 1.0
  used to pass) and `w6_cluster` must be integral (1.9 silently truncated to 1).

### Correctness

- The geometry complex-script screen now reads all `a:t` descendants, closing a bypass
  where field-only RTL text entered width math with the Latin/CJK model and no W18.
- An explicit run color the decoder cannot resolve (hslClr, scrgbClr, sysClr, prstClr,
  transforms) stops color resolution and abstains into W18 instead of falling through
  to an inherited color (reproduced false positive: explicit white hslClr judged with
  inherited black).
- Geometry locations carry `paragraph` indexes and `field: true` for field-only
  paragraphs, matching the run-level location contract.
- SARIF E1's static title no longer claims Hangul-only scope (E1 also covers
  kana/hanzi on fonts with no CJK glyphs), and every rule's `helpUri` points to its
  own page under docs/rules/.

### New

- `archforge rules` (one-line rule list) and `archforge explain CODE` (meaning,
  profiles, fix, doc link) for rule discovery without the README; `archforge lint`
  as an explicit alias for single-file mode; `python -m archforge` entry point
  (removes the runpy RuntimeWarning the Action used to print).
- Per-rule documentation pages generated from the rule registry
  (`docs/rules/E1.md` ... `W18.md`, `scripts/make_rule_docs.py`).
- Project governance docs: GOVERNANCE.md and four ADRs (profiles, incompleteness
  contract, target renderer, baseline identity).
- A PyPI trusted-publishing workflow (`publish.yml`, OIDC; activates once the PyPI
  publisher is registered).
- CALIBRATION gains an honest renderer-coverage matrix (measured vs unknown vs
  known-different per renderer).

### Docs

- The E1 resolution priority now includes the paragraph `pPr/defRPr` step everywhere
  it is described (README en/ko, skill pack, CALIBRATION's derived summary); it was
  measured in probe 7 but had drifted out of two of the four descriptions.
- The bundled skill pack drops the "deep CJK coverage" phrasing for the accurate
  "Hangul-deep, CJK-aware", and its agent loop keys on `summary.pass`.

### Deferred with reasons (unchanged from the review's own sequencing)

- JUnit reporter: assigned to an external contributor (#2); shipping it ourselves
  would take their first contribution.
- Canonical document model, detector decomposition, structured `Finding.data`,
  baseline v3 identity: the 0.7 architecture pass, done as its own release with its
  own verification cycle rather than appended to a contract-fix release.
- Public multi-generator corpus, property/fuzz testing, versioned docs site, worker
  timeouts: 0.8 material; each needs assets or infrastructure that cannot be conjured
  credibly in a hardening release.

## 0.6.0 (2026-07-11)

A hardening release for the integration layer. An external structural review of 0.5.0
and an independent verification session agreed on the diagnosis: the engine is solid,
but the new delivery surfaces (GitHub Action, scan) shipped with trust and failure-
semantics gaps. This release fixes those before wider announcement, plus the geometry
model's largest remaining error sources.

### GitHub Action (rewritten)

- Pinning the action tag now pins the linter: the default install is the action's own
  checked-out source (`source: action`), not PyPI latest. `source: pypi` plus `version`
  remains as an explicit opt-in.
- Inputs travel through env into a Python runner and become argv entries directly;
  nothing is interpolated into a shell script body. `files` and `extra-args` are one
  entry per line, so paths with spaces survive and globs (including `**`) are expanded
  by `archforge scan` itself rather than the shell.
- Deck-folder configs are ignored by default (`allow-config: false` passes
  `--no-config`): a PR cannot weaken the CI gate by committing a config next to its
  deck. Incomplete checks fail by default (`fail-incomplete: true`).
- The action is now itself CI-tested (`uses: ./`): newline file lists, a directory with
  spaces, glob inputs, SARIF output, failure propagation, and the config trust boundary.

### scan batch semantics

- One broken or misconfigured file no longer aborts the batch: usage errors become
  per-file `status: "error"` entries (text, JSON `files[]`, and `summary.error_files`),
  and the remaining files are still checked. Exit stays 1 when anything failed.
- A single CLI `--baseline` across multiple files is refused (fingerprints carry no
  file identity, so deck A's acceptances would suppress deck B's findings); per-deck
  baselines via deck-folder configs are unaffected.
- A deck config's `lang` no longer leaks across later files and the aggregate report
  (the report renders in the invocation language).
- Directory walks sort both directories and files: aggregate output order is now
  deterministic across filesystems.

### Validation and input hardening

- NaN no longer bypasses threshold validation: `--hard-min nan` and a bare `NaN`
  literal in `.archforge.json` (which `json.load` accepts) both exit 2. This closed a
  re-opening of the exact `--hard-min 0` bypass fixed in 0.5.0.
- All flag and config validation now runs before linting and before `--write-baseline`
  records anything (a typo'd `--skip` used to record a baseline as if valid).
- `--config` combined with `--no-config` is a contradiction and now errors instead of
  silently dropping the explicit config.
- A zip preflight bounds entry count (20k), total uncompressed size (2GB), and
  per-entry compression ratio before python-pptx parses untrusted input; violations
  are usage errors with a stated reason. Non-budget image decode failures now count
  into W18 (`image_decode`) instead of being silently swallowed.
- Baseline run-condition metadata is checked on load, not just recorded: a profile or
  tool-version mismatch prints a warning.
- W6/W10 fingerprints no longer embed page numbers (they violated the
  page-independent fingerprint contract).

### Strict policy split

- `--fail-on-warning`, `--fail-incomplete`, and `--e2-no-exemptions` are now separate
  flags; failing every advisory and failing on unchecked spans are different policies.
  `--strict` remains as the union (compatibility alias). `summary.pass` reflects the
  active policy.

### Geometry model

- Text-frame insets (lIns/tIns/rIns/bIns, OOXML defaults applied, group-scale aware)
  now offset and shrink the usable frame area: glyph boxes no longer start at the
  frame edge or overstate usable width.
- Explicit line breaks (`a:br`) start a new visual line in geometry (previously a
  br-split sentence was measured as one overlong line), and field text (`a:fld`)
  occupies real width.
- Table geometry uses real column widths, skips merged-region continuation cells, and
  spans the origin cell across its merged columns/rows; glyph boxes carry
  `cell`=[row, col] into W15-W17 locations (including `related`).
- Rotated text frames are documented as out of scope for geometry by design and do not
  mark the result incomplete (the previous docs implied otherwise).

### SARIF

- Rule `shortDescription` uses static titles (parameterized `%.1fpt` templates no
  longer leak into code-scanning UIs), plus `helpUri` and `defaultConfiguration`.
- Every result carries `partialFingerprints` (the baseline fingerprint), so GitHub
  code scanning can track findings across runs instead of re-opening them each push.

### CLI and metadata

- `archforge --version`.
- PyPI metadata: project URLs (Documentation/Changelog/Issues/Source) and the English
  natural-language classifier alongside Korean.
- CI additionally verifies the built wheel end to end (`demo`, single-file fail/pass,
  `scan` with aggregate JSON and multi-file SARIF) and runs `twine check`.

### Docs

- Removed the "every gate" overclaim from the examples pointer, replaced "deep CJK
  coverage" with the accurate "Hangul-deep, CJK-aware" phrasing, fixed the `W1..W18`
  code-range notation (W2-W4 do not exist), fixed CONTRIBUTING's install command to
  include the test extra, and updated the stale v1 fingerprint description in
  config.py's module docstring.

### Deferred (recorded, not shipped)

- Structured `Finding.data` (numbers instead of pre-rendered message args) and
  location-bucketed fingerprints: both change public contracts and deserve their own
  design pass.
- W7 sharing the effective-glyph geometry and per-run colors; an embedded-font
  cmap-based E1; detector module decomposition; a multi-generator public corpus. These
  are the 0.7 candidates, deliberately not squeezed into a hardening release.

## 0.5.0 (2026-07-10)

This is a deployment and onboarding release. It addresses the bottleneck that all 3
fourth-round reviews flagged in common (delivery, not the engine): 30-second onboarding
(demo), CI integration (scan, GitHub Action, pre-commit), a completed location contract
for agent auto-fix, and a full README overhaul.

### Added

- `archforge demo`: generates a broken.pptx seeded with 6 defect types and a corrected
  fixed.pptx, then lints them on the spot (first-run experience). Added 3 committed
  variants built from the same source (broken/fixed/style_warnings) plus expected-output
  docs to the repo's `examples/`.
- `archforge scan PATHS...`: multi-lint over a mix of files, directories (recursive), and
  globs. Per-file verdicts go through the same path as single-file mode, and it emits
  aggregate JSON (`files[]` plus an aggregate summary) and multi-file SARIF (one run,
  per-file artifactLocation). Exit 1 if any file fails; zero matches is not a silent pass
  but exit 2 (guards against CI footguns). PowerPoint lock files (~$*.pptx) are excluded.
- GitHub Action (composite `action.yml`): `uses: Love-Ash/archforge@v0.5.0` with
  files/profile/strict/sarif/version inputs. Also added a pre-commit hook
  (`.pre-commit-hooks.yaml`).
- Completed the location contract (target for agent auto-fix): absolute bbox with group
  transforms composed in, table cell `cell`=[row, col], automatic field `field: true`,
  W15-W17's effective glyph bbox plus the pair-verdict `related` (counterpart shape), and
  W7 absolute bbox.
- a:fld (slide number and date fields) now passes through the same gates as regular runs
  (E1/E3/E4). Schema-wise, fld also carries rPr+t and renders under the same rules, but
  python-pptx's run traversal skipped it, leaving a blind spot. Fields are excluded from
  ghost/W14 title collection (prevents large page numbers from contaminating results).
  a:br is now treated as a single line-break character in E2 context and offset
  calculation.
- Community docs: CONTRIBUTING (evidence standard: gates are tuned only against
  rendering comparisons), SECURITY (threat model), 3 issue templates (including a
  false-positive report template), and a PR template.

### Fixed (correctness)

- Latent group affine bug: a slide's grpSpPr is in the p: namespace, but the code was
  finding it with a:, so the group transform always fell back to identity. This caused
  W15-W17 geometry and loc bbox to be judged against raw coordinates in moved, desynced
  groups. Fixed with local-name matching and pinned down with absolute-coordinate tests
  (found through measurement during the 0.5.0 loc work).
- Made the sort key for W15/W16/W17's top-2 selection an explicit single key (overflow
  amount / ratio), removing the case where alphabetical text order broke ties.

### Fourth-round review correctness fixes (unreleased batch accumulated on main)

Correctness-defect fixes from the fourth external review (2 strategy/distribution-
perspective reviews plus 9 CONFIRMED findings from a self-verification session). We
accepted the release-discipline feedback and folded these into this release instead of
bumping the version at the time.

- (HIGH) Config-file trust boundary: added `--no-config`, and the applied config path is
  now always shown in JSON `summary.config` and the text footnote. A means to expose and
  block the path where an attacker-controlled config in the deck folder could silently
  weaken the gates.
- (HIGH) Baseline fingerprint v2 (schema 2, requires regeneration): a locale-neutral key
  (fp_key) removes language dependence, dropping the page number lets it survive slide
  insertion, occurrence-count (count) based multiset semantics, and run-condition
  metadata (tool_version, profile, lang) is recorded. v1 files are now rejected with an
  explicit error. Added a suppression footnote to text output (fixes the case where
  suppressed violations were misread as a clean pass). baseline is now labeled beta.
- Fixed `--help`'s default-profile guidance, which still said the old value (full) (a
  documentation drift where --help was wrong ever since the breaking-change release).
- Config fail-safe: unknown keys (typos) now exit 2, value type and range are validated
  (tracebacks removed), lang is validated, and baseline paths are resolved relative to
  the config file. CLI thresholds are also range-validated (closes the workaround where
  `--hard-min 0` silently disabled E3).
- Fixed the library `lint(profile=...)` silently treating a misspelled profile as full;
  it now raises ValueError.
- Documentation drift: aligned the README/SKILL description of E4 with its actual scope
  (Hangul and Hanja, kana excluded), removed the hardcoded version from the JSON example,
  clarified that `--strict` releasing E2 only matters under full, and removed internal
  decision-making language from CHANGELOG.

Deferred / rejected acceptances (decision record): kept the core default itself as-is
(confirmed by the user in the third round, compensated for with visibility). a:fld/a:br
text not being checked, and raw coordinates for loc bbox inside groups, were published
as known limitations and have now been resolved in this release's Added and Fixed items.

### Docs

- Full README (en/ko) overhaul: real-render before/after comparison images and a demo
  GIF (both produced from actual PowerPoint renders of the examples decks), a 30-second
  onboarding section, a CI section, and location-contract documentation. Added the
  scan/demo and location key contract to SKILL.md.
- Dual-language demo decks: `archforge demo` now builds a deck matching the report
  language (ko/en), and each variant is monolingual. The English deck seeds the
  script-independent defects (E2/E3/W15/W16); the Hangul-only defects (E1/E4) are
  demonstrated by the Korean deck. The English README assets are real renders of the
  English deck, and the Korean README assets are real renders of the Korean deck.
  Added broken_en/fixed_en to examples/.
- Bundled docs/assets/social-preview.jpg (1280x640): for uploading to the repo's
  Settings > Social preview.

### Tests

- 85 to 93 tests: fld gate, br offset, table cell loc, group absolute bbox, W15/W16 loc,
  scan/demo CLI, multi-file SARIF, examples contract.

## 0.4.0 (2026-07-10)

A release reflecting the third external review's restructuring roadmap and default-
profile decision. The purpose of the default change is to remove first-run false
positives and improve onboarding (style rules are now opt-in).

### Breaking changes

- The default profile changed from full to core. Running with no options now checks
  only objective defects (E1/E3/E4, W1/W5/W7/W8, W15-W18); AI-tell and house-style rules
  (E2 dashes, W6 repetition, W9-W14) are opt-in via `--profile full`. This fixes the
  first-impression problem where first-time users got exit 1 for normal punctuation.
  Agent loops that lint machine-generated decks should specify full explicitly
  (SKILL.md's default command has been changed accordingly).

### Restructuring

- Finding model (findings.py): the check engine now produces only locale-neutral
  findings (code + message id + args), and language is decided at the reporter stage.
  Backward compatibility with the old 4-tuple unpacking/indexing is retained. Each
  finding now carries a structured location (location: shape_id, shape_name, bbox, part,
  paragraph, run), letting agents pinpoint the fix target without re-parsing the detail
  string.
- Split out a rule registry (rules.py) and reporters (reporters.py: text/json/sarif).
- Config file: .archforge.json in the deck folder / working folder (.yml if PyYAML is
  installed). CLI flags take precedence over config, and unknown keys are ignored with
  a warning.
- baseline: `--write-baseline` accepts an existing deck's violations and reports only
  new violations afterward (W18 is out of scope, and the suppressed count is recorded
  in summary.baseline_suppressed).
- SARIF 2.1.0 output (`--sarif PATH`): integrates with GitHub code scanning.
- Performance budget: images over 25MP skip alpha trimming (documented), and W6/W10
  pairwise comparison is capped at 200 slides (documented).
- python-pptx upper bound (<2): protects internal API dependencies from an unvetted
  major upgrade.
- benchmarks/: a public harness that generates fixtures from a defect matrix and scores
  per-rule precision/recall reproducibly (also serves as a CI gate). Generator coverage
  starts with python-pptx.

Tests: 76 to 82.

## 0.3.1 (2026-07-10)

A consistency release that fully addresses all 5 P0 and 7 P1 findings from the third
external in-depth review of 0.3.0. The theme is not new features but internal engine
consistency: a single effective document model.

### P0 (release-blocking)

- E1 now reads paragraph pPr/defRPr font inheritance. Confirmed by measurement via COM
  probe 7: a paragraph defRPr's a:ea is actually rendered and beats the lstStyle chain
  (priority: run rPr > paragraph defRPr > lstStyle chain > theme). Missing this step was
  producing both false positives (blocking decks that gave a paragraph a Hangul font)
  and false negatives (passing decks where the paragraph overrides with Latin-only while
  only the upper chain was checked) at the same time.
- W15-W17 geometry checks now resolve sizeless runs with the same StyleResolver as E3.
  This fixes the structural inconsistency where the same document had two different
  effective style models (E3 = 40pt, geometry = default 12pt). False negatives for
  off-slide overflow and collisions on titles with a large inherited size are
  eliminated.
- Native table cell text is now included in geometry checks (cell rectangles are
  computed by accumulating column widths and row heights). Tables in machine-generated
  decks were a W15-W17 blind spot.
- --render contract: a missing folder, or a missing render for a page that has images,
  now surfaces as W18/incomplete. Removed the silent zero-checks path (a deck with no
  images is still judged complete even without renders).
- Profile is now an engine execution policy rather than a post-hoc CLI filter: usable
  from the library via lint(profile=...); excluded rules simply do not run, so their
  O(S^2) comparison cost is not paid, and internal failures in excluded rules no longer
  leak into W18.

### P1

- Strengthened --skip validation: nonexistent codes (typos) are rejected with exit 2,
  and W18 (the incompleteness signal) cannot be suppressed.
- Expanded W7 text-color resolution: beyond direct RGB, it now covers paragraph defRPr,
  the lstStyle inheritance chain, and schemeClr (theme clrScheme resolution). Image and
  text coordinates are now absolute, including group transforms.
- E4's message now matches its actual scope (Hangul and Hanja). Hanja-only-run E4 is
  kept as-is but is a candidate for a future profile.
- Extended i18n to detail and stderr (W6 detail's example markers, W10's page/pages, and
  3 diagnostic notes).
- Unified is_cjk as a composite of is_hangul + is_kana + is_hanja (fixes the
  inconsistency where extended Hangul was dropped from W8).
- ghost title collection now prioritizes the title placeholder over size (fixes the
  case where a 60pt KPI big-number pushed out the actual 26pt title).
- Started a JSON version contract: schema_version ("1.0"), tool{name, version},
  target_renderer ("powerpoint-windows"). Updated SKILL's agent stop condition to
  error_count==0 AND incomplete==false.

Tests: 65 to 76.

## 0.3.0 (2026-07-10)

Globalization release: made report language and rule policy separable layers.

- Message-catalog i18n (en/ko): all gate messages, CLI help, and errors moved into
  English/Korean catalogs. Report language follows the user, not the deck: `--lang` >
  `ARCHFORGE_LANG` > system locale > en. Gate codes (E1, W15) and JSON keys remain a
  language-independent stable contract, and a `lang` field was added to JSON. Matching
  format-argument order between the two language templates is pinned by a test.
- Profiles: `--profile full|core|editorial`. core = objective defects only (excludes
  style/convention rules such as E2), editorial = excludes W6/W14. Unlike `--skip`
  (WARN-only), this is a named public policy, so it is allowed to exclude style-oriented
  ERRORs, but the choice is recorded in JSON `summary.profile` and `skipped_codes` (not
  a silent bypass). Reflects the external strategy review's call to "separate objective
  defects from house style".
- Made README English-first and split out README.ko.md (the PyPI page is English too).
  Documented the render model's target renderer (PowerPoint for Windows) explicitly.
- Updated SKILL.md to the 0.3.0 contract (lang, profile, skipped_codes, E2 v2, Hangul
  scope).
- Incorporated confirmed findings from our own third-round adversarial panel (4
  perspectives):
  - Two-tier script layer: the main blocklist (Inter, Arial, and mono-style families)
    lacks CJK entirely, so it is also valid for kana and Hanja judgments, whereas JP/SC
    subsets (Noto Sans JP and the like) are Hangul-judgment-only. Fixed the regression
    where E1/E4 for Hanja-only runs had been entirely dropped in 0.2.1 (affecting Korean
    decks' personal names and legal terms); only runs mixed with kana are excluded from
    E4. Added halfwidth and Hangul Jamo Extended-A/B to the Hangul Unicode range, and
    added Tibetan, Myanmar, and Khmer to the geometry-skip scripts.
  - Closed 4 E2 gaps: a false negative for two consecutive dashes, an attached-
    parenthetical bypass for word+digit mixed tokens (like "conclusion2024"), superscript
    footnote numerals being misread as digits (numeral detection is now limited to ASCII
    and fullwidth), and a false positive for a spaced negative sign.
  - Restricted title eligibility to an explicit size or a title-family placeholder
    (closes the remaining path where a shape lstStyle or master bodyStyle inherited
    size was sweeping body prose into ghost/W14).
  - When --lang is given repeatedly, the last value now applies; `archforge skill
    --lang` is accepted; and the misparse of `--lang ko skill` ordering is fixed.
    Language state is now held in contextvars (safe for threaded embedding).
  - Added `summary.incomplete`: a machine-readable signal for whether W18 occurred.
    Since pass alone reflects only ERRORs, CI should check pass together with
    incomplete, or use --strict (corrects overstated documentation).
- Tests: 56 to 65.

Accepted limitations (not fixed): argparse's own boilerplate text (usage/error) stays
fixed in English; UTF-8 output may appear garbled on legacy cp949 consoles (crash
prevention takes priority).

## 0.2.1 (2026-07-10)

A safety release that fully addresses the second external re-review of 0.2.0 (17
agents, 7 confirmed findings). The theme is not producing false positives regardless of
what language a deck is in (the script layer).

- Fixed a title-collection flood regression: the defaultTextStyle fallback of 18pt is
  still used for gating but is now excluded from title candidacy (distinguishing the
  source of the size). Fixes a new regression introduced in 0.2.0 where --ghost dumped
  the entire body text and W14 misfired on decks with no explicit size.
- Script layer: restricted E1/E4 triggers to text containing Hangul. Kana-only and
  Hanja-only runs are no longer judged using Hangul coverage knowledge (fixes false
  blocking of Japanese/Chinese decks using Noto Sans JP/SC, etc.; Japanese kana
  letter-spacing is normal practice and is not an E4 target). Vertical writing
  (bodyPr@vert) and RTL/complex-shaping scripts (Arabic, Hebrew, Indic, Thai) now skip
  geometry estimation and surface as W18.
- E1 font inheritance chain: font slots missing from run rPr are now resolved through
  the lstStyle chain (shape -> layout ph -> master ph -> master txStyles ->
  defaultTextStyle). Confirmed by measurement via COM probe (docs/CALIBRATION.md probe
  6) that a master lstStyle's a:ea is actually inherited into rendering, and that title
  placeholders pick up the theme's majorFont ea. Fixes the "confirmed Malgun fallback"
  false positive and the unused-mj-ea problem seen in standard corporate templates.
- E2 v2: shifted the discriminating axis from character to function. An en dash passes
  even when spaced if both surrounding tokens are numeric-ish (contain digits: 2020, Q1,
  5%, FY24); if only one side is numeric-ish, it passes only when unspaced (a range);
  spaced parentheticals and word-joining are still blocked. Fixes 4 remaining
  false-positive patterns: spaced ranges, Q1 ranges, percent ranges, and digit+Hangul
  joins.
- (Breaking change) Restricted --skip to WARN codes only (E codes are rejected with
  exit 2), and recorded what was applied in JSON summary.skipped_codes. Fixes the
  footgun where a release-blocking gate could be turned off without a trace. Pipelines
  using `--skip E2` under 0.2.0 should migrate to `--profile core`.
- Extended W18 wiring across the board: W6/W7/W9-W14 guards, theme parsing failures,
  and even deck-level checks (p00). The documented claim that "anything a guard
  swallows surfaces in the output" is now actually true.
- Blocklist: fonts with "cjk" in the name now always pass (fixes the Noto Sans Mono CJK
  KR false positive); added Aptos (current Office default), Courier, and PT Sans/Serif.
- Tests: 47 to 56.

## 0.2.0 (2026-07-10)

A release that fully addresses an external in-depth review (6 perspectives plus
adversarial re-verification), then runs the fixes themselves back through our own
adversarial panel (8-perspective review plus per-finding rebuttal verification, 22
agents), fixing 11 additional confirmed findings.

### Added in the self-adversarial-panel round

- Added W18: surfaces guard-swallowed, unable-to-check regions (corrupt or non-standard
  attributes) as a warning not just to stderr but also in JSON/text output. Fixes the
  silent degradation where CI checking only `summary.pass` misread an incomplete check
  as a full pass. Escalates to exit 1 under `--strict`.
- Moved the core line-gate guard from frame granularity to run granularity: fixes the
  case where one run's garbage attribute was also swallowing a genuine E1 violation in
  a neighboring run of the same frame, producing a false clean.
- Decoupled the geometry cache: removed the shared gate where a glyph-box computation
  failure was also silencing an unrelated image's W16 (now only the failed axis backs
  off, while the surviving axis keeps checking).
- E2's numeric-context exception is now judged at paragraph scope: fixes a false
  positive for a year range split across a run boundary ("2020" + "(U+2013)2024")
  (PowerPoint frequently splits runs at spell-check and formatting boundaries).
- E1 now judges OOXML theme tokens (like "+mn-lt"/"+mj-ea") after resolving them to
  real fonts: fixes the false negative where a Latin-only theme font fallback went
  undetected because the token was matched against the blocklist literally.
- Removed a double-append path in the W6 slide signature (sig and tokens are now each
  guarded independently): fixes W6 page numbers shifting when token collection failed.
- SizeResolver: moved the placeholder-traversal guard from the whole loop to per-item
  (previously one corrupt placeholder aborted the entire lookup), and added
  memoization for layout/master lookups.
- Running `archforge skill` now prints an stderr notice if a file named "skill" exists
  in the current folder (to lint that file, use `archforge ./skill`).
- If a master's theme parsing fails, stderr now announces the fallback (preserves the
  signal at the point where the distinction between a parse failure and an empty slot
  would otherwise be lost).

### E1: font-detection redesign (measured render model)

- Confirmed the CJK font resolution priority by measurement via a PowerPoint COM
  probe: run `a:ea` > a non-empty theme minorFont `a:ea` (takes priority over run
  `a:latin`) > run `a:latin` only when the theme is empty > OS fallback. Recorded in
  `docs/CALIBRATION.md`.
- Retired the old `ea or latin` proxy judgment. This eliminates both (1) the false
  negative caused by a Latin-only latin font in an empty theme, and (2) the false
  positive raised by looking only at latin when a non-empty Hangul theme ea was
  actually being rendered.
- Expanded the blocklist to roughly 60 Latin families commonly used in decks (Inter,
  Arial, Calibri, Segoe UI, Roboto, IBM Plex Sans, Noto Sans, etc.). Blocked false
  positives on Hangul-complete variant prefixes via an exception list (Arial Unicode,
  IBM Plex Sans KR, Noto Sans/Serif KR/CJK).
- Removed NanumGothicCoding from the blocklist (it was misclassifying a
  Hangul-complete monospace font as Latin-only).
- Theme `a:ea` is now resolved per-master via the slide -> master -> theme
  relationship (rels). Fixes E1 misfiring against an unrelated first theme part in
  multi-master decks. Uses XML parsing instead of byte regex.

### E2: legitimate typography exceptions

- An en dash between digits (U+2013 range notation) and a mathematical minus before a
  digit (U+2212 negative) now pass under the default mode. `--strict` still blocks
  everything with no exceptions (to preserve the original blanket-blocking discipline).

### Size gate: inheritance-chain resolution

- Effective size is now resolved through run > paragraph > shape lstStyle > layout
  placeholder > master placeholder/txStyles > `defaultTextStyle`. E3/W1/W8 now
  actually run on placeholder-and-template decks. W5 fires only when the entire chain
  is silent.

### Robustness

- Added frame/block-level guards to the core line gates (E1-E4) and W6/W10
  clustering. Fixes the case where a single garbage attribute from an external
  generator killed the entire lint run and got mislabeled as exit 2 ("could not open").
- `spc` tracking now absorbs OOXML universal measures ("1.5pt") and garbage values
  (the spc-side counterpart of the same pitfall as the autofit-percentage union).
- Clamped the cosine value for perfectly cloned slides, which could exceed 1.0 due to
  floating-point error.

### Genre and tuning

- `--skip CODES`: selectively suppress warnings that don't fit a genre (e.g., W14 for
  editorial decks).
- `--w6-sim` / `--w6-cluster`: tunable thresholds for W6 skeleton-repetition
  detection.
- W14 now recognizes number+unit titles (e.g., "Revenue grew 3x") as an assertive
  claim.

### Packaging and distribution

- Bundled the skill pack into the wheel (`src/archforge/skills/`, package-data). In
  0.1.0, pip users had no way to obtain SKILL.md.
- Added an `archforge skill` subcommand: print / install via `--install [DIR]` /
  `--path`.
- Renamed the skill directory to `archforge-pptx-lint/` to match the frontmatter name
  (per the Agent Skills spec). Parity between the root `skills/` copy and the packaged
  original is pinned by a test.
- Added GitHub Actions CI: an ubuntu 3.9/3.12/3.13 + windows 3.12 matrix, plus a
  wheel-bundling and install smoke job. Declared `[project.optional-dependencies] test`.

### Performance and architecture

- Per-slide glyph and image ink bbox is now computed once and injected into W15-W17
  (previously recomputed 3 times for text and 2 times via PIL decode for images).
- Extracted E1's judgment into `e1_violation()` and E2's into `dash_violations()` (now
  unit-testable). Geometry boxes use a `GlyphBox` NamedTuple instead of magic-index
  tuples. Added type hints to major public functions.

### Docs and tests

- Added `docs/CALIBRATION.md`: corpus composition and methodology, the E1 measurement
  table, per-gate threshold rationale, and recalibration knobs.
- Updated README and SKILL.md to reflect 0.2.0 behavior (the JSON `ghost` field,
  tunable flags, E2 exceptions, the E1 model, and the skill install path).
- Tests: 18 to 42: E1 model matrix, theme relationships, E2 context/strict, spc
  defenses, partial survival, placeholder inheritance (E3/W1), W5 when
  defaultTextStyle is removed, table cell E1, the exit 2 path, text output, `--ghost`,
  `--skip`, W6 tunables, W14 numeric claims, and skill sync/frontmatter/subcommand.

## 0.1.0 (2026-07-10)

Initial public release. E1-E4 / W1-W17 gates, `--json`, an Agent Skills skill pack
(in-repo), 18 pytest tests.
