# Calibration Log

Records why each threshold sits where it does, and where the judgment model came from. In
short: thresholds are not a matter of taste but the output of render comparisons against real
decks and adversarial verification, and any rationale not recorded here remains in code
comments with a date.

## Corpus and method

- Reference corpus: about 50 actually produced/edited decks (consulting-style reports, IR,
  pitch decks, editorial mixes; Hangul-centric plus some English). Since these are
  private/work materials, the originals are not published. In place of publishing them, this
  document records the rationale behind each gate below and exposes the key thresholds as CLI
  flags, so users can recalibrate against their own corpus.
- Procedure: exhaustive scan to collect flags -> compare flagged pages against actual rendered
  PNGs to classify genuine defects vs. false positives -> place the threshold where the two
  diverge -> adversarial verification (constructing and attacking reproduction pptx files that
  induce false positives) -> lock in the surviving thresholds and false-positive suppression
  rules as pytest fixtures.
- Recalibration knobs: `--hard-min` `--body-min` `--small-min` `--w6-sim` `--w6-cluster`
  `--skip` `--strict`.

## E1 font resolution model (measured via PowerPoint COM, 2026-07-10)

Determined by rendering measurements, not by the spec document. A probe deck (pure Hangul
with no spaces, 60pt) was exported to PNG via PowerPoint COM, and after aligning ink bounding
boxes, same-font identity was judged by pixel mean absolute difference (MAD). Because this
pits a serif (Batang) against a sans-serif (Malgun), the verdict is unambiguous.

| Case | run a:latin | run a:ea | theme minorFont a:ea | Actual render |
|---|---|---|---|---|
| a1 | Batang | none | "" (empty slot) | Batang (MAD 0.00 vs a2) |
| a2 | none | Batang | "" | Batang |
| a3 | none | none | "" | Malgun fallback |
| a4 | Consolas | none | "" | Malgun fallback (MAD 0.00 vs a3) |
| b1 | Batang | none | Malgun Gothic | Malgun (MAD 0.00 vs b2: theme beats latin) |
| b2 | none | none | Malgun Gothic | Malgun |
| b3 | none | Batang | Malgun Gothic | Batang (run ea beats theme) |

Additional measurement (probe 6, 2026-07-10, 0.2.1): two font inheritance paths.

| Case | Configuration | Actual render |
|---|---|---|
| P1 | Master body ph lstStyle lvl1 defRPr a:ea="Batang", run has no font, theme ea "" | Batang (master lstStyle ea is inherited) |
| Q1 | Theme majorFont ea="Batang", minorFont ea="Malgun", title ph run has no font | Batang (titles use majorFont ea) |
| Q2 | Same deck's body ph run has no font | Malgun (body uses minorFont ea) |

So starting in 0.2.1, a font slot missing from run rPr is resolved via the lstStyle chain
(shape -> layout ph -> master ph -> master txStyles -> defaultTextStyle), and the theme
fallback uses majorFont (title) / minorFont (everything else) ea depending on the placeholder
family.

Additional measurement (probe 7, 2026-07-10, 0.3.1): paragraph-level default run properties.

| Case | Configuration | Actual render |
|---|---|---|
| c1 | Paragraph pPr/defRPr a:ea="Batang", run has no font, theme empty slot | Batang (paragraph inheritance is real) |
| c2 | Master body ph lstStyle ea="Batang" + paragraph defRPr ea="Malgun", run has no font | Malgun (paragraph beats lstStyle) |

Confirmed priority: run rPr > paragraph pPr/defRPr > lstStyle chain > theme. Side measurement:
a file with lstStyle placed on an ordinary textbox's txBody is refused by PowerPoint on open
(placeholder and master are fine). The 3rd-round external review confirmed this missing
paragraph-level step through false-positive/false-negative reproduction (P0-1). The default
template's defaultTextStyle carries +mn-lt/+mn-ea tokens, which is consistent even down to the
fact that the effective latin of a fontless textbox resolves to Calibri, matching probe s3's
COM font properties.

Derived priority: run `a:ea` > lstStyle inheritance chain > non-empty theme `a:ea` (major/minor
by family) > run `a:latin` (only when theme ea is an empty slot) > OS fallback (Malgun). Two
implications:

1. A deck that assigns a Hangul font via python-pptx's `run.font.name` (which records only
   a:latin) actually renders in that font under the default template (empty theme ea). This is
   a legitimate pattern, so E1 does not flag it.
2. Conversely, a Hangul run given a Latin-only font silently falls back to Malgun. This is what
   E1 targets.

Limitation of the blocklist approach: a Latin-only font not in `LATIN_ONLY_FONTS` is missed (a
false negative). About 60 Latin families common in decks are registered, and Hangul-complete
variants getting caught by a prefix match are blocked via `KOREAN_CAPABLE_EXCEPTIONS` (Arial
Unicode, IBM Plex Sans KR, Noto Sans/Serif KR/CJK). NanumGothicCoding is a Hangul-complete
font, so it was removed from the list in 0.2.0.

Theme ea is resolved per master via the slide -> layout -> master -> theme relationship
(rels). 0.1.0 grabbed the package's first theme part, which could result in judging against a
theme unrelated to the content in multi-master decks.

## E2 dash gate (0.2.1 v2)

- Blocked characters: U+2012, U+2013, U+2014, U+2015, U+2212, U+FF0D, U+2E3A, U+2E3B, U+2500.
- The discriminating axis is function, not character. The same en dash can be a legitimate
  range connector or a sentence-punctuation mark (a telltale AI parenthetical), so it is
  distinguished by whether the neighboring tokens are numeric and whether they are adjacent.
  - Both surrounding tokens are numeric (contain digits: 2020, Q1, 5%, FY24): passes. Even
    with spacing, this is a range.
  - Only one side is numeric + adjacent (e.g. 2020 immediately followed by a dash + "present"):
    passes.
  - Only one side is numeric + spaced: blocked. Patterns like "growth (dash) in 2024" are
    exactly the AI parenthetical pattern.
  - Word-to-word connection (Seoul(dash)Busan): blocked. This cannot be mechanically
    distinguished from a parenthetical, so it is a conservative choice, and we acknowledge this
    is a residual false-positive zone (to be relaxed in a future English profile).
  - U+2212 passes if immediately followed by a digit (negative numbers, formulas).
- Context is evaluated over the full paragraph text, not the run (to prevent false positives
  from PowerPoint's run splitting).
- `--strict`: blocks everything with no exceptions. For houses whose document style rules ban
  dash characters entirely.

## Script layer (0.2.1)

Pairs each check with the script of the text run (Unicode code point, deterministic).

- Font knowledge has two layers (corrected by the 3rd-round panel). The main blocklist (Inter,
  Arial, monospace family) lacks CJK entirely, so it is valid for judging Hangul, kana, and Han
  characters alike, while the JP/SC subset (Noto Sans JP and similar, which have kana and Han)
  is treated as Latin-only only for Hangul judgment. So a Han-only run set in Inter or a
  monospace font is caught by E1 (a genuine fallback), while the same run set in Noto Sans JP
  passes. Empty or unspecified slots for non-Hangul scripts are silent (the OS fallback
  landscape there is unmeasured, so no claim is made).
- E4 excludes only runs mixed with kana (kana tracking/letter-spacing is normal practice in
  Japanese design). Han-only runs remain in scope because personal names and legal terms are
  common in Korean decks. From a Chinese-body-text perspective this judgment may be
  opinionated, so it is a candidate for a future profile.
- One unmeasured extrapolation is called out explicitly: in the family branch of the theme ea
  fallback, grouping all non-title placeholders (footer, date, page number) and ordinary
  textboxes under minorFont is a generalization from probe 6, which measured only title/body.
  This is consistent with the direction of the OOXML spec but is a point that has not been
  verified by COM measurement.
- For vertical writing (bodyPr@vert) and RTL/complex-shaping scripts (Arabic, Hebrew, Indic,
  Thai), the glyph-width approximation table (Latin/CJK dichotomy) is meaningless, so geometry
  checks (W15-W17) are skipped and reported via W18.
- Size, layout, and effect gates (E3, W1, W5, W6, W12, W13, etc.) are language-agnostic, so
  they run unchanged across all scripts.

## Size gates (E3 / W1 / W8 / W5)

- E3 floor of 5.0pt: below this, the render is simply illegible (per corpus render
  comparison).
- W1: under 9pt + frame width over 4in + paragraph 40+ characters = body-level text that is
  too small. Source/caption false positives are suppressed via the length and width
  conditions.
- W8: 5-7.5pt CJK (Hangul, kana, Han) + narrow frame (4in or less) = subtext inside a mockup
  or card. Small Hangul in a wide frame is likely a caption and is excluded (reflecting the
  public hygiene audit).
- Effective size is resolved from autofit (fontScale, including its percent-string form) and
  the entire inheritance chain (run > paragraph > shape lstStyle > layout placeholder lstStyle
  > master placeholder/txStyles > defaultTextStyle). 0.1.0 looked only as far as run and
  paragraph and let everything else fall through to W5, which meant the size gate was
  effectively dead in placeholder/template decks. Now W5 fires only when the entire chain is
  silent (i.e. an external generator omits even defaultTextStyle).

## Geometry gates (W15 / W16 / W17)

- W15 overlap threshold 45%: in render comparison, the 30-35% band was entirely false
  positives (titles beneath big numbers, presumed intrusion between two columns), while 60%
  and above was entirely genuine overlap. The threshold sits between the two, with a cap of 2
  findings per page.
- W16 overflow tolerance: 0.15in for text, 0.12in for pictures. Corner bleed on decorative
  shapes is a standard technique and is not checked (shape checking was rejected based on
  render measurement).
- W17 straddle judgment 25-75%: fully on top (an overlay caption) belongs to W7, fully outside
  is irrelevant. Pictures under 1 square inch are ignored, and if a solid card sits in the
  z-order between the photo and the text and backs 90% or more of the text, it is excluded.
- False-positive suppression fixtures (12 adversarial-verification cases reproduced by
  measurement, 2026-07-03): group off/chOff desync, actual single-line width with wrap=none,
  autofit percent string, P-mode + tRNS transparency, srcRect crop, flipH/V, axis-aligned bbox
  under rotation, drop cap, identical-text echo, caption over a card.

## Composition/copy gates (W6 / W9-W14)

- W6 skeleton repetition: 6x4 grid occupancy vector cosine similarity > 0.90, with 4 or more
  slides sharing the same skeleton (judged by the largest cluster, invariant to deck length).
  Houses that intentionally use a template can tighten via `--w6-sim`/`--w6-cluster` or
  `--skip W6`. The cosine of a perfect clone can exceed 1.0 due to floating-point error, so it
  is clamped to 1.0.
- W9 accent vertical bar: a solid color with saturation 0.55 or higher and lightness
  0.18-0.78, 3 or more vertically aligned stacks, confirmed as a list marker by adjacent text
  to the right. Multicolor is excluded because that indicates a legend (data encoding).
- W12 footer: excluding the cover, the median of the most frequent bucket after 0.05in
  quantization becomes the house baseline, flagging only pages off by 0.03-0.25in. An
  absolute-deviation approach was rejected because it produced false positives in 17 of 50
  decks measured.
- W13 effects: an empty effectLst with no children exists to block inheritance and is not
  counted. Aggregated once per deck (firing repeatedly per page is noise).
- W14 action titles: fires once per deck when 3 or more Hangul titles are noun phrases and
  they make up half or more of all titles. In addition to the sentence-ending heuristic, a
  title containing a number + unit ("3x", "42%", "12 billion") is recognized as an assertion
  even if it ends in a noun (0.2.0). Editorial/portfolio-showcase decks should use
  `--skip W14`.

## Robustness policy

True to its self-declared role as the last line of defense for arbitrary pptx files, per-check
guards absorb unparseable attributes and let the remaining checks continue. Guard granularity
is per run: an in-house adversarial panel reproduced by measurement (2026-07-10) that a
frame-level guard lets damage in one run swallow a genuine violation in a neighboring run,
producing a false "clean" result. For the same reason, geometry base data (glyph boxes,
picture ink boxes) retreat only on the failed axis while checks on the surviving axis continue.

Spans swallowed by a guard are surfaced in the output contract (JSON/text) as W18. Leaving
this in stderr alone would let a CI or agent pipeline that only checks the exit code and
summary.pass misread an incomplete check as a full pass. Under --strict, W18 is also escalated
to exit 1.

OOXML union types (an integer or a universal measure like "1.5pt") are schema-valid input, so
they are absorbed by parsing: the same trap was measured once in autofit fontScale
(2026-07-03) and once in spc tracking (2026-07-10). Theme tokens in a run slot (e.g.
"+mn-lt") are resolved to the per-master theme font before judgment, and E2's numeric-context
exception is judged using the whole paragraph text as context rather than the run (to prevent
false positives from run-boundary splitting).

Known retreat point: for a master whose theme XML opens at the zip level but fails lxml
parsing, the code retreats to the same assumption as an empty slot (Malgun fallback), and
notifies the retreat via stderr. Not distinguishing a parse failure from a confirmed empty
slot in the judgment is a deliberate, conservative choice.

## Known limitations (public specification, incorporating the 4th-round review / partially resolved in 0.5.0)

Resolved in 0.5.0 and locked in as regression fixtures:

- a:fld (slide-number/date fields) now passes through the same gates as an ordinary run
  (E1/E3/E4), and its location carries `field: true`. a:br is treated as a single line-break
  character in E2's context/offset handling. Fields are excluded from ghost/W14 title
  collection (to prevent contamination from large page numbers).
- The bbox in a finding's location is the absolute coordinate (in inches) with group
  transforms composed in, and table-cell findings carry a 0-based `cell`=[row, col] index.
  During this work a latent bug was found and fixed by measurement: a slide's grpSpPr is in
  the p: namespace, but the previous code did a find with a:, so the group affine always fell
  back to identity, and W15-W17 geometry was also being judged with raw coordinates in
  translated/desynced groups (fixed via local-name matching, locked in with absolute-coordinate
  tests).
- W15-W17 now carry location: bbox is the absolute bbox of the effective glyph (or picture
  ink) used to judge the overlap, and pairwise judgments (W15/W17) also identify the
  counterpart shape via `related`.

Resolved in 0.6.0 (external review of 0.5.0, locked in as regression fixtures):

- Geometry now models text-frame insets (lIns/tIns/rIns/bIns; OOXML defaults 0.1in
  left/right, 0.05in top/bottom, scaled under group transforms), explicit line breaks
  (`a:br` starts a new visual line instead of being measured as one overlong line),
  field text (`a:fld` occupies real width), real table column widths, and merged cells
  (continuation cells are skipped; the origin cell spans its merged region). Glyph
  boxes from table cells carry their `cell`=[row, col] identity into W15-W17 locations.
- Rotated text frames are explicitly out of scope for geometry (rotation is
  overwhelmingly decorative practice) and deliberately do not mark the result
  incomplete; the docs now state this contract instead of implying W18 coverage.
- W6/W10 fingerprints no longer embed page numbers (they broke the page-independent
  fingerprint contract on slide insertion); W6 keys on the layout-skeleton signature
  and W10 on the cloned diagram's shared fill tokens.
- NaN could bypass every threshold-range validation (NaN comparisons are always False,
  and json.load accepts a bare NaN literal), silently disabling E3 through --hard-min
  or an attacker-controlled config; finiteness is now validated on both paths.
- A zip preflight bounds entry count, total uncompressed size, and per-entry
  compression ratio before python-pptx parses the package, and non-budget image decode
  failures now surface through W18 (`image_decode`) instead of being swallowed.

Remaining limitations:

- Lifting E2's exceptions via `--strict` is only meaningful in the profile where E2 runs
  (full).
- Paragraph spacing (spaceBefore/spaceAfter), indents, bullets, tab stops, and text
  columns are still not modeled in geometry; list-heavy decks can accumulate vertical
  drift in glyph boxes. Insets and explicit breaks (the two largest error sources) are
  in as of 0.6.0.
- W7 samples the text-frame rectangle and uses the first resolved run color; it does
  not yet share the W15-W17 effective-glyph geometry or evaluate per-run contrast.
- E1 remains a font-name registry (curated Latin-only tables plus Hangul-complete
  exceptions), not a glyph-coverage engine: embedded-font cmaps are not inspected.
- baseline is beta. Fingerprint v2 uses a code + locale-neutral content key
  (page-independent, occurrence-count-managed), which is safe against language changes and
  slide insertion, but full identity for output that is regenerated every time is not
  possible without generator provenance (a source map). Distinct violations with identical
  content are pooled by occurrence count and suppressed without location distinction.
- The autofit annotation in E3 messages and the detail strings of W6, W10, and W16 are fixed
  in the language at generation time (unlike the main message body). The baseline fingerprint
  is decoupled from this issue via fp_key.
