---
name: archforge-pptx-lint
description: Use when building, editing, or reviewing .pptx files - run the archforge linter on the built file before delivery, read its gate codes, fix defects, and re-lint until clean. Catches silent font fallback (Hangul-deep, CJK-aware), tracking damage, unreadable sizes, text collisions, off-canvas bleed, and AI-generated deck tells that are invisible in code review.
---

# Archforge: PPTX quality gate for agents

Archforge is a CLI linter for built `.pptx` files. It inspects the file itself
(no PowerPoint needed), so it works in any headless environment. The most damaging
pptx defects are silent: fonts fall back to an OS default without any error (the
classic case is CJK text on a Latin-only font, where the coverage knowledge runs
deepest), tracking quietly wrecks character spacing, autofit shrinks text past
readability. This linter is how you catch them before a human does.

## When to run

Run it EVERY time, not just when something looks wrong:

1. After every build of a `.pptx` you generated or modified.
2. Before telling the user a deck is done. ERROR count must be 0.
3. After every fix round (fixes routinely introduce new collisions).

## Install and run

```
pip install archforge            # from PyPI
pip install -e <repo-path>       # or from the repo, for development
archforge deck.pptx --profile full --fail-incomplete --json   # THE agent command
                                            # need the AI-tell rules (default is core)
archforge deck.pptx              # objective defects only (core, the 0.4.0 default)
archforge deck.pptx --fail-incomplete   # incomplete checks (W18) fail: use this in CI
archforge deck.pptx --fail-on-warning   # WARNs fail too
archforge deck.pptx --e2-no-exemptions  # E2 numeric exemptions off (full profile)
archforge deck.pptx --strict     # union of the three flags above (compatibility alias)
archforge deck.pptx --ghost      # per-page title list (horizontal-logic review)
archforge deck.pptx --render pages/   # + on-image contrast check (W7); needs p01.png/p02.png-named PNGs
archforge deck.pptx --skip W14,W6     # suppress specific WARNs (WARN-only; recorded in JSON)
archforge deck.pptx --lang en         # report language (codes are language-independent)
archforge deck.pptx --sarif out.sarif # SARIF 2.1.0 for GitHub code scanning
archforge deck.pptx --junit out.xml   # JUnit XML for Jenkins/GitLab test reports
archforge deck.pptx --write-baseline bl.json / --baseline bl.json   # adopt existing decks (beta)
archforge baseline inspect bl.json    # what a baseline suppresses, under which conditions
archforge deck.pptx --no-config       # ignore config files (untrusted decks)
archforge deck.pptx --w6-sim 0.95 --w6-cluster 5   # loosen W6 for template-driven houses
archforge deck.pptx --hard-min 5 --body-min 9 --small-min 7.5   # size gate thresholds (pt)
archforge scan decks/ --profile full --json   # many files/dirs/globs at once; exit 1 if any fails
archforge demo                   # build broken.pptx + fixed.pptx and lint both (sanity tour)
archforge skill --install        # install this skill pack into ./.claude/skills
```

This skill pack ships inside the wheel: after `pip install archforge`, run
`archforge skill` to print it or `archforge skill --install [DIR]` to install it.
No repo clone needed.

`--json` output shape:

```json
{"schema_version": "1.0",
 "tool": {"name": "archforge", "version": "x.y.z"},
 "target_renderer": "powerpoint-windows",
 "file": "...",
 "lang": "en",
 "errors":   [{"page": 3, "code": "E1", "message": "...", "detail": "..."}],
 "warnings": [{"page": 5, "code": "W15", "message": "...", "detail": "..."}],
 "ghost":    [{"page": 1, "title": "..."}],
 "summary":  {"error_count": 0, "warn_count": 2, "pass": true,
              "incomplete": false, "profile": "full", "skipped_codes": []}}
```

Messages follow the report language (`--lang` > `ARCHFORGE_LANG` > OS locale); codes are
stable language-independent identifiers. Key on codes, not message text.

Findings carry a `location` payload when resolvable - use it as the auto-fix target
instead of searching by text: `shape_id`/`shape_name`, `bbox` `[x,y,w,h]` in inches
(absolute; group transforms applied), `paragraph`/`run` indexes into the shape's text
frame, table `cell` as `[row, col]`, `field: true` when the text lives in an `a:fld`
auto field (slide number/date; no `run` index there), `part` (the slide XML part), and
for pair findings (W15 overlap, W17 straddle) a `related` counterpart with the same keys.
`--schema 2` switches to schema 2.0: a single `findings[]` array with `severity` and a
structured `data` object per item (numbers, not a parsed sentence), plus a
`capabilities` map and a structured `abstentions[]`. Schema 1.0 (the default above) is
unchanged. `archforge scan --json` wraps one per-file document per deck plus an
aggregate summary (`summary.file_count`, `failed_files`, `error_files`, `pass`,
`incomplete`). A broken or
misconfigured file becomes a per-file `status: "error"` entry and the scan continues;
`--baseline` under scan requires exactly one matched file (fingerprints carry no file
identity, so a shared baseline would suppress findings across unrelated decks).

`ghost` lists the largest (>=18pt) title on each page that has one, in any language, so
you can read the titles top-to-bottom and check the horizontal logic. Pages with no
title-sized text are omitted, so an empty `ghost` means the deck has no clear titles.

## Code table and how to fix

ERROR = ship-blockers. Fix, rebuild, re-lint. Never deliver with ERROR > 0.

| Code | Meaning | Fix |
|------|---------|-----|
| E1 | The font that will actually render Hangul text is Latin-only (no Hangul glyphs), so Hangul silently falls back to Malgun. Effective font follows the measured PowerPoint model: run `a:ea` > paragraph `pPr/defRPr` > lstStyle inheritance chain (shape > layout ph > master ph > master txStyles > defaultTextStyle) > theme ea (majorFont for title placeholders, minorFont otherwise; non-empty theme ea beats run `a:latin`) > run `a:latin` only when the theme ea slot is empty > OS fallback. Hangul-scoped: kana/hanzi-only runs are never judged with Hangul coverage knowledge | Give Korean runs a CJK-capable font. Setting only `font.name` (= `a:latin`) with a Korean font works when the theme ea slot is empty, but setting `a:ea` explicitly is the robust fix. Keep mono/Latin display fonts for ASCII-only labels |
| E2 | A dash-family character used as sentence punctuation: em dash U+2014, en dash U+2013, figure dash U+2012, horizontal bar U+2015, 2/3-em dashes U+2E3A/U+2E3B, math minus U+2212, fullwidth hyphen U+FF0D, box-drawing line U+2500. Range-function en dashes pass by default: both neighbor tokens numeric-ish (2020, Q1, 5%, FY24; spaces allowed), or one numeric-ish neighbor with the dash attached. Spaced one-sided dashes (the AI parenthetical pattern) and word-to-word joins are blocked. Minus before a digit passes. Context is the full paragraph, so ranges split across runs don't false-positive | For prose dashes, use a colon, comma, parentheses, or line break. Ranges and negative numbers are fine by default; under `--strict`, use `~` for ranges and ASCII hyphen for minus |
| E3 | Effective font size below 5pt (autofit scale, paragraph AND placeholder/layout/master/defaultTextStyle inheritance included): unreadable | Redesign, don't just bump the number: fewer items, one representative element bigger |
| E4 | Positive letter-spacing (tracking) on a run containing Hangul (Hanja counts when mixed with Hangul): letter-spacing damage. Kana-containing runs and Hanja-only runs are exempt (tracked kana and tracked hanzi are normal JP/CN practice) | Set tracking to 0 on Hangul/Hanja runs. Track ASCII-only labels only |

WARN = advisory. Read each one; confirm on a rendered page image when the message
says so. Most are approximation-based, calibrated against rendered output
(see `docs/CALIBRATION.md` in the repo for thresholds and method).

| Code | Meaning | Typical action |
|------|---------|----------------|
| W1 | Body-class frame below 9pt | Ignore for sources/captions; raise otherwise |
| W5 | Font size not found anywhere in the inheritance chain (run, paragraph, placeholder, master, defaultTextStyle) | Set sizes explicitly so gates can measure |
| W6 | Same layout skeleton repeated on 4+ pages (tunable via `--w6-sim`/`--w6-cluster`) | Vary the grid per page; if the repetition is an intentional template system, tune or `--skip W6` |
| W7 | Text over image with contrast < 2.5 (needs `--render`) | Add scrim/darken image side |
| W8 | Small CJK (5-7.5pt) in a narrow frame (<=4in wide; device mockups, cards) | Move labels out of the mockup into callouts |
| W9 | Accent vertical bars repeated as list markers | Use dots or type hierarchy instead of colored bars |
| W10 | Hand-drawn diagram cloned across pages | Redesign or confirm the repetition is intentional |
| W11 | AI-tell copy: buzzwords, stock openings | Rewrite in the deck's own voice |
| W12 | Footer baseline drifts from the dominant baseline | Align footers to one baseline |
| W13 | Native PowerPoint shadow/glow/3D effects | Remove; they read as dated |
| W14 | Most titles are noun phrases, not claims (titles with figures count as claims) | Rewrite as action titles (the `--ghost` list should read as a story). Editorial/portfolio decks: `--skip W14` |
| W15 | Two text frames' estimated glyph areas overlap >45% | Check the rendered page; move/shrink one. Drop caps and echo typography are auto-excluded |
| W16 | Text glyphs or picture ink extend past the canvas edge | Pull content inside; decorative shape bleed is auto-excluded |
| W17 | Text straddles a picture's ink edge (25-75% inside) | Move the caption fully on or off the image; captions on solid cards are auto-excluded |
| W18 | Some spans (page-level, `page` N) or deck-level checks (`page` 0) could not run: malformed/atypical attributes, vertical text, or RTL/complex scripts whose geometry cannot be estimated. Results may be incomplete | Treat the scope as unverified: inspect stderr for what was skipped, fix the malformed source, re-lint. Under `--strict` this fails the build |

## Agent workflow

```
build deck.pptx
loop:
    result = archforge deck.pptx --profile full --fail-incomplete --json
    if result.summary.pass: break   # pass reflects the active policy (summary.policy)
    fix the listed defects (smallest page number first);
    if incomplete, fix the malformed spans W18 points at
    rebuild
for each WARN: decide fix vs. accept, based on the rendered page
render all pages once and eyeball them   # gates catch the mechanical class;
                                          # composition quality is still on you
```

Run with `--fail-incomplete` and gate on `summary.pass == true`. The active failure
policy travels in `summary.policy`, so pass always means "passed under the policy you
asked for"; a result with incomplete=true means part of the deck was never actually
checked, which `--fail-incomplete` turns into a failure for you.

Two rules of thumb learned the hard way:

- A passing lint is not a finished deck. The linter covers the mechanical defect
  class; page composition, narrative, and taste still need a human-style review
  of the rendered pages.
- When a fix round adds or moves elements, re-run the linter. Collisions appear
  in fix rounds more often than in first builds.

## Notes on internals (for debugging surprising results)

- The E1 font-resolution model is not guessed from the spec; it was measured by
  rendering probe decks in PowerPoint via COM (2026-07-10) and pinned as fixtures.
  Theme `a:ea` is resolved per slide through its layout's master relationship, so
  multi-master decks are judged against the theme they actually use. Fonts missing
  from run rPr resolve through the lstStyle inheritance chain (measured: a master
  placeholder lstStyle `a:ea` really does render), and title placeholders resolve
  theme tokens against majorFont. Target renderer: PowerPoint for Windows.
- Script detection is per text run via Unicode code points (deterministic). Hangul
  runs get the Korean font/tracking gates; other scripts are never falsely flagged
  by Korean coverage knowledge. Language-independent gates (sizes, layout, effects,
  geometry) run for every script.
- Effective sizes resolve the full inheritance chain: run > paragraph > shape
  lstStyle > layout placeholder lstStyle > master placeholder / txStyles >
  presentation defaultTextStyle. W5 only fires when the entire chain is silent.
  Sizes inherited from defaultTextStyle don't qualify text as a "title" for
  ghost/W14 (that would flood them on decks with no explicit sizes).
- Text geometry is estimated per paragraph from per-run font sizes, real line
  spacing, autofit (`fontScale` including percent-string form, `lnSpcReduction`),
  wrap mode (`wrap="none"` frames are measured as one long line, which is what
  python-pptx `add_textbox` produces by default), group transforms (off/chOff),
  and alignment. Rotated text frames are skipped.
- Picture "ink" boxes are alpha-trimmed (transparent chart margins don't count),
  crop- and flip-aware, and palette+tRNS transparent PNGs are handled.
- One malformed attribute never kills the whole report: guards are per-run (a
  corrupt run cannot swallow a sibling run's real violation) and per-slide, and
  anything a guard skips is surfaced as a W18 warning plus a machine-readable
  `summary.incomplete` flag. Gate on `pass` AND `incomplete` together (or run
  `--fail-incomplete`, which turns W18 into a failure); `summary.pass` reflects the
  active failure policy recorded in `summary.policy`.
- E2's numeric-context exemptions are evaluated against the full paragraph text,
  so ranges split across run boundaries (PowerPoint does this via spellcheck and
  formatting seams) don't false-positive. Theme font tokens ("+mn-lt" etc.) in
  run properties are resolved to the actual theme fonts before E1 judgment.
- Auto fields (`a:fld`: slide numbers, dates) are gated like regular runs (E1/E3/E4);
  they are excluded from ghost/W14 title collection. Line breaks (`a:br`) count as
  line breaks in E2's paragraph context.
- Geometry accounts for text-frame insets (lIns/tIns/rIns/bIns), explicit line breaks
  (`a:br`), field text (`a:fld`), real table column widths, and merged cells (the
  origin cell spans its merged region; continuation cells are not double-counted).
  Rotated text frames are out of scope for geometry by design and do not mark the
  result incomplete.
- Known limit: placeholders inheriting alignment from a layout's list style are
  measured as left-aligned. If W15/W16 flags look wrong on template placeholders,
  verify on the render before acting.
