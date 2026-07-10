---
name: archforge-pptx-lint
description: Use when building, editing, or reviewing .pptx files (especially Korean/CJK decks) - run the archforge linter on the built file before delivery, read its gate codes, fix defects, and re-lint until clean. Catches silent Korean font fallback, CJK tracking damage, unreadable sizes, text collisions, off-canvas bleed, and AI-generated deck tells that are invisible in code review.
---

# Archforge: PPTX quality gate for agents

Archforge (아치포지) is a CLI linter for built `.pptx` files. It inspects the file itself
(no PowerPoint needed), so it works in any headless environment. Korean typography
is a first-class concern: the most damaging pptx defects for Korean decks are
silent (fonts fall back to Malgun without any error; tracking quietly wrecks
Hangul spacing) and this linter is how you catch them before a human does.

## When to run

Run it EVERY time, not just when something looks wrong:

1. After every build of a `.pptx` you generated or modified.
2. Before telling the user a deck is done. ERROR count must be 0.
3. After every fix round (fixes routinely introduce new collisions).

## Install and run

```
pip install archforge            # from PyPI
pip install -e <repo-path>       # or from the repo, for development
archforge deck.pptx --profile full --json   # THE agent command: machine-made decks
                                            # need the AI-tell rules (default is core)
archforge deck.pptx              # objective defects only (core, the 0.4.0 default)
archforge deck.pptx --strict     # WARNs also fail + numeric-dash exemptions off
archforge deck.pptx --ghost      # per-page title list (horizontal-logic review)
archforge deck.pptx --render pages/   # + on-image contrast check (W7); needs p01.png/p02.png-named PNGs
archforge deck.pptx --skip W14,W6     # suppress specific WARNs (WARN-only; recorded in JSON)
archforge deck.pptx --lang en         # report language (codes are language-independent)
archforge deck.pptx --sarif out.sarif # SARIF 2.1.0 for GitHub code scanning
archforge deck.pptx --write-baseline bl.json / --baseline bl.json   # adopt existing decks
archforge deck.pptx --w6-sim 0.95 --w6-cluster 5   # loosen W6 for template-driven houses
archforge deck.pptx --hard-min 5 --body-min 9 --small-min 7.5   # size gate thresholds (pt)
archforge skill --install        # install this skill pack into ./.claude/skills
```

This skill pack ships inside the wheel: after `pip install archforge`, run
`archforge skill` to print it or `archforge skill --install [DIR]` to install it.
No repo clone needed.

`--json` output shape:

```json
{"schema_version": "1.0",
 "tool": {"name": "archforge", "version": "0.3.1"},
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

`ghost` lists the largest (>=18pt) title on each page that has one, in any language, so
you can read the titles top-to-bottom and check the horizontal logic. Pages with no
title-sized text are omitted, so an empty `ghost` means the deck has no clear titles.

## Code table and how to fix

ERROR = ship-blockers. Fix, rebuild, re-lint. Never deliver with ERROR > 0.

| Code | Meaning | Fix |
|------|---------|-----|
| E1 | The font that will actually render Hangul text is Latin-only (no Hangul glyphs), so Hangul silently falls back to Malgun. Effective font follows the measured PowerPoint model: run `a:ea` > lstStyle inheritance chain (shape > layout ph > master ph > master txStyles > defaultTextStyle) > theme ea (majorFont for title placeholders, minorFont otherwise; non-empty theme ea beats run `a:latin`) > run `a:latin` only when the theme ea slot is empty > OS fallback. Hangul-scoped: kana/hanzi-only runs are never judged with Hangul coverage knowledge | Give Korean runs a CJK-capable font. Setting only `font.name` (= `a:latin`) with a Korean font works when the theme ea slot is empty, but setting `a:ea` explicitly is the robust fix. Keep mono/Latin display fonts for ASCII-only labels |
| E2 | A dash-family character used as sentence punctuation: em dash U+2014, en dash U+2013, figure dash U+2012, horizontal bar U+2015, 2/3-em dashes U+2E3A/U+2E3B, math minus U+2212, fullwidth hyphen U+FF0D, box-drawing line U+2500. Range-function en dashes pass by default: both neighbor tokens numeric-ish (2020, Q1, 5%, FY24; spaces allowed), or one numeric-ish neighbor with the dash attached. Spaced one-sided dashes (the AI parenthetical pattern) and word-to-word joins are blocked. Minus before a digit passes. Context is the full paragraph, so ranges split across runs don't false-positive | For prose dashes, use a colon, comma, parentheses, or line break. Ranges and negative numbers are fine by default; under `--strict`, use `~` for ranges and ASCII hyphen for minus |
| E3 | Effective font size below 5pt (autofit scale, paragraph AND placeholder/layout/master/defaultTextStyle inheritance included): unreadable | Redesign, don't just bump the number: fewer items, one representative element bigger |
| E4 | Positive letter-spacing (tracking) on consecutive Hangul: Hangul spacing breaks. Hangul-scoped (kana tracking is normal Japanese practice) | Set tracking to 0 on Hangul runs. Track ASCII-only labels only |

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
    result = archforge deck.pptx --profile full --json   # full: machine-made decks
    if result.summary.error_count == 0 and not result.summary.incomplete: break
    fix the listed defects (smallest page number first);
    if incomplete, fix the malformed spans W18 points at
    rebuild
for each WARN: decide fix vs. accept, based on the rendered page
render all pages once and eyeball them   # gates catch the mechanical class;
                                          # composition quality is still on you
```

Ship condition is `error_count == 0 AND incomplete == false` - a pass with
incomplete=true means part of the deck was never actually checked.

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
  `--strict`, which turns W18 into a failure); `pass` alone reflects ERRORs only.
- E2's numeric-context exemptions are evaluated against the full paragraph text,
  so ranges split across run boundaries (PowerPoint does this via spellcheck and
  formatting seams) don't false-positive. Theme font tokens ("+mn-lt" etc.) in
  run properties are resolved to the actual theme fonts before E1 judgment.
- Known limit: placeholders inheriting alignment from a layout's list style are
  measured as left-aligned. If W15/W16 flags look wrong on template placeholders,
  verify on the render before acting.
