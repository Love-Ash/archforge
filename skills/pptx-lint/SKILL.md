---
name: jangpyo-pptx-lint
description: Use when building, editing, or reviewing .pptx files (especially Korean/CJK decks) - run the jangpyo linter on the built file before delivery, read its gate codes, fix defects, and re-lint until clean. Catches silent Korean font fallback, CJK tracking damage, unreadable sizes, text collisions, off-canvas bleed, and AI-generated deck tells that are invisible in code review.
---

# jangpyo: PPTX quality gate for agents

jangpyo (장표) is a CLI linter for built `.pptx` files. It inspects the file itself
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
pip install -e <repo-path>     # from the repo (not yet on PyPI)
jangpyo deck.pptx              # human-readable report, exit 1 if any ERROR
jangpyo deck.pptx --json       # machine-readable (recommended for agents)
jangpyo deck.pptx --strict     # WARNs also fail the exit code
jangpyo deck.pptx --ghost      # per-page title list (horizontal-logic review)
jangpyo deck.pptx --render pages/   # + on-image contrast check (W7); needs p01.png/p02.png-named PNGs
```

`--json` output shape:

```json
{"file": "...",
 "errors":   [{"page": 3, "code": "E1", "message": "...", "detail": "..."}],
 "warnings": [{"page": 5, "code": "W15", "message": "...", "detail": "..."}],
 "ghost":    [{"page": 1, "title": "..."}],
 "summary":  {"error_count": 0, "warn_count": 2, "pass": true}}
```

Messages are in Korean; codes are stable identifiers. Use the code table below.

`ghost` lists the largest (>=18pt) title on each page that has one, in any language, so
you can read the titles top-to-bottom and check the horizontal logic. Pages with no
title-sized text are omitted, so an empty `ghost` means the deck has no clear titles.

## Code table and how to fix

ERROR = ship-blockers. Fix, rebuild, re-lint. Never deliver with ERROR > 0.

| Code | Meaning | Fix |
|------|---------|-----|
| E1 | CJK text will render through a Latin-only font slot (e.g. IBM Plex Mono, Consolas) or an empty theme `a:ea` slot: silent Malgun fallback | Give Korean runs a CJK-capable font (Gothic/Myeongjo family). Keep mono fonts for ASCII-only labels. Patch the theme `a:ea` slot if fonts are unset |
| E2 | A dash-family character in rendered text: em/en/figure dash, and also the math minus (U+2212), fullwidth hyphen (U+FF0D), box-drawing line (U+2500) | For dashes, use a colon, comma, parentheses, or line break. For a minus sign (common in copied financial data, e.g. −3.2%), use an ASCII hyphen `-` |
| E3 | Effective font size below 5pt (autofit scale and inheritance included): unreadable | Redesign, don't just bump the number: fewer items, one representative element bigger |
| E4 | Positive letter-spacing (tracking) on consecutive CJK characters: Hangul spacing breaks | Set tracking to 0 on CJK runs. Track ASCII-only labels only |

WARN = advisory. Read each one; confirm on a rendered page image when the message
says so. Most are approximation-based, calibrated against rendered output.

| Code | Meaning | Typical action |
|------|---------|----------------|
| W1 | Body-class frame below 9pt | Ignore for sources/captions; raise otherwise |
| W5 | Font size not found on run/paragraph | Set sizes explicitly so gates can measure |
| W6 | Same layout skeleton repeated on 4+ pages | Vary the grid per page |
| W7 | Text over image with contrast < 2.5 (needs `--render`) | Add scrim/darken image side |
| W8 | Small CJK (5-7.5pt) in a narrow frame (<=4in wide; device mockups, cards) | Move labels out of the mockup into callouts |
| W9 | Accent vertical bars repeated as list markers | Use dots or type hierarchy instead of colored bars |
| W10 | Hand-drawn diagram cloned across pages | Redesign or confirm the repetition is intentional |
| W11 | AI-tell copy: buzzwords, stock openings | Rewrite in the deck's own voice |
| W12 | Footer baseline drifts from the dominant baseline | Align footers to one baseline |
| W13 | Native PowerPoint shadow/glow/3D effects | Remove; they read as dated |
| W14 | Most titles are noun phrases, not claims | Rewrite as action titles (the `--ghost` list should read as a story) |
| W15 | Two text frames' estimated glyph areas overlap >45% | Check the rendered page; move/shrink one. Drop caps and echo typography are auto-excluded |
| W16 | Text glyphs or picture ink extend past the canvas edge | Pull content inside; decorative shape bleed is auto-excluded |
| W17 | Text straddles a picture's ink edge (25-75% inside) | Move the caption fully on or off the image; captions on solid cards are auto-excluded |

## Agent workflow

```
build deck.pptx
loop:
    result = jangpyo deck.pptx --json
    if result.summary.error_count == 0: break
    fix the listed defects (smallest page number first)
    rebuild
for each WARN: decide fix vs. accept, based on the rendered page
render all pages once and eyeball them   # gates catch the mechanical class;
                                          # composition quality is still on you
```

Two rules of thumb learned the hard way:

- A passing lint is not a finished deck. The linter covers the mechanical defect
  class; page composition, narrative, and taste still need a human-style review
  of the rendered pages.
- When a fix round adds or moves elements, re-run the linter. Collisions appear
  in fix rounds more often than in first builds.

## Notes on internals (for debugging surprising results)

- Text geometry is estimated per paragraph from per-run font sizes, real line
  spacing, autofit (`fontScale` including percent-string form, `lnSpcReduction`),
  wrap mode (`wrap="none"` frames are measured as one long line, which is what
  python-pptx `add_textbox` produces by default), group transforms (off/chOff),
  and alignment. Rotated text frames are skipped.
- Picture "ink" boxes are alpha-trimmed (transparent chart margins don't count),
  crop- and flip-aware, and palette+tRNS transparent PNGs are handled.
- Known limit: placeholders inheriting alignment from a layout's list style are
  measured as left-aligned. If W15/W16 flags look wrong on template placeholders,
  verify on the render before acting.
