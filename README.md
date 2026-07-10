<div align="center">

# Archforge

**The preflight linter for AI-generated PowerPoint.**

Catches silent Korean font fallback, unreadable sizes, colliding frames,
off-canvas text, and AI-tell punctuation in built `.pptx` files,
before a human ever sees a render.

![python](https://img.shields.io/badge/python-3.9%2B-3776AB)
![license](https://img.shields.io/badge/license-MIT-green)
![status](https://img.shields.io/badge/status-beta-orange)
[![ci](https://github.com/Love-Ash/archforge/actions/workflows/ci.yml/badge.svg)](https://github.com/Love-Ash/archforge/actions/workflows/ci.yml)

[한국어 README](README.ko.md)

![demo](docs/assets/demo-en.gif)

</div>

PowerPoint opens both of these decks without a single warning. One of them is broken:

![before / after](docs/assets/before-after-en.png)

Code review cannot see any of it, because the defects live in font slots, autofit
scales, and coordinates that only materialize at render time. Archforge reads the
`.pptx` itself (XML, font-resolution chain, geometry, image alpha), so it needs no
PowerPoint installation and runs anywhere your agent or CI runs.

## 30 seconds

```bash
pip install archforge
archforge demo        # builds broken.pptx + fixed.pptx and lints both, in front of you
```

Then point it at your own deck:

```bash
archforge deck.pptx                 # objective defects only (core profile, the default)
archforge deck.pptx --profile full  # + AI-tell / style rules: machine-made decks want this
archforge deck.pptx --json          # machine-readable JSON (agents / CI)
archforge scan decks/ --profile full   # many files, directories, or globs in one run
```

The three decks in [examples/](examples/) show every gate with expected outputs.

## Why

The worst pptx defects are silent. No error is raised when:

- Korean text lands on a Latin-only font and silently falls back to Malgun Gothic
- positive letter-spacing quietly wrecks Hangul spacing
- autofit shrinks text below readable size
- text frames collide, or glyphs run off the canvas

These are exactly the defects machine-generated decks produce, and exactly the ones
an LLM cannot see in its own output. Archforge is the gate between "the build
succeeded" and "a human would sign off on the render."

## Usage

```bash
archforge deck.pptx --strict        # WARNs also fail + numeric-dash exemptions off
archforge deck.pptx --ghost         # per-page title list (horizontal-logic review)
archforge deck.pptx --render pages/ # add on-image contrast check (W7) from p01.png-style renders
archforge deck.pptx --skip W14,W6   # suppress specific WARNs (recorded in JSON)
archforge deck.pptx --lang en       # report language (default: ARCHFORGE_LANG, then OS locale)
archforge deck.pptx --no-config     # ignore config files (linting untrusted decks)
archforge deck.pptx --sarif out.sarif        # SARIF 2.1.0 (GitHub code scanning)
archforge deck.pptx --write-baseline bl.json # adopt an existing deck as-is (beta)
archforge deck.pptx --baseline bl.json       # report only new findings after that
archforge deck.pptx --hard-min 5 --body-min 9 --small-min 7.5   # size gate thresholds
archforge deck.pptx --w6-sim 0.95 --w6-cluster 5                # W6 repetition thresholds
archforge scan out/**/*.pptx --json          # aggregated JSON; exit 1 if any file fails
archforge demo --dir tour                    # regenerate the demo pair anywhere
```

Project defaults live in `.archforge.json` (or `.archforge.yml` with
`pip install archforge[yaml]`) next to the deck or in the working directory; CLI flags
override the config file, and the applied config path is always visible in the output.

```json
{ "profile": "full", "skip": ["W14"], "baseline": ".archforge-baseline.json" }
```

JSON output (single file; `scan --json` wraps one of these per file plus an aggregate
summary):

```json
{
  "schema_version": "1.0",
  "tool": { "name": "archforge", "version": "x.y.z" },
  "target_renderer": "powerpoint-windows",
  "file": "deck.pptx",
  "lang": "en",
  "errors":   [{ "page": 3, "code": "E1", "message": "...", "detail": "...",
                 "location": { "shape_id": 7, "shape_name": "TextBox 6",
                               "bbox": [1.0, 2.4, 5.0, 1.0], "paragraph": 0, "run": 1,
                               "part": "/ppt/slides/slide3.xml" } }],
  "warnings": [{ "page": 5, "code": "W15", "message": "...", "detail": "...",
                 "location": { "shape_id": 4, "bbox": [1.0, 2.4, 3.1, 0.4],
                               "related": { "shape_id": 9, "bbox": [1.2, 2.5, 3.3, 0.4] } } }],
  "ghost":    [{ "page": 1, "title": "..." }],
  "summary":  { "error_count": 1, "warn_count": 2, "pass": false, "incomplete": false,
                "profile": "full", "skipped_codes": [], "baseline_suppressed": 0,
                "config": null }
}
```

`location` is the auto-fix target for agents: shape id/name, absolute bbox in inches
(group transforms applied), paragraph/run indexes, table `cell` as `[row, col]`,
`field: true` for auto fields (slide numbers, dates), and for pair findings
(W15/W17) a `related` counterpart. Rule codes are stable, language-independent
identifiers; messages follow the report language (`--lang` > `ARCHFORGE_LANG` > OS
locale), which follows the user, not the deck.

## CI

GitHub Action (composite, installs from PyPI):

```yaml
jobs:
  deck-lint:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: Love-Ash/archforge@v0.5.0
        with:
          files: decks/
          profile: full
          sarif: archforge.sarif
      - uses: github/codeql-action/upload-sarif@v3
        if: always()
        with:
          sarif_file: archforge.sarif
```

pre-commit:

```yaml
repos:
  - repo: https://github.com/Love-Ash/archforge
    rev: v0.5.0
    hooks:
      - id: archforge
        # args: [--profile, full]
```

## What it catches

ERRORs block shipping (exit 1):

| Code | Meaning |
|:----:|---------|
| `E1` | The font that will actually render Hangul text is Latin-only: silent Malgun fallback. Resolution follows a measured PowerPoint model (see below) |
| `E2` | Dash-family characters used as sentence punctuation (the top AI-generated-deck tell). Numeric ranges (2020 to 2024 with an en dash, Q1 to Q3, 5% to 10%) and minus signs pass by default; `--strict` blocks everything |
| `E3` | Effective size below 5pt after autofit and the full placeholder inheritance chain: unreadable |
| `E4` | Positive tracking on consecutive Hangul/Hanja: letter-spacing damage (kana-containing runs are exempt; tracked kana is legitimate Japanese practice) |

WARNs are advisory:

| Code | Meaning |
|:----:|---------|
| `W1` | Body-class frame below 9pt |
| `W5` | No font size anywhere in the inheritance chain |
| `W6` | Same layout skeleton on 4+ pages (tunable; template systems: tune or skip) |
| `W7` | Low text-over-image contrast (needs `--render`) |
| `W8` | Small CJK in narrow frames (device mockups, cards) |
| `W9` | Accent vertical bars repeated as list markers |
| `W10` | Hand-drawn diagram cloned across pages |
| `W11` | AI-tell copy: buzzwords, stock openings |
| `W12` | Footer baseline drift |
| `W13` | Native PowerPoint shadow/glow/3D effects |
| `W14` | Titles are nominal phrases, not claims (Korean heuristic; numeric titles count as claims) |
| `W15` | Estimated text-on-text overlap |
| `W16` | Text glyphs or picture ink off-canvas |
| `W17` | Text straddling an image ink edge |
| `W18` | Some spans could not be checked (malformed input): results incomplete. Fails under `--strict` |

Profiles separate objective defects from style policy, and since 0.4.0 the default is
`core`: only the mechanical gates (E1/E3/E4, W1/W5/W7/W8, W15-W18) run unless you opt in.
`full` adds the AI-tell and convention rules (E2 dashes, W6 repetition, W9-W14) and is the
right mode for agent build-loops linting machine-generated decks; `editorial` drops W6/W14
for editorial and portfolio decks. Excluded rules are not merely hidden, they are not
executed, and every choice is recorded in the JSON summary, so nothing is silently
bypassed.

## How it works

The E1 font-resolution model is measured, not guessed from the OOXML spec. Probe decks
rendered through PowerPoint COM (pixel-compared, serif vs sans) established the actual
priority: run `a:ea` > lstStyle inheritance chain (shape > layout placeholder > master
placeholder > master txStyles > defaultTextStyle) > theme ea (majorFont for title
placeholders, minorFont otherwise; a non-empty theme ea beats run `a:latin`) > run
`a:latin` only when the theme ea slot is empty > OS fallback. The record lives in
[docs/CALIBRATION.md](docs/CALIBRATION.md). Themes resolve per slide master, so
multi-master decks are judged against the theme they actually use.

Effective sizes walk the same inheritance chain down to `defaultTextStyle`, so size gates
work on template/placeholder decks where nothing sets explicit sizes. Text in auto
fields (`a:fld`: slide numbers, dates) passes the same gates as regular runs, and line
breaks (`a:br`) count as line breaks in punctuation context.

Geometry checks (W15-W17) approximate effective glyph and image-ink areas: per-run sizes,
real line spacing, autofit (including percent-string form), wrap mode, group transforms,
alignment, image alpha trim, crop and flip. Intentional composition (drop caps, echo
typography, decorative bleed, captions on cards) is excluded.

Thresholds are calibrated against rendered output of a ~50-deck real corpus and hardened
with adversarial reproduction fixtures. Methods and per-gate rationale:
[docs/CALIBRATION.md](docs/CALIBRATION.md).

Malformed input never dies silently: guards are per-run and per-slide, and anything a
guard skips surfaces as `W18` in the JSON output plus a machine-readable
`summary.incomplete` flag. Gate CI on `pass` and `incomplete` together (or run
`--strict`, which turns W18 into a failure); `pass` alone reflects ERRORs only.

> Scope notes. Font-coverage knowledge (E1/E4) is currently Hangul-deep: runs written in
> other scripts are never falsely flagged (script detection is per text run, via Unicode
> code points), and per-script coverage tables are the extension path for JP/CN depth.
> Geometry estimation skips vertical text and RTL/complex scripts honestly (W18) instead
> of guessing. The render model targets PowerPoint for Windows; other renderers
> (Mac/web/LibreOffice) may resolve fonts differently.

## Agent integration

Designed for LLM-agent build-lint-fix loops:

```
build deck.pptx
loop:
    result = archforge deck.pptx --profile full --json   # machine-made decks: AI-tell rules on
    if result.summary.error_count == 0 and not result.summary.incomplete: break
    fix listed defects (location payloads point at the exact shape/run), rebuild
review WARNs against renders
```

The Agent Skills pack (standard SKILL.md + YAML frontmatter) teaches this loop and
per-code fixes to any supporting agent (Claude Code, Codex, ...). It ships inside the
wheel: `archforge skill --install`. If you cloned the repo, `skills/archforge-pptx-lint/`
is the same file.

A passing lint is not a finished deck: the linter owns the mechanical defect class;
composition and narrative still need eyes on renders.

## Contributing

False-positive reports with a repro deck are the most valuable contribution; they become
permanent regression fixtures. See [CONTRIBUTING.md](CONTRIBUTING.md) for the evidence
bar (gates are calibrated against renders, not taste) and
[SECURITY.md](SECURITY.md) for the threat model.

## Name

archforge = arch (structure) + forge. A forge where a deck's structure and Korean
typography get hammered straight before shipping.

## License

MIT © Minjae Kwon (Ash)
