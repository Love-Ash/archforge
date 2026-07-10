<div align="center">

# Archforge

**A preflight quality linter for built `.pptx` files, with first-class Korean typography awareness**

archforge = arch(structure) + forge

![python](https://img.shields.io/badge/python-3.9%2B-3776AB)
![license](https://img.shields.io/badge/license-MIT-green)
![status](https://img.shields.io/badge/status-beta-orange)
[![ci](https://github.com/Love-Ash/archforge/actions/workflows/ci.yml/badge.svg)](https://github.com/Love-Ash/archforge/actions/workflows/ci.yml)

[한국어 README](README.ko.md)

</div>

---

Archforge inspects a built `.pptx` before you ship it. It reads the file itself (XML, font
slots, coordinates, image alpha), needs no PowerPoint installation, and is built for the
workflow where code or an AI agent generates decks: stable rule codes, `--json` output,
exit codes, and an agent skill pack that ships inside the wheel.

## Why

The worst pptx defects are silent. No error is raised when:

- Korean text lands on a Latin-only font and silently falls back to Malgun Gothic
- positive letter-spacing quietly wrecks Hangul spacing
- autofit shrinks text below readable size
- text frames collide, or glyphs run off the canvas

Code review cannot see any of this; you normally catch it by eyeballing renders.
Archforge catches the mechanical defect class at build time.

## Install

```bash
pip install archforge
```

Python 3.9+, depends only on python-pptx and Pillow. No PowerPoint needed.

The agent skill pack ships inside the wheel:

```bash
archforge skill                  # print the skill pack (SKILL.md)
archforge skill --install        # install into ./.claude/skills/archforge-pptx-lint/
archforge skill --install DIR    # install anywhere
```

## Usage

```bash
archforge deck.pptx                 # human-readable report, exit 1 on any ERROR
archforge deck.pptx --json          # machine-readable JSON (agents / CI)
archforge deck.pptx --strict        # WARNs also fail + numeric-dash exemptions off
archforge deck.pptx --ghost         # per-page title list (horizontal-logic review)
archforge deck.pptx --render pages/ # add on-image contrast check (W7) from p01.png-style renders
archforge deck.pptx --profile core  # objective defects only (style/convention rules off)
archforge deck.pptx --skip W14,W6   # suppress specific WARNs (recorded in JSON)
archforge deck.pptx --lang en       # report language (default: ARCHFORGE_LANG, then OS locale)
archforge deck.pptx --hard-min 5 --body-min 9 --small-min 7.5   # size gate thresholds
archforge deck.pptx --w6-sim 0.95 --w6-cluster 5                # W6 repetition thresholds
```

JSON output:

```json
{
  "file": "deck.pptx",
  "lang": "en",
  "errors":   [{ "page": 3, "code": "E1", "message": "...", "detail": "..." }],
  "warnings": [{ "page": 5, "code": "W15", "message": "...", "detail": "..." }],
  "ghost":    [{ "page": 1, "title": "..." }],
  "summary":  { "error_count": 1, "warn_count": 2, "pass": false,
                "incomplete": false, "profile": "full", "skipped_codes": [] }
}
```

Rule codes are stable, language-independent identifiers. Messages follow the report
language, which follows the user (not the deck): `--lang` > `ARCHFORGE_LANG` > OS locale.

## What it catches

ERRORs block shipping (exit 1):

| Code | Meaning |
|:----:|---------|
| `E1` | The font that will actually render Hangul text is Latin-only: silent Malgun fallback. Resolution follows a measured PowerPoint model (see below) |
| `E2` | Dash-family characters used as sentence punctuation (the top AI-generated-deck tell). Numeric ranges (2020 to 2024 with an en dash, Q1 to Q3, 5% to 10%) and minus signs pass by default; `--strict` blocks everything |
| `E3` | Effective size below 5pt after autofit and the full placeholder inheritance chain: unreadable |
| `E4` | Positive tracking on consecutive Hangul: letter-spacing damage (Hangul-scoped; kana tracking is legitimate Japanese practice) |

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

Profiles separate objective defects from style policy: `core` keeps only the mechanical
gates (E1/E3/E4, W1/W5/W7/W8, W15-W18); `editorial` drops W6/W14 for editorial and
portfolio decks; `full` (default) runs everything. Profile choices are recorded in the
JSON summary, so nothing is silently bypassed.

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
work on template/placeholder decks where nothing sets explicit sizes.

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
    result = archforge deck.pptx --json
    if result.summary.error_count == 0: break
    fix listed defects, rebuild
review WARNs against renders
```

The Agent Skills pack (standard SKILL.md + YAML frontmatter) teaches this loop and
per-code fixes to any supporting agent (Claude Code, Codex, ...). It ships inside the
wheel: `archforge skill --install`. If you cloned the repo, `skills/archforge-pptx-lint/`
is the same file.

A passing lint is not a finished deck: the linter owns the mechanical defect class;
composition and narrative still need eyes on renders.

## License

MIT © Minjae Kwon (Ash)
