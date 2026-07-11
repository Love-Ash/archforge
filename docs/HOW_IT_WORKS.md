# How it works

The technical detail behind the gates. The README keeps a three-line summary and links
here; this is the page for readers who want the model.

## E1: a measured font-resolution model

The E1 font-resolution model is measured, not guessed from the OOXML spec. Probe decks
rendered through PowerPoint COM (pixel-compared, serif vs sans) established the actual
priority:

```
run a:ea
  > paragraph pPr/defRPr
  > lstStyle inheritance chain (shape > layout placeholder > master placeholder
                                > master txStyles > defaultTextStyle)
  > theme ea (majorFont for title placeholders, minorFont otherwise;
              a non-empty theme ea beats run a:latin)
  > run a:latin (only when the theme ea slot is empty)
  > OS fallback (Malgun Gothic for Hangul)
```

The full record, including the seven COM probe rounds and every threshold's rationale,
is in [CALIBRATION.md](CALIBRATION.md). Themes resolve per slide master, so
multi-master decks are judged against the theme they actually use.

## Effective sizes

Effective sizes walk the same inheritance chain down to `defaultTextStyle`, so size
gates work on template/placeholder decks where nothing sets explicit sizes. Text in
auto fields (`a:fld`: slide numbers, dates) passes the same gates as regular runs, and
line breaks (`a:br`) count as line breaks in punctuation context and in geometry.

## Geometry (W15-W17)

Geometry checks approximate effective glyph and image-ink areas: per-run sizes, real
line spacing, autofit (including percent-string form), wrap mode, group transforms,
text-frame insets, alignment, image alpha trim, crop and flip, and merged table cells.
Intentional composition (drop caps, echo typography, decorative bleed, captions on
cards) is excluded. Rotated text frames are out of scope by design.

## Calibration

Thresholds are calibrated against rendered output of a ~50-deck private corpus and
hardened with adversarial reproduction fixtures. The publicly reproducible part is
[corpus/](../corpus/), where every deck ships with a manifest of its expected findings
and `run_corpus.py` enforces the match in CI. Methods and per-gate rationale live in
[CALIBRATION.md](CALIBRATION.md); the renderer-coverage matrix there states exactly
what has been measured on which renderer.

## Incompleteness is a first-class output

Malformed input never dies silently: guards are per-run and per-slide, and anything a
guard skips surfaces as `W18` in the JSON output plus a machine-readable
`summary.incomplete` flag. `summary.pass` reflects the active failure policy (recorded
in `summary.policy`): with the defaults it covers ERRORs only, so gate CI with
`--fail-incomplete` (the GitHub Action does this for you) and key on `summary.pass`.
See [ADR 002](adr/002-w18-incompleteness-contract.md).

## Scope

Font-coverage knowledge (E1/E4) is currently Hangul-deep: runs written in other scripts
are never falsely flagged (script detection is per text run, via Unicode code points),
and per-script coverage tables are the extension path for JP/CN depth. Geometry
estimation skips vertical text and RTL/complex scripts honestly (W18) instead of
guessing. The render model targets PowerPoint for Windows; other renderers (Mac, web,
LibreOffice) may resolve fonts differently. Archforge performs a resource preflight and
`--timeout` bounds wall-clock time, but it is not a sandbox for hostile documents; see
[SECURITY.md](../SECURITY.md).

## Why the gate is a separate tool

Archforge is deliberately not part of any authoring tool, and that is a design position,
not an accident of history. Three reasons:

1. **A generator grading its own output shares its own blind spots.** If the authoring
   tool's writer leaves the theme ea slot empty, its checker was built by the same
   people with the same model of what "correct" looks like. The corpus has a live
   example: OfficeCLI 1.0.135's default deck with Hangul text ships an E1 blocker
   ([corpus/officecli/](../corpus/officecli/)), and its own effective-font readback
   models the latin slot only. An independent gate answers to a different ground truth:
   PowerPoint's measured render behavior ([CALIBRATION.md](CALIBRATION.md)).
2. **Preview engines are not the target renderer.** Tools that render their own HTML
   preview resolve fonts with their own logic. The defect class E1 goes deepest on
   (a non-empty theme ea beating the run's own a:latin) is a property of PowerPoint's
   resolver specifically; a preview drawing the right font is not evidence PowerPoint
   will.
3. **The gate must outlive any one generator.** Pipelines swap authoring tools;
   the acceptance criterion should not swap with them. Same file in, same exit code
   out, regardless of what wrote the file, and the corpus keeps fixtures from four
   writers (python-pptx, PptxGenJS, PowerPoint's own writer, OfficeCLI) to hold that.

## Per-rule pages

Every rule has its own page under [docs/rules/](rules/) with its meaning, a fix, and the
"if you disagree" pointer; `archforge explain CODE` prints the same, and SARIF
`helpUri` links there.
