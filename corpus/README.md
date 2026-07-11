# Public seed corpus

The externally reproducible accuracy record. Every deck here is committed together
with a manifest stating its expected findings, and `run_corpus.py` enforces the match
in CI. This exists because the thresholds were originally calibrated against a private
real-deck corpus (work materials that cannot be published); this directory is the part
anyone can rerun, attack, and extend.

```bash
python corpus/run_corpus.py     # lint every deck, compare against manifests, exit 1 on drift
```

## What ground truth means here

- `python-pptx/`: defects seeded by construction. The manifest's `expected` is true by
  construction, not by trusting the linter (e.g. Hangul deliberately placed on a
  Latin-only `a:latin` with the default theme's empty ea slot).
- `pptxgenjs/`: the same defect patterns emitted by a different generator (PptxGenJS 4,
  `gen_pptxgenjs.js`), so cross-generator XML shapes are covered. Ground truth is by
  construction plus XML inspection against the measured resolution model
  ([docs/CALIBRATION.md](../docs/CALIBRATION.md)).
- `powerpoint-native/`: the python-pptx decks re-saved by real PowerPoint via COM
  (`gen_powerpoint_native.py`), so files written by PowerPoint's own writer (different
  normalization, part ordering) are covered.
- `officecli/`: decks authored through the OfficeCLI binary (`gen_officecli.py`,
  OfficeCLI 1.0.135), a third-party .NET writer. The defect fixture is OfficeCLI's own
  defaults: its blank-document theme ships empty ea slots and `add shape` writes no run
  `a:ea`, so a Hangul run lands on a Latin-only face (the E1 class). The clean fixture
  sets `--prop font.ea` explicitly, proving the author-with-X, gate-with-archforge
  pipeline composes.
- `malformed/`: inputs that must produce a controlled outcome: a truncated package is
  a usage error, vertical text must mark the report incomplete instead of guessing.

## Manifest format

```json
{
  "generator": "pptxgenjs",
  "profile": "full",
  "expected": {"E1": 1},
  "expect_incomplete": false,
  "notes": "why this expectation is ground truth"
}
```

`expected` counts are exact; an empty object means the deck must be clean.
`expect_exit2: true` marks a file that must be rejected as a usage error.

## Contributing decks

A false-positive reproduction deck is the most valuable thing you can add: a deck that
renders fine but gets flagged, with a manifest of what SHOULD be reported and ideally
a rendered screenshot. Missing coverage that needs outside assets: Google Slides
exports, Canva exports, LibreOffice exports, multi-master corporate templates, and
JP/CN decks reviewed by native readers. Licensing: contributions must be shareable
(CC0 preferred).
