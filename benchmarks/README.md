# Benchmarks

Reproducible per-rule accuracy harness. Instead of committing binary decks, fixtures are
generated from a defect matrix at run time, linted through the installed CLI, and scored
per rule (precision / recall against declared expectations). Non-green expectations exit 1,
so this doubles as a CI gate.

```bash
python benchmarks/run_benchmarks.py            # run and score
python benchmarks/run_benchmarks.py --keep out # keep fixtures for inspection
```

## Honest scope

- Current generator coverage: python-pptx only. PowerPoint / LibreOffice / Google Slides
  export variants are the extension slots (`GENERATORS` in the runner) and the roadmap for
  a real compatibility matrix.
- The ~50-deck real-world corpus used for threshold calibration is private client-style
  material and is not included; its methodology and per-gate rationale are documented in
  [docs/CALIBRATION.md](../docs/CALIBRATION.md). This harness is the public, reproducible
  seed of that evidence, not a replacement for it.
- Font-resolution ground truth comes from PowerPoint COM render probes (also in
  CALIBRATION.md); those require Windows + PowerPoint and are not run in CI.
