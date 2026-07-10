<!-- One behavior change per PR. See CONTRIBUTING.md for the full bar. -->

## What changed and why

## Evidence
<!-- For threshold/gate changes: repro pptx or rendered-page comparison.
     Gates are calibrated against renders, not the spec (docs/CALIBRATION.md). -->

## Checklist
- [ ] Test that fails before and passes after (positive fixture)
- [ ] Negative fixture: the closest legitimate pattern still passes
- [ ] Messages added in both `ko` and `en` (`messages.py`)
- [ ] `docs/CALIBRATION.md` updated if behavior/thresholds changed
- [ ] `python -m pytest tests -q` green, `python benchmarks/run_benchmarks.py` green
