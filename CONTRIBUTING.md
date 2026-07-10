# Contributing

Thanks for taking the time. Archforge is small on purpose; the bar for a change
is that it makes the linter's verdicts more *true*, not just bigger.

> Current phase: development is fast and maintainer-led. To avoid duplicating work,
> **comment to claim an issue before starting substantial code, and it will be held
> for you.** Issues labeled `roadmap` are maintainer-led tracking, not up-for-grabs.
> The single most valuable contribution regardless of phase is a false-positive
> reproduction deck (see below) - those are never at risk of collision.

## The one rule that shapes everything

Gates are calibrated against **rendered output**, not the OOXML spec. If your
change touches a threshold or a font-resolution rule, it needs evidence:
either a reproduction pptx (best) or a rendered-page comparison. The measured
render model and every threshold's rationale live in
[docs/CALIBRATION.md](docs/CALIBRATION.md); PRs that change behavior should
update it in the same commit.

## Reporting a false positive (most valuable contribution)

A linter earns trust by what it does *not* flag. If archforge flags something
that renders fine, open an issue with the **False positive** template and, if
at all possible, attach a minimal .pptx that reproduces it (strip anything
confidential; `python scripts/make_examples.py` shows how small a repro deck
can be). Confirmed false positives become regression fixtures in
`tests/test_lint.py`, so they can never come back.

## Development setup

```bash
git clone https://github.com/Love-Ash/archforge
cd archforge
pip install -e ".[test,yaml]"
python -m pytest tests -q          # full suite; must stay green
python benchmarks/run_benchmarks.py   # regression harness (synthetic corpus)
```

Commit messages are in English, imperative mood, describing the change itself (no
internal process notes).

Python 3.9+ is supported; run the suite on 3.9 semantics if you use newer
syntax. Dependencies stay minimal: python-pptx and Pillow only.

## What a good PR looks like

- One behavior change per PR, with a test that fails before and passes after.
- Positive **and** negative fixtures: prove it catches the defect, then prove
  it does not flag the closest legitimate pattern (an intentional design
  gesture, a Japanese-typography convention, a numeric range).
- Messages go through `messages.py` in both `ko` and `en`; rule codes are
  stable identifiers and never change meaning.
- New rules must state their profile (`core` is for objective render defects
  only; style/AI-tell rules go to `full`) and their severity honestly: if a
  human should look at a render before acting, it is a WARN, not an ERROR.
- No new dependencies without prior discussion in an issue.

## What gets declined

- Thresholds tuned by taste ("45% feels too strict") without render evidence.
- Rules that encode one house style as a universal defect.
- Autofix that rewrites decks: out of scope for now (a linter that edits your
  file needs a level of trust this project hasn't earned yet).

## Release discipline

Maintainer releases are batched (rc first, then one polished release); merged
PRs land in the next batch rather than triggering same-day versions.
