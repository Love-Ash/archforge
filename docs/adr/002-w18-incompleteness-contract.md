# ADR 002: incompleteness is a first-class output, not a log line

Status: accepted (0.2.1, hardened 0.6.0)

## Context

Arbitrary pptx input means guards must swallow malformed spans to keep the report
alive. A CI consumer that only sees exit codes would read a half-checked deck as a
pass.

## Decision

Everything a guard skips surfaces as W18 findings (page-level and deck-level) plus the
machine-readable `summary.incomplete`. W18 cannot be `--skip`ped or baselined away.
`--fail-incomplete` turns incompleteness into failure and is the CI-recommended
default (the GitHub Action ships with it on). Skips that are policy, not damage
(rotated text frames), are documented as out of scope and deliberately do not mark
incompleteness.

## Consequences

`pass` alone is never the whole story unless `--fail-incomplete` is active; the JSON
carries both `pass` and `incomplete`, and since 0.6.1 the active failure policy.
