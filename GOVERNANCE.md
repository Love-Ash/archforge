# Governance

A small document so contributors know how decisions get made here.

## Roles

- Maintainer: Minjae Kwon (Ash, @Love-Ash). Reviews and merges PRs, triages issues,
  cuts releases, owns the rule registry and public contracts.
- Contributors: anyone with a merged PR or an accepted false-positive fixture. Credited
  in release notes.

## What gets accepted

- Gates and thresholds change only with render evidence (see CONTRIBUTING). Taste is
  not evidence; a reproduction deck or a rendered-page comparison is.
- New rules must state their profile honestly: objective render defects go to `core`,
  style and AI-tell policies are opt-in (`full`/`editorial`). A rule a reasonable house
  style could disagree with is never an unconditional ERROR.
- Confirmed false positives become permanent regression fixtures before the fix merges.

## Breaking changes and releases

- Public contracts are the JSON schema, rule codes and their meanings, the baseline
  file schema, CLI exit codes, and the Action inputs. Breaking any of these requires a
  minor version bump before 1.0 (and a major after), a CHANGELOG entry, and where
  feasible a compatibility alias for two minor releases.
- Pre-1.0, releases iterate fast: a 0.x bump can ship the day it is ready, and several
  in one day is expected while contracts are still settling. What is gated is
  announcement-grade visibility, not the version number: a public launch (Show HN,
  broad posts) waits for an RC soak period. From 1.0 on, minor releases are batched and
  an RC precedes each one; same-day churn on stable contracts is the mistake to avoid.

## Security

Private vulnerability reporting per [SECURITY.md](SECURITY.md); fixes ship as patch
releases with credit.

## If the maintainer disappears

If the repository sees no maintainer activity for 6 months, anyone may open an issue
proposing a fork as the continuation point; the README will link to a community fork
that demonstrates releases and test discipline.
