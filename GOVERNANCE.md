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
- Releases are batched: an rc or soak period precedes announcement-grade releases.
  Same-day version churn is a known past mistake we do not repeat.

## Security

Private vulnerability reporting per [SECURITY.md](SECURITY.md); fixes ship as patch
releases with credit.

## If the maintainer disappears

If the repository sees no maintainer activity for 6 months, anyone may open an issue
proposing a fork as the continuation point; the README will link to a community fork
that demonstrates releases and test discipline.
