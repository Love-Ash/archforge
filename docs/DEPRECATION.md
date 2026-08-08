# Deprecation policy

What archforge promises not to break, how it breaks things when it must, and how much
warning you get. This lands with 0.9.0, the release candidate, because a stability promise
made at 1.0 with no stated policy behind it is not a promise.

## What is covered

Three surfaces are public and change under this policy.

**The command line.** Subcommand names, flag names, and the meaning of exit codes 0, 1
and 2. A flag may gain values; it may not silently change what an existing value does.

**The machine outputs.** The JSON report (`--json`, both `schema_version` 1.0 and 2.0),
SARIF, JUnit, and the baseline file format. Their schemas live in
[`schemas/`](../schemas) and are validated in the test suite. Fields may be added. A field
may not change type or meaning, and a documented field may not disappear inside a major
version.

**The rule contract.** A rule's code and its severity class. `E2` will not become a
warning, and a code will not be reused for a different defect.

## What is not covered

**Rule sensitivity.** Whether a specific deck trips W15 can change in any release. The
gates are heuristics measured against a corpus, and improving them means findings move.
This is the main reason `--baseline` exists: it records what you already accept so that a
sharpened rule does not fail your build on day one. Changes that shift findings on the
public corpus are listed in the changelog with the counts.

**Anything named with a leading underscore**, and anything not in `archforge.lint.__all__`.
The Python API surface is what `__all__` lists; the rest is internals and moves without
notice. As of 0.9.0 `__all__` is explicit for the first time, which is what makes this
sentence meaningful.

**Report prose.** Message wording and translations change freely. Match on `code`, never on
the human-readable string. This is why every finding carries a locale-neutral `msg_id`.

## How a break happens

1. **Announced.** The changelog entry says what is deprecated, what replaces it, and the
   release it will be removed in.
2. **Warned.** The old form keeps working and prints a deprecation notice to stderr, not to
   stdout, so piped machine output is unaffected.
3. **Removed.** No sooner than the next minor release, and never in a patch.

Before 1.0, one minor release of warning. After 1.0, a removal waits for the next major.

## Security

A fix for a security issue may skip the warning step. It will still be a separate release
and will say so plainly in the changelog.

## Versioning

Semantic versioning, with the caveat about rule sensitivity above: a MINOR release can
change which findings a deck produces, because a rule is not an API. If that reading is too
loose for your pipeline, pin an exact version and use `--baseline`.

While the major version is 0, MINOR carries the breaking changes and PATCH does not.
