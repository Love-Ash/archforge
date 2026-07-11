# CLI and machine-contract reference

The full flag reference, config file, and JSON output contract. The README keeps the
ten-second version; this page is the complete one.

## Every flag

```bash
archforge deck.pptx --fail-incomplete   # incomplete checks (W18) fail: recommended in CI
archforge deck.pptx --fail-on-warning   # WARNs fail too
archforge deck.pptx --e2-no-exemptions  # E2 numeric-range/minus exemptions off (full profile)
archforge deck.pptx --strict            # union of the three flags above
archforge deck.pptx --ghost         # per-page title list (horizontal-logic review)
archforge deck.pptx --render pages/ # add on-image contrast check (W7) from p01.png-style renders
archforge deck.pptx --skip W14,W6   # suppress specific WARNs (recorded in JSON)
archforge deck.pptx --lang en       # report language (default: ARCHFORGE_LANG, then OS locale)
archforge deck.pptx --no-config     # ignore config files (linting untrusted decks)
archforge deck.pptx --sarif out.sarif        # SARIF 2.1.0 (GitHub code scanning)
archforge deck.pptx --junit out.xml          # JUnit XML (Jenkins/GitLab test reports)
archforge deck.pptx --json --schema 2        # schema 2.0: findings[] + severity + data + capabilities
archforge deck.pptx --timeout 60             # wall-clock limit, isolated in a child process
archforge deck.pptx --write-baseline bl.json # adopt an existing deck as-is (beta)
archforge deck.pptx --baseline bl.json       # report only new findings after that
archforge deck.pptx --hard-min 5 --body-min 9 --small-min 7.5   # size gate thresholds
archforge deck.pptx --w6-sim 0.95 --w6-cluster 5                # W6 repetition thresholds
archforge scan out/**/*.pptx --json          # aggregated JSON; exit 1 if any file fails
                                             # (an input matching nothing exits 2)
archforge rules                              # one-line summary of every rule
archforge explain W15                        # what a rule means and how to fix it
archforge baseline inspect bl.json           # what a baseline suppresses, under which conditions
archforge demo --dir tour                    # regenerate the demo pair anywhere
```

## Config file and JSON contract

Project defaults live in `.archforge.json` (or `.archforge.yml` with
`pip install archforge[yaml]`) next to the deck or in the working directory; CLI flags
override the config file, and the applied config path is always visible in the output.

```json
{ "profile": "full", "skip": ["W14"], "baseline": ".archforge-baseline.json",
  "severity": { "E2": "warning" } }   // severity overrides: policy-layer rules only
```

JSON output (single file; `scan --json` wraps one of these per file plus an aggregate
summary):

```json
{
  "schema_version": "1.0",
  "tool": { "name": "archforge", "version": "x.y.z" },
  "target_renderer": "powerpoint-windows",
  "file": "deck.pptx",
  "lang": "en",
  "errors":   [{ "page": 3, "code": "E1", "message": "...", "detail": "...",
                 "location": { "shape_id": 7, "shape_name": "TextBox 6",
                               "bbox": [1.0, 2.4, 5.0, 1.0], "paragraph": 0, "run": 1,
                               "part": "/ppt/slides/slide3.xml" } }],
  "warnings": [{ "page": 5, "code": "W15", "message": "...", "detail": "...",
                 "location": { "shape_id": 4, "bbox": [1.0, 2.4, 3.1, 0.4],
                               "related": { "shape_id": 9, "bbox": [1.2, 2.5, 3.3, 0.4] } } }],
  "ghost":    [{ "page": 1, "title": "..." }],
  "summary":  { "error_count": 1, "warn_count": 2, "pass": false, "incomplete": false,
                "profile": "full", "skipped_codes": [], "baseline_suppressed": 0,
                "config": null }
}
```

`location` is the auto-fix target for agents: shape id/name, absolute bbox in inches
(group transforms applied), paragraph/run indexes, table `cell` as `[row, col]`,
`field: true` for auto fields (slide numbers, dates), and for pair findings
(W15/W17) a `related` counterpart. Rule codes are stable, language-independent
identifiers; messages follow the report language (`--lang` > `ARCHFORGE_LANG` > OS
locale), which follows the user, not the deck.


Formal, machine-validatable schemas for every JSON shape live in
[../schemas/](../schemas/); tests validate real CLI output against them.
