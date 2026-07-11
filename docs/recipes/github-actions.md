# Recipe: archforge in GitHub Actions

Minimal gate; the action installs its own pinned source, ignores deck-folder configs,
and fails on incomplete checks by default:

```yaml
jobs:
  deck-lint:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: Love-Ash/archforge@v0.8.0
        with:
          files: |
            decks/
          profile: full
          sarif: archforge.sarif
      - uses: github/codeql-action/upload-sarif@v3
        if: always()
        with:
          sarif_file: archforge.sarif
```

Useful extras: `changed-only: true` with `base-ref: origin/main` to lint only changed
decks on big repos; step outputs (`passed`, `error-count`, ...) for follow-up jobs; a
per-file table lands in the job summary automatically. JUnit consumers: add
`extra-args: --junit report.xml` and upload it.
