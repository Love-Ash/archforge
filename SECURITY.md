# Security Policy

## Supported versions

Only the latest release on PyPI is supported with fixes.

## Threat model, briefly

Archforge parses untrusted .pptx files (ZIP + XML + images via python-pptx and
Pillow) and, since 0.4.0, reads `.archforge.json|.yml` config files discovered
next to the deck. Things we treat as vulnerabilities:

- Crafted pptx/config input that causes code execution, path traversal
  (e.g. via baseline/config paths), or resource exhaustion beyond the
  documented budgets (25MP image decode cap, W6/W10 pair caps).
- A config file weakening gates **invisibly** (the applied config path must
  always appear in `summary.config` and the text footnote; `--no-config`
  must fully disable discovery).

## Reporting

Archforge performs a resource preflight (zip entry/size/ratio budgets) and defensive
per-run/per-slide parsing, and `--timeout` bounds wall-clock time in an isolated child
process. It is not a sandbox for hostile Office documents: run untrusted decks with
`--no-config --timeout N` and, for a strong boundary, inside your own container.

Please use GitHub's private vulnerability reporting on this repository
(Security tab -> "Report a vulnerability"). If that is unavailable, open a
plain issue saying only "security contact requested" without details, and the
maintainer will follow up with a private channel.

You can expect an acknowledgement within a week. Fixes ship as a patch release
with credit unless you prefer otherwise.
