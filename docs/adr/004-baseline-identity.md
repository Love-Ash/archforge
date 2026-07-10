# ADR 004: baseline fingerprints are content-keyed and page-free; full identity waits for provenance

Status: accepted (0.5.0 v2), revisit planned with structured Finding data

## Context

Generated decks are rebuilt constantly: slide insertion and message-language changes
invalidated v1's page+detail fingerprints. But a perfect identity for re-generated
artifacts does not exist without generator provenance (a source map).

## Decision

Fingerprint v2 = rule code + a locale-neutral content key, deliberately excluding the
page number; multiple identical findings are managed by occurrence count. Run
conditions (tool version, profile, lang) are recorded and checked with a warning on
mismatch. A single CLI baseline across multiple scan files is refused because the
fingerprint carries no file identity. Baselines are labeled beta.

## Consequences

Same-content findings in different locations pool together (documented limitation).
Location-bucketed fingerprints and artifact identity are deferred to the structured
Finding.data redesign, so baseline users face one more migration, not two.
