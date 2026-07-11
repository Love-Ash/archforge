# ADR 004: baseline fingerprints are content-keyed and page-free; full identity waits for provenance

Status: superseded by the v3 amendment below (0.7.0). Original decision was v2 (0.5.0).

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

## Amendment (v3, 0.7.0)

The deferred migration landed together with structured Finding.data, so baseline users
migrate once, as this ADR promised. Fingerprint v3 = the v2 content key combined with a
page-free structural location bucket (a semantic shape name, or the bbox grid when the
name is a generator auto-name like "TextBox N"), so a defect that moves to a genuinely
different place is treated as new instead of silently re-suppressed. A threshold hash is
recorded and warned on mismatch. v1/v2 files are rejected with a regenerate message.

Still deferred (0.8+): full artifact identity (logical id / template hash / generator)
and `baseline inspect|diff|update` UX. The location bucket closes the most damaging gap
(a moved defect vanishing); artifact identity guards a rarer misuse (a baseline applied
to an unrelated deck) and is tracked, not shipped, here.
