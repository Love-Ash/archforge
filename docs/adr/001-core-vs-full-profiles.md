# ADR 001: core is objective defects; style and AI-tell rules are opt-in

Status: accepted (0.4.0, reaffirmed 0.6.1)

## Context

Early versions ran every rule by default. First-time users got exit 1 for normal
punctuation (E2), which read as false positives and poisoned first impressions. At the
same time, the AI-tell rules are the product's namesake feature for agent build loops.

## Decision

The default profile is `core`: rules a renderer can prove wrong (font fallback,
unreadable sizes, collisions, off-canvas, incompleteness). `full` opts into the
style/AI-tell layer (E2 dashes, W6 repetition, W9-W14); `editorial` drops the two rules
editorial decks legitimately violate. Profiles are engine execution policy: excluded
rules never run and every exclusion is recorded in the JSON summary.

## Consequences

Agent loops must pass `--profile full` explicitly (the skill pack does). Anyone judging
the engine can do so without agreeing with the editorial layer.
