# ADR 003: the render model targets PowerPoint for Windows, by measurement

Status: accepted (0.2.0)

## Context

The OOXML spec defines font slots and inheritance but not the precedence a renderer
actually applies; implementations differ and internet folklore contradicts itself.

## Decision

Font-resolution behavior (E1) is pinned by rendering probe decks through PowerPoint
COM on Windows and pixel-comparing serif-vs-sans outcomes. The measured chain (run
a:ea > paragraph defRPr > lstStyle chain > theme ea > a:latin on an empty theme slot >
OS fallback) is the model; the record lives in docs/CALIBRATION.md. Other renderers
(Mac, web, LibreOffice) are explicitly out of scope until measured the same way.

## Consequences

Claims stay honest ("target renderer: powerpoint-windows" is in the JSON contract),
and extending renderer coverage means running probes there, not guessing.
