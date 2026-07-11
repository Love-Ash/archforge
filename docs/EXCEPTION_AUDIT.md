# Broad-exception audit (0.8.x)

Every `except Exception` path in the engine, classified (external audit: the policy
says "swallowed checks surface as W18", and that claim needs a per-path inventory, not
an assertion). Regenerate the raw list with:

```bash
grep -n "except Exception" src/archforge/*.py
```

## Classes

**A. Guarded into W18 (the policy working as designed).** Page-, frame-, paragraph-,
run-, and deck-level guards that increment a `skipped`/`deck_skipped` counter, which
becomes a W18 finding and `summary.incomplete`. All per-rule detector guards in
`lint()` are in this class (w6_sig, w10_tokens, w7, w9, w12_w13, glyph_boxes,
pic_boxes, w15, w16_w17, frames, frame, para, para_size, run, w6, w10, w11_w14,
image_decode, theme_parse).

**B. Controlled input errors.** Paths that convert the exception into a ValueError /
UsageError the CLI maps to exit 2 (package open, config load, baseline load, zip
preflight).

**C. Legitimate absence.** Optional XML that is genuinely allowed to be missing (a
shape with no name, no fill, no placeholder index, no geometry, an empty theme slot).
Returning None/False/empty here is spec behavior, not a swallowed check: nothing was
checkable in the first place. Most `shape_loc`, style-chain cache misses, and
fill/effect probes are in this class.

**D. Accepted-soft fallbacks (documented trade-offs).** Two survive by decision:

- `_group_xf` malformed group transform -> parent transform. Children are placed
  approximately instead of the whole page's geometry aborting over one odd group.
  Risk: a defect inside such a group can be mis-positioned. Accepted because the
  alternative (raise -> page-level W18) throws away every other frame on the page.
- `_theme_fonts_from_blob` / `_theme_colors_from_blob` parse failure -> None, which IS
  surfaced (theme_parse counter -> W18) at the deck level; listed here because the
  inner helper itself is silent and relies on the caller's accounting.

**E. Fixed in this audit.** `frame_autofit` used to swallow any parse failure into
`(1.0, 0.0)`, which HID a real autofit shrink from E3: a garbled `fontScale` read as
"no shrink" and produced a silent false negative. It now handles absence explicitly
and lets malformed values propagate to the surrounding guards, which record the span
as W18 (test: `test_malformed_autofit_marks_incomplete`).

## Rules for new code

1. A broad except may exist only inside a guard that increments a skip counter, or
   where the caught state is spec-legal absence (class C).
2. Never default a *value* on parse failure (that is class E); default only on
   *absence*.
3. Any new class-D decision gets an entry here with its trade-off.
