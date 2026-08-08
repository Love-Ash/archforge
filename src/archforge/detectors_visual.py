# -*- coding: utf-8 -*-
"""Appearance detectors: the rules that judge how a page is drawn.

  W7   low contrast between text and the image under it (needs rendered pages)
  W9   accent bars repeated as list markers
  W12  footer baselines drifting from the dominant baseline
  W13  native PowerPoint shadow, glow and 3D effects

Plus the two signature helpers the repetition rules read, the layout skeleton and the
fill-token set, which W6 and W10 compare across pages.

Extracted from lint.py for the 0.7 decomposition (#5). W7 is the one detector here that
touches the filesystem, because it reads the PNGs the caller exported; everything else
works off the shape tree. Re-exported from lint for backward compatibility.
"""
import glob
import os
from collections import Counter

from pptx import Presentation

try:
    from .findings import Finding, shape_loc
    from .ooxml import EMU_PER_IN, NS
    from .colors import (_is_accent, _luma, _resolve_run_rgb, _shape_fill_hex,
                         _shape_line_hex, _COLOR_UNKNOWN)
    from .geometry import _geo_rect, _is_pic, iter_shapes, iter_shapes_geo
except ImportError:   # standalone execution
    from findings import Finding, shape_loc
    from ooxml import EMU_PER_IN, NS
    from colors import (_is_accent, _luma, _resolve_run_rgb, _shape_fill_hex,
                        _shape_line_hex, _COLOR_UNKNOWN)
    from geometry import _geo_rect, _is_pic, iter_shapes, iter_shapes_geo

try:
    from PIL import Image
except ImportError:   # pragma: no cover - W7 abstains without Pillow
    Image = None

_EFFECT_TAGS = tuple(NS + t for t in ("outerShdw", "innerShdw", "glow", "reflection"))

_3D_TAGS = tuple(NS + t for t in ("sp3d", "scene3d"))

def accent_vbars_check(slide, si, sw, sh, warns):
    """W9: an AI-generated-deck tell where accent-colored vertical bars are repeated as list
    markers, using color to build item structure (measured in real decks). Reflects the
    2026-07-02 adversarial audit: colors are read via the sp.line/sp.fill accessors to cover
    namespaces and connectors, vertical bars explicitly include zero-width connectors (w~0),
    and adjacent text to the right confirms it is really a list marker."""
    bars, texts = [], []
    for sp in iter_shapes(slide.shapes):
        try:
            L, T, Wd, Ht = sp.left, sp.top, sp.width, sp.height
        except Exception:
            continue
        if None in (L, T, Wd, Ht):
            continue
        x, y, w, h = L / EMU_PER_IN, T / EMU_PER_IN, Wd / EMU_PER_IN, Ht / EMU_PER_IN
        if getattr(sp, "has_text_frame", False) and sp.text_frame.text.strip():
            texts.append((x, y, w, h))
        hexc = _shape_line_hex(sp) or _shape_fill_hex(sp)
        if not _is_accent(hexc):
            continue
        if 0.2 <= h <= 1.0 and (w < 0.05 or h > 3 * w):   # vertical bar (including
                                                           # zero-width connectors)
            bars.append((x, y, w, h, hexc))
    if len(bars) < 3:
        return
    hues = {b[4] for b in bars}
    if len(hues) != 1:                       # multiple colors is legitimate data encoding
                                              # (a legend)
        return
    xs = [b[0] for b in bars]
    if max(xs) - min(xs) > 0.15:             # vertically aligned stack only (horizontal
                                              # spread = chart/divider, excluded)
        return
    rt = 0
    for bx, by, bw, bh, _hx in bars:
        for tx, ty, tw, th in texts:
            if tx > bx and (tx - (bx + bw)) < 0.6 and not (ty > by + bh or ty + th < by):
                rt += 1
                break
    if rt >= len(bars) - 1:
        warns.append(Finding(si, "W9", "w9", (len(bars),),
                         "x=%.2fin hue=%s" % (min(xs), next(iter(hues)))))

def _fill_tokens(slide, sw, sh):
    """Turns the slide's solid-fill shapes into a multiset of (color, 24-grid position/size)
    tokens. Full-bleed backgrounds are excluded."""
    t = Counter()
    for sp in iter_shapes(slide.shapes):
        try:
            L, T, Wd, Ht = sp.left, sp.top, sp.width, sp.height
        except Exception:
            continue
        if None in (L, T, Wd, Ht) or not Wd or not Ht:
            continue
        if Wd > 0.9 * sw and Ht > 0.9 * sh:
            continue
        fh = _shape_fill_hex(sp)
        if fh is None:
            continue
        t[(fh, round(L / sw * 24), round(T / sh * 24), round(Wd / sw * 24), round(Ht / sh * 24))] += 1
    return t

def footer_top(slide, sw, sh):
    """The top (in) of the bottommost text in the slide's bottom band (y>0.88H). None if there
    is no footer."""
    best = None
    for sp in iter_shapes(slide.shapes):
        if not getattr(sp, "has_text_frame", False):
            continue
        try:
            if not sp.text_frame.text.strip():
                continue
            t = sp.top
        except Exception:
            continue
        if t is None or t <= 0.88 * sh:
            continue
        ti = t / EMU_PER_IN
        if best is None or ti > best:
            best = ti
    return best

def footer_check(foot_tops, warns):
    """W12: footer baseline misalignment across pages. In a 50-deck measurement (2026-07-02),
    the absolute-deviation approach mistook cover-page credits and bottom captions for footers,
    producing false positives in 17 decks. Instead this excludes the cover page (p1), treats
    the dominant baseline (mode of the 0.05in quantization buckets, 3+ pages) as the house
    footer, and flags only pages that deviate "slightly" (0.03-0.25in) from it. Anything off
    by more than 0.25in is assumed to be a different element such as a caption or divider and
    is ignored (its existence is not even checked)."""
    tops = [(si, t) for si, t in foot_tops.items() if t is not None and si > 1]
    if len(tops) < 4:
        return
    q = Counter(round(t / 0.05) for _si, t in tops)
    qv, cnt = q.most_common(1)[0]
    if cnt < 3:
        return
    # House baseline = the median of the actual values in the dominant bucket (using the
    # bucket center instead would make a majority of 7.08 values sit exactly 0.02 away from
    # 7.10, a floating-point boundary false positive confirmed by rescanning the 50 decks)
    bvals = sorted(t for _si, t in tops if round(t / 0.05) == qv)
    base = bvals[len(bvals) // 2]
    off = [(si, t) for si, t in tops if 0.03 < abs(t - base) <= 0.25]
    if off:
        ex = " ".join("p%d=%.2f" % (si, t) for si, t in off[:4])
        warns.append(Finding(0, "W12", "w12", (base, len(off)), ex))

def effects_count(slide):
    """The count and kinds of the slide's effective PPT effects (shadow, glow, 3D). Some
    generators leave a childless empty effectLst purely to block inheritance, so this only
    counts elements that actually have effect children (prevents an empty-element false
    positive)."""
    n = 0
    kinds = set()
    for sp in iter_shapes(slide.shapes):
        spPr = getattr(sp._element, "spPr", None)
        if spPr is None:
            continue
        eff = spPr.find(NS + "effectLst")
        if eff is not None:
            for ch in eff:
                if ch.tag in _EFFECT_TAGS:
                    n += 1
                    kinds.add(ch.tag.split("}")[1])
        for tag in _3D_TAGS:
            if spPr.find(tag) is not None:
                n += 1
                kinds.add(tag.split("}")[1])
    return n, kinds

def effects_check_deck(per_page, warns):
    """W13: aggregated once per deck (firing repeatedly per page is noise: measured on the
    50-deck corpus). Could be an intentional neon/glow style, so this is a WARN and the
    judgment is left to human eyes."""
    hits = [(si, n, kinds) for si, (n, kinds) in per_page.items() if n >= 2]
    if not hits:
        return
    total = sum(n for _si, n, _k in hits)
    kinds = sorted(set().union(*[k for _si, _n, k in hits]))
    pages = ",".join("p%d" % si for si, _n, _k in hits[:6])
    warns.append(Finding(0, "W13", "w13", (total, len(hits)),
                         "%s | %s" % (pages, ",".join(kinds))))

def _diagram_clone_marks(inter):
    """Counts "decorative texture clones" (small dots, joints, etc., 1x1 or smaller on the
    24-grid) in the multiset of fill shapes shared between two pages.
    Card-shaped (mid-size block) and table (full-width band) elements are W6's jurisdiction
    and are not counted here, avoiding a false positive on three-column comparison cards
    (reflecting the adversarial audit)."""
    marks = area = 0
    for (fh, gx, gy, gw, gh), c in inter.items():
        if gw <= 1 and gh <= 1:
            marks += c
        if gw >= 1 and gh >= 1:
            area += gw * gh * c
    if marks >= 8 and area / (24.0 * 24.0) >= 0.06:
        return marks
    return 0

def slide_layout_sig(slide, sw, sh, gw=6, gh=4):
    """Turns slide element placement into a gw x gh grid occupancy vector. Full-bleed
    backgrounds are excluded.
    Weights: text=1, shape=0.5, image=2. Used to compare "what is where" (the skeleton)."""
    sig = [0.0] * (gw * gh)
    n = 0
    for sp in iter_shapes(slide.shapes):
        try:
            L, T, Wd, Ht = sp.left, sp.top, sp.width, sp.height
        except Exception:
            continue
        if None in (L, T, Wd, Ht) or not Wd or not Ht:
            continue
        if Wd > 0.9 * sw and Ht > 0.9 * sh:      # excludes full-bleed backgrounds/images
            continue
        cx = (L + Wd / 2) / sw; cy = (T + Ht / 2) / sh
        if not (0 <= cx <= 1.0 and 0 <= cy <= 1.0):
            continue
        gc = min(gw - 1, max(0, int(cx * gw))); gr = min(gh - 1, max(0, int(cy * gh)))
        wgt = 2.0 if _is_pic(sp) else (1.0 if getattr(sp, "has_text_frame", False) else 0.5)
        sig[gr * gw + gc] += wgt
        n += 1
    return sig, n

def contrast_check(slide, si, sw, sh, render_dir, warns, styler=None, thm_colors=None,
                   skipped=None):
    """Detects low-contrast text over an image (approximated from the rendered PNG). Only
    text frames overlapping a picture, once per slide.
    Returns: "no_pics" (nothing to check) / "no_png" (there is a picture but no conventional
    render = incomplete) / "ok" (checked).
    Coordinates are absolute, including the group transform (third review P1: fixes pictures
    and text inside a group being misaligned when using raw coordinates)."""
    from PIL import Image
    pics = []
    for sp, _z, xf in iter_shapes_geo(slide.shapes):
        if _is_pic(sp):
            geo = _geo_rect(sp, xf)
            if geo is not None:
                x, y, w, h, _rot = geo
                pics.append((x * EMU_PER_IN, y * EMU_PER_IN, w * EMU_PER_IN, h * EMU_PER_IN))
    if not pics:
        return "no_pics"
    cand = glob.glob(os.path.join(render_dir, "p%02d.png" % si))
    if not cand:
        return "no_png"
    try:
        im = Image.open(cand[0]).convert("RGB"); px = im.load(); PW, PH = im.size
    except Exception:
        return "no_png"
    for sp, _z, xf in iter_shapes_geo(slide.shapes):
        if not getattr(sp, "has_text_frame", False):
            continue
        geo = _geo_rect(sp, xf)
        if geo is None:
            continue
        gx, gy, gw_, gh_, _rot = geo
        L, T, Wd, Ht = gx * EMU_PER_IN, gy * EMU_PER_IN, gw_ * EMU_PER_IN, gh_ * EMU_PER_IN
        over = any(not (L + Wd <= p0 or L >= p0 + pw or T + Ht <= q0 or T >= q0 + ph)
                   for p0, q0, pw, ph in pics)
        if not over:
            continue
        rgbs = [_resolve_run_rgb(r, para, sp.text_frame, sp, slide, styler, thm_colors)
                for para in sp.text_frame.paragraphs for r in para.runs if r.text.strip()]
        if any(c is _COLOR_UNKNOWN for c in rgbs):
            # An explicit color we cannot decode: judging with an inherited color instead
            # produced false positives (0.6.1). Abstain on this frame and surface it.
            if skipped is not None:
                skipped["w7_color_unknown"] += 1
            continue
        rgbs = [c for c in rgbs if c]
        if not rgbs:
            continue
        txt_rgb = rgbs[0]
        x0 = max(0, int(L / sw * PW)); y0 = max(0, int(T / sh * PH))
        x1 = min(PW, int((L + Wd) / sw * PW)); y1 = min(PH, int((T + Ht) / sh * PH))
        if x1 <= x0 or y1 <= y0:
            continue
        sx = max(1, (x1 - x0) // 24); sy = max(1, (y1 - y0) // 24)
        lumas = sorted(_luma(px[x, y]) for x in range(x0, x1, sx) for y in range(y0, y1, sy))
        if not lumas:
            continue
        L_txt = _luma(txt_rgb)
        # Background luma is taken from the quantile at the opposite extreme of the text
        # color: this measures worst-case local contrast (WCAG is based on the worst case)
        # and avoids mean contamination from mixed-in text ink (dark 15th percentile for light
        # text, light 85th percentile for dark text).
        if L_txt >= 0.5:
            L_bg = lumas[int(len(lumas) * 0.15)]
        else:
            L_bg = lumas[min(len(lumas) - 1, int(len(lumas) * 0.85))]
        hi = max(L_bg, L_txt); lo = min(L_bg, L_txt)
        ratio = (hi + 0.05) / (lo + 0.05)
        if ratio < 2.5:
            warns.append(Finding(si, "W7", "w7", (ratio,),
                                 "text=%r" % sp.text_frame.text[:20],
                                 data={"confidence": "measured",
                                       "evidence_source": "render",
                                       "render_confirmed": True},
                                 loc=shape_loc(sp, bbox=[gx, gy, gw_, gh_])))
            return "ok"
    return "ok"   # found and checked the render PNG for this page (regardless of whether
                  # W7 actually fired)
