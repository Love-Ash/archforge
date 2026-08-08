# -*- coding: utf-8 -*-
"""Colour resolution: shape fills and lines, theme colour maps, and the effective RGB of
a text run after the inheritance chain.

Extracted from lint.py for the 0.7 decomposition (#5), in the same shape as the earlier
fonts.py and dashes.py kernels: pure functions, no Finding, no I/O, no CLI. W7's contrast
gate consumes these; the gate itself stays in the detector layer because it emits findings.

The names are re-exported from lint for backward compatibility.
"""
import colorsys
from typing import Dict, Optional

from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

try:
    from .ooxml import NS
except ImportError:   # standalone execution
    from ooxml import NS

# Sentinel: an explicit color exists but the decoder cannot resolve it (hslClr, scrgbClr,
# sysClr, prstClr, tint/shade transforms...). Falling through to inherited colors produced
# W7 false positives, e.g. an explicit white hslClr run judged with an inherited black
# (0.6.1, external review): an unknown explicit color must stop resolution, not be skipped.
# Distinct from None, which means "resolved to no explicit color".
_COLOR_UNKNOWN = object()


def _shape_fill_hex(sp):
    """Solid fill hex (6-digit uppercase) or None. Because this uses the python-pptx accessor,
    namespaces and connector internals are handled automatically."""
    try:
        f = sp.fill
        if f.type == 1:   # MSO_FILL_TYPE.SOLID
            c = f.fore_color
            if c.type is not None and c.rgb is not None:
                return str(c.rgb).upper()
    except Exception:
        pass
    return None


def _shape_line_hex(sp):
    """Line/outline color hex or None. Also covers the line color of connectors (cxnSp)
    (measured: real-deck connector vertical bars were caught passing through this path)."""
    try:
        c = sp.line.color
        if c.type is not None and c.rgb is not None:
            return str(c.rgb).upper()
    except Exception:
        pass
    return None


def _is_accent(hexc):
    """Judges a semantic accent color: HSV saturation >= 0.55 and lightness 0.18-0.78.
    Excludes backgrounds, rule lines, body text, and low-saturation secondary colors."""
    if not hexc or len(hexc) < 6:
        return False
    try:
        r = int(hexc[0:2], 16) / 255.0
        g = int(hexc[2:4], 16) / 255.0
        b = int(hexc[4:6], 16) / 255.0
    except Exception:
        return False
    mx, mn = max(r, g, b), min(r, g, b)
    sv = 0.0 if mx == 0 else (mx - mn) / mx
    _, l, _ = colorsys.rgb_to_hls(r, g, b)
    return sv >= 0.55 and 0.18 <= l <= 0.78


def _theme_colors_from_blob(blob: bytes) -> Optional[Dict[str, str]]:
    """Maps theme clrScheme color names -> RRGGBB (sysClr uses lastClr). Used to resolve
    schemeClr for W7 (third external review P1: reinforces a gap where reading only direct
    RGB left theme-colored text out of the check)."""
    try:
        from lxml import etree
        root = etree.fromstring(blob)
        scheme = root.find(".//" + NS + "clrScheme")
        if scheme is None:
            return None
        out = {}
        for el in scheme:
            name = el.tag.split("}")[1]
            srgb = el.find(NS + "srgbClr")
            if srgb is not None and srgb.get("val"):
                out[name] = srgb.get("val").upper()
                continue
            sysc = el.find(NS + "sysClr")
            if sysc is not None and sysc.get("lastClr"):
                out[name] = sysc.get("lastClr").upper()
        # schemeClr reference-name mapping (standard clrMap default): tx1->dk1, tx2->dk2,
        # bg1->lt1, bg2->lt2
        for ref, base in (("tx1", "dk1"), ("tx2", "dk2"), ("bg1", "lt1"), ("bg2", "lt2")):
            if base in out:
                out.setdefault(ref, out[base])
        return out
    except Exception:
        return None


def theme_colors_by_master(prs) -> Dict[str, Optional[Dict[str, str]]]:
    """A map of theme colors per slide master. Key = master partname string."""
    out = {}
    try:
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT
        for master in prs.slide_masters:
            try:
                theme_part = master.part.part_related_by(RT.THEME)
                out[str(master.part.partname)] = _theme_colors_from_blob(theme_part.blob)
            except Exception:
                out[str(master.part.partname)] = None
    except Exception:
        pass
    return out


def _cosv(a, b):
    da = sum(x * x for x in a) ** 0.5; db = sum(x * x for x in b) ** 0.5
    if da == 0 or db == 0:
        return 0.0
    # Prevents identical vectors from slightly exceeding 1.0 due to floating-point error and
    # breaching the w6_sim=1.0 ceiling
    return min(1.0, sum(x * y for x, y in zip(a, b)) / (da * db))


def _luma(rgb):
    def lin(c):
        c = c / 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * lin(rgb[0]) + 0.7152 * lin(rgb[1]) + 0.0722 * lin(rgb[2])


def _run_rgb(run):
    try:
        c = run.font.color
        if c is not None and c.type is not None and c.rgb is not None:
            v = c.rgb
            return (v[0], v[1], v[2])
    except Exception:
        pass
    return None


def _resolve_run_rgb(run, para, tframe, sp, slide, styler=None, thm_colors=None):
    """Resolves text color: run rPr direct RGB -> paragraph defRPr -> lstStyle inheritance
    chain -> resolving schemeClr against the theme clrScheme (third review P1: fixes theme
    colored text being missing from W7 when only direct RGB was read). Returns an RGB
    tuple, None (no explicit color anywhere), or _COLOR_UNKNOWN (explicit but
    undecodable; the caller must not guess)."""
    def from_el(el):
        if el is None:
            return None
        fill = el.find(NS + "solidFill")
        if fill is None:
            return None
        srgb = fill.find(NS + "srgbClr")
        if srgb is not None and srgb.get("val"):
            v = srgb.get("val")
            return (int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16))
        sch = fill.find(NS + "schemeClr")
        if sch is not None and thm_colors:
            v = (thm_colors or {}).get(sch.get("val") or "")
            if v:
                return (int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16))
        # solidFill present but not a decodable srgb/scheme color: explicit-but-unknown
        return _COLOR_UNKNOWN

    direct = _run_rgb(run)
    if direct:
        return direct
    try:
        c = from_el(run._r.find(NS + "rPr"))
        if c is not None:
            return c
    except Exception:
        pass
    try:
        pPr = para._p.find(NS + "pPr")
        c = from_el(pPr.find(NS + "defRPr") if pPr is not None else None)
        if c is not None:
            return c
    except Exception:
        pass
    if styler is not None:
        try:
            for el, _src in styler._chain(tframe, sp, slide, getattr(para, "level", 0)):
                c = from_el(el)
                if c is not None:
                    return c
        except Exception:
            pass
    return None
