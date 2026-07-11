# -*- coding: utf-8 -*-
"""E1 font kernel (0.8, #5 decomposition): the Latin-only blocklist, run/paragraph
font-slot parsing, theme-token resolution, and the measured E1 decision itself.
Pure functions over lxml elements and plain dicts; no Finding, no I/O.
Re-exported from lint for backward compatibility."""
import re
from typing import Dict, Optional

NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

# Latin-only fonts that have no Hangul glyphs, so even when assigned to Hangul text they don't
# apply (prefix matching, so it also covers weight variants). To narrow the limitation of the
# blocklist approach (fonts not on this list are a false negative), Latin families common in
# decks are registered broadly (reflecting the 2026-07-10 external review's E1 false-negative
# finding: families like Inter, Arial, and Calibri were missing).
LATIN_ONLY_FONTS = (
    # Monospace
    "ibm plex mono", "courier", "consolas", "cascadia mono", "cascadia code",
    "roboto mono", "jetbrains mono", "fira code", "source code pro", "pt mono",
    "dejavu sans mono", "dejavu sans", "menlo", "monaco", "sf mono", "space mono",
    # Sans-serif (common in web/design decks)
    "switzer", "inter", "arial", "helvetica", "calibri", "aptos", "segoe ui", "roboto",
    "lato", "montserrat", "poppins", "open sans", "source sans", "raleway",
    "nunito", "work sans", "dm sans", "manrope", "karla", "figtree", "sora",
    "outfit", "plus jakarta", "barlow", "oswald", "rubik", "mulish",
    "tahoma", "verdana", "trebuchet", "century gothic", "franklin gothic",
    "gill sans", "futura", "candara", "corbel", "ibm plex sans", "noto sans",
    "pt sans", "pt serif",
    # Serif
    "times new roman", "georgia", "garamond", "palatino", "cambria", "constantia",
    "playfair", "merriweather", "lora", "libre baskerville", "crimson", "noto serif",
)

# Exceptions that match the prefixes above but actually have Hangul glyphs (checked first to
# prevent false positives).
# Example: "IBM Plex Sans KR" matches the "ibm plex sans" prefix but is a fully Hangul-capable
# font.
KOREAN_CAPABLE_EXCEPTIONS = (
    "arial unicode", "ibm plex sans kr", "noto sans kr", "noto sans cjk",
    "noto serif kr", "noto serif cjk",
)

# JP/SC/TC subset fonts that match a blocklist prefix but do have kana and Hanja: treated as
# Latin-only (no Hangul glyphs) for Hangul judgment, but pass for kana/Hanja judgment.
# Separating these two layers lets us catch a genuine fallback for Hanja text (Inter, mono
# families = no CJK at all) while avoiding false positives on JP fonts (third adversarial
# panel).
JP_CN_CAPABLE_PREFIXES = (
    "noto sans jp", "noto sans sc", "noto sans tc", "noto sans hk",
    "noto serif jp", "noto serif sc", "noto serif tc", "ibm plex sans jp",
)


def run_fonts(run):
    """The a:latin/a:ea/a:cs typeface of the run rPr. run.font.name only looks at a:latin,
    so the XML is read directly."""
    out = {}
    rPr = run._r.find(NS + "rPr")
    if rPr is not None:
        for slot in ("latin", "ea", "cs"):
            el = rPr.find(NS + slot)
            if el is not None:
                tf = el.get("typeface")
                if tf:
                    out[slot] = tf
    return out


_UNIVERSAL_PER_PT = {"pt": 1.0, "in": 72.0, "cm": 72.0 / 2.54, "mm": 72.0 / 25.4,
                     "pc": 12.0, "pi": 12.0}


def _text_point_attr(v: Optional[str]) -> Optional[int]:
    """ST_TextPoint union: both '150' (1/100pt integer) and a universal measure like '1.5pt'
    are schema-valid (ECMA-376). int() alone dies with ValueError on the latter: the spc
    version of the same union trap as _pct_attr (measured in adversarial verification,
    2026-07-03). Unparseable garbage values return None."""
    if v is None:
        return None
    v = v.strip()
    try:
        return int(v)
    except ValueError:
        pass
    m = re.match(r"^(-?[0-9]+(?:\.[0-9]+)?)(mm|cm|in|pt|pc|pi)$", v)
    if not m:
        return None
    return int(round(float(m.group(1)) * _UNIVERSAL_PER_PT[m.group(2)] * 100))


def para_fonts(para) -> Dict[str, str]:
    """The latin/ea typeface of the paragraph pPr/defRPr. In OOXML this is the paragraph's
    default run property, ranking immediately after run rPr (measured with COM probe 7,
    2026-07-10: the paragraph defRPr ea is actually rendered and beats the lstStyle chain.
    Third external review P0-1: missing this step caused both E1 false positives and false
    negatives)."""
    out = {}
    try:
        pPr = para._p.find(NS + "pPr")
        if pPr is not None:
            d = pPr.find(NS + "defRPr")
            if d is not None:
                for slot in ("latin", "ea"):
                    el = d.find(NS + slot)
                    if el is not None:
                        tf = el.get("typeface")
                        if tf:
                            out[slot] = tf
    except Exception:
        pass
    return out


def run_track(run) -> Optional[int]:
    """The run's tracking (spc, in units of 1/100 pt). None if absent or unparseable."""
    rPr = run._r.find(NS + "rPr")
    if rPr is None:
        return None
    return _text_point_attr(rPr.get("spc"))


def is_latin_only_font(name: Optional[str], script: str = "hangul") -> bool:
    """Is this a font that lacks glyphs for the given script (prefix matching, capability
    exceptions checked first)?
    script="hangul": judged by Hangul glyphs (default). script="cjk_other": judged by
    kana/Hanja, passing JP/SC subset fonts (which have kana/Hanja). The main blocklist body
    (Inter, Arial, mono families) has no CJK at all, so both criteria apply to it (third
    adversarial panel: Hanja regression fix).
    A font whose name contains "cjk" (Noto/Source Han superfont families) always passes."""
    if not name:
        return False
    low = name.strip().lower()
    if "cjk" in low:
        return False
    if script != "hangul" and any(low.startswith(x) for x in JP_CN_CAPABLE_PREFIXES):
        return False
    if any(low.startswith(x) for x in KOREAN_CAPABLE_EXCEPTIONS):
        return False
    return any(low.startswith(x) for x in LATIN_ONLY_FONTS)


def e1_violation(text: str, fonts: Dict[str, str], thm_ea: Optional[str],
                 script: str = "hangul"):
    """E1 judgment: is the effective render font for CJK text Latin-only, or otherwise
    undecidable (Malgun fallback)?

    Measured-by-render model (PowerPoint COM probe, 2026-07-10, docs/CALIBRATION.md):
      1) If run a:ea exists, that font draws the Hangul.
      2) Otherwise, the theme minorFont a:ea. If not empty, it takes priority over run
         a:latin.
      3) Only when the theme ea slot is empty does run a:latin draw the Hangul (if it has
         Hangul glyphs).
      4) If none of the above, or it is Latin-only, fall back to the OS (Windows Malgun).
    Returns: (message, detail) or None. Addresses two points from the external review
    (2026-07-10) at once: the false negatives that an "ea-or-latin" substitute produced, and
    the false positive that an ea-only judgment would have produced on a legitimate pattern
    (font.name = a Hangul font that actually renders because the theme ea is empty). The
    measured-by-render model resolves both simultaneously."""
    if script != "hangul":
        # Kana/Hanja mode (third adversarial panel: Hanja regression fix): uses the same
        # slot resolution, but only fires when the specified font explicitly lacks the
        # relevant glyphs. An empty slot or unspecified font stays silent (the OS fallback
        # terrain for non-Hangul scripts is unmeasured, so it is not asserted).
        theme_ea = (thm_ea or "").strip()
        cand = fonts.get("ea") or theme_ea or fonts.get("latin")
        if cand and is_latin_only_font(cand, script):
            return ("e1_cjk_other", "font=%r text=%r" % (cand, text[:24]),
                    {"script": script, "effective_font": cand,
                     "font_source": ("run.ea" if fonts.get("ea") else
                                     "theme.ea" if theme_ea else "run.latin")})
        return None
    run_ea = fonts.get("ea")
    if run_ea:
        if is_latin_only_font(run_ea):
            return ("e1_run_ea", "font=%r text=%r" % (run_ea, text[:24]),
                    {"script": "hangul", "effective_font": run_ea,
                     "font_source": "run.ea", "fallback_font": "Malgun Gothic"})
        return None
    theme_ea = (thm_ea or "").strip()
    if theme_ea:
        if is_latin_only_font(theme_ea):
            return ("e1_theme_ea", "theme=%r text=%r" % (thm_ea, text[:24]),
                    {"script": "hangul", "effective_font": theme_ea,
                     "font_source": "theme.ea", "fallback_font": "Malgun Gothic"})
        return None   # a non-empty Hangul-capable theme ea gets rendered (run latin cannot
                       # beat it: measured)
    run_latin = fonts.get("latin")
    if run_latin:
        if is_latin_only_font(run_latin):
            return ("e1_latin_empty_theme", "font=%r text=%r" % (run_latin, text[:24]),
                    {"script": "hangul", "effective_font": run_latin,
                     "font_source": "run.latin", "fallback_font": "Malgun Gothic"})
        return None   # with an empty theme ea, a Hangul-capable latin font actually draws the
                       # Hangul (measured)
    return ("e1_nofont", "text=%r" % text[:24],
            {"script": "hangul", "effective_font": None,
             "font_source": "none", "fallback_font": "Malgun Gothic"})


def resolve_font_tokens(fonts: Dict[str, str], thm_fonts: Optional[Dict[str, str]]) -> Dict[str, str]:
    """Replaces OOXML theme tokens (types like "+mn-lt"/"+mj-ea") in run slot values with
    the actual font name.
    If a token cannot be resolved (theme parse failure, empty value), that slot is emptied
    (= left to the inheritance chain).
    Confirmed in the adversarial panel (2026-07-10): comparing the token literally against the
    blocklist never matches, so a run that falls back to a Latin-only theme font silently
    passed E1."""
    out = dict(fonts)
    for slot in ("latin", "ea"):
        v = out.get(slot)
        if not v or not v.startswith("+"):
            continue
        resolved = (thm_fonts or {}).get(v[1:].strip().lower())
        if resolved:
            out[slot] = resolved
        else:
            out.pop(slot, None)
    return out
