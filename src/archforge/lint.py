# -*- coding: utf-8 -*-
"""
Archforge: a Hangul-specialized quality linter that takes a built .pptx and mechanically
blocks recurring defects.

Reason to exist: even when rules are written down, the same mistakes repeat every time a deck
is made. Defects that a machine can check should be blocked by this linter, not by human eyes.
It catches the points where Hangul decks silently break (fallback to Latin-only fonts, CJK
tracking, undersized text) and the tells specific to AI-generated decks (long dashes, cliche
copy, recycled layouts) in the build output. Because it is the last line of defense for an
arbitrary pptx, it also covers bypass paths such as inheritance, slots, autofit, tables, and
groups (reflecting adversarial audit findings).

ERROR = blocks deployment:
  E1 The effective render font for Hangul text is Latin-only (no Hangul glyphs) = a silent
     Malgun fallback.
     The effective font is resolved with the measured-by-render COM model (2026-07-10,
     docs/CALIBRATION.md):
     run a:ea > lstStyle inheritance chain (shape -> layout -> master, measured in probe 6)
     > theme ea
     (title family uses majorFont, everything else uses minorFont; if not empty it takes
     priority over run a:latin)
     > (only when the theme ea slot is empty) run a:latin > OS fallback (Malgun).
     The trigger is Hangul-only: kana- or Hanja-only runs are not judged using Hangul coverage
     knowledge (0.2.1).
  E2 A dash-class character appears in the rendered text (em/en/figure dash plus math minus
     U+2212, fullwidth hyphen-minus U+FF0D, box-drawing horizontal line U+2500, etc.).
     The axis of judgment is function, not character: an en dash passes if both neighboring
     tokens are numeric (a range), passes only when directly adjacent if just one side is
     numeric, and is blocked for a spaced-out parenthetical or a word-to-word join. U+2212
     passes when immediately followed by a digit.
     Context is read from the whole paragraph, not the run (to prevent false positives from
     run splitting). --strict blocks everything with no exceptions.
  E3 The effective font (reflecting autofit fontScale, paragraph inheritance, and the
     placeholder inheritance chain) is below HARD_MIN (default 5.0pt) = unreadable.
  E4 Two or more consecutive Hangul characters carry meaningfully positive tracking (spc>50 =
     0.5pt) = tracking has spread apart.
     Hangul-only: kana tracking is normal practice in Japanese design, so it cannot be declared
     a defect outright (0.2.1).
WARN = informational (deployment still passes):
  W1 The effective font of a body-level frame (wide, high character count) is below BODY_MIN
     (default 9.0pt) = body text is too small
  W8 Small CJK text in a narrow frame (effective size in [HARD_MIN, SMALL_MIN=7.5)pt) = risk
     of unreadable text inside device mockups or cards
     (the gray zone above the E3 unreadable floor and below the W1 body-level threshold.
     Surfaces sub-text inside phone mockups during preflight)
  W6 Recycled layout skeleton: shape bbox signature cosine similarity > 0.90, cluster of 4+
     pages (no render needed, always checked)
  W7 Low-contrast text over an image, ratio<2.5 (only when --render <pages> is given, compared
     pixel-by-pixel against the rendered PNG)
  W9 3+ accent-colored vertical bars repeated as list markers = using color to build structure
     (a Claude tell). Covers connectors and zero-width bars
  W10 A hand-drawn diagram (e.g. cross-section, decorative texture) repeated almost identically
      across multiple pages = flags whether it is recycling or intentional for human judgment
  W11 AI-tell copy: buzzwords (narrow dictionary, all pages), cliche stock openings (p1-3 only)
  W12 Footer baseline misalignment: pages deviating 0.03-0.25in from the dominant baseline
      (cover page excluded, median of 0.05in buckets)
  W13 2+ native PPT shadow/glow/3D effects (an old-fashioned tell. An empty effectLst with no
      children is not counted)
  W14 A majority of titles (3+, half or more) are descriptive noun phrases = not action titles.
      --ghost also prints the ghost deck (title listing)
  W15 Estimated text-on-text overlap (approximate effective glyph bbox, intersection > 45% of
      the smaller box) = occlusion or collision, at most 2 findings per page
  W16 Off-canvas overflow: text = 0.15in+ outside the effective glyph box boundary, images =
      0.12in+ outside the ink bbox (alpha-trimmed)
      (full-bleed 70%+ excluded). Corner bleed on decorative shapes is standard technique and
      is not checked (rejected after measured-by-render testing)
  W17 Text straddles the ink boundary of a non-background image (only 25-75% of the glyph is
      inside) = looks cropped. Fully on top of the image is W7's jurisdiction
  W5 Font size found nowhere in the run, paragraph, or inheritance chain (only when the whole
     chain is silent)
  W18 Some region could not be checked due to corrupted or non-standard attributes: the result
      may be incomplete (surfaces the region a guard swallowed in the output contract, not just
      stderr; promoted to exit 1 under --strict)

W15-W17 geometry robustness (fixed after reproducing 12 findings from adversarial
verification, 2026-07-03): group off/chOff affine transform, wrap=none (word_wrap=False, the
python-pptx add_textbox default) single-line actual width, autofit percentage strings
('62.5%') and lnSpcReduction, per-paragraph alignment, empty-paragraph endParaRPr size, rotated
shapes (axis-aligned expanded bbox, rotated text is skipped), picture srcRect crop, flipH/V, P
mode tRNS transparency, and W17 suppresses a solid card (by z order) between a photo and its
caption. Remaining limitation: a placeholder's alignment inherited from layout lstStyle falls
back to left alignment.

Usage: archforge <built.pptx> [--hard-min 5.0] [--body-min 9.0] [--strict] [--render <pages>]
              [--ghost] [--json] [--skip CODES] [--w6-sim 0.90] [--w6-cluster 3]
  --strict: WARN also causes exit 1 + lifts the E2 numeric-context exception. --render:
  enables W7.
  --ghost: lists titles (for eyeballing horizontal logic). --skip: suppresses the given codes
  (e.g. --skip W14,W6 for an editorial deck).
Returns: exit 1 if there is any ERROR, otherwise 0.

Subcommand: archforge skill [--install [DIR]] [--path]
  Prints the bundled agent skill pack (SKILL.md), or installs it to DIR (default
  ./.claude/skills).
"""
import os
import re
import sys
import glob
import math
import argparse
import colorsys
from collections import Counter, namedtuple
from typing import Dict, List, Optional, Tuple

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

try:
    from .messages import M, set_lang, get_lang
    from .findings import Finding, shape_loc
    from .rules import RULES, ALL_CODES, PROFILES, DEFAULT_PROFILE, severity
    from . import config as _config
    from . import reporters as _reporters
except ImportError:   # fallback for standalone file execution (python lint.py)
    from messages import M, set_lang, get_lang
    from findings import Finding, shape_loc
    from rules import RULES, ALL_CODES, PROFILES, DEFAULT_PROFILE, severity
    import config as _config
    import reporters as _reporters

EMU_PER_IN = 914400
NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
NS_P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

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

# Long dash class + fullwidth hyphen-minus + 2x/3x em + box-drawing horizontal line (built
# from code points: no dash character appears literally in this source)
LONG_DASHES = {chr(c) for c in (0x2012, 0x2013, 0x2014, 0x2015, 0x2212, 0xFF0D, 0x2E3A, 0x2E3B, 0x2500)}
_EN_DASH = chr(0x2013)
_MINUS = chr(0x2212)

# Rule metadata, profiles, and the full code list have moved to the rules.py registry
# (0.4.0 decomposition).
# PROFILES/ALL_CODES continue to be exposed from this module via the import above (backward
# compatibility).


# Per-run Unicode script detection moved to scripts.py (0.7 decomposition: parsing layer).
# Re-exported here so existing callers (jl.is_hangul etc.) and this module's own detectors
# keep working unchanged. _geometry_unsupported keeps its underscore alias for internal use.
try:
    from .scripts import (is_hangul, has_hangul, is_kana, is_hanja, is_cjk, has_cjk,
                          geometry_unsupported as _geometry_unsupported)
except ImportError:
    from scripts import (is_hangul, has_hangul, is_kana, is_hanja, is_cjk, has_cjk,
                         geometry_unsupported as _geometry_unsupported)


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


class _FldRun:
    """Adapter for a:fld (auto fields such as slide number, date). Since CT_TextField also has
    an rPr+t structure per the schema, PowerPoint renders it with the same rules as a normal
    run, but python-pptx's para.runs only returns a:r, so field text was a blind spot for
    E1/E3/E4 (carried over from the fourth review, 0.5.0). Exposes only the minimal interface
    the checking code uses: ._r for run_fonts/run_track, .font.size for size."""
    __slots__ = ("_r", "text")

    class _Pt:
        __slots__ = ("pt",)

        def __init__(self, pt):
            self.pt = pt

    def __init__(self, fld):
        self._r = fld
        t = fld.find(NS + "t")
        self.text = (t.text or "") if t is not None else ""

    @property
    def font(self):
        return self   # only the .size access is used

    @property
    def size(self):
        try:
            v = self._r.find(NS + "rPr").get("sz")
        except Exception:
            return None
        if not v:
            return None
        try:
            return self._Pt(int(v) / 100.0)
        except (TypeError, ValueError):
            return None


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
            return ("e1_cjk_other", "font=%r text=%r" % (cand, text[:24]))
        return None
    run_ea = fonts.get("ea")
    if run_ea:
        if is_latin_only_font(run_ea):
            return ("e1_run_ea", "font=%r text=%r" % (run_ea, text[:24]))
        return None
    theme_ea = (thm_ea or "").strip()
    if theme_ea:
        if is_latin_only_font(theme_ea):
            return ("e1_theme_ea", "theme=%r text=%r" % (thm_ea, text[:24]))
        return None   # a non-empty Hangul-capable theme ea gets rendered (run latin cannot
                       # beat it: measured)
    run_latin = fonts.get("latin")
    if run_latin:
        if is_latin_only_font(run_latin):
            return ("e1_latin_empty_theme", "font=%r text=%r" % (run_latin, text[:24]))
        return None   # with an empty theme ea, a Hangul-capable latin font actually draws the
                       # Hangul (measured)
    return ("e1_nofont", "text=%r" % text[:24])


def _is_digit_ch(c: str) -> bool:
    """Digit judgment covers ASCII and fullwidth digits only: isdigit() also returns True for
    superscript footnote marks (U+00B9) and circled numbers (U+2460), which let a word with a
    footnote mark like "revenue-superscript-1" be mistaken as numeric and slip through the
    exception (measured in the adversarial panel)."""
    return "0" <= c <= "9" or 0xFF10 <= ord(c) <= 0xFF19


def _dash_neighbor(text: str, i: int, step: int) -> Tuple[str, bool]:
    """The neighboring token (up to 12 chars) of the dash at position i, and whether there is
    whitespace. step=-1 for left, +1 for right."""
    j = i + step
    spaced = False
    while 0 <= j < len(text) and text[j].isspace():
        spaced = True
        j += step
    buf = []
    while 0 <= j < len(text) and not text[j].isspace() and text[j] not in LONG_DASHES:
        buf.append(text[j])
        j += step
        if len(buf) >= 12:
            break
    # A left scan (step=-1) accumulates in reverse order, so reverse it back to original text
    # order: this makes the digit-leading check look at the token's actual first character
    # rather than the character adjacent to the dash (fix measured in the adversarial panel)
    tok = "".join(reversed(buf)) if step < 0 else "".join(buf)
    return tok, spaced


def dash_violations(text: str, strict: bool = False,
                    span: Optional[Tuple[int, int]] = None) -> List[str]:
    """Extracts E2 violation characters. The axis of judgment is function, not the character
    itself: it separates range connectors (legitimate typography) from punctuation (an AI
    parenthetical tell) by whether the neighboring tokens are numeric and whether they are
    joined without a space (0.2.1 v2, fixes 4 remaining false-positive shapes from the second
    external re-check).

    en dash (U+2013):
      - both neighboring tokens are numeric (containing a digit: 2020, Q1, 5%, FY24) -> passes
        (a range even with spaces)
      - only one side numeric + the dash is directly adjacent -> passes (a "2020-present"
        style range)
      - only one side numeric + there is a space -> blocked (an AI-style parenthetical like
        "growth - in 2024")
      - word~word -> blocked (cannot be machine-distinguished from a parenthetical, a
        conservative choice. "Seoul~Busan" type cases remain a residual false positive)
    Math minus (U+2212): passes if immediately followed by a digit (negative number/formula).
    em dash and the rest of the dash class: blocked in every context (the signature function).
    strict=True blocks everything with no exceptions.

    If span=(start,end) is given, only characters in that range are reported, but neighboring
    context is read from the whole text. When the caller passes text=whole paragraph and
    span=the run's range, ranges split across run boundaries are not falsely flagged
    (confirmed in the adversarial panel, 2026-07-10)."""
    lo, hi = span if span is not None else (0, len(text))
    bad = []
    for i in range(lo, min(hi, len(text))):
        c = text[i]
        if c not in LONG_DASHES:
            continue
        if not strict:
            # A consecutive dash (an adjacent character before/after is also a dash) has no
            # place in any writing convention: block unconditionally.
            # Closes a hole where, in an exact double sequence like '2020--2024', the two
            # dashes see each other as neighbors, the token becomes empty, and the
            # one-side-numeric exception was slipping through (measured in the adversarial
            # panel).
            prev_ch = text[i - 1] if i > 0 else ""
            next_ch = text[i + 1] if i + 1 < len(text) else ""
            if prev_ch in LONG_DASHES or next_ch in LONG_DASHES:
                bad.append(c)
                continue
            if c == _EN_DASH:
                lt, lsp = _dash_neighbor(text, i, -1)
                rt, rsp = _dash_neighbor(text, i, +1)
                lnum = any(_is_digit_ch(ch) for ch in lt)
                rnum = any(_is_digit_ch(ch) for ch in rt)
                if lnum and rnum:
                    continue
                # The one-side-numeric exception only applies to tokens that start with a
                # digit: closes a bypass where a word+digit mixed token like "conclusion2024"
                # let an adjacent parenthetical pass through (measured in the adversarial
                # panel)
                l_lead = bool(lt) and _is_digit_ch(lt[0])
                r_lead = bool(rt) and _is_digit_ch(rt[0])
                if (l_lead or r_lead) and not lsp and not rsp:
                    continue
            elif c == _MINUS:
                # Allow whitespace: fixes a false positive on financial notation with a spaced
                # sign, like "- 3.2%" (adversarial panel)
                j = i + 1
                while j < len(text) and text[j].isspace():
                    j += 1
                if j < len(text) and _is_digit_ch(text[j]):
                    continue
        bad.append(c)
    return bad


def _pct_attr(v, default):
    """OOXML percentage union type: both '62500' (1/1000 %) and '62.5%' (string form) are
    valid (ST_TextFontScalePercentOrPercentString). int() alone dies with ValueError on the
    latter, and a blanket except swallowed it to 1.0 (measured in adversarial verification,
    2026-07-03)."""
    if not v:
        return default
    v = v.strip()
    if v.endswith("%"):
        return float(v[:-1]) / 100.0
    return int(v) / 100000.0


def frame_autofit(tf):
    """The (fontScale, lnSpcReduction) ratio pair. (1.0, 0.0) if there is no normAutofit."""
    try:
        bodyPr = tf._txBody.find(NS + "bodyPr")
        if bodyPr is not None:
            na = bodyPr.find(NS + "normAutofit")
            if na is not None:
                return (_pct_attr(na.get("fontScale"), 1.0),
                        _pct_attr(na.get("lnSpcReduction"), 0.0))
    except Exception:
        pass
    return 1.0, 0.0


def frame_font_scale(tf):
    """The text frame's autofit fontScale (ratio). Kept for backward compatibility with
    existing consumers such as E3."""
    return frame_autofit(tf)[0]


def _group_xf(sp, xf):
    """Composes the group shape's off/ext vs chOff/chExt affine with the parent xf.
    Coefficients (ax,bx,ay,by): abs = a*raw + b (EMU). Falls back to the parent xf unchanged
    (identity fallback) if parsing fails.

    grpSpPr is in the p: namespace at the slide level (only the inner xfrm/off elements are
    a:). Fixes a latent bug where the previous code did find(a:grpSpPr), which always
    returned None, causing a silent identity fallback (measured in 0.5.0: found by
    reproducing a case where a moved, desynced group's loc bbox came out in raw
    coordinates)."""
    try:
        gsp = None
        for ch in sp._element:
            if isinstance(ch.tag, str) and ch.tag.endswith("}grpSpPr"):
                gsp = ch
                break
        x = gsp.find(NS + "xfrm")
        off, ext = x.find(NS + "off"), x.find(NS + "ext")
        cho, che = x.find(NS + "chOff"), x.find(NS + "chExt")
        ox, oy = int(off.get("x")), int(off.get("y"))
        ew, eh = int(ext.get("cx")), int(ext.get("cy"))
        cx, cy = int(cho.get("x")), int(cho.get("y"))
        cw_, ch_ = int(che.get("cx")) or ew, int(che.get("cy")) or eh
        sx, sy = ew / float(cw_), eh / float(ch_)
        ax, bx, ay, by = xf
        return (ax * sx, ax * (ox - cx * sx) + bx,
                ay * sy, ay * (oy - cy * sy) + by)
    except Exception:
        return xf


def collect_frames(shapes, xf=(1.0, 0.0, 1.0, 0.0)):
    """A list of (text_frame, width_emu, owner_shape, cell_rc, xf). Recurses into groups and
    includes native table cells. cell_rc is (row, col), 0-based, for table cells, otherwise
    None. xf is the group absolute-coordinate affine (same coefficients as iter_shapes_geo),
    used to turn the loc bbox of a run-level finding into real slide coordinates instead of
    the group's chOff coordinate space (carried over from the fourth review, 0.5.0)."""
    out = []
    for sp in shapes:
        try:
            st = sp.shape_type
        except Exception:
            st = None
        if st == MSO_SHAPE_TYPE.GROUP:
            out += collect_frames(sp.shapes, _group_xf(sp, xf))
            continue
        if getattr(sp, "has_table", False):
            tbl = sp.table
            ncol = len(tbl.columns) or 1
            try:
                col_w = [(c.width or 0) for c in tbl.columns]
            except Exception:
                col_w = []
            for ri, row in enumerate(tbl.rows):
                for ci, cell in enumerate(row.cells):
                    # Merged regions (0.6.0): continuation cells mirror their origin's
                    # text frame, so walking them double-counted the same runs. The
                    # origin's usable width spans the merged columns (real column widths
                    # when available; the old even-split approximation as fallback).
                    try:
                        if cell.is_spanned:
                            continue
                        w_emu = sum(col_w[ci:ci + max(1, cell.span_width)]) \
                            or (sp.width or 0) // ncol
                    except Exception:
                        w_emu = (sp.width or 0) // ncol
                    out.append((cell.text_frame, w_emu, sp, (ri, ci), xf))
            continue
        if sp.has_text_frame:
            out.append((sp.text_frame, sp.width or 0, sp, None, xf))
    return out


def iter_shapes(shapes):
    """Flattens all shapes into a single traversal, recursing into groups."""
    for sp in shapes:
        try:
            if sp.shape_type == MSO_SHAPE_TYPE.GROUP:
                for inner in iter_shapes(sp.shapes):
                    yield inner
                continue
        except Exception:
            pass
        yield sp


def _is_pic(sp):
    try:
        return sp.shape_type == MSO_SHAPE_TYPE.PICTURE
    except Exception:
        return False


# Absolute-coordinate traversal for W15-W17 geometry consumers. A group child's raw left/top
# is in the group's chOff coordinate space, which drifts from slide coordinates in a pptx
# where the group has been moved or resized (off!=chOff desync, standard behavior when
# dragging in PowerPoint) (measured in adversarial verification, 2026-07-03). Composes the
# off/ext vs chOff/chExt affine and yields (shape, z-order, absolute xfrm function
# coefficients). xf=(ax,bx,ay,by): abs = a*raw + b (EMU).
def iter_shapes_geo(shapes, xf=(1.0, 0.0, 1.0, 0.0), _z=None):
    if _z is None:
        _z = [0]
    for sp in shapes:
        try:
            is_grp = sp.shape_type == MSO_SHAPE_TYPE.GROUP
        except Exception:
            is_grp = False
        if is_grp:
            # Group-internal coordinate g: composes abs_g = ox + (g - cx)*sx with the parent
            # xf (_group_xf)
            for inner in iter_shapes_geo(sp.shapes, _group_xf(sp, xf), _z):
                yield inner
            continue
        z = _z[0]
        _z[0] += 1
        yield sp, z, xf


def _geo_rect(sp, xf):
    """The absolute bbox (in) with xf applied. A rotated shape is expanded to an axis-aligned
    bbox around a fixed center. None on failure."""
    try:
        L, T, Wd, Ht = sp.left, sp.top, sp.width, sp.height
    except Exception:
        return None
    if None in (L, T, Wd, Ht):
        return None
    ax, bx, ay, by = xf
    x = (ax * L + bx) / EMU_PER_IN
    y = (ay * T + by) / EMU_PER_IN
    w = ax * Wd / EMU_PER_IN
    h = ay * Ht / EMU_PER_IN
    rot = 0.0
    try:
        rot = float(sp.rotation or 0.0)
    except Exception:
        pass
    if rot % 360.0:
        import math
        r = math.radians(rot)
        w2 = abs(w * math.cos(r)) + abs(h * math.sin(r))
        h2 = abs(w * math.sin(r)) + abs(h * math.cos(r))
        x, y, w, h = x + (w - w2) / 2, y + (h - h2) / 2, w2, h2
    return x, y, w, h, (rot % 360.0 != 0.0)


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


def _theme_fonts_from_blob(blob: bytes) -> Optional[Dict[str, str]]:
    """The 4 major/minor font slots from the theme XML: {"mn-ea","mn-lt","mj-ea","mj-lt"}.
    Because this uses XML parsing, it is unaffected by quote serialization or attribute order
    (fixes a vulnerability from the earlier byte-regex era, external review 2026-07-10). Also
    used to resolve theme tokens like "+mn-lt" in run rPr (confirmed in the adversarial panel:
    treating the token as a literal font name causes an E1 false negative). None on parse
    failure."""
    try:
        from lxml import etree
        root = etree.fromstring(blob)
        out = {}
        for prefix, tag in (("mn", "minorFont"), ("mj", "majorFont")):
            base = root.find(".//" + NS + "fontScheme/" + NS + tag)
            if base is None:
                continue
            for suffix, slot in (("ea", "ea"), ("lt", "latin")):
                el = base.find(NS + slot)
                if el is not None:
                    out["%s-%s" % (prefix, suffix)] = el.get("typeface") or ""
        return out
    except Exception:
        return None


def theme_fonts_by_master(prs) -> Dict[str, Optional[Dict[str, str]]]:
    """A map of theme font slots per slide master. Key = master partname string, value = None
    on parse failure. Fixes an issue where, in a multi-master deck, grabbing the first theme
    from iter_parts falsely fired E1 based on an empty slot from an unrelated master: this is
    resolved via the master-to-theme relationship (rels) instead (external review,
    2026-07-10)."""
    out = {}
    try:
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT
        for master in prs.slide_masters:
            try:
                theme_part = master.part.part_related_by(RT.THEME)
                out[str(master.part.partname)] = _theme_fonts_from_blob(theme_part.blob)
            except Exception:
                out[str(master.part.partname)] = None
    except Exception:
        pass
    return out


def theme_ea_by_master(prs) -> Dict[str, Optional[str]]:
    """Backward compatibility: a map with only the minorFont a:ea per master (value None =
    parse failure, "" = empty slot)."""
    return {k: (v.get("mn-ea", "") if v is not None else None)
            for k, v in theme_fonts_by_master(prs).items()}


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


def theme_ea_font(prs) -> Optional[str]:
    """Backward-compatible entry point: the theme a:ea of the first master (relationship-based
    resolution).
    Returns: font name / "" (empty slot = Malgun fallback on Windows) / None (theme parse
    failure)."""
    ea_map = theme_ea_by_master(prs)
    for v in ea_map.values():
        if v is not None:
            return v
    return None


def _sz_from_defrpr(d) -> Optional[float]:
    """Converts defRPr@sz (1/100pt integer) to pt. None if absent or garbage."""
    if d is None:
        return None
    sz = d.get("sz")
    if sz is None:
        return None
    try:
        return int(sz) / 100.0
    except ValueError:
        return None


def _lst_defrpr(lst_el, lvl: int):
    """The defRPr element for the given level from an lstStyle-type container (a:lvlXpPr
    children)."""
    if lst_el is None:
        return None
    lvl = min(max(int(lvl or 0), 0), 8)
    p = lst_el.find(NS + "lvl%dpPr" % (lvl + 1))
    if p is None:
        return None
    return p.find(NS + "defRPr")


def _lst_sz_pt(lst_el, lvl: int) -> Optional[float]:
    """Backward-compatible shim: the defRPr sz (pt) at the given level in lstStyle."""
    return _sz_from_defrpr(_lst_defrpr(lst_el, lvl))


class StyleResolver:
    """When a run or paragraph has no explicit attribute, resolves the effective style via the
    OOXML inheritance chain. Size and font (a:ea/a:latin) are resolved through the same
    chain, but each attribute is searched independently.

    Chain (ECMA-376 text style hierarchy): shape txBody lstStyle -> (if a placeholder) the
    layout's same-idx placeholder lstStyle -> master placeholder lstStyle -> master
    txStyles (title/body/other) -> presentation defaultTextStyle.

    Size: None (= W5) if absent everywhere. 0.1.0 only looked at run/paragraph and let
    everything fall through to W5, which effectively killed E3/W1/W8 in placeholder-based
    decks (external review, 2026-07-10).
    Font: confirmed by measured-by-render COM probing that the master lstStyle's a:ea is
    actually inherited into rendering (2026-07-10, probe 6, docs/CALIBRATION.md). 0.2.0 only
    looked at run rPr and produced a "confirmed Malgun fallback" false positive on standard
    corporate templates (confirmed in the second external re-check)."""

    def __init__(self, prs):
        self._prs = prs
        self._default_el = None
        self._default_loaded = False
        # An inheritance-dependent deck has thousands of runs repeatedly querying the same
        # layout/master: memoize the defRPr element keyed by partname (fixes an asymmetry the
        # adversarial panel flagged, where every run re-traversed the chain)
        self._default_cache: Dict[int, object] = {}
        self._layout_cache: Dict[Tuple[str, int, int], object] = {}
        self._master_ph_cache: Dict[Tuple[str, str, int], object] = {}
        self._master_tx_cache: Dict[Tuple[str, str, int], object] = {}

    def _default_defrpr(self, lvl: int):
        if not self._default_loaded:
            self._default_loaded = True
            try:
                from lxml import etree
                root = etree.fromstring(self._prs.part.blob)
                self._default_el = root.find(NS_P + "defaultTextStyle")
            except Exception:
                self._default_el = None
        if lvl not in self._default_cache:
            self._default_cache[lvl] = _lst_defrpr(self._default_el, lvl)
        return self._default_cache[lvl]

    @staticmethod
    def _ph_family(ph_type) -> str:
        try:
            from pptx.enum.shapes import PP_PLACEHOLDER as PH
            if ph_type in (PH.TITLE, PH.CENTER_TITLE, PH.VERTICAL_TITLE):
                return "title"
            if ph_type in (PH.BODY, PH.SUBTITLE, PH.VERTICAL_BODY, PH.OBJECT):
                return "body"
        except Exception:
            pass
        return "other"

    @staticmethod
    def _ph_lst(shape_like):
        try:
            txBody = shape_like.text_frame._txBody
            return txBody.find(NS + "lstStyle")
        except Exception:
            return None

    def ph_family_of(self, sp) -> Optional[str]:
        """title/body/other if a placeholder, otherwise None. Used for E1's majorFont branch
        (the title family uses the theme majorFont ea: measured in probe 6, Q1)."""
        try:
            if getattr(sp, "is_placeholder", False):
                return self._ph_family(sp.placeholder_format.type)
        except Exception:
            pass
        return None

    def _layout_ph_defrpr(self, slide, idx: int, lvl: int):
        # The guard is per placeholder: wrapping the whole loop in one try lets a single
        # corrupted placeholder abort the entire search, missing the real style at a later
        # index (confirmed in the adversarial panel, 2026-07-10)
        try:
            layout = slide.slide_layout
            key = (str(layout.part.partname), idx, lvl)
        except Exception:
            return None
        if key in self._layout_cache:
            return self._layout_cache[key]
        found = None
        try:
            phs = list(layout.placeholders)
        except Exception:
            phs = []
        for ph in phs:
            try:
                if ph.placeholder_format.idx == idx:
                    found = _lst_defrpr(self._ph_lst(ph), lvl)
                    break
            except Exception:
                continue
        self._layout_cache[key] = found
        return found

    def _master_ph_defrpr(self, slide, family: str, lvl: int):
        try:
            master = slide.slide_layout.slide_master
            key = (str(master.part.partname), family, lvl)
        except Exception:
            return None
        if key in self._master_ph_cache:
            return self._master_ph_cache[key]
        found = None
        try:
            phs = list(master.placeholders)
        except Exception:
            phs = []
        for ph in phs:
            try:
                if self._ph_family(ph.placeholder_format.type) == family:
                    found = _lst_defrpr(self._ph_lst(ph), lvl)
                    break
            except Exception:
                continue
        self._master_ph_cache[key] = found
        return found

    def _master_tx_defrpr(self, slide, family: str, lvl: int):
        try:
            master = slide.slide_layout.slide_master
            key = (str(master.part.partname), family, lvl)
        except Exception:
            return None
        if key in self._master_tx_cache:
            return self._master_tx_cache[key]
        found = None
        try:
            tx = master.element.find(NS_P + "txStyles")
            if tx is not None:
                tag = {"title": "titleStyle", "body": "bodyStyle"}.get(family, "otherStyle")
                found = _lst_defrpr(tx.find(NS_P + tag), lvl)
        except Exception:
            found = None
        self._master_tx_cache[key] = found
        return found

    def _chain(self, tf, sp, slide, lvl: int):
        """Iterates the inheritance chain as (defRPr element, source). Nodes with no
        element are skipped."""
        try:
            own = _lst_defrpr(tf._txBody.find(NS + "lstStyle"), lvl)
            if own is not None:
                yield own, "own"
        except Exception:
            pass
        idx = ph_type = None
        try:
            if getattr(sp, "is_placeholder", False):
                pf = sp.placeholder_format
                idx, ph_type = pf.idx, pf.type
        except Exception:
            idx = ph_type = None
        if idx is not None:
            el = self._layout_ph_defrpr(slide, idx, lvl)
            if el is not None:
                yield el, "layout"
            family = self._ph_family(ph_type)
            el = self._master_ph_defrpr(slide, family, lvl)
            if el is not None:
                yield el, "master_ph"
            el = self._master_tx_defrpr(slide, family, lvl)
            if el is not None:
                yield el, "master_tx"
        el = self._default_defrpr(lvl)
        if el is not None:
            yield el, "default"

    def resolve_size(self, tf, sp, slide, lvl: int) -> Tuple[Optional[float], Optional[str]]:
        """(effective size in pt, source) or (None, None). The source distinction exists to
        prevent the title collector from flooding: an 18pt defaultTextStyle fallback is valid
        for gating purposes but is not "the title size this deck intended," so it is excluded
        from title candidacy (fixes a regression confirmed in the second external
        re-check)."""
        for el, src in self._chain(tf, sp, slide, lvl):
            v = _sz_from_defrpr(el)
            if v is not None:
                return v, src
        return None, None

    def resolve_font(self, tf, sp, slide, lvl: int, slot: str) -> Optional[str]:
        """The a:{slot} (ea/latin) typeface of defRPr from the inheritance chain. May be a
        theme token."""
        for el, _src in self._chain(tf, sp, slide, lvl):
            f = el.find(NS + slot)
            if f is not None:
                name = f.get("typeface")
                if name:
                    return name
        return None

    def resolve(self, tf, sp, slide, lvl: int) -> Optional[float]:
        """Backward-compatible entry point: effective size (pt) only. tf = the text frame
        being checked (including table cells), sp = the owning shape."""
        return self.resolve_size(tf, sp, slide, lvl)[0]


SizeResolver = StyleResolver   # backward-compatible alias (0.2.0 public name)


# AI-tell copy: only obvious cliches (a narrow dictionary = suppresses false positives).
# General words that could be legitimate in context are not included.
BUZZWORDS = (
    "시너지", "패러다임", "게임체인저", "게임 체인저", "혁신을 가속", "가치를 극대화",
    "미래를 선도", "새로운 지평", "무한한 가능성", "홀리스틱", "엔드투엔드", "엔드 투 엔드",
    "초격차", "글로벌 리더로 도약", "위대한 여정",
    "synergy", "paradigm shift", "game-changer", "game changer", "cutting-edge",
    "state-of-the-art", "seamless", "revolutionize", "leverage synerg", "holistic",
    "unlock the potential", "empower",
)
STALE_OPENINGS = (
    "오늘날", "급변하는", "4차 산업혁명 시대", "바야흐로", "현대 사회에서",
    "디지털 전환의 시대", "디지털 전환의 물결", "알아보겠습니다", "살펴보겠습니다",
    "in today's", "in the rapidly changing", "in this presentation",
)


def copy_cliche_check(page_texts, warns):
    """W11: AI-tell copy. Buzzwords are checked on every page, cliche openings only in the
    intro (p1-3)."""
    for si in sorted(page_texts):
        blob = " ".join(page_texts[si])
        low = blob.lower()
        hits = sorted({b for b in BUZZWORDS if b.lower() in low})
        if hits:
            warns.append(Finding(si, "W11", "w11_buzz", (len(hits),), ", ".join(hits[:5])))
        if si <= 3:
            op = sorted({o for o in STALE_OPENINGS if o.lower() in low})
            if op:
                warns.append(Finding(si, "W11", "w11_open", (), ", ".join(op[:5])))


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


_EFFECT_TAGS = tuple(NS + t for t in ("outerShdw", "innerShdw", "glow", "reflection"))
_3D_TAGS = tuple(NS + t for t in ("sp3d", "scene3d"))


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


# W15 text overlap: the most common defect axis in generated decks (elements pile up with
# every revision round), but the frame bbox is drawn generously by convention and can't be
# used, so this approximates the effective glyph width instead.
_W_CJK, _W_LAT, _W_SP = 0.96, 0.52, 0.28   # character-width/font-size ratio approximation
                                           # (conservative: suppresses false positives)

# Effective glyph bbox (in) per paragraph. Turned a magic index tuple into named fields
# (external review, 2026-07-10).
# sp (owning shape) is for the loc payload of W15-W17 findings (0.5.0); the coordinates are
# already the group's absolute coordinates.
GlyphBox = namedtuple("GlyphBox", "x0 y0 x1 y1 rep max_pt frame_id sp cell para field")
GlyphBox.__new__.__defaults__ = (None, None, None, False)


def _glyph_w(s, size_pt):
    w = 0.0
    for ch in s:
        if ch == " ":
            w += _W_SP
        elif is_cjk(ch) or ord(ch) > 0x2E80:
            w += _W_CJK
        else:
            w += _W_LAT
    return w * size_pt / 72.0


def _empty_para_pt(para, default_pt):
    """The effective size of an empty paragraph (a spacer): prioritizes endParaRPr/defRPr sz
    (fixes a phantom-height issue where a 4pt spacer was counted as 12pt, measured in
    adversarial verification)."""
    try:
        if para.font.size is not None:
            return para.font.size.pt
    except Exception:
        pass
    try:
        epr = para._p.find(NS + "endParaRPr")
        if epr is not None and epr.get("sz"):
            return int(epr.get("sz")) / 100.0
    except Exception:
        pass
    return default_pt


def _text_glyph_boxes(slide, default_pt=12.0, skipped=None, styler=None):
    """Approximates the effective glyph bbox (in) per paragraph. Returns
    [(x0,y0,x1,y1,representative text,max_pt,frame_id)].
    Width is summed from each run's actual size; line height reflects the actual
    line_spacing value (1.2 if absent) combined with autofit lnSpcReduction. x is placed
    using per-paragraph alignment (the frame's first explicit value if unset, otherwise
    left), which reduces misplacement in frames with mixed alignment. wrap=none
    (word_wrap=False, the python-pptx add_textbox default) extends past the frame in a
    single line, so no wrap folding is applied and the actual width is used as-is. A rotated
    frame is skipped since estimation would be invalid. Groups are converted to absolute
    coordinates via iter_shapes_geo.
    Calibrated against real-deck render comparisons plus reproduced adversarial-verification
    measurements.
    Script layer (0.2.1): vertical writing (bodyPr@vert) and frames containing RTL or
    complex-shaping scripts are skipped, since glyph-width approximation is meaningless for
    them, and if a skipped Counter is passed it tallies these and surfaces them via W18.
    0.3.1 (third external review, P0): if given a styler (StyleResolver), a run with no
    explicit size is resolved through the same inheritance chain as E3 (fixes an
    inconsistency where two different effective-style models existed in a single document).
    Native tables compute each cell's rectangle by accumulating column widths and row
    heights, so cell text is included too.
    Known limitations: a placeholder that inherits alignment from layout lstStyle falls back
    to left alignment, and skipping rotated frames is excluded from the W18 tally since it is
    standard decorative practice."""
    import math
    out = []

    def emit_frame(tframe, fx, fy, fw, fh, fid, owner_sp, cell=None, sx=1.0, sy=1.0):
        try:
            bodyPr = tframe._txBody.find(NS + "bodyPr")
            vert = bodyPr.get("vert") if bodyPr is not None else None
            if vert not in (None, "horz"):
                if skipped is not None:
                    skipped["vertical_text"] += 1
                return
        except Exception:
            bodyPr = None
        # Text-frame insets (0.6.0, external review): glyphs start inside the frame, not
        # at its edge. OOXML defaults are lIns/rIns 91440 EMU (0.1in) and tIns/bIns 45720
        # (0.05in), the same order of magnitude as W16's 0.15in tolerance, so ignoring
        # them shifted every glyph box left/up and overstated usable width. Insets live
        # in shape-local units, so group scale (sx/sy) applies.
        li, ri_, ti, bi = 91440, 91440, 45720, 45720
        try:
            if bodyPr is not None:
                li = int(bodyPr.get("lIns", li))
                ri_ = int(bodyPr.get("rIns", ri_))
                ti = int(bodyPr.get("tIns", ti))
                bi = int(bodyPr.get("bIns", bi))
        except Exception:
            pass
        fx += sx * li / EMU_PER_IN
        fy += sy * ti / EMU_PER_IN
        fw = max(fw - sx * (li + ri_) / EMU_PER_IN, 0.0)
        fh = max(fh - sy * (ti + bi) / EMU_PER_IN, 0.0)
        try:
            # All a:t descendants, not para.runs: field text (a:fld) participates in
            # geometry, so it must participate in the complex-script screen too. A
            # field-only Arabic frame used to bypass this check and get measured with
            # the Latin/CJK width model without any W18 (0.6.1, external review).
            frame_text = "".join(t.text or "" for t in tframe._txBody.iter(NS + "t"))
            if _geometry_unsupported(frame_text):
                if skipped is not None:
                    skipped["complex_script"] += 1
                return
        except Exception:
            pass
        fw2 = max(fw, 0.05)
        scale, lnred = frame_autofit(tframe)
        wrap = tframe.word_wrap is not False   # None (no attribute) = OOXML default
                                               # square = wrap
        frame_align = None
        for para in tframe.paragraphs:
            if para.alignment is not None:
                frame_align = para.alignment
                break
        paras = []   # (line_w, pmx, ptxt, factor, n, align, p_idx, field_only)
        for p_idx, para in enumerate(tframe.paragraphs):
            pmx, ptxt = 0.0, ""
            saw_field, saw_run_text = False, False
            segs = [0.0]   # widths of a:br-separated visual lines
            # Document-order walk over a:r / a:fld / a:br (0.6.0, external review): fld
            # text occupies real width and an explicit line break starts a new visual
            # line. Before this, a br-split sentence was measured as one overlong line
            # (width overstated, height understated). Falls back to para.runs on any
            # structural surprise.
            items = []
            try:
                runs_l = list(para.runs)
                r_seen = 0
                for child in para._p:
                    tag = child.tag
                    if tag == NS + "r":
                        if r_seen < len(runs_l):
                            items.append(runs_l[r_seen])
                            r_seen += 1
                    elif tag == NS + "br":
                        items.append(None)
                    elif tag == NS + "fld":
                        items.append(_FldRun(child))
                if r_seen != len(runs_l):
                    items = list(runs_l)
            except Exception:
                items = list(para.runs)
            for r in items:
                if r is None:   # a:br: next visual line
                    segs.append(0.0)
                    ptxt += " "
                    continue
                t = r.text
                if not t:
                    continue
                if isinstance(r, _FldRun):
                    saw_field = True
                else:
                    saw_run_text = True
                if r.font.size is not None:
                    sz = r.font.size.pt
                elif para.font.size is not None:
                    sz = para.font.size.pt
                else:
                    sz = None
                    if styler is not None:
                        try:
                            sz = styler.resolve(tframe, owner_sp, slide, getattr(para, "level", 0))
                        except Exception:
                            sz = None
                    if sz is None:
                        sz = default_pt
                sz *= scale
                segs[-1] += _glyph_w(t, sz)
                ptxt += t
                if sz > pmx:
                    pmx = sz
            if not ptxt.strip():
                paras.append((0.0, _empty_para_pt(para, default_pt) * scale, "", 1.2, 1,
                              None, p_idx, False))
                continue
            ls = para.line_spacing
            if ls is None:
                factor = 1.2
            elif isinstance(ls, float):
                factor = ls
            else:
                try:
                    factor = ls.pt / pmx if pmx else 1.2
                except Exception:
                    factor = 1.2
            if wrap:
                n = sum(max(1, math.ceil(w / (fw2 * 1.04))) for w in segs)
            else:
                n = len(segs)
            line_w = max(segs)
            al = para.alignment if para.alignment is not None else frame_align
            paras.append((line_w, pmx, ptxt, factor, n, al, p_idx,
                          saw_field and not saw_run_text))
        gh_total = sum(n * pmx * max(f, 0.95) * (1.0 - lnred) / 72.0
                       for (pw, pmx, ptxt, f, n, al, _pi, _fo) in paras)
        if gh_total <= 0:
            return
        va = str(tframe.vertical_anchor) if tframe.vertical_anchor is not None else ""
        if "MIDDLE" in va:
            cy = fy + max(0.0, (fh - gh_total) / 2)
        elif "BOTTOM" in va:
            cy = fy + max(0.0, fh - gh_total)
        else:
            cy = fy
        for (pw, pmx, ptxt, factor, n, al, p_idx, field_only) in paras:
            ph = n * pmx * max(factor, 0.95) * (1.0 - lnred) / 72.0
            if ptxt.strip():
                gw = pw if not wrap else min(pw, fw2)
                a = str(al) if al is not None else ""
                if "CENTER" in a:
                    x0 = fx + (fw2 - gw) / 2
                elif "RIGHT" in a:
                    x0 = fx + fw2 - gw
                else:
                    x0 = fx
                out.append(GlyphBox(x0, cy, x0 + gw, cy + ph, ptxt[:24], pmx, fid,
                                    owner_sp, cell, p_idx, field_only))
            cy += ph

    for sp, _z, xf in iter_shapes_geo(slide.shapes):
        if getattr(sp, "has_table", False):
            # Native table cells (third external review, P0: tables in auto-generated decks
            # were a geometry blind spot).
            # Cell rectangle = table origin + accumulated column widths/row heights (xf
            # scaling applied to EMU).
            geo = _geo_rect(sp, xf)
            if geo is None or geo[4]:
                continue
            tx, ty = geo[0], geo[1]
            ax = xf[0]
            ay = xf[2]
            try:
                tbl = sp.table
                col_w = [(c.width or 0) for c in tbl.columns]
                row_h = [(r.height or 0) for r in tbl.rows]
            except Exception:
                continue
            cy_off = 0
            for ri, row in enumerate(tbl.rows):
                cx_off = 0
                for ci, cell in enumerate(row.cells):
                    cw = col_w[ci] if ci < len(col_w) else 0
                    rh = row_h[ri] if ri < len(row_h) else 0
                    try:
                        # Merged regions (0.6.0, external review): continuation cells are
                        # covered by their origin; the origin's rectangle spans the merged
                        # column widths and row heights instead of a single grid cell.
                        if cell.is_spanned:
                            cx_off += cw
                            continue
                        span_w = sum(col_w[ci:ci + max(1, cell.span_width)]) or cw
                        span_h = sum(row_h[ri:ri + max(1, cell.span_height)]) or rh
                    except Exception:
                        span_w, span_h = cw, rh
                    try:
                        ctf = cell.text_frame
                    except Exception:
                        cx_off += cw
                        continue
                    emit_frame(ctf,
                               tx + ax * cx_off / EMU_PER_IN,
                               ty + ay * cy_off / EMU_PER_IN,
                               ax * span_w / EMU_PER_IN,
                               ay * span_h / EMU_PER_IN,
                               id(cell._tc), sp, cell=(ri, ci), sx=ax, sy=ay)
                    cx_off += cw
                cy_off += row_h[ri] if ri < len(row_h) else 0
            continue
        if not getattr(sp, "has_text_frame", False):
            continue
        geo = _geo_rect(sp, xf)
        if geo is None:
            continue
        fx, fy, fw, fh, rotated = geo
        if rotated:
            continue
        emit_frame(sp.text_frame, fx, fy, fw, fh, id(sp), sp, sx=xf[0], sy=xf[2])
    return out


def text_overlap_check(slide, si, warns, boxes: Optional[List[GlyphBox]] = None):
    """W15: the effective glyph regions of two different text frames overlap meaningfully
    (occlusion/collision). This is approximation-based, hence WARN. Fires only when the
    intersection area exceeds 45% of the smaller box, at most 2 findings per page.
    The 45% threshold is measured against render comparisons: cases in the 30-35% range are
    all false positives (estimated as a title under a big number, or bleed-over between two
    columns), while 60%+ are all real overlaps (e.g. a chart overrunning an exhibit label, a
    caption chip sitting on a legend).
    Intentional layering is excluded: an echo of identical text (an afterimage typography
    effect), and 1-2 character oversized glyphs (drop caps, chapter numerals like I/II).
    If boxes is given, it is used as-is without recomputation (a once-per-slide computation
    cache)."""
    if boxes is None:
        boxes = _text_glyph_boxes(slide)
    hits = []
    for i in range(len(boxes)):
        for j in range(i + 1, len(boxes)):
            a, b = boxes[i], boxes[j]
            if a.frame_id == b.frame_id:      # paragraphs of the same frame are stacked,
                                               # so excluded
                continue
            if a.rep.strip() == b.rep.strip():
                continue
            if any(len(x.rep.strip()) <= 2 and x.max_pt >= 28 for x in (a, b)):
                continue
            ix = min(a.x1, b.x1) - max(a.x0, b.x0)
            iy = min(a.y1, b.y1) - max(a.y0, b.y0)
            if ix <= 0.02 or iy <= 0.02:
                continue
            area = ix * iy
            amin = min((a.x1 - a.x0) * (a.y1 - a.y0), (b.x1 - b.x0) * (b.y1 - b.y0))
            if amin > 0 and area > 0.45 * amin:
                hits.append((area / amin, a, b))
    # Sort key is frac only: GlyphBox's sp field isn't comparable, so a full tuple comparison
    # would fail (0.5.0)
    for frac, a, b in sorted(hits, key=lambda h: h[0], reverse=True)[:2]:
        # loc: not the frame's raw bbox, but the effective glyph bbox (absolute, in) actually
        # used to judge the overlap, as-is.
        # related carries the counterpart frame so an agent can pinpoint which pair to move
        # (0.5.0).
        loc = shape_loc(a.sp, bbox=[a.x0, a.y0, a.x1 - a.x0, a.y1 - a.y0], cell=a.cell,
                        paragraph=a.para, field=a.field) or {}
        rel = shape_loc(b.sp, bbox=[b.x0, b.y0, b.x1 - b.x0, b.y1 - b.y0], cell=b.cell,
                        paragraph=b.para, field=b.field)
        if rel:
            loc["related"] = rel
        warns.append(Finding(si, "W15", "w15", (frac * 100,), "%r ~ %r" % (a.rep, b.rep),
                             loc=loc or None))


def _pic_boxes(slide, sw_in, sh_in, skipped=None):
    """The effective ink bbox (in) and z-order of non-background pictures. Returns
    [(x0,y0,x1,y1,z)].
    Full-bleed or mesh backgrounds covering 70%+ of the slide are excluded. A transparent PNG
    (e.g. a matplotlib chart) has a frame bbox much larger than its ink, which is a source of
    false positives, so it is trimmed to the alpha-opaque bbox.
    Performance budget (0.4.0, third review): for images over 25MP, alpha trimming is skipped,
    the frame bbox is used as-is, and this is disclosed via the skipped counter (prevents
    decode blowup on large batches).
    Reflects adversarial-verification measurements (2026-07-03): P mode + tRNS is converted
    to RGBA before trimming, srcRect crop is mapped by narrowing to the visible source
    window, flipH/flipV mirrors within that window, a rotated picture uses the axis-aligned
    expanded bbox but skips ink trimming since it would be invalid, and group children are
    converted to absolute coordinates."""
    out = []
    for sp, z, xf in iter_shapes_geo(slide.shapes):
        if getattr(sp, "shape_type", None) != MSO_SHAPE_TYPE.PICTURE:
            continue
        geo = _geo_rect(sp, xf)
        if geo is None:
            continue
        x, y, w, h, rotated = geo
        if w * h >= 0.7 * sw_in * sh_in:
            continue
        if not rotated:
            try:
                from PIL import Image
                import io as _io
                im = Image.open(_io.BytesIO(sp.image.blob))
                if im.width * im.height > 25_000_000:
                    if skipped is not None:
                        skipped["image_decode_budget"] += 1
                    raise RuntimeError("image decode budget")   # the except below keeps the
                                                                 # frame bbox
                if im.mode == "P" and "transparency" in im.info:
                    im = im.convert("RGBA")
                if "A" in im.getbands():
                    bb = im.getchannel("A").point(lambda a: 255 if a > 16 else 0).getbbox()
                    if bb is None:
                        continue   # fully transparent = no ink
                    iw, ih = im.size
                    wl, wt = iw * float(sp.crop_left or 0), ih * float(sp.crop_top or 0)
                    wr = iw * (1.0 - float(sp.crop_right or 0))
                    wb = ih * (1.0 - float(sp.crop_bottom or 0))
                    l, t, r, b = bb
                    l, r = max(l, wl), min(r, wr)
                    t, b = max(t, wt), min(b, wb)
                    if l >= r or t >= b:
                        continue   # no ink within the crop window
                    try:
                        x2 = sp._element.spPr.find(NS + "xfrm")
                        if x2 is not None and x2.get("flipH") == "1":
                            l, r = wl + (wr - r), wl + (wr - l)
                        if x2 is not None and x2.get("flipV") == "1":
                            t, b = wt + (wb - b), wt + (wb - t)
                    except Exception:
                        pass
                    ww, wh = (wr - wl) or 1.0, (wb - wt) or 1.0
                    x, y, w, h = (x + w * (l - wl) / ww, y + h * (t - wt) / wh,
                                  w * (r - l) / ww, h * (b - t) / wh)
            except Exception as ex:
                # The frame bbox is kept as the fallback either way, but a real decode
                # failure is no longer silent (0.6.0, external review: "nothing dies
                # silently" only held for the budget path). The budget path already
                # tallied itself above.
                if skipped is not None and str(ex) != "image decode budget":
                    skipped["image_decode"] += 1
        out.append((x, y, x + w, y + h, z, sp))
    return out


def overflow_check(slide, si, sw_in, sh_in, warns,
                   boxes: Optional[List[GlyphBox]] = None, pics: Optional[list] = None):
    """W16: off-canvas overflow. Using the frame bbox as the criterion was previously
    rejected because of the large false-positive rate from the generous-frame convention, but
    W15's effective glyph bbox resolved that objection (2026-07-03): text fires only when the
    actual character area breaches the boundary. For non-text, only pictures (ink bbox,
    trimmed) are checked: bleed where a decorative shape (e.g. a glow circle) spills off a
    corner is standard technique and not a defect (checking shapes was rejected after
    corpus render measurements)."""
    TOL_T, TOL_S = 0.15, 0.12
    if boxes is None:
        boxes = _text_glyph_boxes(slide)
    if pics is None:
        pics = _pic_boxes(slide, sw_in, sh_in)
    hits = []
    for gb in boxes:
        over = max(-gb.x0, -gb.y0, gb.x1 - sw_in, gb.y1 - sh_in)
        if over > TOL_T:
            hits.append((over, M("w16_text") % gb.rep, "t|%r" % gb.rep,
                         shape_loc(gb.sp, bbox=[gb.x0, gb.y0, gb.x1 - gb.x0, gb.y1 - gb.y0],
                                   cell=gb.cell, paragraph=gb.para, field=gb.field)))
    for (px0, py0, px1, py1, _z, psp) in pics:
        over = max(-px0, -py0, px1 - sw_in, py1 - sh_in)
        if over > TOL_S:
            hits.append((over, M("w16_pic") % (px1 - px0, py1 - py0),
                         "p|%.1fx%.1f" % (px1 - px0, py1 - py0),
                         shape_loc(psp, bbox=[px0, py0, px1 - px0, py1 - py0])))
    # Sort key is over only: a loc dict isn't comparable, so a full tuple comparison would
    # fail (0.5.0)
    for over, what, fpk, loc in sorted(hits, key=lambda h: h[0], reverse=True)[:2]:
        # fp_key: detail (what) is a locale-dependent string, so it's excluded from the
        # baseline fingerprint (fourth review)
        warns.append(Finding(si, "W16", "w16", (over,), what, fp_key=fpk, loc=loc))


def _occluder_boxes(slide, sw_in, sh_in):
    """The bbox and z of solid-fill shapes (cards, panels) sitting on top of a picture. Used
    to suppress a legitimate layout, a caption card over a photo, that was falsely caught by
    W17 (measured in adversarial verification)."""
    out = []
    for sp, z, xf in iter_shapes_geo(slide.shapes):
        if getattr(sp, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE:
            continue
        if getattr(sp, "has_text_frame", False) and sp.text_frame.text.strip():
            continue
        try:
            if sp.fill.type is None or "SOLID" not in str(sp.fill.type):
                continue
        except Exception:
            continue
        geo = _geo_rect(sp, xf)
        if geo is None:
            continue
        x, y, w, h, _rot = geo
        if w * h >= 0.9 * sw_in * sh_in:
            continue
        out.append((x, y, x + w, y + h, z))
    return out


def text_image_straddle_check(slide, si, sw_in, sh_in, warns,
                              boxes: Optional[List[GlyphBox]] = None, pics: Optional[list] = None):
    """W17: text straddles the ink boundary of a non-background picture (only 25-75% of the
    glyph is inside the image) = half on, half off, so it looks cropped or the background
    looks split. Fully on top (an overlay caption) is the contrast-gate's (W7's) jurisdiction,
    and fully off is irrelevant. Pictures under 1 square inch (icon/logo scale) are ignored.
    If a solid card sits in z-order between the photo and the text and backs 90%+ of the text
    area, this is excluded as a caption on a card rather than a straddle.
    At most 2 findings per page."""
    if pics is None:
        pics = _pic_boxes(slide, sw_in, sh_in)
    if not pics:
        return
    occl = _occluder_boxes(slide, sw_in, sh_in)
    if boxes is None:
        boxes = _text_glyph_boxes(slide)
    hits = []
    for gb in boxes:
        rep = gb.rep
        if len(rep.strip()) < 3:
            continue
        ta = (gb.x1 - gb.x0) * (gb.y1 - gb.y0)
        if ta <= 0:
            continue
        for (px0, py0, px1, py1, pz, psp) in pics:
            if (px1 - px0) * (py1 - py0) < 1.0:
                continue
            ix = min(gb.x1, px1) - max(gb.x0, px0)
            iy = min(gb.y1, py1) - max(gb.y0, py0)
            if ix <= 0 or iy <= 0:
                continue
            frac = (ix * iy) / ta
            if not (0.25 <= frac <= 0.75):
                continue
            carded = False
            for (ox0, oy0, ox1, oy1, oz) in occl:
                if oz <= pz:
                    continue
                cx = min(gb.x1, ox1) - max(gb.x0, ox0)
                cy2 = min(gb.y1, oy1) - max(gb.y0, oy0)
                if cx > 0 and cy2 > 0 and cx * cy2 >= 0.9 * ta:
                    carded = True
                    break
            if not carded:
                hits.append((frac, gb, (px0, py0, px1, py1, psp)))
    # Sort key is frac only (sp field isn't comparable, 0.5.0)
    for frac, gb, pic in sorted(hits, key=lambda h: h[0], reverse=True)[:2]:
        loc = shape_loc(gb.sp, bbox=[gb.x0, gb.y0, gb.x1 - gb.x0, gb.y1 - gb.y0],
                        cell=gb.cell, paragraph=gb.para, field=gb.field) or {}
        rel = shape_loc(pic[4], bbox=[pic[0], pic[1], pic[2] - pic[0], pic[3] - pic[1]])
        if rel:
            loc["related"] = rel
        warns.append(Finding(si, "W17", "w17", (frac * 100,), "%r" % gb.rep, loc=loc or None))


def action_title_check(titles, warns):
    """W14: a majority of titles are descriptive noun phrases (e.g. "Market Overview,"
    "Competitive Analysis") = not action titles (the MBB idea that reading only the titles
    should carry the argument). Because this uses a sentence-ending heuristic, there can be
    false negatives that misjudge a phrase as a claim (e.g. a noun ending in "da" like "bada"
    [sea]), but false positives are kept narrow: fires once per deck only when 3+ Hangul
    titles are noun phrases and they make up at least half of all titles."""
    entries = []
    for si in sorted(titles):
        txt = " ".join(titles[si][1]).strip()
        chars = [c for c in txt if not c.isspace()]
        cjk_n = sum(1 for c in chars if is_cjk(c))
        # Excludes titles with fewer than 3 Hangul characters or under 30% Hangul share: this
        # filters out big-stat numbers ("3M+ 18.3%") and short brand names that were being
        # mistaken for titles (measured in the 50-deck scan), as well as English titles.
        if len(chars) < 4 or cjk_n < 3 or cjk_n < 0.3 * len(chars):
            continue
        core = txt.rstrip(" ?!.…”’")
        # A title with a number+unit embedded (e.g. "Revenue Grows 3x") is a claim-style
        # headline even if it ends in a noun: fixes a case where a judgment based on the
        # sentence ending alone misclassified such titles as noun phrases (external review,
        # 2026-07-10).
        numeric_claim = bool(re.search(r"[0-9][0-9,.]*\s*(%|배|억|조|만|천|pp|bp|x|X|원|건|명|개)", txt))
        claim = ("?" in txt or "!" in txt or numeric_claim
                 or core.endswith(("다", "까", "요", "자", "죠", "함", "임")))
        entries.append((si, txt, claim))
    nominal = [(si, t) for si, t, c in entries if not c]
    if len(nominal) >= 3 and len(nominal) * 2 >= len(entries):
        ex = " ".join("p%d'%s'" % (si, t[:14]) for si, t in nominal[:4])
        warns.append(Finding(0, "W14", "w14", (len(nominal), len(entries)), ex))
    return entries


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


# Sentinel: an explicit color exists but the decoder cannot resolve it (hslClr, scrgbClr,
# sysClr, prstClr, tint/shade transforms...). Falling through to inherited colors produced
# W7 false positives, e.g. an explicit white hslClr run judged with an inherited black
# (0.6.1, external review): an unknown explicit color must stop resolution, not be skipped.
_COLOR_UNKNOWN = object()


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
                                 loc=shape_loc(sp, bbox=[gx, gy, gw_, gh_])))
            return "ok"
    return "ok"   # found and checked the render PNG for this page (regardless of whether
                  # W7 actually fired)


def _zip_preflight(path):
    """A cheap decompression-bomb screen before python-pptx parses the package (0.6.0,
    external review: the linter takes untrusted input, and the only budgets so far were
    per-image decode and pair caps). Entry count, total uncompressed size, and per-entry
    compression ratio are bounded. A violation raises ValueError (a hostile or broken
    file is a usage error, not a deck finding); budgets sit far above any real deck."""
    import zipfile
    try:
        with zipfile.ZipFile(path) as z:
            infos = z.infolist()
    except (OSError, zipfile.BadZipFile) as e:
        raise ValueError("not a readable pptx package: %s" % e)
    if len(infos) > 20000:
        raise ValueError("zip preflight: too many entries (%d > 20000)" % len(infos))
    total = sum(i.file_size for i in infos)
    if total > 2_000_000_000:
        raise ValueError("zip preflight: uncompressed size %d exceeds the 2GB budget" % total)
    for i in infos:
        if i.file_size > 10_000_000 and i.compress_size \
                and i.file_size / i.compress_size > 200:
            raise ValueError("zip preflight: suspicious compression ratio in %r" % i.filename)


def lint(path, hard_min=5.0, body_min=9.0, small_min=7.5, render_dir=None, ghost=None,
         strict=False, w6_sim=0.90, w6_min_cluster=3, profile=DEFAULT_PROFILE):
    """profile is an execution policy (third external review P0: applied at the engine stage,
    not as a CLI post-filter).
    Excluded rules are not run at all, so there is no O(S^2) comparison cost either, internal
    failures of an excluded rule cannot leak into W18, and library callers can use profiles
    too.

    0.4.0 breaking change: the default profile is now core (objective defects only). To also
    get AI-tell and house-style rules (E2, W6, W9-W14), pass profile="full" explicitly. Fixes
    a first-impression problem where a first-time user's no-option run exited 1 over normal
    punctuation (external strategy review, confirmed by the user)."""
    if profile not in PROFILES:
        # Fixes a typo'd profile silently behaving like full (an empty exclusion set)
        # (fourth review)
        raise ValueError("unknown profile %r (choices: %s)" % (profile, ", ".join(sorted(PROFILES))))
    _zip_preflight(path)
    prs = Presentation(path)
    sw, sh = prs.slide_width, prs.slide_height
    errors, warns = [], []
    sigs = []
    toks = {}
    page_texts = {}
    foot_tops = {}
    fx_pp = {}
    titles = {}
    excl = PROFILES.get(profile, frozenset())
    fonts_map = theme_fonts_by_master(prs)
    colors_map = theme_colors_by_master(prs) if render_dir else {}
    colors_default = next((v for v in colors_map.values() if v is not None), None)
    thm_default = next((v for v in fonts_map.values() if v is not None), None)
    deck_skipped = Counter()   # deck-level unable-to-check (used to fire W18 p00)
    if render_dir and not os.path.isdir(render_dir):
        # The path where W7 silently ran zero checks if the --render folder didn't exist
        # (third review P0): surfaced as incomplete instead
        deck_skipped["render_dir_missing"] += 1
        print(M("note_render_dir_missing") % render_dir, file=sys.stderr)
    theme_fails = sum(1 for v in fonts_map.values() if v is None)
    if theme_fails:
        # Parse failure (None) and a confirmed empty slot ("") fall back to the same
        # fallback assumption in the E1 branch.
        # This is the point where the distinction disappears, so it's surfaced via W18
        # (adversarial panel, second re-check: prevents silent collapse).
        deck_skipped["theme_parse"] = theme_fails
        print(M("note_theme_parse"), file=sys.stderr)
    styler = StyleResolver(prs)

    render_png_hits = 0
    for si, slide in enumerate(prs.slides, 1):
        # theme font slots of the master this slide uses (prevents E1 false firing in
        # multi-master decks)
        thm_fonts = thm_default
        thm_colors = colors_default
        try:
            pn = str(slide.slide_layout.slide_master.part.partname)
            thm_fonts = fonts_map.get(pn, thm_default)
            thm_colors = colors_map.get(pn, colors_default)
        except Exception:
            pass
        skipped = Counter()   # W18: tallies unable-to-check regions (surfaces silent
                               # degradation in the JSON)
        try:
            slide_part = str(slide.part.partname)   # for the part field of a finding
                                                     # location
        except Exception:
            slide_part = None
        # sig and toks are each guarded separately: wrapping them in one try would, if sig
        # succeeds but toks fails, cause the except's re-append to push sigs out of order and
        # misalign W6 page numbers (reproduced and measured in the adversarial panel).
        # Collection and checking for a rule excluded by the profile is not run at all
        # (third review P0).
        if "W6" not in excl:
            try:
                sig = slide_layout_sig(slide, sw, sh)
            except Exception as e:
                sig = ([0.0] * 24, 0)   # preserves position (sigs is indexed by page order)
                skipped["w6_sig"] += 1
                print("W6 sig skipped p%02d: %s" % (si, e), file=sys.stderr)
            sigs.append(sig)
        if "W10" not in excl:
            try:
                toks[si] = _fill_tokens(slide, sw, sh)
            except Exception as e:
                toks[si] = Counter()
                skipped["w10_tokens"] += 1
                print("W10 tokens skipped p%02d: %s" % (si, e), file=sys.stderr)
        if render_dir:
            try:
                r7 = contrast_check(slide, si, sw, sh, render_dir, warns,
                                    styler=styler, thm_colors=thm_colors, skipped=skipped)
                if r7 == "ok":
                    render_png_hits += 1
                elif r7 == "no_png":
                    # a page with a picture but no conventional render = W7 not run for this
                    # page (incomplete)
                    skipped["w7_no_render"] += 1
            except Exception as e:
                skipped["w7"] += 1
                print("W7 skipped p%02d: %s" % (si, e), file=sys.stderr)
        if "W9" not in excl:
            try:
                accent_vbars_check(slide, si, sw, sh, warns)
            except Exception as e:
                skipped["w9"] += 1
                print("W9 skipped p%02d: %s" % (si, e), file=sys.stderr)
        if "W12" not in excl or "W13" not in excl:
            try:
                foot_tops[si] = footer_top(slide, sw, sh)
                fx_pp[si] = effects_count(slide)
            except Exception as e:
                skipped["w12_w13"] += 1
                print("W12/W13 skipped p%02d: %s" % (si, e), file=sys.stderr)
        # Base geometry data is computed only once per slide (previously W15-W17 each
        # recomputed it independently: text boxes 3 times, PIL picture decode twice. Fixes
        # a large-deck performance issue, external review 2026-07-10)
        sw_in, sh_in = sw / EMU_PER_IN, sh / EMU_PER_IN
        try:
            tboxes = _text_glyph_boxes(slide, skipped=skipped, styler=styler)
        except Exception as e:
            tboxes = None
            skipped["glyph_boxes"] += 1
            print("W15/W16/W17 glyph boxes skipped p%02d: %s" % (si, e), file=sys.stderr)
        try:
            pboxes = _pic_boxes(slide, sw_in, sh_in, skipped=skipped)
        except Exception as e:
            pboxes = None
            skipped["pic_boxes"] += 1
            print("W16/W17 pic boxes skipped p%02d: %s" % (si, e), file=sys.stderr)
        # Only the axis that failed falls back to an empty list; the axis that is still alive
        # keeps being checked: fixes a shared gate where a tboxes failure was also silencing
        # the unrelated picture W16 check (2 cases reproduced and measured in the adversarial
        # panel).
        if tboxes is not None:
            try:
                text_overlap_check(slide, si, warns, boxes=tboxes)
            except Exception as e:
                skipped["w15"] += 1
                print("W15 skipped p%02d: %s" % (si, e), file=sys.stderr)
        try:
            overflow_check(slide, si, sw_in, sh_in, warns,
                           boxes=tboxes if tboxes is not None else [],
                           pics=pboxes if pboxes is not None else [])
            if tboxes is not None and pboxes is not None:
                text_image_straddle_check(slide, si, sw_in, sh_in, warns, boxes=tboxes, pics=pboxes)
        except Exception as e:
            skipped["w16_w17"] += 1
            print("W16/W17 skipped p%02d: %s" % (si, e), file=sys.stderr)
        # Core-line gates (E1-E4, W1/W5/W8). The guard is per run: a per-frame guard let one
        # run's garbage attribute swallow a real violation in a neighboring run of the same
        # frame, producing a false pass (reproduced and measured in the adversarial panel,
        # 2026-07-10). Swallowed regions are surfaced in the JSON via W18.
        try:
            frames = collect_frames(slide.shapes)
        except Exception as e:
            frames = []
            skipped["frames"] += 1
            print("E1-E4 frames skipped p%02d: %s" % (si, e), file=sys.stderr)
        for tf, w_emu, owner_sp, cell_rc, sp_xf in frames:
            try:
                fw_in = w_emu / EMU_PER_IN
                scale = frame_font_scale(tf)
                paragraphs = list(tf.paragraphs)
                frame_fam = styler.ph_family_of(owner_sp)
            except Exception as e:
                skipped["frame"] += 1
                print("E1-E4 skipped p%02d frame: %s" % (si, e), file=sys.stderr)
                continue
            for pi, para in enumerate(paragraphs):
                try:
                    runs = list(para.runs)
                    # In-document-order items (run_like, index into para.runs or None,
                    # is_fld):
                    # a:fld (an auto field) is rendered with the same rPr as a normal run, so
                    # it must pass the same gates, and a:br is treated as a single line-break
                    # character in E2 context/offsets (carried over from the fourth review,
                    # 0.5.0). If the a:r count doesn't match, falls back to the previous runs
                    # path.
                    items = []
                    try:
                        r_seen = 0
                        for child in para._p:
                            tag = child.tag
                            if tag == NS + "r":
                                if r_seen < len(runs):
                                    items.append((runs[r_seen], r_seen, False))
                                    r_seen += 1
                            elif tag == NS + "br":
                                items.append((None, None, False))
                            elif tag == NS + "fld":
                                items.append((_FldRun(child), None, True))
                        if r_seen != len(runs):
                            items = [(r, i, False) for i, r in enumerate(runs)]
                    except Exception:
                        items = [(r, i, False) for i, r in enumerate(runs)]
                    run_offs = []
                    pos = 0
                    pieces = []
                    for run_like, _ri, _isf in items:
                        piece = "\n" if run_like is None else run_like.text
                        run_offs.append(pos)
                        pos += len(piece)
                        pieces.append(piece)
                    ptext = "".join(pieces)
                    p_fonts = para_fonts(para)   # paragraph defRPr: ranks right after run
                                                 # rPr (probe 7)
                    try:
                        para_size = para.font.size
                    except Exception:
                        para_size = None
                        skipped["para_size"] += 1
                except Exception as e:
                    skipped["para"] += 1
                    print("E1-E4 skipped p%02d para: %s" % (si, e), file=sys.stderr)
                    continue
                for ii, (run, ri, is_fld) in enumerate(items):
                    try:
                        if run is None:   # a:br: contributes context only, not itself checked
                            continue
                        t = run.text
                        if not t:
                            continue
                        if not is_fld:
                            page_texts.setdefault(si, []).append(t)
                        lvl = getattr(para, "level", 0)

                        # E1: effective render font judgment (measured-by-render model, see
                        # e1_violation).
                        # Per-script judgment: Hangul uses the full measured model; kana and
                        # Hanja only fire on fonts with no CJK at all (Inter, mono families)
                        # and pass JP/SC subset fonts (third panel).
                        # A slot missing from run rPr is filled in from the lstStyle
                        # inheritance chain (measured in probe 6: the master lstStyle's a:ea
                        # is actually inherited into rendering); the title family uses the
                        # theme majorFont ea.
                        script = None
                        if has_hangul(t):
                            script = "hangul"
                        elif any(is_kana(c) or is_hanja(c) for c in t):
                            script = "cjk_other"
                        if script:
                            fonts = run_fonts(run)
                            for slot in ("ea", "latin"):
                                if slot not in fonts and slot in p_fonts:
                                    fonts[slot] = p_fonts[slot]   # paragraph defRPr
                                                                  # (probe 7: beats lstStyle)
                            for slot in ("ea", "latin"):
                                if slot not in fonts:
                                    try:
                                        inh = styler.resolve_font(tf, owner_sp, slide, lvl, slot)
                                    except Exception:
                                        inh = None
                                    if inh:
                                        fonts[slot] = inh
                            fonts = resolve_font_tokens(fonts, thm_fonts)
                            eff_thm_ea = (thm_fonts or {}).get("mj-ea" if frame_fam == "title" else "mn-ea")
                            v = e1_violation(t, fonts, eff_thm_ea, script)
                            if v is not None:
                                errors.append(Finding(si, "E1", v[0], (), v[1],
                                                      loc=shape_loc(owner_sp, paragraph=pi, run=ri, part=slide_part, cell=cell_rc, xf=sp_xf, field=is_fld)))

                        # E2: long-dash class. Context is read from the whole paragraph
                        # (ptext) but only this run's span is reported: prevents a false
                        # positive on a '2020'/'-2021' range split across run boundaries
                        # (confirmed in the adversarial panel)
                        bad = [] if "E2" in excl else \
                            dash_violations(ptext, strict=strict,
                                            span=(run_offs[ii], run_offs[ii] + len(t)))
                        if bad:
                            errors.append(Finding(si, "E2", "e2", (),
                                                  "cp=%s text=%r" % (",".join("U+%04X" % ord(c) for c in sorted(set(bad))), t[:24]),
                                                  loc=shape_loc(owner_sp, paragraph=pi, run=ri, part=slide_part, cell=cell_rc, xf=sp_xf, field=is_fld)))

                        # E3 / W1 / W5: effective font size (run -> paragraph -> placeholder
                        # inheritance chain, reflecting autofit).
                        # size_src is used to judge title-collection eligibility: a
                        # defaultTextStyle fallback (18pt) is valid for gating purposes but is
                        # not an intended title size, and would flood ghost/W14 (regression
                        # fix).
                        size_src = "explicit"
                        if run.font.size is not None:
                            base_pt = run.font.size.pt
                        elif para_size is not None:
                            base_pt = para_size.pt
                        else:
                            try:
                                base_pt, size_src = styler.resolve_size(tf, owner_sp, slide, lvl)
                            except Exception:
                                base_pt, size_src = None, None
                        if base_pt is not None:
                            eff = base_pt * scale
                            # Title eligibility: explicit size only, or a title-family
                            # placeholder. If an inherited size from lstStyle or the master
                            # bodyStyle exceeds 18pt, body prose gets swept into ghost/W14
                            # (third adversarial panel: reproduced and measured with a 20pt
                            # body using its own lstStyle)
                            # An auto field (e.g. slide number) is not a title even if its
                            # size is large (0.5.0)
                            if eff >= 18 and t.strip() and not is_fld \
                                    and (size_src == "explicit" or frame_fam == "title"):
                                # If a title placeholder exists, it is the title regardless
                                # of size: fixes a 60pt KPI big-number pushing out the real
                                # 26pt title (third review)
                                is_title_ph = frame_fam == "title"
                                cur = titles.get(si)
                                if cur is None or (is_title_ph and not cur[2]) \
                                        or (is_title_ph == cur[2] and eff > cur[0] + 0.1):
                                    titles[si] = (eff, [t], is_title_ph)
                                elif is_title_ph == cur[2] and abs(eff - cur[0]) <= 0.1:
                                    cur[1].append(t)
                            if eff < hard_min:
                                note = "" if scale == 1.0 else M("e3_note") % (base_pt, scale)
                                errors.append(Finding(si, "E3", "e3", (eff, hard_min, note), "text=%r" % t[:24],
                                                      loc=shape_loc(owner_sp, paragraph=pi, run=ri, part=slide_part, cell=cell_rc, xf=sp_xf, field=is_fld)))
                            elif eff < body_min and fw_in > 4.0 and len(ptext) >= 40:
                                warns.append(Finding(si, "W1", "w1", (eff, body_min),
                                                     "w=%.1fin len=%d text=%r" % (fw_in, len(ptext), ptext[:24]),
                                                     loc=shape_loc(owner_sp, paragraph=pi, run=ri, part=slide_part, cell=cell_rc, xf=sp_xf, field=is_fld)))
                            elif eff < small_min and has_cjk(t) and fw_in <= 4.0:
                                # Narrow frames only (<=4in): small Hangul in a wide frame is
                                # more likely to be a caption or annotation rather than being
                                # inside a mockup or card, which would conflict with the
                                # message (mockup assumption) (reflects the public hygiene
                                # audit).
                                warns.append(Finding(si, "W8", "w8", (eff, small_min),
                                                     "w=%.1fin text=%r" % (fw_in, t[:24]),
                                                     loc=shape_loc(owner_sp, paragraph=pi, run=ri, part=slide_part, cell=cell_rc, xf=sp_xf, field=is_fld)))
                        else:
                            warns.append(Finding(si, "W5", "w5", (), "text=%r" % t[:24],
                                                 loc=shape_loc(owner_sp, paragraph=pi, run=ri, part=slide_part, cell=cell_rc, xf=sp_xf, field=is_fld)))

                        # E4: 2+ consecutive Hangul/Hanja characters + meaningfully positive
                        # tracking. A run mixed with kana is excluded as Japanese (spreading
                        # kana tracking is normal design practice). 0.6.1 (external review):
                        # the run must actually contain Hangul. Tracking on a Hanja-only run
                        # is legitimate convention in Chinese typography, and flagging it as
                        # a universal ERROR contradicted the "other scripts are never falsely
                        # flagged" scope promise. Hanja still counts toward the consecutive
                        # requirement when mixed with Hangul (Korean names, legal terms).
                        if not any(is_kana(c) for c in t) \
                                and any(is_hangul(c) for c in t) \
                                and sum(1 for c in t if is_hangul(c) or is_hanja(c)) >= 2:
                            tr = run_track(run)
                            if tr is not None and tr > 50:
                                errors.append(Finding(si, "E4", "e4", (tr,), "text=%r" % t[:24],
                                                      loc=shape_loc(owner_sp, paragraph=pi, run=ri, part=slide_part, cell=cell_rc, xf=sp_xf, field=is_fld)))
                    except Exception as e:
                        skipped["run"] += 1
                        print("E1-E4 skipped p%02d run: %s" % (si, e), file=sys.stderr)

        # W18: surfaces regions a guard swallowed on this page into the output contract
        # (JSON, text).
        # If it only stayed in stderr, a CI that only looks at the exit code and JSON summary
        # would misread an incomplete check as a pass (confirmed in the adversarial panel,
        # 2026-07-10). Promoted to exit 1 under --strict.
        if skipped:
            det = ", ".join("%s=%d" % (k, v) for k, v in sorted(skipped.items()))
            warns.append(Finding(si, "W18", "w18_page", (), det))

    # If there are zero matches for the naming convention (p01.png) but PNGs with other names
    # exist, hint at the naming convention.
    # Incompleteness itself is surfaced via W18/incomplete by the per-page w7_no_render
    # counter (0.3.1).
    if render_dir and os.path.isdir(render_dir) and render_png_hits == 0:
        anypng = glob.glob(os.path.join(render_dir, "*.png"))
        if anypng:
            print(M("note_render_naming") % (render_dir, os.path.basename(anypng[0])), file=sys.stderr)

    # W6: recycled layout skeleton. Sparse slides (dividers, covers, etc. with fewer than 4
    # non-empty cells) are excluded, and a warning fires if a content slide has w6_min_cluster
    # or more other slides similar to it (>w6_sim).
    # Invariant to deck length (an overall pair ratio would dilute and miss local recycling in
    # a large deck, so this judges by the largest cluster instead).
    # Thresholds are CLI-tunable (--w6-sim/--w6-cluster): a house with strong intentional
    # template consistency can tighten them to suppress this (fixes a genre-blindness
    # complaint, external review 2026-07-10; total suppression is --skip W6).
    try:
        content = [] if "W6" in excl else \
            [(i + 1, sig) for i, (sig, n) in enumerate(sigs) if n >= 3]
        if len(content) > 200:
            # Performance budget (0.4.0): a cap on the O(S^2) pairwise comparison. The
            # truncation is disclosed.
            deck_skipped["w6_capped"] += 1
            content = content[:200]
        if len(content) >= w6_min_cluster + 1:
            adj = {p: [] for p, _ in content}
            for a in range(len(content)):
                for b in range(a + 1, len(content)):
                    pa, sa = content[a]; pb, sb = content[b]
                    sim = _cosv(sa, sb)
                    if sim > w6_sim:
                        adj[pa].append((pb, sim)); adj[pb].append((pa, sim))
            worst_p, worst = max(adj.items(), key=lambda kv: len(kv[1]))
            if len(worst) >= w6_min_cluster:
                ex = " ".join("p%d~p%d(%.2f)" % (worst_p, b, s) for b, s in sorted(worst, key=lambda x: -x[1])[:4])
                # fp_key from the skeleton signature itself, not the page list: page
                # numbers in the fingerprint broke baseline suppression on slide
                # insertion, the exact failure fingerprint v2 exists to prevent
                # (0.6.0, external verification finding)
                sig_map = dict(content)
                fpk = "w6|" + ",".join("%.2f" % v for v in sig_map.get(worst_p, ()))
                warns.append(Finding(0, "W6", "w6", (len(worst) + 1,), M("w6_detail") % ex,
                                     fp_key=fpk))
    except Exception as e:
        deck_skipped["w6"] += 1
        print("W6 skipped: %s" % e, file=sys.stderr)

    # W11 copy cliches, W12 footer alignment, W13 effects, W14 action titles (all deck-level,
    # profile-gated)
    try:
        if "W11" not in excl:
            copy_cliche_check(page_texts, warns)
        if "W12" not in excl:
            footer_check(foot_tops, warns)
        if "W13" not in excl:
            effects_check_deck(fx_pp, warns)
        if "W14" not in excl:
            action_title_check(titles, warns)
        if ghost is not None:
            # ghost is every title (18pt+) collected regardless of W14's Hangul filter: since
            # it is meant for reviewing horizontal logic by reading only titles even in an
            # English/numeric deck, it is pulled from the raw titles rather than the filtered
            # entries.
            ghost.extend((si, " ".join(titles[si][1]).strip()) for si in sorted(titles))
    except Exception as e:
        deck_skipped["w11_w14"] += 1
        print("W11/W12/W13/W14 skipped: %s" % e, file=sys.stderr)

    # W10: clone recycling of a hand-drawn diagram (e.g. cross-section). Specialized to the
    # decorative-texture (marks) path so that three-column card layouts are left to W6 and
    # not counted here (adversarial audit 2026-07-02: avoids a card-shaped subblock false
    # positive). Fills the gap W6 misses, sub-0.90 similarity single pairs (measured on
    # repeated cross-section diagrams), using pptx shapes alone.
    try:
        cadj = {} if "W10" in excl else {p: [] for p in toks}
        nums_t = sorted(cadj)
        if len(nums_t) > 200:
            deck_skipped["w10_capped"] += 1
            nums_t = nums_t[:200]
        for ia in range(len(nums_t)):
            for ib in range(ia + 1, len(nums_t)):
                pa, pb = nums_t[ia], nums_t[ib]
                if _diagram_clone_marks(toks[pa] & toks[pb]):
                    cadj[pa].append(pb); cadj[pb].append(pa)
        if cadj:
            cw, cwl = max(cadj.items(), key=lambda kv: len(kv[1]))
            if len(cwl) >= 1:
                grp = sorted({cw, *cwl})
                pages_key = ",".join("p%d" % p for p in grp)
                # fp_key from the cloned diagram's shared fill tokens, not the page list
                # (page-independent fingerprint contract, 0.6.0)
                try:
                    inter = set(toks[grp[0]])
                    for p in grp[1:]:
                        inter &= set(toks[p])
                    fpk = "w10|" + ",".join(sorted(str(t) for t in inter))[:120]
                except Exception:
                    fpk = "w10|n=%d" % len(grp)
                warns.append(Finding(0, "W10", "w10", (len(grp),),
                                     M("w10_detail") % pages_key, fp_key=fpk))
    except Exception as e:
        deck_skipped["w10"] += 1
        print("W10 skipped: %s" % e, file=sys.stderr)

    # Deck-level W18: also surfaces unable-to-check outcomes from deck-level checks (W6, W10,
    # W11-W14) and theme parsing into the output contract. 0.2.0 only tallied the geometry
    # and core gates, contradicting the documentation's "surfaces everything" claim
    # (confirmed in the second external re-check: fixes the partial W18 implementation).
    if deck_skipped:
        det = ", ".join("%s=%d" % (k, v) for k, v in sorted(deck_skipped.items()))
        warns.append(Finding(0, "W18", "w18_deck", (), det))

    return errors, warns


def _skill_res():
    """The SKILL.md skill pack bundled with the package (importlib.resources, py3.9+).
    Multi-argument joinpath requires 3.11+, so this uses a chain of / instead, which is safe
    down to the 3.9 zip loader."""
    from importlib import resources
    return resources.files("archforge") / "skills" / "archforge-pptx-lint" / "SKILL.md"


def skill_main(argv=None):
    """`archforge skill`: prints the bundled skill pack or installs it into an agent skill
    folder.
    Fixes a distribution gap where the skill pack wasn't shipped to pip users (external
    review, 2026-07-10): SKILL.md is now bundled in the wheel and can be fetched directly via
    this subcommand."""
    ap = argparse.ArgumentParser(prog="archforge skill", description=M("skill_desc"))
    ap.add_argument("--install", nargs="?", const="", metavar="DIR", help=M("help_skill_install"))
    ap.add_argument("--path", action="store_true", help=M("help_skill_path"))
    # Program-wide flag: the value was already applied by main()'s prescan, so this just
    # accepts it (without this, `archforge skill --lang ko` would die as unrecognized: measured
    # in the third adversarial panel)
    ap.add_argument("--lang", default=None, choices=("ko", "en"), help=M("help_lang"))
    a = ap.parse_args(argv)
    src = _skill_res()
    if a.path:
        print(str(src))
        return 0
    text = src.read_text(encoding="utf-8")
    if a.install is not None:
        root = a.install or os.path.join(".claude", "skills")
        dst_dir = os.path.join(root, "archforge-pptx-lint")
        os.makedirs(dst_dir, exist_ok=True)
        dst = os.path.join(dst_dir, "SKILL.md")
        with open(dst, "w", encoding="utf-8", newline="\n") as f:
            f.write(text)
        print(M("skill_installed") % dst)
        return 0
    sys.stdout.write(text)
    return 0


def _timeout_reexec(argv):
    """--timeout SECONDS: run the whole invocation in a child process with a wall clock,
    so a hostile or pathological deck cannot hang CI (0.6.x, external review: a real
    resource bound needs process isolation, and signal.alarm is POSIX-only). Re-execs
    sys.argv minus --timeout with a sentinel env var so the child does not recurse;
    subprocess timeout is portable to Windows. Returns an exit code, or None if no
    timeout was requested or this is already the child."""
    if os.environ.get("_ARCHFORGE_TIMEOUT_CHILD") == "1":
        return None
    secs = None
    rest = []
    it = iter(range(len(argv)))
    i = 0
    while i < len(argv):
        tok = argv[i]
        if tok == "--timeout" and i + 1 < len(argv):
            secs = argv[i + 1]; i += 2; continue
        if tok.startswith("--timeout="):
            secs = tok.split("=", 1)[1]; i += 1; continue
        rest.append(tok); i += 1
    if secs is None:
        return None
    try:
        secs_f = float(secs)
        if not (secs_f > 0):
            raise ValueError
    except ValueError:
        print(M("err_config") % ("--timeout must be a positive number of seconds, got %r"
                                 % secs), file=sys.stderr)
        return 2
    import subprocess
    env = dict(os.environ)
    env["_ARCHFORGE_TIMEOUT_CHILD"] = "1"
    try:
        return subprocess.call([sys.executable, "-m", "archforge"] + rest,
                               env=env, timeout=secs_f)
    except subprocess.TimeoutExpired:
        print(M("err_timeout") % secs_f, file=sys.stderr)
        return 124


def main():
    # Reconfigured before parser creation so Korean messages and argparse help don't die with
    # UnicodeEncodeError on non-UTF-8 stdout (pipes, cp949/cp1252); --help is printed during
    # parse_args, so the order here matters.
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass
    argv = sys.argv[1:]
    rc = _timeout_reexec(argv)
    if rc is not None:
        sys.exit(rc)
    # --lang must be finalized before the --help string, so this prescans before the parser
    # is created.
    # When given multiple times, the last value wins per argparse convention (third
    # adversarial panel: fixes a first-match-wins bug).
    lang_arg = None
    for i, tok in enumerate(argv):
        if tok == "--lang" and i + 1 < len(argv):
            lang_arg = argv[i + 1]
        elif tok.startswith("--lang="):
            lang_arg = tok.split("=", 1)[1]
    if lang_arg:
        set_lang(lang_arg)
    # Detecting the skill subcommand looks past leading --lang-style flags:
    # fixes `archforge --lang ko skill` being misinterpreted as "lint a file named skill"
    # (third panel)
    rest = list(argv)
    while rest and (rest[0] == "--lang" or rest[0].startswith("--lang=")):
        rest = rest[2:] if rest[0] == "--lang" else rest[1:]
    if rest and rest[0] == "skill":
        # Warns about the conflict with wanting to lint a file literally named "skill"
        # (adversarial panel finding: prevents silent misbehavior). To lint that file, call it
        # by path, e.g. `archforge ./skill`.
        if os.path.exists("skill"):
            print(M("skill_conflict"), file=sys.stderr)
        sys.exit(skill_main(rest[1:]))
    if rest and rest[0] in ("scan", "demo", "rules", "explain"):
        if os.path.exists(rest[0]) and os.path.isfile(rest[0]):
            print(M("subcmd_conflict") % (rest[0], rest[0]), file=sys.stderr)
        dispatch = {"scan": scan_main, "demo": demo_main,
                    "rules": rules_main, "explain": explain_main}
        sys.exit(dispatch[rest[0]](rest[1:]))
    if rest and rest[0] == "lint":
        # Explicit alias for single-file mode (`archforge lint deck.pptx`), so scripts
        # can be unambiguous about the subcommand (0.6.1). Drops just the token; the
        # leading --lang prefix (already prescanned) is preserved for parsing.
        if os.path.exists("lint") and os.path.isfile("lint"):
            print(M("subcmd_conflict") % ("lint", "lint"), file=sys.stderr)
        argv = argv[:len(argv) - len(rest)] + rest[1:]
    ap = argparse.ArgumentParser(prog="archforge", description=M("prog_desc"))
    ap.add_argument("pptx")
    ap.add_argument("--render", default=None, help=M("help_render"))
    ap.add_argument("--write-baseline", default=None, metavar="PATH", help=M("help_write_baseline"))
    _add_common_flags(ap)
    a = ap.parse_args(argv)

    try:
        res = _lint_one(a.pptx, a)
    except UsageError as e:
        print(str(e), file=sys.stderr)
        sys.exit(2)
    if res is None:   # --write-baseline: exit after recording
        sys.exit(0)

    if a.sarif:
        import json
        sarif_doc = _reporters.build_sarif(a.pptx, res["errors"], res["warns"])
        with open(a.sarif, "w", encoding="utf-8", newline="\n") as f:
            json.dump(sarif_doc, f, ensure_ascii=False, indent=2)

    if a.junit:
        xml_text = _reporters.build_junit_multi(
            [(a.pptx, res["errors"], res["warns"],
              set(res["summary"]["skipped_codes"]),
              res["summary"]["policy"], None)])
        with open(a.junit, "w", encoding="utf-8", newline="\n") as f:
            f.write(xml_text)

    if a.json:
        import json
        doc = _reporters.build_json_doc(a.pptx, res["errors"], res["warns"],
                                        res["ghost"], res["summary"],
                                        schema=res["schema"],
                                        capabilities=res["capabilities"],
                                        abstentions=res["abstentions"],
                                        invocation=res["invocation"],
                                        rules_split=res["rules_split"])
        print(json.dumps(doc, ensure_ascii=False, indent=2))
        sys.exit(1 if res["fail"] else 0)

    for line in _reporters.render_text(a.pptx, res["errors"], res["warns"], res["ghost"],
                                       res["profile"], res["profile_excl"], res["skip"],
                                       config_path=res["cfg_path"],
                                       baseline_suppressed=res["baseline_suppressed"],
                                       baseline_path=res["baseline_path"]):
        print(line)

    sys.exit(1 if res["fail"] else 0)


# Maps a W18 skip-reason key (machine string, stable) to the rules it prevented and the
# capability it degrades (0.7 schema 2.0). Reasons not listed default to no affected rules
# and the "meta" capability. Keys mirror the Counter keys written at each guard.
_REASON_RULES = {
    "vertical_text": (["W15", "W16", "W17"], "geometry"),
    "complex_script": (["W15", "W16", "W17"], "geometry"),
    "glyph_boxes": (["W15", "W16", "W17"], "geometry"),
    "pic_boxes": (["W16", "W17"], "geometry"),
    "w15": (["W15"], "geometry"),
    "w16_w17": (["W16", "W17"], "geometry"),
    "image_decode": (["W16", "W17"], "geometry"),
    "image_decode_budget": (["W16", "W17"], "geometry"),
    "frames": (["E1", "E3", "E4", "W1", "W5", "W8"], "typography"),
    "frame": (["E1", "E3", "E4", "W1", "W5", "W8"], "typography"),
    "para": (["E1", "E2", "E3", "E4"], "typography"),
    "para_size": (["E3"], "typography"),
    "run": (["E1", "E2", "E3", "E4"], "typography"),
    "w7": (["W7"], "render"),
    "w7_no_render": (["W7"], "render"),
    "w7_color_unknown": (["W7"], "render"),
    "render_dir_missing": (["W7"], "render"),
    "w9": (["W9"], "structure"),
    "w6_sig": (["W6"], "structure"),
    "w6": (["W6"], "structure"),
    "w6_capped": (["W6"], "structure"),
    "w10_tokens": (["W10"], "structure"),
    "w10": (["W10"], "structure"),
    "w10_capped": (["W10"], "structure"),
    "w11_w14": (["W11", "W12", "W13", "W14"], "structure"),
    "w12_w13": (["W12", "W13"], "structure"),
    "theme_parse": (["E1"], "typography"),
}

# Every skip-reason key a detector can emit must be registered above, so a structural
# abstention never lands as ([], "meta") with structure still reported "complete"
# (0.7.1, external review P0). test_reason_registry_covers_all_keys enforces this.
KNOWN_REASON_KEYS = frozenset(_REASON_RULES)


def _capabilities_and_abstentions(warns, render_requested):
    """Turns the W18 findings into a structured capabilities map and abstentions list
    (0.7 schema 2.0). W18 detail is a machine-key Counter string ('vertical_text=1, ...'),
    so parsing it back is deterministic. Verdict is untouched; this is a richer view of the
    same incompleteness signal."""
    abstentions = []
    degraded = set()
    for f in warns:
        if f.code != "W18":
            continue
        for part in (f.detail or "").split(","):
            part = part.strip()
            if "=" not in part:
                continue
            key, _, cnt = part.partition("=")
            key = key.strip()
            try:
                count = int(cnt)
            except ValueError:
                count = 1
            rules, cap = _REASON_RULES.get(key, ([], "meta"))
            degraded.add(cap)
            abstentions.append({"reason": key, "page": f.page, "count": count,
                                "affected_rules": rules})
    caps = {}
    caps["typography"] = "partial" if "typography" in degraded else "complete"
    caps["geometry"] = "partial" if "geometry" in degraded else "complete"
    caps["structure"] = "partial" if "structure" in degraded else "complete"
    caps["render_contrast"] = ("partial" if "render" in degraded
                               else ("complete" if render_requested else "not_requested"))
    # An unregistered reason (should not happen: enforced by test) still surfaces here
    # rather than vanishing into a "complete" verdict.
    if "meta" in degraded:
        caps["meta"] = "partial"
    return caps, abstentions


def _validate_cli_globals(a):
    """Validation of CLI-supplied values that are identical for every file in a scan.
    Returns an error string (caller exits 2) or None. Deck-config-supplied values stay
    per-file (a bad config next to one deck is that file's problem, not the batch's)."""
    if a.config and a.no_config:
        return M("err_config") % "--config conflicts with --no-config"
    if a.config and not os.path.exists(a.config):
        return M("err_config") % a.config
    for v, name in ((a.hard_min, "hard_min"), (a.body_min, "body_min"),
                    (a.small_min, "small_min")):
        if v is not None and (not math.isfinite(v) or v <= 0):
            return M("err_config") % ("threshold out of range: %s=%r" % (name, v))
    if a.w6_sim is not None and (not math.isfinite(a.w6_sim) or not (0 < a.w6_sim <= 1)):
        return M("err_config") % ("threshold out of range: w6_sim=%r" % a.w6_sim)
    if a.w6_cluster is not None and a.w6_cluster < 1:
        return M("err_config") % ("threshold out of range: w6_cluster=%r" % a.w6_cluster)
    if a.skip:
        codes_ = {c.strip().upper() for c in a.skip.split(",") if c.strip()}
        bad = sorted(c for c in codes_ if not c.startswith("W"))
        if bad:
            return M("err_skip_e") % ",".join(bad)
        unknown = sorted(c for c in codes_ if c not in ALL_CODES)
        if unknown:
            return M("err_skip_unknown") % ",".join(unknown)
        if "W18" in codes_:
            return M("err_skip_w18")
    return None


class UsageError(Exception):
    """A per-file usage/config error (missing file, bad config, invalid pptx, bad flags).

    Single-file mode prints it and exits 2, preserving the existing contract. scan mode
    converts it into a per-file result and keeps scanning: one broken deck must not kill
    the batch or swallow the aggregate report (0.6.0, external review P0)."""


def _pkg_version():
    try:
        from importlib.metadata import version
        return version("archforge")
    except Exception:
        return "unknown"


def _add_common_flags(ap):
    """Flags shared by single-file mode and scan mode (prevents duplicate-definition drift,
    0.5.0)."""
    ap.add_argument("--version", action="version", version="archforge " + _pkg_version())
    ap.add_argument("--hard-min", type=float, default=None, help=M("help_hard_min"))
    ap.add_argument("--body-min", type=float, default=None, help=M("help_body_min"))
    ap.add_argument("--strict", action="store_true", help=M("help_strict"))
    # --strict split into orthogonal policies (0.6.0, external review): failing every
    # advisory warning, failing on incomplete checks, and lifting E2's numeric exemptions
    # are three different decisions. --strict remains as the union for compatibility.
    ap.add_argument("--fail-on-warning", action="store_true", help=M("help_fail_on_warning"))
    ap.add_argument("--fail-incomplete", action="store_true", help=M("help_fail_incomplete"))
    ap.add_argument("--e2-no-exemptions", action="store_true", help=M("help_e2_no_exemptions"))
    ap.add_argument("--small-min", type=float, default=None, help=M("help_small_min"))
    ap.add_argument("--ghost", action="store_true", help=M("help_ghost"))
    ap.add_argument("--json", action="store_true", help=M("help_json"))
    ap.add_argument("--schema", default="1.0", choices=("1.0", "2.0", "1", "2"),
                    help=M("help_schema"))
    ap.add_argument("--skip", default=None, metavar="CODES", help=M("help_skip"))
    ap.add_argument("--profile", default=None, choices=sorted(PROFILES), help=M("help_profile"))
    ap.add_argument("--lang", default=None, choices=("ko", "en"), help=M("help_lang"))
    ap.add_argument("--w6-sim", type=float, default=None, help=M("help_w6_sim"))
    ap.add_argument("--w6-cluster", type=int, default=None, help=M("help_w6_cluster"))
    ap.add_argument("--config", default=None, metavar="PATH", help=M("help_config"))
    ap.add_argument("--no-config", action="store_true", help=M("help_no_config"))
    ap.add_argument("--sarif", default=None, metavar="PATH", help=M("help_sarif"))
    ap.add_argument("--junit", default=None, metavar="PATH", help=M("help_junit"))
    ap.add_argument("--timeout", default=None, metavar="SECONDS", help=M("help_timeout"))
    ap.add_argument("--baseline", default=None, metavar="PATH", help=M("help_baseline"))


def _lint_one(path, a):
    """One file's config resolution -> check -> filter -> summary. Shared by main (single
    mode) and scan_main.
    Usage errors (missing file, bad config, invalid pptx, bad flags) raise UsageError:
    single-file mode turns that into exit 2 (unchanged contract), scan mode turns it into
    a per-file result and keeps going (0.6.0). All flag/config validation happens BEFORE
    the lint run and before --write-baseline recording (0.6.0: a typo'd --skip used to
    record a baseline as if nothing were wrong). If --write-baseline is set, returns None
    after recording. Returned dict: errors/warns/ghost/summary/fail/profile/profile_excl/
    skip/cfg_path/baseline_suppressed/baseline_path."""
    if not os.path.exists(path):
        raise UsageError(M("err_notfound") % path)

    # Explicit --config combined with --no-config used to silently drop the explicit
    # config; that contradiction is now an error (0.6.0, external review).
    if a.config and a.no_config:
        raise UsageError(M("err_config") % "--config conflicts with --no-config")

    # Config file (.archforge.json/.yml): CLI flags always win (0.4.0).
    # Trust boundary (fourth review): since a config file in the deck folder could weaken the
    # gate, the applied config path is recorded in the output contract (JSON summary.config,
    # a text footnote) and can be turned off with --no-config.
    cfg = {}
    cfg_path = None if a.no_config else _config.find_config(path, a.config)
    if a.config and not a.no_config and cfg_path is None:
        raise UsageError(M("err_config") % a.config)
    if cfg_path:
        try:
            cfg, cfg_warns = _config.load_config(cfg_path)
            for wmsg in cfg_warns:
                print("archforge: %s (%s)" % (wmsg, cfg_path), file=sys.stderr)
        except Exception as e:
            raise UsageError(M("err_config") % ("%s (%s)" % (cfg_path, e)))

    def pick(cli_val, cfg_key, default):
        if cli_val is not None:
            return cli_val
        return cfg.get(cfg_key, default)

    try:
        hard_min = float(pick(a.hard_min, "hard_min", 5.0))
        body_min = float(pick(a.body_min, "body_min", 9.0))
        small_min = float(pick(a.small_min, "small_min", 7.5))
        w6_sim = float(pick(a.w6_sim, "w6_sim", 0.90))
        w6_cluster = int(pick(a.w6_cluster, "w6_cluster", 3))
    except (TypeError, ValueError) as e:
        raise UsageError(M("err_config") % ("threshold: %s" % e))
    # Range validation: closes a bypass where --hard-min 0 silently disabled E3 (the
    # unreadable-text block) (formerly X1). NaN slips through ordinary comparisons
    # (NaN <= 0 is False), which re-opened the same bypass via --hard-min nan or a bare
    # NaN literal in .archforge.json (json.load accepts it), so finiteness is checked
    # explicitly (0.6.0, external verification finding).
    if not all(math.isfinite(v) for v in (hard_min, body_min, small_min, w6_sim)) \
            or hard_min <= 0 or body_min <= 0 or small_min <= 0 \
            or not (0 < w6_sim <= 1) or w6_cluster < 1:
        raise UsageError(M("err_config") % (
            "threshold out of range (hard_min/body_min/small_min > 0, "
            "0 < w6_sim <= 1, w6_cluster >= 1, all finite)"))
    profile = pick(a.profile, "profile", DEFAULT_PROFILE)
    if profile not in PROFILES:
        raise UsageError(M("err_config") % ("profile=%r" % profile))
    lang_final = pick(a.lang, "lang", None)
    if lang_final:
        set_lang(lang_final)
    baseline_path = pick(a.baseline, "baseline", None)
    skip_raw = pick(a.skip, "skip", "")
    if isinstance(skip_raw, list):
        skip_raw = ",".join(str(c) for c in skip_raw)

    # --skip validation happens BEFORE the run and before baseline recording (0.6.0).
    # --skip is WARN-only: silently swallowing even E codes would be a footgun that turns
    # off the deployment-blocking gate without a trace (second external re-check). Unknown
    # codes are typos that would make CI look normal (third review P1), and W18 is an
    # incompleteness signal, not something to suppress (third review P1).
    skip = {c.strip().upper() for c in skip_raw.split(",") if c.strip()}
    bad_skip = sorted(c for c in skip if not c.startswith("W"))
    if bad_skip:
        raise UsageError(M("err_skip_e") % ",".join(bad_skip))
    unknown_skip = sorted(c for c in skip if c not in ALL_CODES)
    if unknown_skip:
        raise UsageError(M("err_skip_unknown") % ",".join(unknown_skip))
    if "W18" in skip:
        raise UsageError(M("err_skip_w18"))

    # --strict = the union of the three orthogonal policies (compatibility alias)
    fail_on_warning = a.strict or a.fail_on_warning
    fail_incomplete = a.strict or a.fail_incomplete
    e2_no_exemptions = a.strict or a.e2_no_exemptions

    ghost = [] if (a.ghost or a.json) else None
    try:
        errors, warns = lint(path, hard_min, body_min, small_min, render_dir=a.render,
                             ghost=ghost, strict=e2_no_exemptions, w6_sim=w6_sim,
                             w6_min_cluster=w6_cluster, profile=profile)
    except Exception as e:
        # ValueError carries an intentional diagnosis (zip preflight budgets); other
        # exception types stay name-only to avoid echoing arbitrary parser internals
        reason = type(e).__name__
        if isinstance(e, ValueError) and str(e):
            reason = "%s: %s" % (reason, e)
        raise UsageError(M("err_open") % (path, reason))

    # Baseline recording mode: saves current violations (excluding the W18 incompleteness
    # signal) as a fingerprint and exits. Runs only after all validation above.
    if getattr(a, "write_baseline", None):
        n = _config.write_baseline(a.write_baseline,
                                   [f for f in list(errors) + list(warns) if f.code != "W18"],
                                   profile=profile, lang=get_lang(),
                                   thresholds={"hard_min": hard_min, "body_min": body_min,
                                               "small_min": small_min, "w6_sim": w6_sim,
                                               "w6_cluster": w6_cluster})
        print(M("baseline_written") % (n, a.write_baseline))
        return None

    # Incompleteness is determined before filtering: even if W18 is --skip'd, the
    # machine-readable signal remains
    has_w18 = any(w[1] == "W18" for w in warns)
    baseline_suppressed = 0
    if baseline_path:
        try:
            known = _config.load_baseline(baseline_path)
        except Exception as e:
            raise UsageError(M("err_config") % ("baseline %s (%s)" % (baseline_path, e)))
        # Recorded run conditions are checked, not just stored (0.6.0, external review):
        # a baseline made under a different profile or tool version suppresses different
        # things than the reader expects. Warning, not error: baselines are beta.
        meta = _config.load_baseline_meta(baseline_path)
        if meta.get("profile") not in (None, "", profile):
            print(M("note_baseline_meta") % ("profile", meta.get("profile"), profile),
                  file=sys.stderr)
        rec_v = str(meta.get("tool_version") or "")
        cur_v = _pkg_version()
        if rec_v and cur_v != "unknown" and rec_v.split(".")[:2] != cur_v.split(".")[:2]:
            print(M("note_baseline_meta") % ("tool_version", rec_v, cur_v), file=sys.stderr)
        rec_thr = meta.get("threshold_hash")
        if rec_thr:
            import hashlib as _hl
            cur_thr = {"hard_min": hard_min, "body_min": body_min, "small_min": small_min,
                       "w6_sim": w6_sim, "w6_cluster": w6_cluster}
            cur_hash = _hl.sha1((",".join("%s=%r" % (k, cur_thr[k]) for k in sorted(cur_thr)))
                                .encode("utf-8")).hexdigest()[:12]
            if cur_hash != rec_thr:
                print(M("note_baseline_meta") % ("thresholds", rec_thr, cur_hash),
                      file=sys.stderr)
        errors, s1 = _config.apply_baseline(errors, known)
        warns, s2 = _config.apply_baseline(warns, known)
        baseline_suppressed = s1 + s2
    # Profile exclusions were already not run at the engine stage (0.3.1). Only --skip is
    # filtered here; the applied skip is recorded in the JSON summary to leave a trace.
    profile_excl = PROFILES[profile]
    excluded = skip | profile_excl
    if skip:
        warns = [w for w in warns if w[1] not in skip]

    schema = "2.0" if str(getattr(a, "schema", "1.0")) in ("2", "2.0") else "1.0"
    caps, abstentions = _capabilities_and_abstentions(warns, bool(a.render))
    # schema 2.0 invocation + rule accounting: skipped_codes mixed profile exclusion and
    # --skip, which mean different things to a consumer (0.7.1, external review section 5).
    invocation = {"profile": profile,
                  "policy": {"fail_on_warning": bool(fail_on_warning),
                             "fail_incomplete": bool(fail_incomplete),
                             "e2_no_exemptions": bool(e2_no_exemptions)},
                  "config": cfg_path,
                  "thresholds": {"hard_min": hard_min, "body_min": body_min,
                                 "small_min": small_min, "w6_sim": w6_sim,
                                 "w6_cluster": w6_cluster}}
    rules_split = {"executed": sorted(ALL_CODES - excluded),
                   "profile_excluded": sorted(profile_excl),
                   "user_suppressed": sorted(skip)}
    fail = bool(errors or (fail_on_warning and warns) or (fail_incomplete and has_w18))
    summary = {"error_count": len(errors), "warn_count": len(warns),
               "pass": not fail,
               # The active failure policy travels with the verdict (0.6.1, external
               # review): identical counts can pass or fail depending on flags, and a
               # JSON consumer could not tell why. Gate on summary.pass.
               "policy": {"fail_on_warning": bool(fail_on_warning),
                          "fail_incomplete": bool(fail_incomplete),
                          "e2_no_exemptions": bool(e2_no_exemptions)},
               "incomplete": has_w18,
               "profile": profile,
               "skipped_codes": sorted(excluded),
               "baseline_suppressed": baseline_suppressed,
               "config": cfg_path}   # always makes visible which config adjusted the gate
                                     # (trust boundary)

    return {"errors": errors, "warns": warns, "ghost": ghost, "summary": summary,
            "fail": fail,
            "profile": profile, "profile_excl": profile_excl, "skip": skip,
            "cfg_path": cfg_path, "baseline_suppressed": baseline_suppressed,
            "baseline_path": baseline_path,
            "schema": schema, "capabilities": caps, "abstentions": abstentions,
            "invocation": invocation, "rules_split": rules_split}


def _expand_scan_paths(patterns):
    """Expands scan arguments: a directory recurses for .pptx files, a glob pattern is
    globbed, and anything else is taken as a literal file path.
    PowerPoint lock files (~$*.pptx) are excluded, and duplicates are removed while
    preserving order. Returns (files, per_pattern_counts): match counts are tracked per
    input so a typo'd second pattern cannot hide behind a first pattern that matched
    (0.6.1, external review P0: zero-match was only detected for the whole set)."""
    out = []
    counts = []
    for pat in patterns:
        n0 = len(out)
        if os.path.isdir(pat):
            for root, dirs, files in os.walk(pat):
                # os.walk's directory order is filesystem-dependent; sorting both lists
                # makes aggregate output deterministic across machines (0.6.0, external
                # review: matters for snapshot tests and CI diffs)
                dirs.sort()
                for fn in sorted(files):
                    if fn.lower().endswith(".pptx") and not fn.startswith("~$"):
                        out.append(os.path.join(root, fn))
        elif any(ch in pat for ch in "*?["):
            for p in sorted(glob.glob(pat, recursive=True)):
                if p.lower().endswith(".pptx") and not os.path.basename(p).startswith("~$"):
                    out.append(p)
        else:
            # A literal path counts as matched here; a missing file becomes a per-file
            # error entry downstream (visible, not silent)
            out.append(pat)
        counts.append((pat, len(out) - n0))
    seen, uniq = set(), []
    for p in out:
        k = os.path.normcase(os.path.abspath(p))
        if k not in seen:
            seen.add(k)
            uniq.append(p)
    return uniq, counts


def scan_main(argv=None):
    """`archforge scan PATHS...`: lints multiple files, directories, and globs in one run
    (0.5.0, for CI/pre-commit use). Per-file judgment goes through the same path as single
    mode (_lint_one). Batch failure semantics (0.6.0, external review P0): a broken or
    misconfigured file becomes a per-file "error" entry and the scan continues; it never
    aborts the batch or swallows the aggregate report. Exit is 1 if any file fails or
    errors. Zero matches is not a silent pass; it exits 2 (prevents a CI footgun)."""
    ap = argparse.ArgumentParser(prog="archforge scan", description=M("scan_desc"))
    ap.add_argument("paths", nargs="+", help=M("help_scan_paths"))
    ap.add_argument("--allow-empty-pattern", action="store_true",
                    help=M("help_allow_empty_pattern"))
    _add_common_flags(ap)
    # Flags exclusive to single-file mode are not supported: only defaults for _lint_one
    # compatibility are seeded here
    # (--render doesn't make sense in scan since the page render folder differs per deck)
    ap.set_defaults(render=None, write_baseline=None)
    a = ap.parse_args(argv)

    # Global usage errors are not per-file errors (0.6.1, external review): a bad CLI
    # flag must exit 2 up front, not degrade into N identical per-file entries.
    err = _validate_cli_globals(a)
    if err:
        print(err, file=sys.stderr)
        return 2

    files, pattern_counts = _expand_scan_paths(a.paths)
    empty = [pat for pat, n in pattern_counts if n == 0]
    if empty and not a.allow_empty_pattern:
        # One input matching nothing must not hide behind another that matched: a typo'd
        # glob or a broken build directory silently passing is the exact CI footgun the
        # whole-set zero check missed (0.6.1, external review P0)
        print(M("err_scan_pattern_empty") % "; ".join(empty), file=sys.stderr)
        return 2
    if not files:
        print(M("err_scan_none") % " ".join(a.paths), file=sys.stderr)
        return 2

    # A single CLI baseline applied across many decks is unsafe: fingerprints carry no
    # file identity, so a finding accepted in deck A would suppress the same finding in
    # deck B (0.6.0, external review P0). Per-deck baselines still work via each deck
    # folder's config file.
    if a.baseline and len(files) > 1:
        print(M("err_scan_baseline"), file=sys.stderr)
        return 2

    # A deck folder's config may set "lang"; without restoration it would leak into every
    # later file and, because messages render lazily, into the whole aggregate report
    # (0.6.0, external verification finding). The scan report renders in the language the
    # scan was invoked with; per-deck config lang is a single-file-mode feature.
    lang0 = get_lang()
    results = []   # (path, res_dict_or_None, usage_error_message_or_None)
    for path in files:
        try:
            res = _lint_one(path, a)
            results.append((path, res, None))
        except UsageError as e:
            print(str(e), file=sys.stderr)
            results.append((path, None, str(e)))
        finally:
            set_lang(lang0)

    ok = [(p, r) for (p, r, _e) in results if r is not None]
    errored = [(p, e) for (p, r, e) in results if r is None]
    failed = sum(1 for (_p, r) in ok if r["fail"]) + len(errored)

    if a.sarif:
        import json
        sarif_doc = _reporters.build_sarif_multi(
            [(p, r["errors"], r["warns"]) for (p, r) in ok])
        with open(a.sarif, "w", encoding="utf-8", newline="\n") as f:
            json.dump(sarif_doc, f, ensure_ascii=False, indent=2)

    if a.junit:
        junit_items = []
        for p, r, e in results:
            if r is None:
                junit_items.append((p, [], [], set(),
                                    {"fail_on_warning": False, "fail_incomplete": False}, e))
            else:
                junit_items.append((p, r["errors"], r["warns"],
                                    set(r["summary"]["skipped_codes"]),
                                    r["summary"]["policy"], None))
        with open(a.junit, "w", encoding="utf-8", newline="\n") as f:
            f.write(_reporters.build_junit_multi(junit_items))

    if a.json:
        import json
        # The scan report is its own document type. Its root schema_version tracks the
        # per-file schema so a consumer that keys on the root does not misparse a v2 file
        # object under a v1 root (0.7.1, external review P0). file_schema_version names the
        # per-file shape explicitly.
        file_schema = "2.0" if str(getattr(a, "schema", "1.0")) in ("2", "2.0") else "1.0"
        root_schema = "scan-2.0" if file_schema == "2.0" else "scan-1.0"
        docs = []
        for p, r, e in results:
            if r is None:
                docs.append({"file": p, "status": "error", "error": e})
            else:
                doc = _reporters.build_json_doc(p, r["errors"], r["warns"], r["ghost"],
                                                r["summary"], schema=r["schema"],
                                                capabilities=r["capabilities"],
                                                abstentions=r["abstentions"],
                                                invocation=r["invocation"],
                                                rules_split=r["rules_split"])
                doc["status"] = "fail" if r["fail"] else "pass"
                docs.append(doc)
        agg = {"schema_version": root_schema,
               "kind": "scan-report",
               "file_schema_version": file_schema,
               "tool": {"name": "archforge", "version": _reporters._tool_version()},
               "lang": get_lang(),
               "scan": {"inputs": [{"pattern": pat, "matches": n}
                                   for pat, n in pattern_counts]},
               "files": docs,
               "summary": {"file_count": len(results), "failed_files": failed,
                           "error_files": len(errored),
                           "error_count": sum(len(r["errors"]) for (_p, r) in ok),
                           "warn_count": sum(len(r["warns"]) for (_p, r) in ok),
                           "pass": failed == 0,
                           "incomplete": any(r["summary"]["incomplete"] for (_p, r) in ok)}}
        print(json.dumps(agg, ensure_ascii=False, indent=2))
        return 1 if failed else 0

    for p, r, e in results:
        if r is None:
            print(M("scan_file_error") % (p, e))
            print()
            continue
        for line in _reporters.render_text(p, r["errors"], r["warns"], r["ghost"],
                                           r["profile"], r["profile_excl"], r["skip"],
                                           config_path=r["cfg_path"],
                                           baseline_suppressed=r["baseline_suppressed"],
                                           baseline_path=r["baseline_path"]):
            print(line)
        print()
    print(M("scan_summary") % (len(results), failed))
    return 1 if failed else 0


def rules_main(argv=None):
    """`archforge rules`: one line per rule so users can discover the gate set without
    opening the README (0.6.1, external review)."""
    ap = argparse.ArgumentParser(prog="archforge rules", description=M("rules_desc"))
    ap.add_argument("--json", action="store_true", help=M("help_json"))
    ap.add_argument("--lang", default=None, choices=("ko", "en"), help=M("help_lang"))
    a = ap.parse_args(argv)
    from .rules import TITLES, category
    rows = []
    for code in sorted(ALL_CODES, key=lambda c: (c[0] != "E", int(c[1:]))):
        profiles = sorted(p for p, excl in PROFILES.items() if code not in excl)
        rows.append({"code": code, "severity": severity(code), "category": category(code),
                     "title": TITLES.get(code, code), "profiles": profiles})
    if a.json:
        import json
        print(json.dumps(rows, ensure_ascii=False, indent=2))
        return 0
    for r in rows:
        print("%-4s %-8s %-11s %s" % (r["code"], r["severity"], r["category"], r["title"]))
    return 0


def explain_main(argv=None):
    """`archforge explain CODE`: what a rule means, when it fires, and how to fix it,
    from the same fix guidance the agent skill pack teaches (0.6.1)."""
    ap = argparse.ArgumentParser(prog="archforge explain", description=M("explain_desc"))
    ap.add_argument("code")
    ap.add_argument("--json", action="store_true", help=M("help_json"))
    ap.add_argument("--lang", default=None, choices=("ko", "en"), help=M("help_lang"))
    a = ap.parse_args(argv)
    from .rules import TITLES, category
    code = a.code.strip().upper()
    if code not in ALL_CODES:
        print(M("err_skip_unknown") % code, file=sys.stderr)
        return 2
    profiles = sorted(p for p, excl in PROFILES.items() if code not in excl)
    doc = {"code": code, "severity": severity(code), "category": category(code),
           "title": TITLES.get(code, code), "profiles": profiles,
           "fix": M("fix_" + code.lower()),
           "help_uri": "https://github.com/Love-Ash/archforge/blob/main/docs/rules/%s.md" % code}
    if a.json:
        import json
        print(json.dumps(doc, ensure_ascii=False, indent=2))
        return 0
    print("%s  %s" % (code, doc["title"]))
    print("  severity: %s | category: %s | profiles: %s"
          % (doc["severity"], doc["category"], ",".join(profiles)))
    print("  fix: %s" % doc["fix"])
    print("  docs: %s" % doc["help_uri"])
    return 0


def demo_main(argv=None):
    """`archforge demo`: generates a deck seeded with defects and its corrected version, then
    lints them on the spot (0.5.0 onboarding).
    Serves as a first-run experience that shows, within 30 seconds of installing, exactly
    what the tool catches."""
    ap = argparse.ArgumentParser(prog="archforge demo", description=M("demo_desc"))
    ap.add_argument("--dir", default="archforge-demo", help=M("help_demo_dir"))
    ap.add_argument("--lang", default=None, choices=("ko", "en"), help=M("help_lang"))
    a = ap.parse_args(argv)
    try:
        from . import demo as _demo
    except ImportError:
        import demo as _demo
    os.makedirs(a.dir, exist_ok=True)
    broken = os.path.join(a.dir, "broken.pptx")
    fixed = os.path.join(a.dir, "fixed.pptx")
    # Deck text also follows the report language (an English user gets an English demo deck,
    # 0.5.0)
    deck_lang = get_lang() if get_lang() in ("ko", "en") else "en"
    _demo.build_broken(broken, lang=deck_lang)
    _demo.build_fixed(fixed, lang=deck_lang)
    print(M("demo_built") % a.dir)
    print()
    rc = 0
    for path in (broken, fixed):
        errors, warns = lint(path, profile="full")
        for line in _reporters.render_text(path, errors, warns, None,
                                           "full", PROFILES["full"], set()):
            print(line)
        print()
        if path == fixed and (errors or warns):
            rc = 1   # the corrected version must always be clean (a contract pinned by tests)
    print(M("demo_next") % broken)
    return rc


if __name__ == "__main__":
    main()
