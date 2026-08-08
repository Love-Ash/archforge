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

This module is the engine: it walks a deck and returns findings. The command line that
wraps it (flags, subcommands, exit codes, report files) lives in cli.py.

The functional dependency runs one way, cli -> lint: nothing here calls the CLI. The import
graph is not one-way, though, and saying so would be false. The module __getattr__ at the
bottom imports cli to forward the old names, so `import archforge` traverses
lint -> cli -> lint. PEP 562 defers that edge past module execution; it does not remove it,
and a static import checker will report the cycle.
"""
import os
import sys
import glob
from collections import Counter

from pptx import Presentation

try:
    from .messages import M, set_lang, get_lang
    from .findings import Finding, shape_loc
    from .rules import RULES, ALL_CODES, PROFILES, DEFAULT_PROFILE, severity
except ImportError:   # fallback for standalone file execution (python lint.py)
    from messages import M, set_lang, get_lang
    from findings import Finding, shape_loc
    from rules import RULES, ALL_CODES, PROFILES, DEFAULT_PROFILE, severity

try:
    from .geometry import (_pct_attr,
                     frame_autofit,
                     frame_font_scale,
                     _group_xf,
                     collect_frames,
                     iter_shapes,
                     _is_pic,
                     iter_shapes_geo,
                     _geo_rect)
except ImportError:   # standalone execution
    from geometry import (_pct_attr,
                    frame_autofit,
                    frame_font_scale,
                    _group_xf,
                    collect_frames,
                    iter_shapes,
                    _is_pic,
                    iter_shapes_geo,
                    _geo_rect)
try:
    from .resolution import (_theme_fonts_from_blob,
                     theme_fonts_by_master,
                     theme_ea_by_master,
                     theme_ea_font,
                     _sz_from_defrpr,
                     _lst_defrpr,
                     _lst_sz_pt,
                     StyleResolver, SizeResolver)
except ImportError:   # standalone execution
    from resolution import (_theme_fonts_from_blob,
                    theme_fonts_by_master,
                    theme_ea_by_master,
                    theme_ea_font,
                    _sz_from_defrpr,
                    _lst_defrpr,
                    _lst_sz_pt,
                    StyleResolver, SizeResolver)
try:
    from .detectors_text import (BUZZWORDS,
                     STALE_OPENINGS,
                     copy_cliche_check,
                     _EN_CLAIM_VERBS,
                     _EN_STRUCTURAL_TITLES,
                     _en_title_key,
                     _en_eligible_title,
                     _en_is_claim,
                     _KO_STRUCTURAL_TITLES,
                     _KO_NOUN_FINALS,
                     _KO_INTERROGATIVE_SUFFIX,
                     _KO_SENTENCE_ENDINGS,
                     _KO_NUMERIC_CLAIM,
                     _ko_title_key,
                     _ko_is_claim,
                     action_title_check)
except ImportError:   # standalone execution
    from detectors_text import (BUZZWORDS,
                    STALE_OPENINGS,
                    copy_cliche_check,
                    _EN_CLAIM_VERBS,
                    _EN_STRUCTURAL_TITLES,
                    _en_title_key,
                    _en_eligible_title,
                    _en_is_claim,
                    _KO_STRUCTURAL_TITLES,
                    _KO_NOUN_FINALS,
                    _KO_INTERROGATIVE_SUFFIX,
                    _KO_SENTENCE_ENDINGS,
                    _KO_NUMERIC_CLAIM,
                    _ko_title_key,
                    _ko_is_claim,
                    action_title_check)
try:
    from .detectors_visual import (_EFFECT_TAGS,
                     _3D_TAGS,
                     accent_vbars_check,
                     _fill_tokens,
                     footer_top,
                     footer_check,
                     effects_count,
                     effects_check_deck,
                     _diagram_clone_marks,
                     slide_layout_sig,
                     contrast_check)
except ImportError:   # standalone execution
    from detectors_visual import (_EFFECT_TAGS,
                    _3D_TAGS,
                    accent_vbars_check,
                    _fill_tokens,
                    footer_top,
                    footer_check,
                    effects_count,
                    effects_check_deck,
                    _diagram_clone_marks,
                    slide_layout_sig,
                    contrast_check)
try:
    from .detectors_geometry import (_FldRun,
                     GlyphBox,
                     _glyph_w,
                     _empty_para_pt,
                     _text_glyph_boxes,
                     text_overlap_check,
                     _pic_boxes,
                     overflow_check,
                     _occluder_boxes,
                     text_image_straddle_check)
except ImportError:   # standalone execution
    from detectors_geometry import (_FldRun,
                    GlyphBox,
                    _glyph_w,
                    _empty_para_pt,
                    _text_glyph_boxes,
                    text_overlap_check,
                    _pic_boxes,
                    overflow_check,
                    _occluder_boxes,
                    text_image_straddle_check)
try:
    from .ooxml import EMU_PER_IN, NS, NS_P
    from .colors import (_shape_fill_hex, _shape_line_hex, _is_accent,
                         _theme_colors_from_blob, theme_colors_by_master,
                         _cosv, _luma, _run_rgb, _resolve_run_rgb,
                         _COLOR_UNKNOWN)
except ImportError:   # standalone execution
    from ooxml import EMU_PER_IN, NS, NS_P
    from colors import (_shape_fill_hex, _shape_line_hex, _is_accent,
                        _theme_colors_from_blob, theme_colors_by_master,
                        _cosv, _luma, _run_rgb, _resolve_run_rgb,
                        _COLOR_UNKNOWN)

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

# E1 font kernel and E2 dash kernel moved to fonts.py / dashes.py (0.8, #5
# decomposition: resolution/detection kernels as pure modules). Re-exported here so
# existing callers (jl.e1_violation, jl.dash_violations, deck_lint's star-import)
# keep working unchanged.
try:
    from .fonts import (LATIN_ONLY_FONTS, KOREAN_CAPABLE_EXCEPTIONS,
                        JP_CN_CAPABLE_PREFIXES, run_fonts, para_fonts, run_track,
                        _text_point_attr, _UNIVERSAL_PER_PT, is_latin_only_font,
                        e1_violation, resolve_font_tokens)
    from .dashes import (LONG_DASHES, _EN_DASH, _MINUS, _is_digit_ch,
                         _dash_neighbor, dash_violations)
except ImportError:
    from fonts import (LATIN_ONLY_FONTS, KOREAN_CAPABLE_EXCEPTIONS,
                       JP_CN_CAPABLE_PREFIXES, run_fonts, para_fonts, run_track,
                       _text_point_attr, _UNIVERSAL_PER_PT, is_latin_only_font,
                       e1_violation, resolve_font_tokens)
    from dashes import (LONG_DASHES, _EN_DASH, _MINUS, _is_digit_ch,
                        _dash_neighbor, dash_violations)


def collect_wireframe(path):
    """Per-slide shape geometry for the annotated HTML report (0.8.x): slide size in
    inches and, per slide, every shape's absolute bbox with a coarse kind
    (text/picture/other) and a short text sample. Read-only and best-effort: a shape
    whose geometry cannot be resolved is simply omitted (the report draws what it can;
    findings carry their own bboxes independently)."""
    prs = Presentation(path)
    sw, sh = prs.slide_width / EMU_PER_IN, prs.slide_height / EMU_PER_IN
    slides = []
    for slide in prs.slides:
        shapes = []
        for sp, _z, xf in iter_shapes_geo(slide.shapes):
            geo = _geo_rect(sp, xf)
            if geo is None:
                continue
            x, y, w, h, _rot = geo
            if _is_pic(sp):
                kind = "picture"
            elif getattr(sp, "has_text_frame", False) and sp.text_frame.text.strip():
                kind = "text"
            else:
                kind = "other"
            sample = ""
            try:
                if kind == "text":
                    sample = sp.text_frame.text.strip().splitlines()[0][:40]
            except Exception:
                pass
            shapes.append({"x": x, "y": y, "w": w, "h": h, "kind": kind,
                           "text": sample})
        slides.append(shapes)
    return {"sw": sw, "sh": sh, "slides": slides}


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
    # A package that passes the zip preflight can still be un-parseable (a truncated or
    # attribute-corrupted part). python-pptx raises lxml/opc errors that are not
    # ValueError; the library contract is a controlled ValueError the CLI maps to a clean
    # exit, never a raw traceback (0.7.1, external review: the fuzzer must see no other
    # exception class escape).
    try:
        prs = Presentation(path)
        sw, sh = prs.slide_width, prs.slide_height
    except ValueError:
        raise
    except Exception as e:
        raise ValueError("could not open %r as a valid .pptx: %s: %s"
                         % (path, type(e).__name__, e))
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
                                                      data=v[2] if len(v) > 2 else None,
                                                      loc=shape_loc(owner_sp, paragraph=pi, run=ri, part=slide_part, cell=cell_rc, xf=sp_xf, field=is_fld)))

                        # E2: long-dash class. Context is read from the whole paragraph
                        # (ptext) but only this run's span is reported: prevents a false
                        # positive on a '2020'/'-2021' range split across run boundaries
                        # (confirmed in the adversarial panel)
                        bad = [] if "E2" in excl else \
                            dash_violations(ptext, strict=strict,
                                            span=(run_offs[ii], run_offs[ii] + len(t)))
                        if bad:
                            cps = ["U+%04X" % ord(c) for c in sorted(set(bad))]
                            errors.append(Finding(si, "E2", "e2", (),
                                                  "cp=%s text=%r" % (",".join(cps), t[:24]),
                                                  data={"characters": cps,
                                                        "function": "sentence_punctuation",
                                                        "strict": bool(strict)},
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
                                                      data={"nominal_pt": base_pt,
                                                            "autofit_scale": scale,
                                                            "size_source": size_src},
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
                                # OOXML spc is hundredths of a point: state the unit
                                # explicitly (0.8, external review: raw "tracking": 200
                                # left the consumer guessing pt vs raw)
                                errors.append(Finding(si, "E4", "e4", (tr,), "text=%r" % t[:24],
                                                      data={"tracking_raw_hundredths_pt": tr,
                                                            "tracking_pt": tr / 100.0},
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


# The command-line layer moved to cli.py (0.8, #5 decomposition). It imports the engine
# from here, so re-exporting it back with a plain import would close a cycle; PEP 562
# resolves the names on first attribute access instead, once both modules are loaded.
# This keeps `archforge.lint:main`, `python -m archforge.lint` and the test suite's
# `jl.KNOWN_REASON_KEYS` working unchanged.
_CLI_NAMES = frozenset({
    "_skill_res",
    "skill_main",
    "_timeout_reexec",
    "main",
    "_REASON_RULES",
    "KNOWN_REASON_KEYS",
    "_capabilities_and_abstentions",
    "_check_out_dir",
    "_html_thumbs",
    "_write_html_report",
    "_validate_cli_globals",
    "UsageError",
    "_pkg_version",
    "_add_common_flags",
    "_lint_one",
    "_expand_scan_paths",
    "scan_main",
    "rules_main",
    "fix_main",
    "baseline_main",
    "explain_main",
    "demo_main",
})


def __getattr__(name):
    if name in _CLI_NAMES:
        try:
            from . import cli
        except ImportError:   # standalone execution
            import cli
        return getattr(cli, name)
    raise AttributeError("module %r has no attribute %r" % (__name__, name))


def __dir__():
    return sorted(set(globals()) | _CLI_NAMES)


if __name__ == "__main__":   # `python -m archforge.lint`, used by the test harness
    try:
        from .cli import main as _main
    except ImportError:   # standalone execution
        from cli import main as _main
    _main()
