# -*- coding: utf-8 -*-
"""E1-E4 typography and font-resolution rules (split from test_lint.py, 0.8.x)."""
from helpers import *   # noqa: F401,F403
from helpers import (_by, _add_fld, _repo_root, _require_repo_file,
                     patch_theme_fonts)   # noqa: F401


def test_e1_nofont_empty_theme_slot(tmp_path):
    """Font unspecified + empty theme ea = Malgun fallback E1. Since 0.2.1, the
    +mn-lt token in defaultTextStyle is resolved through the inheritance chain
    (effective latin=Calibri), so it's caught via a more accurate branch; pin it
    to the code rather than the message."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "폰트 미지정 한글", size=12)
    errors, _w = lint_full(save(p, tmp_path, "fx.pptx"))
    assert any(c == "E1" for (_si, c, m, _d) in errors), errors

def test_e1_render_model_slots(tmp_path):
    """Measured render model regression (2026-07-10 COM probe):
    p1 run ea is latin-only -> E1. p2 with an empty theme, latin is latin-only
    (Inter, previously a missed detection) -> E1.
    p3 with an empty theme, latin is a Hangul font -> that font actually renders
    the Hangul (a legitimate pattern) -> passes.
    p4 if run ea is a Hangul font, ea wins even when latin is latin-only -> passes."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "이에이 슬롯 모노", font="Arial", ea="IBM Plex Mono", size=12)
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "인터 폴백 한글", font="Inter", size=12)
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "원티드 산스 한글", font="Wanted Sans", size=12)
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "이에이 승리 한글", font="IBM Plex Mono", ea="Wanted Sans", size=12)
    errors, _w = lint_full(save(p, tmp_path, "fx.pptx"))
    e1_pages = [si for (si, m, d) in by_code(errors, "E1")]
    assert 1 in e1_pages and 2 in e1_pages
    assert 3 not in e1_pages and 4 not in e1_pages

def test_e1_theme_ea_korean_suppresses_latin(tmp_path):
    """If the theme ea is a Hangul font, the theme ea handles the render even when
    run latin is latin-only (measured): pins the regression for the E1 false
    positive that the old ea-or-latin substitute used to emit."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "테마가 받는 한글", font="IBM Plex Mono", size=12)
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "폰트 미지정 한글", size=12)
    path = patch_theme_ea(save(p, tmp_path, "fx.pptx"), "Malgun Gothic")
    errors, _w = lint_full(path)
    assert not by_code(errors, "E1"), errors

def test_e1_theme_ea_latin_only_flags(tmp_path):
    """If the theme ea itself is latin-only, font-unspecified Hangul is E1. Since
    0.2.1, the +mn-ea token in defaultTextStyle is resolved through the chain and
    caught as effective ea=Consolas (branch-independent, pinned by font name)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "폰트 미지정 한글", size=12)
    path = patch_theme_ea(save(p, tmp_path, "fx.pptx"), "Consolas")
    errors, _w = lint_full(path)
    assert any(c == "E1" and "Consolas" in d for (_si, c, m, d) in errors), errors

def test_theme_ea_by_master_relationship(tmp_path):
    """Whether theme resolution comes from the master->theme relationship rather
    than iter_parts order (key = master partname)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "관계 해석", font="Wanted Sans", size=12)
    prs = Presentation(save(p, tmp_path, "fx.pptx"))
    ea_map = jl.theme_ea_by_master(prs)
    assert len(ea_map) == 1
    key = next(iter(ea_map))
    assert "slideMaster" in key
    assert ea_map[key] == ""   # default template = empty slot

def test_e1_unit_matrix():
    """e1_violation decision matrix (a unit regression that doubles as documentation
    of the measured render model)."""
    v = jl.e1_violation
    assert v("한글", {"ea": "IBM Plex Mono"}, "") is not None          # run ea latin-only
    assert v("한글", {"ea": "맑은 고딕", "latin": "Consolas"}, "") is None   # run ea Hangul
    assert v("한글", {"latin": "IBM Plex Mono"}, "맑은 고딕") is None   # theme ea handles the render
    assert v("한글", {}, "Consolas") is not None                        # theme ea latin-only
    assert v("한글", {"latin": "Wanted Sans"}, "") is None              # empty theme + latin Hangul font
    assert v("한글", {"latin": "Inter"}, "") is not None                # empty theme + latin latin-only (previously a missed detection)
    assert v("한글", {}, "") is not None                                # nothing at all = Malgun confirmed

def test_latin_only_font_boundaries():
    """Blocklist prefix-match boundary: a Hangul-complete variant must not get
    swept up by a prefix match."""
    f = jl.is_latin_only_font
    assert f("Inter") and f("Arial") and f("Calibri") and f("IBM Plex Sans") and f("Noto Sans")
    assert not f("NanumGothicCoding")     # Hangul-complete monospace: pins the regression for an old misclassification
    assert not f("IBM Plex Sans KR")
    assert not f("Noto Sans KR") and not f("Noto Serif KR")
    assert not f("Arial Unicode MS")
    assert not f("Wanted Sans") and not f("Pretendard") and not f("Malgun Gothic")
    assert not f("") and not f(None)

def test_e2_numeric_context(tmp_path):
    """E2 context exception: a numeric-range en dash and a minus sign U+2212 pass in
    default mode; an em dash is always blocked."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "FY2020" + EN_DASH + "2024 실적", font="Wanted Sans", size=12)
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, MINUS + "3.2% 하락", font="Wanted Sans", size=12)
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "이건" + EM_DASH + "안 됨", font="Wanted Sans", size=12)
    path = save(p, tmp_path, "fx.pptx")
    errors, _w = lint_full(path)
    e2_pages = [si for (si, m, d) in by_code(errors, "E2")]
    assert e2_pages == [3], errors
    # strict: exceptions lifted, all three pages blocked
    errors_s, _w = lint_full(path, strict=True)
    assert sorted(si for (si, m, d) in by_code(errors_s, "E2")) == [1, 2, 3]

def test_dash_violations_unit():
    dv = jl.dash_violations
    assert dv("2020" + EN_DASH + "2024") == []
    assert dv("2020" + EN_DASH + "2024", strict=True) == [EN_DASH]
    assert dv(MINUS + "3.2%") == []
    assert dv("A" + EN_DASH + "B") == [EN_DASH]        # not a numeric context -> stays blocked
    assert dv("끝" + MINUS) == [MINUS]                  # no digit follows -> stays blocked
    assert dv("a" + EM_DASH + "b") == [EM_DASH]

def test_e4_universal_measure_spc(tmp_path):
    """spc's universal measure ('1.5pt') and a garbage value: without crashing, the
    former should be E4 and the latter should be ignored."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "자간 벌어진 한글", font="Wanted Sans", size=12, spc="1.5pt")
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "정상 자간 한글", font="Wanted Sans", size=12, spc="garbage")
    errors, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    e4_pages = [si for (si, m, d) in by_code(errors, "E4")]
    assert e4_pages == [1], errors

def test_e1_theme_token_resolution(tmp_path):
    """Resolve OOXML theme tokens (the +mn-lt kind) to actual fonts (adversarial
    panel confirmed: matching the token against the blocklist literally causes an
    E1 missed detection). The default template's theme minor latin=Calibri
    (latin-only), so a Hangul run with latin="+mn-lt" must be E1."""
    p = new_prs()
    s = add_slide(p)
    box = tb(s, 1, 1, 5, 0.5, "토큰 한글", size=12)
    rPr = box.text_frame.paragraphs[0].runs[0]._r.get_or_add_rPr()
    rPr.append(rPr.makeelement(qn("a:latin"), {"typeface": "+mn-lt"}))
    errors, _w = lint_full(save(p, tmp_path, "fx.pptx"))
    assert any("Calibri" in d for (_si, m, d) in by_code(errors, "E1")), errors

def test_resolve_font_tokens_unit():
    rft = jl.resolve_font_tokens
    thm = {"mn-lt": "Calibri", "mn-ea": "맑은 고딕", "mj-lt": "Georgia", "mj-ea": ""}
    assert rft({"latin": "+mn-lt"}, thm) == {"latin": "Calibri"}
    assert rft({"ea": "+mn-ea"}, thm) == {"ea": "맑은 고딕"}
    assert rft({"latin": "+mj-lt", "ea": "Wanted Sans"}, thm) == {"latin": "Georgia", "ea": "Wanted Sans"}
    assert rft({"ea": "+mj-ea"}, thm) == {}          # if resolved to an empty value, remove the slot (leave it to the chain)
    assert rft({"latin": "+mn-lt"}, None) == {}       # if theme parsing fails, remove the slot
    assert rft({"latin": "Batang"}, thm) == {"latin": "Batang"}   # leave as-is if not a token

def test_e2_run_boundary_split(tmp_path):
    """False-positive regression for a numeric range split across a run boundary
    (adversarial panel confirmed, high): even when '2020'/'-2021' are different
    runs, the exception applies via the paragraph context. An em dash stays
    blocked even when split."""
    p = new_prs()
    s = add_slide(p)   # p1: en dash range split at a run boundary
    box = s.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(0.5))
    para = box.text_frame.paragraphs[0]
    r1 = para.add_run(); r1.text = "2020"; r1.font.size = Pt(14)
    r2 = para.add_run(); r2.text = EN_DASH + "2024 실적"; r2.font.size = Pt(14)
    s = add_slide(p)   # p2: em dash split (always blocked)
    box = s.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(0.5))
    para = box.text_frame.paragraphs[0]
    r1 = para.add_run(); r1.text = "말"; r1.font.size = Pt(14)
    r2 = para.add_run(); r2.text = EM_DASH + "이음"; r2.font.size = Pt(14)
    path = save(p, tmp_path, "fx.pptx")
    errors, _w = lint_full(path)
    e2_pages = [si for (si, m, d) in by_code(errors, "E2")]
    assert e2_pages == [2], errors
    errors_s, _w = lint_full(path, strict=True)   # strict blocks everything regardless of split
    assert sorted(set(si for (si, m, d) in by_code(errors_s, "E2"))) == [1, 2]

def test_size_inheritance_placeholder_gates(tmp_path):
    """Placeholder inheritance chain activation regression: placeholder text with
    no explicit size doesn't leak into W5 but is resolved to the layout lstStyle
    size, so E3/W1 actually run (the old version disarmed everything into W5)."""
    p = Presentation()
    p.slide_width = Inches(13.333)
    p.slide_height = Inches(7.5)
    lay = p.slide_layouts[1]   # Title and Content
    for ph in lay.placeholders:
        sz = {0: "400", 1: "800"}.get(ph.placeholder_format.idx)   # title 4pt, body 8pt
        if sz is None:
            continue
        txBody = ph.text_frame._txBody
        lst = txBody.find(qn("a:lstStyle"))
        if lst is None:
            lst = txBody.makeelement(qn("a:lstStyle"), {})
            txBody.insert(1, lst)   # after bodyPr
        lvl1 = lst.makeelement(qn("a:lvl1pPr"), {})
        defR = lst.makeelement(qn("a:defRPr"), {"sz": sz})
        lvl1.append(defR)
        lst.append(lvl1)
    s = p.slides.add_slide(lay)
    s.shapes.title.text_frame.text = "판독 불가 제목"
    s.placeholders[1].text_frame.text = "본문이 레이아웃 상속 크기로 해석되는지 보는 사십자 이상의 충분히 긴 문장입니다"
    errors, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    assert not by_code(warns, "W5"), warns
    assert any("4.0pt" in m for (_si, m, _d) in by_code(errors, "E3")), errors       # title 4pt
    assert any("8.0pt" in m for (_si, m, _d) in by_code(warns, "W1")), warns          # body 8pt

def test_size_inheritance_master_txstyles(tmp_path):
    """Master txStyles path: default template resolves title=44pt, body(OBJECT)=32pt."""
    p = Presentation()
    s = p.slides.add_slide(p.slide_layouts[1])
    s.shapes.title.text_frame.text = "제목"
    s.placeholders[1].text_frame.text = "본문"
    path = save(p, tmp_path, "fx.pptx")
    prs = Presentation(path)
    sl = prs.slides[0]
    sizer = jl.SizeResolver(prs)
    got = {}
    for sp in sl.shapes:
        if sp.has_text_frame and sp.is_placeholder:
            got[sp.placeholder_format.idx] = sizer.resolve(sp.text_frame, sp, sl, 0)
    assert got[0] == 44.0 and got[1] == 32.0, got

def test_table_cell_e1(tmp_path):
    """Native table-cell path (a headline feature that had zero coverage): also
    catches font defects on CJK inside a cell."""
    p = new_prs()
    s = add_slide(p)
    gf = s.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(8), Inches(2))
    cell = gf.table.cell(0, 0)
    cell.text = "표 안 폴백 한글"
    r = cell.text_frame.paragraphs[0].runs[0]
    r.font.name = "Consolas"
    r.font.size = Pt(12)
    ok = gf.table.cell(1, 1)
    ok.text = "정상 셀 한글"
    r2 = ok.text_frame.paragraphs[0].runs[0]
    r2.font.name = "Wanted Sans"
    r2.font.size = Pt(12)
    errors, _w = lint_full(save(p, tmp_path, "fx.pptx"))
    e1 = by_code(errors, "E1")
    assert len(e1) == 1 and "Consolas" in e1[0][2], errors


# ---------------------------------------------------------------- layout-level gates

def test_autofit_parser_forms(tmp_path):
    p = new_prs()
    s = add_slide(p)
    b = tb(s, 1, 1, 4, 1, "autofit test", font="Wanted Sans", size=12)
    bodyPr = b.text_frame._txBody.find(qn("a:bodyPr"))
    na = bodyPr.makeelement(qn("a:normAutofit"),
                            {"fontScale": "62.5%", "lnSpcReduction": "20%"})
    bodyPr.append(na)
    fs, lr = jl.frame_autofit(b.text_frame)
    assert abs(fs - 0.625) < 1e-6 and abs(lr - 0.20) < 1e-6
    na.set("fontScale", "62500")
    na.set("lnSpcReduction", "20000")
    fs, lr = jl.frame_autofit(b.text_frame)
    assert abs(fs - 0.625) < 1e-6 and abs(lr - 0.20) < 1e-6


# ---------------------------------------------------------------- negative + CLI

def test_title_collection_keeps_placeholder_sizes(tmp_path):
    """Placeholder inherited size (master titleStyle 44pt) is an intentional title,
    so it must still be collected."""
    p = Presentation()
    s = p.slides.add_slide(p.slide_layouts[1])
    s.shapes.title.text_frame.text = "진짜 제목"
    s.placeholders[1].text_frame.text = "본문"
    ghost = []
    lint_full(save(p, tmp_path, "fx.pptx"), ghost=ghost)
    assert any("진짜 제목" in t for _si, t in ghost), ghost

def test_script_layer_two_tier(tmp_path):
    """The script layer's two-tier structure (third panel: resolves both the Hanja
    regression and the JP font false positive at once).
    p1 kana+Noto Sans JP -> silent (font has kana). p2 Hangul+Noto Sans JP ->
    E1 (no Hangul). p3 kana+positive tracking -> no E4 (Japanese convention).
    p4 Hanja-only+Georgia -> E1 (font has no CJK at all). p5 Hanja-only+positive
    tracking -> no E4 since 0.6.1 (tracked hanzi is legitimate Chinese typography;
    the run must contain Hangul). p6 kana+IBM Plex Mono -> E1. p7 mixed
    Hangul+Hanja+tracking -> E4 (Korean names/legal terms keep coverage)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "こんにちは", font="Noto Sans JP", size=12)
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "한글 텍스트", font="Noto Sans JP", size=12)
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "あいうえお", font="Noto Sans JP", size=12, spc=100)
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "漢字専用", font="Georgia", size=12)
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "大韓民國", font="맑은 고딕", size=12, spc=100)
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "フリーレン", font="IBM Plex Mono", size=12)
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "金九 선생", font="맑은 고딕", size=12, spc=100)
    errors, _w = lint_full(save(p, tmp_path, "fx.pptx"))
    e1_pages = sorted(si for (si, m, d) in by_code(errors, "E1"))
    assert e1_pages == [2, 4, 6], errors
    e4_pages = sorted(si for (si, m, d) in by_code(errors, "E4"))
    assert e4_pages == [7], errors

def test_hangul_range_extensions():
    """Closes the missed detection for halfwidth Hangul and the Jamo Extension
    blocks (third panel)."""
    assert jl.is_hangul(chr(0xFFA1))    # halfwidth giyeok
    assert jl.is_hangul(chr(0xA960))    # Jamo Extended-A
    assert jl.is_hangul(chr(0xD7B0))    # Jamo Extended-B
    assert jl._geometry_unsupported(chr(0x0F40))   # Tibetan
    assert jl._geometry_unsupported(chr(0x1000))   # Myanmar
    assert jl._geometry_unsupported(chr(0x1780))   # Khmer

def test_latin_only_font_additions():
    f = jl.is_latin_only_font
    assert f("Aptos") and f("Courier") and f("PT Sans") and f("PT Serif")
    assert not f("Noto Sans Mono CJK KR")   # contains cjk = Hangul-complete (second re-review fixed a false positive)
    assert not f("Source Han Sans CJK KR")

def test_e2_v2_adversarial_edges():
    """Pins the regression for 4 cases found by the E2 v2 adversarial panel
    (2026-07-10, third round)."""
    dv = jl.dash_violations
    MINUS = chr(0x2212)
    assert dv("전일대비 " + MINUS + " 3.2%") == []            # spaced minus sign (financial notation)
    assert dv("끝 " + MINUS) == [MINUS]                        # still blocked when no digit follows
    assert dv("2020" + EN_DASH + EN_DASH + "2024") == [EN_DASH, EN_DASH]   # closes the missed detection for 2 consecutive dashes
    assert dv("결론2024" + EN_DASH + "우리는") == [EN_DASH]    # closes the bypass via a word+digit mixed token
    assert dv("매출" + chr(0x00B9) + EN_DASH + "영업이익 증가") == [EN_DASH]   # superscript footnote
    assert dv("2020" + EN_DASH + "현재") == []                 # keeps the digit-starting-token range rule
    assert dv("Q1" + EN_DASH + "Q3") == []                     # keeps the both-sides-numeric rule

def test_e2_v2_range_forms(tmp_path):
    """E2 v2: passes the remaining 4 false-positive forms + keeps blocking a
    spaced one-sided-digit case (an AI-inserted-clause pattern)."""
    dv = jl.dash_violations
    assert dv("2020 " + EN_DASH + " 2024") == []          # spaced numeric range
    assert dv("Q1" + EN_DASH + "Q3") == []                 # alphanumeric token
    assert dv("5%" + EN_DASH + "10%") == []                # percent range
    assert dv("2020" + EN_DASH + "현재") == []             # attached one-sided digit
    assert dv("성장 " + EN_DASH + " 2024년에는") == [EN_DASH]   # a spaced inserted clause stays blocked
    assert dv("서울" + EN_DASH + "부산") == [EN_DASH]       # word-to-word connection stays conservatively blocked
    assert dv("2020" + EN_DASH + "현재", strict=True) == [EN_DASH]
    # integration: also pinned via a deck fixture
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 6, 0.5, "매출 5%" + EN_DASH + "10% 구간", font="Wanted Sans", size=14)
    s = add_slide(p)
    tb(s, 1, 1, 6, 0.5, "성장 " + EN_DASH + " 2024년에는 확대", font="Wanted Sans", size=14)
    errors, _w = lint_full(save(p, tmp_path, "fx.pptx"))
    assert [si for (si, m, d) in by_code(errors, "E2")] == [2], errors

def test_e1_master_lststyle_font_inheritance(tmp_path):
    """Reflects Probe 6's measured finding: the master ph lstStyle's a:ea is the
    effective render font. Latin-only means E1, a Hangul font means it passes
    (fixes the "Malgun fallback confirmed" false positive)."""
    def build(ea_font):
        p = Presentation()
        master = p.slide_masters[0]
        for ph in master.placeholders:
            if ph.placeholder_format.idx == 1:
                txBody = ph.text_frame._txBody
                lst = txBody.find(qn("a:lstStyle"))
                if lst is None:
                    lst = txBody.makeelement(qn("a:lstStyle"), {})
                    txBody.insert(1, lst)
                lvl1 = lst.makeelement(qn("a:lvl1pPr"), {})
                defR = lst.makeelement(qn("a:defRPr"), {})
                defR.append(defR.makeelement(qn("a:ea"), {"typeface": ea_font}))
                lvl1.append(defR)
                lst.append(lvl1)
        s = p.slides.add_slide(p.slide_layouts[1])
        s.shapes.title.text_frame.text = ""
        body = s.placeholders[1]
        body.text_frame.text = "마스터 상속 한글"
        return save(p, tmp_path, "fx_%s.pptx" % ("bad" if "Mono" in ea_font else "ok"))

    errors, _w = lint_full(build("IBM Plex Mono"))
    assert any("IBM Plex Mono" in d for (_si, m, d) in by_code(errors, "E1")), errors
    errors, _w = lint_full(build("맑은 고딕"))
    assert not by_code(errors, "E1"), errors

def test_e1_title_uses_major_font(tmp_path):
    """Reflects Probe 6 Q1/Q2's measured finding: the title ph rides the theme's
    majorFont ea, and the body rides the minorFont ea."""
    p = Presentation()
    s = p.slides.add_slide(p.slide_layouts[1])
    s.shapes.title.text_frame.text = "제목 한글"
    s.placeholders[1].text_frame.text = "본문 한글"
    path = save(p, tmp_path, "fx.pptx")
    patch_theme_fonts(path, major_ea="Consolas", minor_ea="맑은 고딕")
    errors, _w = lint_full(path)
    e1 = by_code(errors, "E1")
    assert any("제목" in d for (_si, m, d) in e1), errors     # major=Consolas -> E1
    assert not any("본문" in d for (_si, m, d) in e1), errors  # minor=Malgun Gothic -> passes

def test_vertical_and_complex_script_geometry_skip(tmp_path):
    """Vertical writing (bodyPr@vert) and complex-typesetting scripts can't be
    geometrically estimated: skip them but surface it via W18."""
    p = new_prs()
    s = add_slide(p)   # p1: a vertical-writing frame goes off-screen -> must have no W16 FP
    box = tb(s, 12.0, 1.0, 4.0, 0.6, "vertical long text overflowing edge",
             font="Wanted Sans", size=18)
    bodyPr = box.text_frame._txBody.find(qn("a:bodyPr"))
    bodyPr.set("vert", "eaVert")
    s = add_slide(p)   # p2: Arabic text -> geometry skip + W18
    tb(s, 11.0, 1.0, 4.0, 0.6, "مرحبا بالعالم",
       font="Arial", size=18)
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    w16_text = [d for (_si, m, d) in by_code(warns, "W16") if "텍스트" in d]
    assert not w16_text, warns
    w18 = by_code(warns, "W18")
    assert any(si == 1 and "vertical_text" in d for (si, m, d) in w18), warns
    assert any(si == 2 and "complex_script" in d for (si, m, d) in w18), warns


# ---------------------------------------------------------------- 0.3.0 i18n + profiles

def test_e1_para_defrpr_inheritance(tmp_path):
    """P0-1 (Probe 7 measured): the paragraph pPr/defRPr font is inherited at the
    rank right after run, and wins over lstStyle.
    Case A: paragraph ea=Hangul font -> there must be no E1 false positive.
    Case B: master lstStyle ea=Hangul font but paragraph ea=Consolas -> the
    paragraph wins, so there must be no E1 missed detection."""
    def add_para_ea(para, ea):
        pPr = para._p.get_or_add_pPr()
        defR = pPr.makeelement(qn("a:defRPr"), {})
        defR.append(defR.makeelement(qn("a:ea"), {"typeface": ea}))
        pPr.append(defR)

    # Case A: verify no false positive
    p = new_prs()
    s = add_slide(p)
    box = s.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(0.6))
    para = box.text_frame.paragraphs[0]
    add_para_ea(para, "맑은 고딕")
    r = para.add_run(); r.text = "문단 상속 한글"; r.font.size = Pt(14)
    errors, _w = lint_full(save(p, tmp_path, "a.pptx"))
    assert not by_code(errors, "E1"), errors

    # Case B: verify no missed detection (master lstStyle Hangul font + paragraph Consolas)
    p = Presentation()
    for ph in p.slide_masters[0].placeholders:
        if ph.placeholder_format.idx == 1:
            txBody = ph.text_frame._txBody
            lst = txBody.find(qn("a:lstStyle"))
            if lst is None:
                lst = txBody.makeelement(qn("a:lstStyle"), {})
                txBody.insert(1, lst)
            lvl1 = lst.makeelement(qn("a:lvl1pPr"), {})
            defR = lst.makeelement(qn("a:defRPr"), {})
            defR.append(defR.makeelement(qn("a:ea"), {"typeface": "맑은 고딕"}))
            lvl1.append(defR)
            lst.append(lvl1)
    s = p.slides.add_slide(p.slide_layouts[1])
    s.shapes.title.text_frame.text = ""
    body = s.placeholders[1]
    para = body.text_frame.paragraphs[0]
    add_para_ea(para, "Consolas")
    r = para.add_run(); r.text = "문단이 이기는 한글"; r.font.size = Pt(14)
    errors, _w = lint_full(save(p, tmp_path, "b.pptx"))
    assert any("Consolas" in d for (_si, m, d) in by_code(errors, "E1")), errors

def test_e4_hanja_message(tmp_path):
    """E4 fires on mixed Hangul+Hanja runs (Korean names/legal terms) and the message
    names the judged scope; a Hanja-only run is exempt since 0.6.1 (tracked hanzi is
    legitimate Chinese typography, external review)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "大韓民國 헌법", font="맑은 고딕", size=12, spc=100)
    errors, _w = lint_full(save(p, tmp_path, "fx.pptx"))
    e4 = by_code(errors, "E4")
    assert e4 and "한자" in e4[0][1], e4
    p2 = new_prs()
    s2 = add_slide(p2)
    tb(s2, 1, 1, 5, 0.5, "中华人民共和国", font="맑은 고딕", size=12, spc=200)
    errors2, _w2 = lint_full(save(p2, tmp_path, "fx2.pptx"))
    assert not by_code(errors2, "E4"), errors2

def test_w8_extended_hangul(tmp_path):
    """P1: with is_cjk unified, halfwidth Hangul is also caught by the W8
    small-CJK judgment."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 2, 0.4, chr(0xFFA1) * 6, font="Wanted Sans", size=6)   # halfwidth giyeok x6
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    assert by_code(warns, "W8"), warns

def test_fld_field_runs_gated(tmp_path):
    """a:fld also renders with the same rPr as a normal run, so it's subject to
    the same gates (carried over from the fourth review): an Arial Hangul field
    is E1, a positive-tracking field is E4. loc carries a field marker, with no
    run index."""
    p = new_prs()
    s = add_slide(p)
    box = s.shapes.add_textbox(Inches(1), Inches(6.8), Inches(3), Inches(0.4))
    _add_fld(box.text_frame.paragraphs[0], "3 페이지", sz=1200, latin="Arial", spc=200)
    errors, _w = lint_full(save(p, tmp_path, "fld.pptx"))
    e1 = _by(errors, "E1")
    assert e1, codes(errors)
    assert e1[0].loc.get("field") is True and "run" not in e1[0].loc
    assert _by(errors, "E4"), codes(errors)

def test_field_only_complex_script_marks_incomplete(tmp_path):
    """Arabic living entirely inside an a:fld used to bypass the complex-script screen
    (built from para.runs) while still entering width math: now the screen sees field
    text and the frame is skipped into W18 (external review P0)."""
    p = new_prs()
    s = add_slide(p)
    box = s.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(0.6))
    _add_fld(box.text_frame.paragraphs[0], "مرحبا بك",
             sz=1800)
    _e, warns = lint_full(save(p, tmp_path, "rtl_fld.pptx"))
    assert "W18" in codes(warns), codes(warns)

def test_malformed_autofit_marks_incomplete(tmp_path):
    """0.8.x exception audit: a garbled normAutofit fontScale must not silently read as
    scale 1.0 (hiding a real shrink from E3); the span aborts into W18 instead."""
    p = new_prs()
    s = add_slide(p)
    box = tb(s, 1, 1, 5, 1, "autofit garbage frame", size=20, ea="맑은 고딕")
    bodyPr = box.text_frame._txBody.find(qn("a:bodyPr"))
    na = bodyPr.makeelement(qn("a:normAutofit"), {"fontScale": "garbage%"})
    bodyPr.append(na)
    deck = save(p, tmp_path, "af.pptx")
    doc = json.loads(run_cli([deck, "--json"]).stdout)
    assert doc["summary"]["incomplete"] is True, doc["summary"]
    assert any(w["code"] == "W18" for w in doc["warnings"])
