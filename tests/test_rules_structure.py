# -*- coding: utf-8 -*-
"""Deck-structure and style rules (W6-W14, render) (split from test_lint.py, 0.8.x)."""
from helpers import *   # noqa: F401,F403
from helpers import (_by, _add_fld, _repo_root, _require_repo_file,
                     patch_theme_fonts)   # noqa: F401


def test_w6_page_numbers_survive_token_failure(tmp_path, monkeypatch):
    """sigs double-append regression (adversarial panel confirmed): even if
    _fill_tokens dies, the sigs positions don't shift, so W6 page numbers point to
    the actual slide number."""
    def boom(slide, sw, sh):
        raise RuntimeError("token fail")
    monkeypatch.setattr(jl, "_fill_tokens", boom)
    p = new_prs()
    for i in range(6):
        s = add_slide(p)
        tb(s, 1, 0.8, 8, 0.6, "Title block %d" % i, font="Wanted Sans", size=24)
        tb(s, 1, 2.0, 6, 2.5, "Body block", font="Wanted Sans", size=12)
        tb(s, 8, 2.0, 4, 2.5, "Side block", font="Wanted Sans", size=12)
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    w6 = by_code(warns, "W6")
    assert w6, warns
    pages = [int(m) for m in re.findall(r"p(\d+)", w6[0][2])]
    assert pages and max(pages) <= 6, w6   # in the days of double-append, p7-p11 used to show up
    # 0.2.1: token-collection failure is also surfaced as W18 (guard wiring extended everywhere)
    assert any("w10_tokens" in d for (_si, m, d) in by_code(warns, "W18")), warns

def test_w6_layout_clones(tmp_path):
    p = new_prs()
    for i in range(6):
        s = add_slide(p)
        tb(s, 1, 0.8, 8, 0.6, "Title block %d" % i, font="Wanted Sans", size=24)
        tb(s, 1, 2.0, 6, 2.5, "Body block", font="Wanted Sans", size=12)
        tb(s, 8, 2.0, 4, 2.5, "Side block", font="Wanted Sans", size=12)
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    assert "W6" in codes(warns)

def test_w6_tunables_suppress(tmp_path):
    """W6 tunable: raising the cluster threshold makes the same deck pass (an
    intentional relief knob for template houses)."""
    p = new_prs()
    for i in range(6):
        s = add_slide(p)
        tb(s, 1, 0.8, 8, 0.6, "Title block %d" % i, font="Wanted Sans", size=24)
        tb(s, 1, 2.0, 6, 2.5, "Body block", font="Wanted Sans", size=12)
        tb(s, 8, 2.0, 4, 2.5, "Side block", font="Wanted Sans", size=12)
    path = save(p, tmp_path, "fx.pptx")
    _e, warns = lint_full(path, w6_min_cluster=10)
    assert "W6" not in codes(warns)
    # an exact clone has cosine exactly 1.00, so the sim threshold must be set to 1.0 to exclude it (confirms the parameter passes through)
    _e, warns = lint_full(path, w6_sim=1.0)
    assert "W6" not in codes(warns)

def test_w7_low_contrast_needs_render(tmp_path):
    from PIL import Image
    img = os.path.join(str(tmp_path), "bg.jpg")
    Image.new("RGB", (1536, 864), (240, 240, 240)).save(img)
    p = new_prs()
    s = add_slide(p)
    s.shapes.add_picture(img, 0, 0, Inches(13.333), Inches(7.5))
    tb(s, 2, 3, 9, 1, "white on white low contrast", font="Wanted Sans", size=40, color="F5F5F5")
    path = save(p, tmp_path, "fx.pptx")
    pages = os.path.join(str(tmp_path), "pages")
    os.makedirs(pages)
    Image.new("RGB", (1600, 900), (240, 240, 240)).save(os.path.join(pages, "p01.png"))
    _e, warns = lint_full(path, render_dir=pages)
    assert "W7" in codes(warns)

def test_render_on_textonly_deck_no_crash(tmp_path):
    """Linting a text-only deck with not a single image via --render must not crash.
    Regression: back when glob existed only as a local import inside
    contrast_check, glob.glob in the render_png_hits==0 branch died with a
    NameError, which got mislabeled as "could not open pptx" (2026-07-04)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 0.8, 9, 0.8, "제목만 있는 페이지", font="Wanted Sans", size=28)
    tb(s, 1, 2.2, 10, 3, "본문 텍스트, 이미지는 없다.", font="Wanted Sans", size=13)
    pages = os.path.join(str(tmp_path), "pages")
    os.makedirs(pages)
    from PIL import Image
    Image.new("RGB", (1600, 900), (240, 240, 240)).save(os.path.join(pages, "slide-1.png"))  # a name that breaks the convention
    errors, warns = lint_full(save(p, tmp_path, "fx.pptx"), render_dir=pages)  # fails if NameError occurs
    assert isinstance(errors, list) and isinstance(warns, list)

def test_w9_accent_vbars(tmp_path):
    p = new_prs()
    s = add_slide(p)
    for i in range(4):
        y = 1.5 + i * 1.1
        vconn(s, 5.0, y, y + 0.65, "E6A94E")
        tb(s, 5.2, y, 5.5, 0.6, "callout item %d body" % i, font="Wanted Sans", size=12)
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    assert "W9" in codes(warns)

def test_w10_diagram_clone_marks(tmp_path):
    p = new_prs()
    for _ in range(2):
        s = add_slide(p)
        rect(s, 2.0, 2.0, 8.0, 3.0, "2E3540")
        for i in range(9):
            rect(s, 2.5 + i * 0.8, 3.0, 0.18, 0.18, "5B6470")
        tb(s, 1, 0.8, 8, 0.6, "diagram page", font="Wanted Sans", size=20)
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    assert "W10" in codes(warns)

def test_w11_ai_copy(tmp_path):
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 10, 0.8, "오늘날 급변하는 시장 환경에서", font="Wanted Sans", size=24)
    s = add_slide(p)
    tb(s, 1, 1, 10, 0.8, "두 사업의 시너지 효과를 극대화", font="Wanted Sans", size=14)
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    w11 = by_code(warns, "W11")
    assert any("오프닝" in m for (_si, m, _d) in w11)
    assert any("버즈워드" in m for (_si, m, _d) in w11)

def test_w12_footer_baseline(tmp_path):
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 3, 8, 1, "Cover page", font="Wanted Sans", size=30)
    for i in range(6):
        s = add_slide(p)
        tb(s, 1, 1, 8, 0.6, "Content %d" % i, font="Wanted Sans", size=20)
        fy = 7.05 if i == 3 else 6.9
        tb(s, 1, fy, 5, 0.3, "footer text", font="Wanted Sans", size=8)
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    assert "W12" in codes(warns)

def test_w13_native_effects(tmp_path):
    p = new_prs()
    s = add_slide(p)
    for i in range(2):
        sp = rect(s, 1 + i * 3, 2, 2, 1.2, "2E3540")
        spPr = sp._element.spPr
        eff = spPr.makeelement(qn("a:effectLst"), {})
        sh = spPr.makeelement(qn("a:outerShdw"), {"blurRad": "40000", "dist": "20000"})
        eff.append(sh)
        spPr.append(eff)
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    assert "W13" in codes(warns)

def test_w14_nominal_titles_and_ghost(tmp_path):
    p = new_prs()
    for t in ("시장 현황", "경쟁 구도 분석", "제품 라인업 개요", "사업 확장 전략",
              "재무 운용 계획", "향후 추진 방안"):
        s = add_slide(p)
        tb(s, 1, 0.8, 9, 0.8, t, font="Wanted Sans", size=26)
        tb(s, 1, 2.2, 10, 3, "본문 내용", font="Wanted Sans", size=12)
    ghost = []
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"), ghost=ghost)
    assert "W14" in codes(warns)
    assert len(ghost) == 6

def test_w14_numeric_claim_titles_pass(tmp_path):
    """A number+unit title is a claim even when it ends in a noun: W14 must not
    false-fire (mitigates genre bias)."""
    p = new_prs()
    for t in ("매출 3배 성장", "점유율 42% 달성", "비용 120억 절감", "가입자 500만 돌파",
              "마진 8pp 개선", "리텐션 2x 향상"):
        s = add_slide(p)
        tb(s, 1, 0.8, 9, 0.8, t, font="Wanted Sans", size=26)
        tb(s, 1, 2.2, 10, 3, "본문 내용", font="Wanted Sans", size=12)
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    assert "W14" not in codes(warns), by_code(warns, "W14")


# ---------------------------------------------------------------- geometry gates

def test_cli_text_output_and_ghost(tmp_path):
    """Human-readable text output + the --ghost section (untested in the old version)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 0.8, 9, 0.8, "고스트 타이틀 페이지", font="Wanted Sans", size=26)
    tb(s, 1, 2, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    path = save(p, tmp_path, "fx.pptx")
    r = run_cli([path, "--ghost"])
    assert r.returncode == 1
    assert "=== ARCHFORGE LINT" in r.stdout
    assert "ghost deck" in r.stdout and "고스트 타이틀" in r.stdout
    assert "ERROR p01 [E1]" in r.stdout
    assert re.search(r"--- ERROR \d+, WARN \d+ ---", r.stdout)

def test_title_flood_regression(tmp_path):
    """Pins the regression where the defaultTextStyle fallback 18pt got collected as
    a title candidate, flooding ghost/W14 (second external re-review confirmed).
    A deck with only size-unspecified body text must have an empty ghost."""
    p = new_prs()
    s = add_slide(p)
    for i, txt in enumerate(("본문 첫 문장입니다", "본문 둘째 문장입니다", "표 안 텍스트 셋째")):
        tb(s, 1, 1 + i, 8, 0.5, txt, font="Wanted Sans", no_size=True)
    ghost = []
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"), ghost=ghost)
    assert ghost == [], ghost
    assert "W14" not in codes(warns)

def test_title_own_lststyle_not_flooded(tmp_path):
    """Body prose in a shape's own lstStyle 20pt is not swept into ghost (closes a
    case reproduced measured by the third panel). Title eligibility = an explicit
    size or a title-family placeholder."""
    p = new_prs()
    s = add_slide(p)
    box = tb(s, 1, 1, 9, 0.6, "이것은 본문 카피 문장입니다 절대 제목이 아닙니다",
             font="Wanted Sans", no_size=True)
    txBody = box.text_frame._txBody
    lst = txBody.makeelement(qn("a:lstStyle"), {})
    lvl1 = lst.makeelement(qn("a:lvl1pPr"), {})
    defR = lst.makeelement(qn("a:defRPr"), {"sz": "2000"})
    lvl1.append(defR)
    lst.append(lvl1)
    txBody.insert(1, lst)
    ghost = []
    lint_full(save(p, tmp_path, "fx.pptx"), ghost=ghost)
    assert ghost == [], ghost

def test_profile_editorial_drops_w14(tmp_path):
    p = new_prs()
    for t in ("시장 현황", "경쟁 구도 분석", "제품 라인업 개요", "사업 확장 전략",
              "재무 운용 계획", "향후 추진 방안"):
        s = add_slide(p)
        tb(s, 1, 0.8, 9, 0.8, t, font="Wanted Sans", size=26)
        tb(s, 1, 2.2, 10, 3, "본문 내용", font="Wanted Sans", size=12)
    path = save(p, tmp_path, "fx.pptx")
    doc = json.loads(run_cli([path, "--json", "--profile", "editorial"]).stdout)
    assert not any(w["code"] == "W14" for w in doc["warnings"])
    assert doc["summary"]["profile"] == "editorial"


# ---------------------------------------------------------------- 0.3.1 (third external review P0/P1)

def test_render_dir_contract(tmp_path):
    """P0-4: a missing --render folder or missing page renders are surfaced as
    incomplete; a deck with no pictures is unaffected."""
    p = new_prs()
    s = add_slide(p)
    s.shapes.add_picture(png(tmp_path, "pic.png"), Inches(2), Inches(2), Inches(3), Inches(2))
    tb(s, 1, 1, 5, 0.5, "clean text", font="Wanted Sans", size=20)
    deck = save(p, tmp_path, "fx.pptx")
    doc = json.loads(run_cli([deck, "--json", "--render",
                              os.path.join(str(tmp_path), "no_such_dir")]).stdout)
    assert doc["summary"]["incomplete"] is True, doc["summary"]
    empty = os.path.join(str(tmp_path), "empty_pages")
    os.makedirs(empty)
    doc = json.loads(run_cli([deck, "--json", "--render", empty]).stdout)
    assert doc["summary"]["incomplete"] is True, doc["summary"]   # missing render for a page that has a picture
    # a deck with no pictures is complete even with no render (there's no W7 target to check at all)
    p2 = new_prs()
    s2 = add_slide(p2)
    tb(s2, 1, 1, 5, 0.5, "text only", font="Wanted Sans", size=20)
    doc = json.loads(run_cli([save(p2, tmp_path, "t.pptx"), "--json", "--render", empty]).stdout)
    assert doc["summary"]["incomplete"] is False, doc["summary"]

def test_ghost_prefers_title_placeholder(tmp_path):
    """When a title placeholder exists, a bigger KPI big-number must not push the
    ghost title out (third review)."""
    p = Presentation()
    s = p.slides.add_slide(p.slide_layouts[1])
    s.shapes.title.text_frame.text = "시장 규모는 빠르게 확대된다"
    for r in s.shapes.title.text_frame.paragraphs[0].runs:
        r.font.size = Pt(26)
    s.placeholders[1].text_frame.text = ""
    box = s.shapes.add_textbox(Inches(4), Inches(3), Inches(5), Inches(1.5))
    r = box.text_frame.paragraphs[0].add_run()
    r.text = "300억"
    r.font.size = Pt(60)
    ghost = []
    lint_full(save(p, tmp_path, "fx.pptx"), ghost=ghost)
    assert ghost and "시장 규모" in ghost[0][1], ghost

def test_w6_detail_english(tmp_path):
    """P1: i18n extends even to detail (in English mode, W6's detail starts with "e.g.")."""
    p = new_prs()
    for i in range(6):
        s = add_slide(p)
        tb(s, 1, 0.8, 8, 0.6, "Title block %d" % i, font="Wanted Sans", size=24)
        tb(s, 1, 2.0, 6, 2.5, "Body block", font="Wanted Sans", size=12)
        tb(s, 8, 2.0, 4, 2.5, "Side block", font="Wanted Sans", size=12)
    doc = json.loads(run_cli([save(p, tmp_path, "fx.pptx"), "--json", "--profile", "full"], lang="en").stdout)
    w6 = [w for w in doc["warnings"] if w["code"] == "W6"]
    assert w6 and w6[0]["detail"].startswith("e.g."), w6


# ---------------------------------------------------------------- 0.4.0 structural overhaul

def test_br_offsets_and_fld_not_title(tmp_path):
    """a:br contributes only a single line-break character in the E2 context
    (offset preserved: an em dash in the run after br is still caught). A large
    auto field (a 40pt page number) is not collected as a ghost title."""
    p = new_prs()
    s = add_slide(p)
    box = tb(s, 1, 1, 8, 1, "첫 줄", size=14)
    para = box.text_frame.paragraphs[0]
    para._p.append(para._p.makeelement(qn("a:br"), {}))
    r2 = para.add_run()
    r2.text = "결론" + EM_DASH + "정리"
    r2.font.size = Pt(14)
    box2 = s.shapes.add_textbox(Inches(11), Inches(6.5), Inches(2), Inches(0.8))
    _add_fld(box2.text_frame.paragraphs[0], "01", sz=4000)
    ghost = []
    errors, _w = lint_full(save(p, tmp_path, "br.pptx"), ghost=ghost)
    assert "E2" in codes(errors), codes(errors)
    assert not any("01" in str(g) for g in ghost), ghost

def test_w7_unknown_explicit_color_abstains(tmp_path):
    """An explicit color the decoder cannot resolve (hslClr) must stop resolution and
    abstain, not fall through to an inherited color (external review P1: a white hslClr
    run was judged with inherited black = false positive)."""
    from PIL import Image
    p = new_prs()
    s = add_slide(p)
    d = os.path.join(str(tmp_path), "pages")
    os.makedirs(d)
    Image.new("RGB", (1333, 750), (10, 10, 10)).save(os.path.join(d, "p01.png"))
    s.shapes.add_picture(png(tmp_path, "bg.png", size=(800, 500)),
                         Inches(1), Inches(1), Inches(8), Inches(4))
    box = tb(s, 2, 2, 5, 1, "white text via hslClr", size=20)
    r = box.text_frame.paragraphs[0].runs[0]
    rPr = r._r.get_or_add_rPr()
    fill = rPr.makeelement(qn("a:solidFill"), {})
    hsl = rPr.makeelement(qn("a:hslClr"), {"hue": "0", "sat": "0%", "lum": "100%"})
    fill.append(hsl)
    rPr.insert(0, fill)
    _e, warns = lint_full(save(p, tmp_path, "hsl.pptx"), render_dir=d)
    assert "W7" not in codes(warns), codes(warns)
    assert "W18" in codes(warns), codes(warns)
