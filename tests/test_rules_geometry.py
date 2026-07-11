# -*- coding: utf-8 -*-
"""W15-W18 geometry, overlap, and abstention rules (split from test_lint.py, 0.8.x)."""
from helpers import *   # noqa: F401,F403
from helpers import (_by, _add_fld, _repo_root, _require_repo_file,
                     patch_theme_fonts)   # noqa: F401


def test_w18_geometry_decoupling(tmp_path):
    """A tboxes calculation failure (a corrupted anchor) must not also silence an
    unrelated picture's W16 (pins the regression for 2 cases reproduced measured
    by the adversarial panel). W16 survives and W18 reports the incompleteness."""
    p = new_prs()
    s = add_slide(p)
    box = tb(s, 1, 1, 4, 0.5, "손상 앵커 텍스트", font="Wanted Sans", size=14)
    bodyPr = box.text_frame._txBody.find(qn("a:bodyPr"))
    bodyPr.set("anchor", "bogusvalue")   # a value with no MSO_VERTICAL_ANCHOR mapping -> tboxes exception
    s.shapes.add_picture(png(tmp_path, "over.png"), Inches(12.8), Inches(2.0),
                         Inches(2.0), Inches(1.5))   # overflows 1.47in on the right
    errors, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    assert any("그림" in d for (_si, m, d) in by_code(warns, "W16")), warns
    assert by_code(warns, "W18"), warns

def test_w15_overlap_and_intentional_layers(tmp_path):
    p = new_prs()
    s = add_slide(p)  # p1: genuine overlap (label over a value)
    tb(s, 1, 2.0, 4, 1.2, "12,400원", font="Wanted Sans", size=44)
    tb(s, 1.2, 2.15, 3, 0.5, "+11.7% 전년 대비", font="Wanted Sans", size=14)
    s = add_slide(p)  # p2: drop cap = intentional
    tb(s, 1.0, 1.0, 1.2, 1.1, "피", font="Wanted Sans", size=60)
    tb(s, 1.0, 1.05, 5, 0.5, "지컬 AI", font="Wanted Sans", size=13)
    s = add_slide(p)  # p3: identical text echo = intentional
    tb(s, 2, 2, 4, 1, "떠난 뒤에야", font="Wanted Sans", size=30)
    tb(s, 2.05, 2.05, 4, 1, "떠난 뒤에야", font="Wanted Sans", size=30)
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    w15 = by_code(warns, "W15")
    assert any(si == 1 for (si, m, d) in w15)
    assert not any(si == 2 for (si, m, d) in w15)
    assert not any(si == 3 for (si, m, d) in w15)

def test_w16_overflow_and_negatives(tmp_path):
    p = new_prs()
    s = add_slide(p)  # p1: text overflows right (wrap=none, single line)
    tb(s, 11.0, 1.0, 4.0, 0.6, "Quarterly revenue grew twenty two percent",
       font="Wanted Sans", size=18)
    s = add_slide(p)  # p2: long Hangul in a word_wrap=True frame overflows the bottom
    bx = tb(s, 1.0, 7.0, 3.0, 0.4,
            "긴 본문이 프레임 계산보다 길어져 바닥을 뚫고 내려가는 경우를 재현한다",
            font="Wanted Sans", size=16)
    bx.text_frame.word_wrap = True
    s = add_slide(p)  # p3: picture clipped on the right
    s.shapes.add_picture(png(tmp_path, "op.png"), Inches(12.5), Inches(2.0), Inches(2.0), Inches(1.5))
    s = add_slide(p)  # p4 (negative): full bleed + a chart boundary straddling transparent margin
    s.shapes.add_picture(png(tmp_path, "bleed.png"), Inches(0), Inches(0), Inches(13.4), Inches(7.55))
    s.shapes.add_picture(png(tmp_path, "chart.png", opaque_box=(0.1, 0.1, 0.5, 0.5)),
                         Inches(9.0), Inches(4.0), Inches(5.0), Inches(3.6))
    tb(s, 1, 1, 6, 0.6, "음성 페이지 본문", font="Wanted Sans", size=14)
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    w16 = by_code(warns, "W16")
    assert any(si == 1 for (si, m, d) in w16)
    assert any(si == 2 for (si, m, d) in w16)
    assert any(si == 3 for (si, m, d) in w16)
    assert not any(si == 4 for (si, m, d) in w16)

def test_w17_straddle_and_negatives(tmp_path):
    p = new_prs()
    s = add_slide(p)  # p1: straddling
    s.shapes.add_picture(png(tmp_path, "p1.png"), Inches(4.0), Inches(2.0), Inches(4.0), Inches(3.0))
    tb(s, 7.0, 3.0, 4.0, 0.4, "잘린 캡션 텍스트 예시", font="Wanted Sans", size=14)
    s = add_slide(p)  # p2 (negative): full overlay
    s.shapes.add_picture(png(tmp_path, "p2.png"), Inches(4.0), Inches(2.0), Inches(4.0), Inches(3.0))
    tb(s, 4.4, 3.0, 3.0, 0.4, "오버레이 캡션", font="Wanted Sans", size=14)
    s = add_slide(p)  # p3 (negative): straddling over a transparent margin (alpha-trimmed)
    s.shapes.add_picture(png(tmp_path, "p3.png", opaque_box=(0.0, 0.0, 0.45, 1.0)),
                         Inches(4.0), Inches(2.0), Inches(4.0), Inches(3.0))
    tb(s, 6.5, 3.0, 3.0, 0.4, "투명부 위 캡션 예시", font="Wanted Sans", size=14)
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    w17 = by_code(warns, "W17")
    assert any(si == 1 for (si, m, d) in w17)
    assert not any(si == 2 for (si, m, d) in w17)
    assert not any(si == 3 for (si, m, d) in w17)

def test_summary_incomplete_flag(tmp_path):
    """summary.incomplete: a machine-readable signal for whether checking was
    impossible (W18) (third panel: corrects documentation overclaiming)."""
    p = new_prs()
    s = add_slide(p)
    box = tb(s, 1, 1, 5, 0.5, "쓰레기 크기 한글", font="Wanted Sans", no_size=True)
    box.text_frame.paragraphs[0].runs[0]._r.get_or_add_rPr().set("sz", "notanumber")
    doc = json.loads(run_cli([save(p, tmp_path, "bad.pptx"), "--json"]).stdout)
    assert doc["summary"]["incomplete"] is True
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "clean", font="Wanted Sans", size=20)
    doc = json.loads(run_cli([save(p, tmp_path, "good.pptx"), "--json"]).stdout)
    assert doc["summary"]["incomplete"] is False

def test_geometry_uses_style_resolver(tmp_path):
    """P0-2: geometry checks use the same inherited size as E3. If layout lstStyle
    40pt text overflows on the right, it's W16 (the old version computed with a
    default 12pt and missed it)."""
    p = Presentation()
    p.slide_width = Inches(13.333)
    p.slide_height = Inches(7.5)
    lay = p.slide_layouts[1]
    for ph in lay.placeholders:
        if ph.placeholder_format.idx == 1:
            txBody = ph.text_frame._txBody
            lst = txBody.find(qn("a:lstStyle"))
            if lst is None:
                lst = txBody.makeelement(qn("a:lstStyle"), {})
                txBody.insert(1, lst)
            lvl1 = lst.makeelement(qn("a:lvl1pPr"), {})
            defR = lst.makeelement(qn("a:defRPr"), {"sz": "4000"})
            lvl1.append(defR)
            lst.append(lvl1)
    s = p.slides.add_slide(lay)
    s.shapes.title.text_frame.text = ""
    body = s.placeholders[1]
    body.left, body.top = Inches(12.0), Inches(2.0)
    body.width, body.height = Inches(1.2), Inches(0.8)
    body.text_frame.word_wrap = False
    body.text_frame.text = "Revenue ABC"   # at 12pt it's 0.95in (passes), at 40pt it's 3.2in (overflows)
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    assert any("텍스트" in d for (_si, m, d) in by_code(warns, "W16")), warns

def test_table_cell_geometry(tmp_path):
    """P0-3: native table-cell text is included in geometry checks (a table straying
    off the right -> W16)."""
    p = new_prs()
    s = add_slide(p)
    gf = s.shapes.add_table(1, 2, Inches(11.0), Inches(2.0), Inches(4.0), Inches(1.0))
    cell = gf.table.cell(0, 1)   # the right column starts around x=13in -> off canvas
    cell.text = "overflowing table cell text"
    r = cell.text_frame.paragraphs[0].runs[0]
    r.font.size = Pt(24)
    cell.text_frame.word_wrap = False
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    assert any("텍스트" in d for (_si, m, d) in by_code(warns, "W16")), warns

def test_group_child_loc_bbox_absolute(tmp_path):
    """Under group-move desync (off!=chOff), a run-level finding's loc bbox is in
    slide absolute coordinates, not the group's chOff coordinate system (carried
    over from the fourth review). Child raw x=12.5in, group moved to 7.5in ->
    absolute x is 7.5in."""
    p = new_prs()
    s = add_slide(p)
    grp = s.shapes.add_group_shape()
    box = grp.shapes.add_textbox(Inches(12.5), Inches(3.0), Inches(1.5), Inches(0.4))
    r = box.text_frame.paragraphs[0].add_run()
    r.text = "깨알 주석"
    r.font.size = Pt(3)   # E3
    grp.left, grp.top = Inches(7.5), Inches(3.0)
    errors, _w = lint_full(save(p, tmp_path, "grp_loc.pptx"))
    e3 = _by(errors, "E3")
    assert e3, codes(errors)
    bbox = e3[0].loc.get("bbox")
    assert bbox and abs(bbox[0] - 7.5) < 0.05, e3[0].loc

def test_w15_w16_locations(tmp_path):
    """W15/W16 carry loc (the effective glyph's absolute bbox), and W15 identifies
    the counterpart frame via related (carried over from the fourth review: W15-17
    lacked location)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1.0, 1.0, 5.0, 1.0, "겹침 판정 대상 문장 하나", size=24)
    tb(s, 1.1, 1.05, 5.0, 1.0, "겹침 판정 대상 문장 둘", size=24)
    tb(s, 12.8, 3.0, 3.0, 0.5, "화면 밖으로 넘어가는 문장", size=18)
    _e, warns = lint_full(save(p, tmp_path, "w15loc.pptx"))
    w15 = _by(warns, "W15")
    assert w15, codes(warns)
    assert w15[0].loc and "bbox" in w15[0].loc and "related" in w15[0].loc, w15[0].loc
    assert "bbox" in w15[0].loc["related"]
    w16 = _by(warns, "W16")
    assert w16, codes(warns)
    assert w16[0].loc and "bbox" in w16[0].loc, w16[0].loc


# ---------------------------------------------------------------- 0.5.0: scan/demo subcommands

def test_br_starts_new_visual_line_in_geometry(tmp_path):
    """An explicit a:br splits the visual line: without the split, this no-wrap line
    would run past the canvas edge and fire W16 (external review: br was in E2 context
    but lost in geometry)."""
    long_half = "wide segment text " * 2
    p = new_prs()
    s = add_slide(p)
    box = tb(s, 8.0, 3.0, 2.0, 1.0, long_half, size=18)
    para = box.text_frame.paragraphs[0]
    para._p.append(para._p.makeelement(qn("a:br"), {}))
    r2 = para.add_run()
    r2.text = long_half
    r2.font.size = Pt(18)
    _e, warns = lint_full(save(p, tmp_path, "br_geo.pptx"))
    assert "W16" not in codes(warns), codes(warns)
    # Control: the same text without the break does overflow
    p2 = new_prs()
    s2 = add_slide(p2)
    tb(s2, 8.0, 3.0, 2.0, 1.0, long_half + " " + long_half, size=18)
    _e2, warns2 = lint_full(save(p2, tmp_path, "br_geo_ctrl.pptx"))
    assert "W16" in codes(warns2), codes(warns2)

def test_insets_shift_glyph_origin(tmp_path):
    """A huge lIns pushes the glyph start right of the frame edge: only inset-aware
    geometry sees this short text cross the canvas (external review: insets ignored)."""
    p = new_prs()
    s = add_slide(p)
    box = tb(s, 11.0, 3.0, 2.3, 0.6, "clipped by inset", size=16)
    bodyPr = box.text_frame._txBody.find(qn("a:bodyPr"))
    bodyPr.set("lIns", str(int(1.5 * 914400)))
    _e, warns = lint_full(save(p, tmp_path, "inset.pptx"))
    assert "W16" in codes(warns), codes(warns)

def test_merged_cells_single_count(tmp_path):
    """Continuation cells of a merged region mirror the origin's text: walking them
    double-counted findings (external review)."""
    p = new_prs()
    s = add_slide(p)
    gfx = s.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(6), Inches(2))
    tbl = gfx.table
    tbl.cell(0, 0).merge(tbl.cell(0, 1))
    tbl.cell(0, 0).text = "한글"
    r = tbl.cell(0, 0).text_frame.paragraphs[0].runs[0]
    r.font.name = "Arial"
    r.font.size = Pt(12)
    errors, _w = lint_full(save(p, tmp_path, "merge.pptx"))
    e1 = [f for f in errors if f.code == "E1"]
    assert len(e1) == 1, codes(errors)

def test_w15_loc_carries_paragraph_and_cell_fields(tmp_path):
    """Geometry locations gained paragraph indexes and (for fld-only paragraphs) the
    field marker (external review P1)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1.0, 1.0, 5.0, 1.0, "overlap sentence one here", size=24)
    tb(s, 1.1, 1.05, 5.0, 1.0, "overlap sentence two here", size=24)
    _e, warns = lint_full(save(p, tmp_path, "w15p.pptx"))
    w15 = [f for f in warns if f.code == "W15"]
    assert w15 and "paragraph" in w15[0].loc, w15[0].loc if w15 else warns

def test_junit_fail_incomplete_matches_exit(tmp_path):
    """0.7.1 P0: under --fail-incomplete the CLI exits 1 for a W18-only deck, and the
    JUnit report must agree (W18 as <failure>), not read green while the run failed."""
    import xml.etree.ElementTree as ET
    # a vertical-text frame abstains -> W18, no ERROR
    p = new_prs()
    s = add_slide(p)
    box = s.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(4))
    r = box.text_frame.paragraphs[0].add_run()
    r.text = "세로쓰기 텍스트"
    r.font.size = Pt(18)
    r._r.get_or_add_rPr().append(box._element.makeelement(qn("a:ea"), {"typeface": "맑은 고딕"}))
    box.text_frame._txBody.find(qn("a:bodyPr")).set("vert", "eaVert")
    deck = save(p, tmp_path, "vt.pptx")
    out = os.path.join(str(tmp_path), "v.xml")
    r1 = run_cli([deck, "--profile", "full", "--junit", out])
    w18 = [c for c in ET.parse(out).getroot().iter("testcase")
           if c.get("name").startswith("W18")]
    assert r1.returncode == 0 and (not w18 or w18[0].find("failure") is None)
    r2 = run_cli([deck, "--profile", "full", "--fail-incomplete", "--junit", out])
    w18 = [c for c in ET.parse(out).getroot().iter("testcase")
           if c.get("name").startswith("W18")]
    assert r2.returncode == 1 and w18 and w18[0].find("failure") is not None

def test_geometry_findings_carry_confidence(tmp_path):
    """0.8.x: geometry findings state their evidence class (xml_geometry estimate,
    not render-confirmed) so a consumer can rank them below measured findings."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 12.8, 3.0, 3.0, 0.5, "off canvas words here", size=18, ea="맑은 고딕")
    deck = save(p, tmp_path, "conf.pptx")
    v2 = json.loads(run_cli([deck, "--schema", "2", "--json"]).stdout)
    w16 = [f for f in v2["findings"] if f["code"] == "W16"][0]
    assert w16["data"]["confidence"] == "estimate"
    assert w16["data"]["evidence_source"] == "xml_geometry"
    assert w16["data"]["render_confirmed"] is False
