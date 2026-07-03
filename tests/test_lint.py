# -*- coding: utf-8 -*-
"""aro 게이트 회귀: 양성(결함을 심은 pptx가 잡히는가)과 음성(클린·의도 연출이
오탐 없이 통과하는가)을 게이트별 픽스처로 고정한다.

긴 대시 등 린터가 차단하는 문자는 소스에 리터럴로 두지 않고 chr(코드포인트)로 만든다:
소스는 깨끗하고 검사 대상 pptx 안에만 금지문자가 존재해야 게이트 테스트가 된다.
"""
import os
import json
import subprocess
import sys

import pytest
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.oxml.ns import qn

import aro.lint as jl


# ---------------------------------------------------------------- helpers
def new_prs():
    p = Presentation()
    p.slide_width = Inches(13.333)
    p.slide_height = Inches(7.5)
    return p


def add_slide(p):
    return p.slides.add_slide(p.slide_layouts[6])


def tb(s, x, y, w, h, text, font=None, size=12, color="222222", spc=None, no_size=False):
    box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    r = box.text_frame.paragraphs[0].add_run()
    r.text = text
    if font:
        r.font.name = font
    if not no_size:
        r.font.size = Pt(size)
    r.font.color.rgb = RGBColor.from_string(color)
    if spc is not None:
        rPr = r._r.get_or_add_rPr()
        rPr.set("spc", str(spc))
    return box


def rect(s, x, y, w, h, fill_hex):
    sp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.fill.solid()
    sp.fill.fore_color.rgb = RGBColor.from_string(fill_hex)
    sp.line.fill.background()
    return sp


def vconn(s, x, y0, y1, line_hex):
    cn = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y0), Inches(x), Inches(y1))
    cn.line.color.rgb = RGBColor.from_string(line_hex)
    cn.line.width = Pt(2)
    return cn


def set_autofit_scale(box, pct):
    bodyPr = box.text_frame._txBody.find(qn("a:bodyPr"))
    for tag in ("a:normAutofit", "a:spAutoFit", "a:noAutofit"):
        e = bodyPr.find(qn(tag))
        if e is not None:
            bodyPr.remove(e)
    na = bodyPr.makeelement(qn("a:normAutofit"), {"fontScale": str(int(pct * 1000))})
    bodyPr.append(na)


def png(d, name, size=(400, 300), opaque_box=None):
    """전체 불투명 RGB 또는 opaque_box=(l,t,r,b 비율) 영역만 불투명한 RGBA PNG."""
    from PIL import Image, ImageDraw
    path = os.path.join(str(d), name)
    if opaque_box is None:
        Image.new("RGB", size, (40, 60, 90)).save(path)
    else:
        im = Image.new("RGBA", size, (0, 0, 0, 0))
        dr = ImageDraw.Draw(im)
        l, t, r, b = opaque_box
        dr.rectangle([size[0] * l, size[1] * t, size[0] * r, size[1] * b], fill=(40, 60, 90, 255))
        im.save(path)
    return path


def codes(items):
    return [c for (_si, c, _m, _d) in items]


def by_code(items, code):
    return [(si, m, d) for (si, c, m, d) in items if c == code]


def save(p, d, name):
    path = os.path.join(str(d), name)
    p.save(path)
    return path


# ---------------------------------------------------------------- line-level gates
def test_line_gates_positive(tmp_path):
    p = new_prs()
    s = add_slide(p)   # p1: E1 한글을 라틴 전용 폰트로
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    s = add_slide(p)   # p2: E2 긴 대시(전각 하이픈)
    tb(s, 1, 1, 5, 0.5, "dash" + chr(0xFF0D) + "test", font="Wanted Sans", size=12)
    s = add_slide(p)   # p3: E3 직접 4pt
    tb(s, 1, 1, 5, 0.5, "tiny text 4pt", font="Wanted Sans", size=4)
    s = add_slide(p)   # p4: E3 autofit 40pt * 10% = 4pt
    box = tb(s, 1, 1, 5, 1.0, "autofit shrink", font="Wanted Sans", size=40)
    set_autofit_scale(box, 10.0)
    s = add_slide(p)   # p5: E4 CJK 양수 트래킹
    tb(s, 1, 1, 5, 0.5, "자간 벌어진 한글", font="Wanted Sans", size=12, spc=100)
    s = add_slide(p)   # p6: W1 넓은 프레임 본문 8.5pt 40자+
    tb(s, 1, 1, 11, 0.6, "0123456789" * 5, font="Wanted Sans", size=8.5)
    s = add_slide(p)   # p7: W5 크기 미상
    tb(s, 1, 1, 5, 0.5, "no size run", font="Wanted Sans", no_size=True)
    s = add_slide(p)   # p8: W8 좁은 프레임 소형 한글 6pt
    tb(s, 1, 1, 2, 0.4, "목업 안 작은 한글", font="Wanted Sans", size=6)
    errors, warns = jl.lint(save(p, tmp_path, "fx.pptx"))
    ec = codes(errors)
    assert "E1" in ec and by_code(errors, "E1")[0][0] == 1
    assert "E2" in ec and by_code(errors, "E2")[0][0] == 2
    e3 = [si for (si, m, d) in by_code(errors, "E3")]
    assert 3 in e3 and 4 in e3
    assert "E4" in ec and by_code(errors, "E4")[0][0] == 5
    assert any(si == 6 for (si, m, d) in by_code(warns, "W1"))
    assert any(si == 7 for (si, m, d) in by_code(warns, "W5"))
    assert any(si == 8 for (si, m, d) in by_code(warns, "W8"))


def test_e1_nofont_empty_theme_slot(tmp_path):
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "폰트 미지정 한글", size=12)
    errors, _w = jl.lint(save(p, tmp_path, "fx.pptx"))
    assert any(c == "E1" and "미지정" in m for (_si, c, m, _d) in errors)


# ---------------------------------------------------------------- layout-level gates
def test_w6_layout_clones(tmp_path):
    p = new_prs()
    for i in range(6):
        s = add_slide(p)
        tb(s, 1, 0.8, 8, 0.6, "Title block %d" % i, font="Wanted Sans", size=24)
        tb(s, 1, 2.0, 6, 2.5, "Body block", font="Wanted Sans", size=12)
        tb(s, 8, 2.0, 4, 2.5, "Side block", font="Wanted Sans", size=12)
    _e, warns = jl.lint(save(p, tmp_path, "fx.pptx"))
    assert "W6" in codes(warns)


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
    _e, warns = jl.lint(path, render_dir=pages)
    assert "W7" in codes(warns)


def test_w9_accent_vbars(tmp_path):
    p = new_prs()
    s = add_slide(p)
    for i in range(4):
        y = 1.5 + i * 1.1
        vconn(s, 5.0, y, y + 0.65, "E6A94E")
        tb(s, 5.2, y, 5.5, 0.6, "callout item %d body" % i, font="Wanted Sans", size=12)
    _e, warns = jl.lint(save(p, tmp_path, "fx.pptx"))
    assert "W9" in codes(warns)


def test_w10_diagram_clone_marks(tmp_path):
    p = new_prs()
    for _ in range(2):
        s = add_slide(p)
        rect(s, 2.0, 2.0, 8.0, 3.0, "2E3540")
        for i in range(9):
            rect(s, 2.5 + i * 0.8, 3.0, 0.18, 0.18, "5B6470")
        tb(s, 1, 0.8, 8, 0.6, "diagram page", font="Wanted Sans", size=20)
    _e, warns = jl.lint(save(p, tmp_path, "fx.pptx"))
    assert "W10" in codes(warns)


def test_w11_ai_copy(tmp_path):
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 10, 0.8, "오늘날 급변하는 시장 환경에서", font="Wanted Sans", size=24)
    s = add_slide(p)
    tb(s, 1, 1, 10, 0.8, "두 사업의 시너지 효과를 극대화", font="Wanted Sans", size=14)
    _e, warns = jl.lint(save(p, tmp_path, "fx.pptx"))
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
    _e, warns = jl.lint(save(p, tmp_path, "fx.pptx"))
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
    _e, warns = jl.lint(save(p, tmp_path, "fx.pptx"))
    assert "W13" in codes(warns)


def test_w14_nominal_titles_and_ghost(tmp_path):
    p = new_prs()
    for t in ("시장 현황", "경쟁 구도 분석", "제품 라인업 개요", "사업 확장 전략",
              "재무 운용 계획", "향후 추진 방안"):
        s = add_slide(p)
        tb(s, 1, 0.8, 9, 0.8, t, font="Wanted Sans", size=26)
        tb(s, 1, 2.2, 10, 3, "본문 내용", font="Wanted Sans", size=12)
    ghost = []
    _e, warns = jl.lint(save(p, tmp_path, "fx.pptx"), ghost=ghost)
    assert "W14" in codes(warns)
    assert len(ghost) == 6


# ---------------------------------------------------------------- geometry gates
def test_w15_overlap_and_intentional_layers(tmp_path):
    p = new_prs()
    s = add_slide(p)  # p1: 진짜 겹침(값 위 라벨)
    tb(s, 1, 2.0, 4, 1.2, "12,400원", font="Wanted Sans", size=44)
    tb(s, 1.2, 2.15, 3, 0.5, "+11.7% 전년 대비", font="Wanted Sans", size=14)
    s = add_slide(p)  # p2: 드롭캡 = 의도
    tb(s, 1.0, 1.0, 1.2, 1.1, "피", font="Wanted Sans", size=60)
    tb(s, 1.0, 1.05, 5, 0.5, "지컬 AI", font="Wanted Sans", size=13)
    s = add_slide(p)  # p3: 동일 텍스트 에코 = 의도
    tb(s, 2, 2, 4, 1, "떠난 뒤에야", font="Wanted Sans", size=30)
    tb(s, 2.05, 2.05, 4, 1, "떠난 뒤에야", font="Wanted Sans", size=30)
    _e, warns = jl.lint(save(p, tmp_path, "fx.pptx"))
    w15 = by_code(warns, "W15")
    assert any(si == 1 for (si, m, d) in w15)
    assert not any(si == 2 for (si, m, d) in w15)
    assert not any(si == 3 for (si, m, d) in w15)


def test_w16_overflow_and_negatives(tmp_path):
    p = new_prs()
    s = add_slide(p)  # p1: 텍스트 우측 뚫음(wrap=none 한 줄)
    tb(s, 11.0, 1.0, 4.0, 0.6, "Quarterly revenue grew twenty two percent",
       font="Wanted Sans", size=18)
    s = add_slide(p)  # p2: word_wrap=True 프레임의 긴 한글이 바닥 뚫음
    bx = tb(s, 1.0, 7.0, 3.0, 0.4,
            "긴 본문이 프레임 계산보다 길어져 바닥을 뚫고 내려가는 경우를 재현한다",
            font="Wanted Sans", size=16)
    bx.text_frame.word_wrap = True
    s = add_slide(p)  # p3: 그림 우측 잘림
    s.shapes.add_picture(png(tmp_path, "op.png"), Inches(12.5), Inches(2.0), Inches(2.0), Inches(1.5))
    s = add_slide(p)  # p4(음성): 풀블리드 + 투명 여백 차트 경계 걸침
    s.shapes.add_picture(png(tmp_path, "bleed.png"), Inches(0), Inches(0), Inches(13.4), Inches(7.55))
    s.shapes.add_picture(png(tmp_path, "chart.png", opaque_box=(0.1, 0.1, 0.5, 0.5)),
                         Inches(9.0), Inches(4.0), Inches(5.0), Inches(3.6))
    tb(s, 1, 1, 6, 0.6, "음성 페이지 본문", font="Wanted Sans", size=14)
    _e, warns = jl.lint(save(p, tmp_path, "fx.pptx"))
    w16 = by_code(warns, "W16")
    assert any(si == 1 for (si, m, d) in w16)
    assert any(si == 2 for (si, m, d) in w16)
    assert any(si == 3 for (si, m, d) in w16)
    assert not any(si == 4 for (si, m, d) in w16)


def test_w17_straddle_and_negatives(tmp_path):
    p = new_prs()
    s = add_slide(p)  # p1: 걸침
    s.shapes.add_picture(png(tmp_path, "p1.png"), Inches(4.0), Inches(2.0), Inches(4.0), Inches(3.0))
    tb(s, 7.0, 3.0, 4.0, 0.4, "잘린 캡션 텍스트 예시", font="Wanted Sans", size=14)
    s = add_slide(p)  # p2(음성): 완전 오버레이
    s.shapes.add_picture(png(tmp_path, "p2.png"), Inches(4.0), Inches(2.0), Inches(4.0), Inches(3.0))
    tb(s, 4.4, 3.0, 3.0, 0.4, "오버레이 캡션", font="Wanted Sans", size=14)
    s = add_slide(p)  # p3(음성): 투명 여백 위 걸침(알파 트림)
    s.shapes.add_picture(png(tmp_path, "p3.png", opaque_box=(0.0, 0.0, 0.45, 1.0)),
                         Inches(4.0), Inches(2.0), Inches(4.0), Inches(3.0))
    tb(s, 6.5, 3.0, 3.0, 0.4, "투명부 위 캡션 예시", font="Wanted Sans", size=14)
    _e, warns = jl.lint(save(p, tmp_path, "fx.pptx"))
    w17 = by_code(warns, "W17")
    assert any(si == 1 for (si, m, d) in w17)
    assert not any(si == 2 for (si, m, d) in w17)
    assert not any(si == 3 for (si, m, d) in w17)


def test_geo_robustness_no_false_positives(tmp_path):
    """적대 검증이 실측 재현한 오탐 시나리오 6종: 전부 플래그 0이어야 한다."""
    from PIL import Image, ImageDraw
    p = new_prs()
    s = add_slide(p)  # p1: wrap=none 팬텀 랩
    tb(s, 1.0, 1.0, 1.0, 0.4, "Quarterly revenue grew 18% on cloud momentum",
       font="Wanted Sans", size=12)
    tb(s, 1.0, 1.4, 2.0, 0.3, "Source: 10-K", font="Wanted Sans", size=9)
    s = add_slide(p)  # p2: 그룹 이동 desync
    grp = s.shapes.add_group_shape()
    box = grp.shapes.add_textbox(Inches(12.5), Inches(3.0), Inches(1.5), Inches(0.4))
    r = box.text_frame.paragraphs[0].add_run()
    r.text = "group label"
    r.font.size = Pt(12)
    grp.left, grp.top = Inches(7.5), Inches(3.0)
    s = add_slide(p)  # p3: 회전 그림(회전 전 bbox는 밖, 렌더는 안)
    pic = s.shapes.add_picture(png(tmp_path, "rot.png", size=(400, 50)),
                               Inches(11.0), Inches(3.5), Inches(4.0), Inches(0.5))
    pic.rotation = 90
    s = add_slide(p)  # p4: flipH(잉크 미러로 화면 안)
    pic = s.shapes.add_picture(png(tmp_path, "flip.png", opaque_box=(0.0, 0.0, 0.5, 1.0)),
                               Inches(-0.8), Inches(3.0), Inches(2.0), Inches(1.0))
    pic._element.spPr.find(qn("a:xfrm")).set("flipH", "1")
    s = add_slide(p)  # p5: P모드+tRNS 투명
    im = Image.new("RGBA", (400, 400), (0, 0, 0, 0))
    ImageDraw.Draw(im).rectangle([150, 150, 250, 250], fill=(40, 60, 90, 255))
    pmode = os.path.join(str(tmp_path), "pmode.png")
    im.convert("P").save(pmode, transparency=0)
    s.shapes.add_picture(pmode, Inches(10.33), Inches(2.0), Inches(4.0), Inches(4.0))
    s = add_slide(p)  # p6: 사진 위 솔리드 카드 위 캡션
    s.shapes.add_picture(png(tmp_path, "photo.png"), Inches(4.0), Inches(0.5), Inches(6.0), Inches(6.5))
    rect(s, 6.0, 2.5, 4.0, 1.5, "FFFFFF")
    tb(s, 6.3, 3.0, 3.4, 0.4, "카드 위 캡션 문장", font="Wanted Sans", size=14)
    _e, warns = jl.lint(save(p, tmp_path, "fx.pptx"))
    geo = [x for x in warns if x[1] in ("W15", "W16", "W17")]
    assert not geo, geo


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
def test_clean_deck_no_flags(tmp_path):
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 0.8, 9, 0.8, "Clean title page", font="Wanted Sans", size=30)
    tb(s, 1, 2.2, 10, 3, "Readable body text at normal size.", font="Wanted Sans", size=13)
    s = add_slide(p)
    tb(s, 1, 0.8, 6, 0.6, "Second layout", font="Wanted Sans", size=22)
    rect(s, 7.5, 1.0, 4.5, 5.0, "20242B")
    tb(s, 1, 3.5, 5.5, 2.5, "Different grid here.", font="Wanted Sans", size=13)
    s = add_slide(p)
    tb(s, 4, 3.2, 6, 0.8, "Closing page", font="Wanted Sans", size=26)
    errors, warns = jl.lint(save(p, tmp_path, "fx.pptx"))
    assert not errors and not warns, (codes(errors), codes(warns))


def test_cli_json_and_exit_codes(tmp_path):
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    bad = save(p, tmp_path, "bad.pptx")
    # stdin=DEVNULL: pytest 캡처 하의 Windows 핸들 상속 실패(WinError 6) 회피
    r = subprocess.run([sys.executable, "-m", "aro.lint", bad, "--json"],
                       capture_output=True, text=True, encoding="utf-8",
                       stdin=subprocess.DEVNULL)
    doc = json.loads(r.stdout)
    assert r.returncode == 1
    assert doc["summary"]["error_count"] >= 1 and not doc["summary"]["pass"]
    assert any(e["code"] == "E1" for e in doc["errors"])

    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.6, "clean page", font="Wanted Sans", size=20)
    good = save(p, tmp_path, "good.pptx")
    r = subprocess.run([sys.executable, "-m", "aro.lint", good, "--json"],
                       capture_output=True, text=True, encoding="utf-8",
                       stdin=subprocess.DEVNULL)
    doc = json.loads(r.stdout)
    assert r.returncode == 0 and doc["summary"]["pass"]


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-q"]))
