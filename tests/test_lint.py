# -*- coding: utf-8 -*-
"""archforge 게이트 회귀: 양성(결함을 심은 pptx가 잡히는가)과 음성(클린·의도 연출이
오탐 없이 통과하는가)을 게이트별 픽스처로 고정한다.

긴 대시 등 린터가 차단하는 문자는 소스에 리터럴로 두지 않고 chr(코드포인트)로 만든다:
소스는 깨끗하고 검사 대상 pptx 안에만 금지문자가 존재해야 게이트 테스트가 된다.

E1 픽스처의 기대값은 PowerPoint COM 렌더 실측 모델(2026-07-10, docs/CALIBRATION.md)을
따른다: run a:ea > 테마 minorFont a:ea(비어있지 않으면 run a:latin보다 우선)
> (테마 ea 빈 슬롯일 때만) run a:latin > OS 폴백(Malgun).
"""
import os
import re
import json
import shutil
import subprocess
import sys
import zipfile

import pytest
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.oxml.ns import qn

import archforge.lint as jl

EN_DASH = chr(0x2013)
EM_DASH = chr(0x2014)
MINUS = chr(0x2212)
FW_HYPHEN = chr(0xFF0D)


# ---------------------------------------------------------------- helpers
def new_prs():
    p = Presentation()
    p.slide_width = Inches(13.333)
    p.slide_height = Inches(7.5)
    return p


def add_slide(p):
    return p.slides.add_slide(p.slide_layouts[6])


def tb(s, x, y, w, h, text, font=None, size=12, color="222222", spc=None, no_size=False, ea=None):
    box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    r = box.text_frame.paragraphs[0].add_run()
    r.text = text
    if font:
        r.font.name = font
    if ea:
        rPr = r._r.get_or_add_rPr()
        el = rPr.makeelement(qn("a:ea"), {"typeface": ea})
        rPr.append(el)
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


def patch_theme_ea(path, typeface):
    """저장된 pptx의 테마 a:ea 빈 슬롯을 typeface로 교체(zip 재작성). 멀티마스터 없는
    기본 템플릿용: E1 테마 분기 픽스처를 만드는 유일하게 안정적인 방법."""
    tmp = path + ".patched.pptx"
    new = ('<a:ea typeface="%s"/>' % typeface).encode("utf-8")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.search(r"theme/theme\d+\.xml$", item.filename):
                data = data.replace(b'<a:ea typeface=""/>', new)
            zout.writestr(item, data)
    shutil.move(tmp, path)
    return path


def run_cli(args):
    """CLI 실행 헬퍼. stdin=DEVNULL: pytest 캡처 하의 Windows 핸들 상속 실패(WinError 6) 회피."""
    return subprocess.run([sys.executable, "-m", "archforge.lint"] + args,
                          capture_output=True, text=True, encoding="utf-8",
                          stdin=subprocess.DEVNULL)


# ---------------------------------------------------------------- line-level gates
def test_line_gates_positive(tmp_path):
    p = new_prs()
    s = add_slide(p)   # p1: E1 한글을 라틴 전용 폰트로
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    s = add_slide(p)   # p2: E2 긴 대시(전각 하이픈)
    tb(s, 1, 1, 5, 0.5, "dash" + FW_HYPHEN + "test", font="Wanted Sans", size=12)
    s = add_slide(p)   # p3: E3 직접 4pt
    tb(s, 1, 1, 5, 0.5, "tiny text 4pt", font="Wanted Sans", size=4)
    s = add_slide(p)   # p4: E3 autofit 40pt * 10% = 4pt
    box = tb(s, 1, 1, 5, 1.0, "autofit shrink", font="Wanted Sans", size=40)
    set_autofit_scale(box, 10.0)
    s = add_slide(p)   # p5: E4 CJK 양수 트래킹
    tb(s, 1, 1, 5, 0.5, "자간 벌어진 한글", font="Wanted Sans", size=12, spc=100)
    s = add_slide(p)   # p6: W1 넓은 프레임 본문 8.5pt 40자+
    tb(s, 1, 1, 11, 0.6, "0123456789" * 5, font="Wanted Sans", size=8.5)
    s = add_slide(p)   # p7: 크기 미지정 -> defaultTextStyle 18pt로 해석(W5 아님: 상속 해석 도입)
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
    # p7: 이전 버전은 W5(상속 미해석)였다. 상속 체인 도입으로 defaultTextStyle 18pt가 잡혀
    # W5도, 크기 게이트 오발화도 없어야 한다.
    assert not any(si == 7 for (si, m, d) in by_code(warns, "W5"))
    assert not any(si == 7 for (si, m, d) in by_code(errors, "E3"))
    assert any(si == 8 for (si, m, d) in by_code(warns, "W8"))


def test_e1_nofont_empty_theme_slot(tmp_path):
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "폰트 미지정 한글", size=12)
    errors, _w = jl.lint(save(p, tmp_path, "fx.pptx"))
    assert any(c == "E1" and "미지정" in m for (_si, c, m, _d) in errors)


def test_e1_render_model_slots(tmp_path):
    """실측 렌더 모델 회귀(2026-07-10 COM 프로브):
    p1 run ea가 라틴 전용 -> E1. p2 빈 테마에서 latin이 라틴 전용(Inter, 구 미탐) -> E1.
    p3 빈 테마에서 latin이 한글 폰트 -> 실제로 그 폰트가 한글을 그림(합법 패턴) -> 통과.
    p4 run ea가 한글 폰트면 latin이 라틴 전용이어도 ea가 이김 -> 통과."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "이에이 슬롯 모노", font="Arial", ea="IBM Plex Mono", size=12)
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "인터 폴백 한글", font="Inter", size=12)
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "원티드 산스 한글", font="Wanted Sans", size=12)
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "이에이 승리 한글", font="IBM Plex Mono", ea="Wanted Sans", size=12)
    errors, _w = jl.lint(save(p, tmp_path, "fx.pptx"))
    e1_pages = [si for (si, m, d) in by_code(errors, "E1")]
    assert 1 in e1_pages and 2 in e1_pages
    assert 3 not in e1_pages and 4 not in e1_pages


def test_e1_theme_ea_korean_suppresses_latin(tmp_path):
    """테마 ea가 한글 폰트면 run latin이 라틴 전용이어도 렌더는 테마 ea가 받는다(실측):
    구버전 ea-or-latin 대용이 내던 E1 오탐의 회귀 고정."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "테마가 받는 한글", font="IBM Plex Mono", size=12)
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "폰트 미지정 한글", size=12)
    path = patch_theme_ea(save(p, tmp_path, "fx.pptx"), "Malgun Gothic")
    errors, _w = jl.lint(path)
    assert not by_code(errors, "E1"), errors


def test_e1_theme_ea_latin_only_flags(tmp_path):
    """테마 ea 자체가 라틴 전용이면 폰트 미지정 CJK는 E1(테마 분기)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "폰트 미지정 한글", size=12)
    path = patch_theme_ea(save(p, tmp_path, "fx.pptx"), "Consolas")
    errors, _w = jl.lint(path)
    assert any(c == "E1" and "테마" in m for (_si, c, m, _d) in errors), errors


def test_theme_ea_by_master_relationship(tmp_path):
    """테마 해석이 iter_parts 순서가 아니라 마스터->테마 관계로 나오는지(키=마스터 partname)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "관계 해석", font="Wanted Sans", size=12)
    prs = Presentation(save(p, tmp_path, "fx.pptx"))
    ea_map = jl.theme_ea_by_master(prs)
    assert len(ea_map) == 1
    key = next(iter(ea_map))
    assert "slideMaster" in key
    assert ea_map[key] == ""   # 기본 템플릿 = 빈 슬롯


def test_e1_unit_matrix():
    """e1_violation 판정 매트릭스(실측 렌더 모델 문서화를 겸한 유닛 회귀)."""
    v = jl.e1_violation
    assert v("한글", {"ea": "IBM Plex Mono"}, "") is not None          # run ea 라틴 전용
    assert v("한글", {"ea": "맑은 고딕", "latin": "Consolas"}, "") is None   # run ea 한글
    assert v("한글", {"latin": "IBM Plex Mono"}, "맑은 고딕") is None   # 테마 ea가 렌더를 받음
    assert v("한글", {}, "Consolas") is not None                        # 테마 ea 라틴 전용
    assert v("한글", {"latin": "Wanted Sans"}, "") is None              # 빈 테마 + latin 한글폰트
    assert v("한글", {"latin": "Inter"}, "") is not None                # 빈 테마 + latin 라틴 전용(구 미탐)
    assert v("한글", {}, "") is not None                                # 전부 없음 = Malgun 확정


def test_latin_only_font_boundaries():
    """블록리스트 접두 매칭 경계: 한글 완비 변형이 접두에 말려들지 않는다."""
    f = jl.is_latin_only_font
    assert f("Inter") and f("Arial") and f("Calibri") and f("IBM Plex Sans") and f("Noto Sans")
    assert not f("NanumGothicCoding")     # 한글 완비 고정폭: 구버전 오분류의 회귀 고정
    assert not f("IBM Plex Sans KR")
    assert not f("Noto Sans KR") and not f("Noto Serif KR")
    assert not f("Arial Unicode MS")
    assert not f("Wanted Sans") and not f("Pretendard") and not f("Malgun Gothic")
    assert not f("") and not f(None)


def test_e2_numeric_context(tmp_path):
    """E2 맥락 예외: 숫자 범위 en dash와 음수 부호 U+2212는 기본 모드 통과, em dash는 항상 차단."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "FY2020" + EN_DASH + "2024 실적", font="Wanted Sans", size=12)
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, MINUS + "3.2% 하락", font="Wanted Sans", size=12)
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "이건" + EM_DASH + "안 됨", font="Wanted Sans", size=12)
    path = save(p, tmp_path, "fx.pptx")
    errors, _w = jl.lint(path)
    e2_pages = [si for (si, m, d) in by_code(errors, "E2")]
    assert e2_pages == [3], errors
    # strict: 예외 해제, 세 페이지 전부 차단
    errors_s, _w = jl.lint(path, strict=True)
    assert sorted(si for (si, m, d) in by_code(errors_s, "E2")) == [1, 2, 3]


def test_dash_violations_unit():
    dv = jl.dash_violations
    assert dv("2020" + EN_DASH + "2024") == []
    assert dv("2020" + EN_DASH + "2024", strict=True) == [EN_DASH]
    assert dv(MINUS + "3.2%") == []
    assert dv("A" + EN_DASH + "B") == [EN_DASH]        # 숫자 맥락 아님 -> 차단 유지
    assert dv("끝" + MINUS) == [MINUS]                  # 뒤에 숫자 없음 -> 차단 유지
    assert dv("a" + EM_DASH + "b") == [EM_DASH]


def test_e4_universal_measure_spc(tmp_path):
    """spc의 universal measure('1.5pt')와 쓰레기값: 크래시 없이 전자는 E4, 후자는 무시."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "자간 벌어진 한글", font="Wanted Sans", size=12, spc="1.5pt")
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "정상 자간 한글", font="Wanted Sans", size=12, spc="garbage")
    errors, warns = jl.lint(save(p, tmp_path, "fx.pptx"))
    e4_pages = [si for (si, m, d) in by_code(errors, "E4")]
    assert e4_pages == [1], errors


def test_text_point_attr_unit():
    tpa = jl._text_point_attr
    assert tpa("150") == 150
    assert tpa("1.5pt") == 150
    assert tpa("0.5in") == 3600
    assert tpa("abc") is None
    assert tpa(None) is None


def test_partial_salvage_on_bad_frame(tmp_path):
    """가드 입도 회귀(적대 패널 확정 2026-07-10): 같은 프레임 안에서 run1이 쓰레기 속성으로
    죽어도 run2의 진짜 E1 위반은 잡혀야 하고(per-run 가드), 삼켜진 구간은 W18로 JSON에
    표면화돼야 한다. 프레임 단위 가드 시절엔 run2 위반까지 통째로 사라져 거짓 clean이 났다."""
    p = new_prs()
    s = add_slide(p)   # p1: 한 프레임에 쓰레기 run + 진짜 E1 run
    box = tb(s, 1, 1, 5, 0.5, "쓰레기 크기 한글", font="Wanted Sans", no_size=True)
    r1 = box.text_frame.paragraphs[0].runs[0]
    r1._r.get_or_add_rPr().set("sz", "notanumber")
    r2 = box.text_frame.paragraphs[0].add_run()
    r2.text = "모노 폴백 한글"
    r2.font.name = "IBM Plex Mono"
    from pptx.util import Pt as _Pt
    r2.font.size = _Pt(12)
    s = add_slide(p)   # p2: 정상 E1 결함(페이지 간 생존도 함께 고정)
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    errors, warns = jl.lint(save(p, tmp_path, "fx.pptx"))
    e1_pages = [si for (si, m, d) in by_code(errors, "E1")]
    assert 1 in e1_pages, errors          # 같은 프레임 이웃 run의 위반 생존
    assert 2 in e1_pages, errors
    w18 = by_code(warns, "W18")
    assert any(si == 1 for (si, m, d) in w18), warns   # 삼켜진 구간이 출력 계약에 표면화


def test_w18_geometry_decoupling(tmp_path):
    """tboxes 계산 실패(손상 anchor)가 무관한 그림 W16까지 침묵시키지 않는다(적대 패널
    실측 재현 2건의 회귀 고정). W16은 살아남고 W18이 불완전성을 알린다."""
    p = new_prs()
    s = add_slide(p)
    box = tb(s, 1, 1, 4, 0.5, "손상 앵커 텍스트", font="Wanted Sans", size=14)
    bodyPr = box.text_frame._txBody.find(qn("a:bodyPr"))
    bodyPr.set("anchor", "bogusvalue")   # MSO_VERTICAL_ANCHOR 매핑 없는 값 -> tboxes 예외
    s.shapes.add_picture(png(tmp_path, "over.png"), Inches(12.8), Inches(2.0),
                         Inches(2.0), Inches(1.5))   # 우측 1.47in 넘침
    errors, warns = jl.lint(save(p, tmp_path, "fx.pptx"))
    assert any("그림" in d for (_si, m, d) in by_code(warns, "W16")), warns
    assert by_code(warns, "W18"), warns


def test_w6_page_numbers_survive_token_failure(tmp_path, monkeypatch):
    """sigs 이중 append 회귀(적대 패널 확정): _fill_tokens가 죽어도 sigs 위치가 밀리지 않아
    W6 페이지 번호가 실제 슬라이드 번호를 가리킨다."""
    def boom(slide, sw, sh):
        raise RuntimeError("token fail")
    monkeypatch.setattr(jl, "_fill_tokens", boom)
    p = new_prs()
    for i in range(6):
        s = add_slide(p)
        tb(s, 1, 0.8, 8, 0.6, "Title block %d" % i, font="Wanted Sans", size=24)
        tb(s, 1, 2.0, 6, 2.5, "Body block", font="Wanted Sans", size=12)
        tb(s, 8, 2.0, 4, 2.5, "Side block", font="Wanted Sans", size=12)
    _e, warns = jl.lint(save(p, tmp_path, "fx.pptx"))
    w6 = by_code(warns, "W6")
    assert w6, warns
    pages = [int(m) for m in re.findall(r"p(\d+)", w6[0][2])]
    assert pages and max(pages) <= 6, w6   # 이중 append 시절엔 p7~p11이 나왔다


def test_e1_theme_token_resolution(tmp_path):
    """OOXML 테마 토큰(+mn-lt 류)을 실폰트로 해석(적대 패널 확정: 토큰을 문자 그대로
    블록리스트에 대조하면 E1 미탐). 기본 템플릿 테마 minor latin=Calibri(라틴 전용)라
    latin="+mn-lt" 한글 런은 E1이어야 한다."""
    p = new_prs()
    s = add_slide(p)
    box = tb(s, 1, 1, 5, 0.5, "토큰 한글", size=12)
    rPr = box.text_frame.paragraphs[0].runs[0]._r.get_or_add_rPr()
    rPr.append(rPr.makeelement(qn("a:latin"), {"typeface": "+mn-lt"}))
    errors, _w = jl.lint(save(p, tmp_path, "fx.pptx"))
    assert any("Calibri" in d for (_si, m, d) in by_code(errors, "E1")), errors


def test_resolve_font_tokens_unit():
    rft = jl.resolve_font_tokens
    thm = {"mn-lt": "Calibri", "mn-ea": "맑은 고딕", "mj-lt": "Georgia", "mj-ea": ""}
    assert rft({"latin": "+mn-lt"}, thm) == {"latin": "Calibri"}
    assert rft({"ea": "+mn-ea"}, thm) == {"ea": "맑은 고딕"}
    assert rft({"latin": "+mj-lt", "ea": "Wanted Sans"}, thm) == {"latin": "Georgia", "ea": "Wanted Sans"}
    assert rft({"ea": "+mj-ea"}, thm) == {}          # 빈 값으로 해석되면 슬롯 제거(체인에 맡김)
    assert rft({"latin": "+mn-lt"}, None) == {}       # 테마 파싱 실패면 슬롯 제거
    assert rft({"latin": "Batang"}, thm) == {"latin": "Batang"}   # 토큰 아니면 그대로


def test_e2_run_boundary_split(tmp_path):
    """run 경계로 쪼개진 숫자 범위 오탐 회귀(적대 패널 확정, high): '2020'/'-2021'이
    서로 다른 run이어도 문단 컨텍스트로 예외가 적용된다. em dash는 분리돼도 차단 유지."""
    p = new_prs()
    s = add_slide(p)   # p1: en dash 범위가 run 경계에서 분리
    box = s.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(0.5))
    para = box.text_frame.paragraphs[0]
    r1 = para.add_run(); r1.text = "2020"; r1.font.size = Pt(14)
    r2 = para.add_run(); r2.text = EN_DASH + "2024 실적"; r2.font.size = Pt(14)
    s = add_slide(p)   # p2: em dash 분리(항상 차단)
    box = s.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(0.5))
    para = box.text_frame.paragraphs[0]
    r1 = para.add_run(); r1.text = "말"; r1.font.size = Pt(14)
    r2 = para.add_run(); r2.text = EM_DASH + "이음"; r2.font.size = Pt(14)
    path = save(p, tmp_path, "fx.pptx")
    errors, _w = jl.lint(path)
    e2_pages = [si for (si, m, d) in by_code(errors, "E2")]
    assert e2_pages == [2], errors
    errors_s, _w = jl.lint(path, strict=True)   # strict는 분리 여부 무관 전부 차단
    assert sorted(set(si for (si, m, d) in by_code(errors_s, "E2"))) == [1, 2]


def test_size_inheritance_placeholder_gates(tmp_path):
    """placeholder 상속 체인 활성화 회귀: 명시 크기 없는 placeholder 텍스트가 W5로 새지 않고
    레이아웃 lstStyle 크기로 해석돼 E3/W1이 실제로 돈다(구버전은 전부 W5로 무력)."""
    p = Presentation()
    p.slide_width = Inches(13.333)
    p.slide_height = Inches(7.5)
    lay = p.slide_layouts[1]   # Title and Content
    for ph in lay.placeholders:
        sz = {0: "400", 1: "800"}.get(ph.placeholder_format.idx)   # 제목 4pt, 본문 8pt
        if sz is None:
            continue
        txBody = ph.text_frame._txBody
        lst = txBody.find(qn("a:lstStyle"))
        if lst is None:
            lst = txBody.makeelement(qn("a:lstStyle"), {})
            txBody.insert(1, lst)   # bodyPr 다음
        lvl1 = lst.makeelement(qn("a:lvl1pPr"), {})
        defR = lst.makeelement(qn("a:defRPr"), {"sz": sz})
        lvl1.append(defR)
        lst.append(lvl1)
    s = p.slides.add_slide(lay)
    s.shapes.title.text_frame.text = "판독 불가 제목"
    s.placeholders[1].text_frame.text = "본문이 레이아웃 상속 크기로 해석되는지 보는 사십자 이상의 충분히 긴 문장입니다"
    errors, warns = jl.lint(save(p, tmp_path, "fx.pptx"))
    assert not by_code(warns, "W5"), warns
    assert any("4.0pt" in m for (_si, m, _d) in by_code(errors, "E3")), errors       # 제목 4pt
    assert any("8.0pt" in m for (_si, m, _d) in by_code(warns, "W1")), warns          # 본문 8pt


def test_size_inheritance_master_txstyles(tmp_path):
    """마스터 txStyles 경로: 기본 템플릿 title=44pt, body(OBJECT)=32pt 해석."""
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


def test_w5_when_chain_fully_absent(tmp_path):
    """defaultTextStyle까지 제거하면(외부 생성기 산출물 시뮬레이션) 크기 미상이 W5로 표면화."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "크기 미상 한글", font="Wanted Sans", no_size=True)
    pres_el = p.slides._sldIdLst.getparent()
    dts = pres_el.find(qn("p:defaultTextStyle"))
    assert dts is not None
    pres_el.remove(dts)
    _e, warns = jl.lint(save(p, tmp_path, "fx.pptx"))
    assert by_code(warns, "W5"), warns


def test_table_cell_e1(tmp_path):
    """네이티브 표 셀 경로(헤드라인 기능인데 커버리지 0이던 것): 셀 안 CJK의 폰트 결함도 잡는다."""
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
    errors, _w = jl.lint(save(p, tmp_path, "fx.pptx"))
    e1 = by_code(errors, "E1")
    assert len(e1) == 1 and "Consolas" in e1[0][2], errors


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


def test_w6_tunables_suppress(tmp_path):
    """W6 튜너블: 클러스터 임계를 올리면 같은 덱이 통과(의도적 템플릿 하우스용 완화 손잡이)."""
    p = new_prs()
    for i in range(6):
        s = add_slide(p)
        tb(s, 1, 0.8, 8, 0.6, "Title block %d" % i, font="Wanted Sans", size=24)
        tb(s, 1, 2.0, 6, 2.5, "Body block", font="Wanted Sans", size=12)
        tb(s, 8, 2.0, 4, 2.5, "Side block", font="Wanted Sans", size=12)
    path = save(p, tmp_path, "fx.pptx")
    _e, warns = jl.lint(path, w6_min_cluster=10)
    assert "W6" not in codes(warns)
    # 완전 클론은 코사인이 정확히 1.00이라 sim 임계는 1.0으로 줘야 배제된다(파라미터 관통 확인)
    _e, warns = jl.lint(path, w6_sim=1.0)
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
    _e, warns = jl.lint(path, render_dir=pages)
    assert "W7" in codes(warns)


def test_render_on_textonly_deck_no_crash(tmp_path):
    """이미지가 한 장도 없는 텍스트 전용 덱을 --render로 린트해도 크래시하지 않아야 한다.
    회귀: glob 이 contrast_check 지역 import 로만 있던 시절, render_png_hits==0 분기의
    glob.glob 이 NameError 로 죽고 그게 'pptx 못 엶'으로 오라벨됐다(2026-07-04)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 0.8, 9, 0.8, "제목만 있는 페이지", font="Wanted Sans", size=28)
    tb(s, 1, 2.2, 10, 3, "본문 텍스트, 이미지는 없다.", font="Wanted Sans", size=13)
    pages = os.path.join(str(tmp_path), "pages")
    os.makedirs(pages)
    from PIL import Image
    Image.new("RGB", (1600, 900), (240, 240, 240)).save(os.path.join(pages, "slide-1.png"))  # 규약과 다른 이름
    errors, warns = jl.lint(save(p, tmp_path, "fx.pptx"), render_dir=pages)  # NameError 나면 실패
    assert isinstance(errors, list) and isinstance(warns, list)


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


def test_w14_numeric_claim_titles_pass(tmp_path):
    """숫자+단위 타이틀은 명사로 끝나도 주장형: W14가 오발화하지 않는다(장르 편향 완화)."""
    p = new_prs()
    for t in ("매출 3배 성장", "점유율 42% 달성", "비용 120억 절감", "가입자 500만 돌파",
              "마진 8pp 개선", "리텐션 2x 향상"):
        s = add_slide(p)
        tb(s, 1, 0.8, 9, 0.8, t, font="Wanted Sans", size=26)
        tb(s, 1, 2.2, 10, 3, "본문 내용", font="Wanted Sans", size=12)
    _e, warns = jl.lint(save(p, tmp_path, "fx.pptx"))
    assert "W14" not in codes(warns), by_code(warns, "W14")


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
    r = run_cli([bad, "--json"])
    doc = json.loads(r.stdout)
    assert r.returncode == 1
    assert doc["summary"]["error_count"] >= 1 and not doc["summary"]["pass"]
    assert any(e["code"] == "E1" for e in doc["errors"])
    assert "ghost" in doc and isinstance(doc["ghost"], list)   # 문서화된 JSON 계약

    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.6, "clean page", font="Wanted Sans", size=20)
    good = save(p, tmp_path, "good.pptx")
    r = run_cli([good, "--json"])
    doc = json.loads(r.stdout)
    assert r.returncode == 0 and doc["summary"]["pass"]


def test_cli_exit2_missing_and_corrupt(tmp_path):
    """exit 2 경로 회귀(구버전 미테스트): 파일 없음과 손상 pptx."""
    r = run_cli([os.path.join(str(tmp_path), "no_such.pptx")])
    assert r.returncode == 2
    assert "찾을 수 없습니다" in r.stderr
    corrupt = os.path.join(str(tmp_path), "corrupt.pptx")
    with open(corrupt, "wb") as f:
        f.write(b"this is not a pptx file at all")
    r = run_cli([corrupt])
    assert r.returncode == 2
    assert "열 수 없습니다" in r.stderr


def test_cli_text_output_and_ghost(tmp_path):
    """사람용 텍스트 출력 + --ghost 섹션(구버전 미테스트)."""
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


def test_cli_strict_gates(tmp_path):
    """--strict: WARN만 있는 덱이 exit 1로 바뀌고, E2 숫자 맥락 예외가 해제된다."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 11, 0.6, "0123456789" * 5, font="Wanted Sans", size=8.5)   # W1만
    warn_only = save(p, tmp_path, "warn.pptx")
    assert run_cli([warn_only]).returncode == 0
    assert run_cli([warn_only, "--strict"]).returncode == 1

    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 6, 0.5, "FY2020" + EN_DASH + "2024 실적", font="Wanted Sans", size=14)
    ranged = save(p, tmp_path, "range.pptx")
    assert run_cli([ranged]).returncode == 0
    r = run_cli([ranged, "--strict", "--json"])
    doc = json.loads(r.stdout)
    assert r.returncode == 1
    assert any(e["code"] == "E2" for e in doc["errors"])


def test_cli_skip_codes(tmp_path):
    """--skip: 장르 무관 경고(W14 등)를 선택 억제."""
    p = new_prs()
    for t in ("시장 현황", "경쟁 구도 분석", "제품 라인업 개요", "사업 확장 전략",
              "재무 운용 계획", "향후 추진 방안"):
        s = add_slide(p)
        tb(s, 1, 0.8, 9, 0.8, t, font="Wanted Sans", size=26)
        tb(s, 1, 2.2, 10, 3, "본문 내용", font="Wanted Sans", size=12)
    path = save(p, tmp_path, "fx.pptx")
    doc = json.loads(run_cli([path, "--json"]).stdout)
    assert any(w["code"] == "W14" for w in doc["warnings"])
    doc = json.loads(run_cli([path, "--json", "--skip", "W14"]).stdout)
    assert not any(w["code"] == "W14" for w in doc["warnings"])


# ---------------------------------------------------------------- skill packaging
def _repo_root():
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def test_skill_pack_sync():
    """리포 루트 skills/(발견용)와 패키지 동봉본(정본)이 갈라지지 않는다."""
    root = _repo_root()
    a = os.path.join(root, "skills", "archforge-pptx-lint", "SKILL.md")
    b = os.path.join(root, "src", "archforge", "skills", "archforge-pptx-lint", "SKILL.md")
    with open(a, "rb") as fa, open(b, "rb") as fb:
        assert fa.read() == fb.read(), "루트 skills/ 사본과 패키지 정본이 다릅니다"


def test_skill_frontmatter_name_matches_dir():
    """Agent Skills 규격: frontmatter name == 스킬 디렉터리명(외부 리뷰 지적 회귀)."""
    root = _repo_root()
    path = os.path.join(root, "skills", "archforge-pptx-lint", "SKILL.md")
    with open(path, encoding="utf-8") as f:
        head = f.read(500)
    m = re.search(r"^name:\s*(\S+)", head, re.M)
    assert m and m.group(1) == "archforge-pptx-lint"


def test_cli_skill_subcommand(tmp_path):
    """archforge skill: 출력·설치 두 경로 모두 패키지 동봉본과 일치."""
    r = run_cli(["skill"])
    assert r.returncode == 0
    assert "name: archforge-pptx-lint" in r.stdout
    r = run_cli(["skill", "--install", str(tmp_path)])
    assert r.returncode == 0
    dst = os.path.join(str(tmp_path), "archforge-pptx-lint", "SKILL.md")
    assert os.path.exists(dst)
    with open(dst, encoding="utf-8") as f:
        installed = f.read()
    assert "name: archforge-pptx-lint" in installed


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-q"]))
