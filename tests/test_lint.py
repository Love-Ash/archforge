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
from archforge import messages as jmsg


@pytest.fixture(autouse=True)
def _force_korean_messages(monkeypatch):
    """기존 한국어 문구 어서션 유지를 위해 테스트는 ko 고정(0.3.0 i18n).
    영어 출력은 test_lang_* 에서 명시적으로 검증한다."""
    monkeypatch.setenv("ARCHFORGE_LANG", "ko")
    jmsg.set_lang("ko")
    yield
    jmsg.set_lang(None)


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


def lint_full(*args, **kw):
    """0.4.0에서 기본 프로파일이 core로 바뀌어, 스타일 규칙까지 전제하는 기존
    픽스처들은 full을 명시한다(기본값 자체는 test_default_profile_core가 검증)."""
    kw.setdefault("profile", "full")
    return jl.lint(*args, **kw)


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


def run_cli(args, lang="ko"):
    """CLI 실행 헬퍼. stdin=DEVNULL: pytest 캡처 하의 Windows 핸들 상속 실패(WinError 6) 회피.
    기본 ko 고정(어서션이 한국어 문구), 영어 검증은 lang="en"으로."""
    env = dict(os.environ)
    env["ARCHFORGE_LANG"] = lang
    return subprocess.run([sys.executable, "-m", "archforge.lint"] + args,
                          capture_output=True, text=True, encoding="utf-8",
                          stdin=subprocess.DEVNULL, env=env)


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
    errors, warns = lint_full(save(p, tmp_path, "fx.pptx"))
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
    """폰트 미지정 + 빈 테마 ea = Malgun 폴백 E1. 0.2.1부터는 defaultTextStyle의
    +mn-lt 토큰이 상속 체인으로 해석돼(실효 latin=Calibri) 더 정확한 분기로 잡히므로
    메시지가 아니라 코드로 고정한다."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "폰트 미지정 한글", size=12)
    errors, _w = lint_full(save(p, tmp_path, "fx.pptx"))
    assert any(c == "E1" for (_si, c, m, _d) in errors), errors


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
    errors, _w = lint_full(save(p, tmp_path, "fx.pptx"))
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
    errors, _w = lint_full(path)
    assert not by_code(errors, "E1"), errors


def test_e1_theme_ea_latin_only_flags(tmp_path):
    """테마 ea 자체가 라틴 전용이면 폰트 미지정 한글은 E1. 0.2.1부터 defaultTextStyle의
    +mn-ea 토큰이 체인으로 해석돼 실효 ea=Consolas로 잡힌다(분기 무관, 폰트명으로 고정)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "폰트 미지정 한글", size=12)
    path = patch_theme_ea(save(p, tmp_path, "fx.pptx"), "Consolas")
    errors, _w = lint_full(path)
    assert any(c == "E1" and "Consolas" in d for (_si, c, m, d) in errors), errors


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
    errors, _w = lint_full(path)
    e2_pages = [si for (si, m, d) in by_code(errors, "E2")]
    assert e2_pages == [3], errors
    # strict: 예외 해제, 세 페이지 전부 차단
    errors_s, _w = lint_full(path, strict=True)
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
    errors, warns = lint_full(save(p, tmp_path, "fx.pptx"))
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
    errors, warns = lint_full(save(p, tmp_path, "fx.pptx"))
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
    errors, warns = lint_full(save(p, tmp_path, "fx.pptx"))
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
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    w6 = by_code(warns, "W6")
    assert w6, warns
    pages = [int(m) for m in re.findall(r"p(\d+)", w6[0][2])]
    assert pages and max(pages) <= 6, w6   # 이중 append 시절엔 p7~p11이 나왔다
    # 0.2.1: 토큰 수집 실패도 W18로 표면화된다(가드 배선 전체 확장)
    assert any("w10_tokens" in d for (_si, m, d) in by_code(warns, "W18")), warns


def test_e1_theme_token_resolution(tmp_path):
    """OOXML 테마 토큰(+mn-lt 류)을 실폰트로 해석(적대 패널 확정: 토큰을 문자 그대로
    블록리스트에 대조하면 E1 미탐). 기본 템플릿 테마 minor latin=Calibri(라틴 전용)라
    latin="+mn-lt" 한글 런은 E1이어야 한다."""
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
    errors, _w = lint_full(path)
    e2_pages = [si for (si, m, d) in by_code(errors, "E2")]
    assert e2_pages == [2], errors
    errors_s, _w = lint_full(path, strict=True)   # strict는 분리 여부 무관 전부 차단
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
    errors, warns = lint_full(save(p, tmp_path, "fx.pptx"))
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
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
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
    errors, _w = lint_full(save(p, tmp_path, "fx.pptx"))
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
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
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
    _e, warns = lint_full(path, w6_min_cluster=10)
    assert "W6" not in codes(warns)
    # 완전 클론은 코사인이 정확히 1.00이라 sim 임계는 1.0으로 줘야 배제된다(파라미터 관통 확인)
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
    errors, warns = lint_full(save(p, tmp_path, "fx.pptx"), render_dir=pages)  # NameError 나면 실패
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
    """숫자+단위 타이틀은 명사로 끝나도 주장형: W14가 오발화하지 않는다(장르 편향 완화)."""
    p = new_prs()
    for t in ("매출 3배 성장", "점유율 42% 달성", "비용 120억 절감", "가입자 500만 돌파",
              "마진 8pp 개선", "리텐션 2x 향상"):
        s = add_slide(p)
        tb(s, 1, 0.8, 9, 0.8, t, font="Wanted Sans", size=26)
        tb(s, 1, 2.2, 10, 3, "본문 내용", font="Wanted Sans", size=12)
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
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
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
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
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
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
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
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
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
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
    errors, warns = lint_full(save(p, tmp_path, "fx.pptx"))
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
    assert run_cli([ranged, "--profile", "full"]).returncode == 0
    r = run_cli([ranged, "--strict", "--profile", "full", "--json"])
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
    doc = json.loads(run_cli([path, "--json", "--profile", "full"]).stdout)
    assert any(w["code"] == "W14" for w in doc["warnings"])
    doc = json.loads(run_cli([path, "--json", "--profile", "full", "--skip", "W14"]).stdout)
    assert not any(w["code"] == "W14" for w in doc["warnings"])


# ---------------------------------------------------------------- 0.2.1 script layer + fixes
def patch_theme_fonts(path, major_ea=None, minor_ea=None):
    """테마 major/minor ea 빈 슬롯을 각각 다른 값으로 패치(zip 재작성)."""
    tmp = path + ".patched.pptx"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.search(r"theme/theme\d+\.xml$", item.filename):
                if major_ea is not None:
                    m = re.search(rb"<a:majorFont>.*?</a:majorFont>", data, re.S)
                    seg = m.group(0).replace(b'<a:ea typeface=""/>',
                                             ('<a:ea typeface="%s"/>' % major_ea).encode("utf-8"))
                    data = data[:m.start()] + seg + data[m.end():]
                if minor_ea is not None:
                    m = re.search(rb"<a:minorFont>.*?</a:minorFont>", data, re.S)
                    seg = m.group(0).replace(b'<a:ea typeface=""/>',
                                             ('<a:ea typeface="%s"/>' % minor_ea).encode("utf-8"))
                    data = data[:m.start()] + seg + data[m.end():]
            zout.writestr(item, data)
    shutil.move(tmp, path)
    return path


def test_title_flood_regression(tmp_path):
    """defaultTextStyle 폴백 18pt가 제목 후보로 수집돼 ghost/W14가 범람하던 회귀의 고정
    (2차 외부 재점검 확정). 크기 미지정 본문만 있는 덱의 ghost는 비어야 한다."""
    p = new_prs()
    s = add_slide(p)
    for i, txt in enumerate(("본문 첫 문장입니다", "본문 둘째 문장입니다", "표 안 텍스트 셋째")):
        tb(s, 1, 1 + i, 8, 0.5, txt, font="Wanted Sans", no_size=True)
    ghost = []
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"), ghost=ghost)
    assert ghost == [], ghost
    assert "W14" not in codes(warns)


def test_title_collection_keeps_placeholder_sizes(tmp_path):
    """placeholder 상속 크기(마스터 titleStyle 44pt)는 의도된 제목이라 계속 수집돼야 한다."""
    p = Presentation()
    s = p.slides.add_slide(p.slide_layouts[1])
    s.shapes.title.text_frame.text = "진짜 제목"
    s.placeholders[1].text_frame.text = "본문"
    ghost = []
    lint_full(save(p, tmp_path, "fx.pptx"), ghost=ghost)
    assert any("진짜 제목" in t for _si, t in ghost), ghost


def test_script_layer_two_tier(tmp_path):
    """스크립트 레이어 2층 구조(3차 패널: 한자 회귀와 JP 폰트 오탐 동시 해소).
    p1 가나+Noto Sans JP -> 침묵(가나 보유 폰트). p2 한글+Noto Sans JP -> E1(한글 없음).
    p3 가나+양수 트래킹 -> E4 없음(일본어 관행). p4 한자 전용+Georgia -> E1(CJK 전무 폰트).
    p5 한자 전용+양수 트래킹 -> E4(한국 덱 인명·법률용어). p6 가나+IBM Plex Mono -> E1."""
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
    errors, _w = lint_full(save(p, tmp_path, "fx.pptx"))
    e1_pages = sorted(si for (si, m, d) in by_code(errors, "E1"))
    assert e1_pages == [2, 4, 6], errors
    e4_pages = sorted(si for (si, m, d) in by_code(errors, "E4"))
    assert e4_pages == [5], errors


def test_hangul_range_extensions():
    """반각 한글·자모 확장 블록 미탐 봉합(3차 패널)."""
    assert jl.is_hangul(chr(0xFFA1))    # 반각 ㄱ
    assert jl.is_hangul(chr(0xA960))    # 자모 확장 A
    assert jl.is_hangul(chr(0xD7B0))    # 자모 확장 B
    assert jl._geometry_unsupported(chr(0x0F40))   # 티베트
    assert jl._geometry_unsupported(chr(0x1000))   # 미얀마
    assert jl._geometry_unsupported(chr(0x1780))   # 크메르


def test_title_own_lststyle_not_flooded(tmp_path):
    """도형 자체 lstStyle 20pt 본문 산문이 ghost에 쓸리지 않는다(3차 패널 실측 재현 봉합).
    제목 자격 = 명시 크기 또는 제목 패밀리 placeholder."""
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


def test_cli_lang_edge_cases(tmp_path):
    """--lang CLI 경계(3차 패널): 반복 지정은 마지막이 이김, skill 서브커맨드와 조합 안전."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "clean", font="Wanted Sans", size=20)
    path = save(p, tmp_path, "fx.pptx")
    doc = json.loads(run_cli([path, "--json", "--lang", "en", "--lang", "ko"]).stdout)
    assert doc["lang"] == "ko"
    r = run_cli(["skill", "--lang", "ko", "--path"])
    assert r.returncode == 0 and "SKILL.md" in r.stdout
    r = run_cli(["--lang", "ko", "skill", "--path"])   # 선행 플래그 + 서브커맨드
    assert r.returncode == 0 and "SKILL.md" in r.stdout


def test_summary_incomplete_flag(tmp_path):
    """summary.incomplete: 검사 불능(W18) 여부의 기계 판독 신호(3차 패널: 문서 과장 교정)."""
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


def test_latin_only_font_additions():
    f = jl.is_latin_only_font
    assert f("Aptos") and f("Courier") and f("PT Sans") and f("PT Serif")
    assert not f("Noto Sans Mono CJK KR")   # cjk 포함 = 한글 완비(2차 재점검 오탐 교정)
    assert not f("Source Han Sans CJK KR")


def test_e2_v2_adversarial_edges():
    """E2 v2 적대 패널 발견 4건의 회귀 고정(2026-07-10 3차)."""
    dv = jl.dash_violations
    MINUS = chr(0x2212)
    assert dv("전일대비 " + MINUS + " 3.2%") == []            # 띄어진 음수 부호(재무 표기)
    assert dv("끝 " + MINUS) == [MINUS]                        # 뒤에 숫자 없으면 여전히 차단
    assert dv("2020" + EN_DASH + EN_DASH + "2024") == [EN_DASH, EN_DASH]   # 2연속 대시 미탐 봉합
    assert dv("결론2024" + EN_DASH + "우리는") == [EN_DASH]    # 단어+숫자 혼합 토큰 우회 봉합
    assert dv("매출" + chr(0x00B9) + EN_DASH + "영업이익 증가") == [EN_DASH]   # 위첨자 각주
    assert dv("2020" + EN_DASH + "현재") == []                 # 숫자 시작 토큰 범위는 유지
    assert dv("Q1" + EN_DASH + "Q3") == []                     # 양쪽 숫자성 규칙 유지


def test_e2_v2_range_forms(tmp_path):
    """E2 v2: 잔존 오탐 4형태 통과 + 띄운 한쪽 숫자(AI 삽입구) 차단 유지."""
    dv = jl.dash_violations
    assert dv("2020 " + EN_DASH + " 2024") == []          # 띄운 숫자 범위
    assert dv("Q1" + EN_DASH + "Q3") == []                 # 알파숫자 토큰
    assert dv("5%" + EN_DASH + "10%") == []                # 퍼센트 범위
    assert dv("2020" + EN_DASH + "현재") == []             # 붙은 한쪽 숫자
    assert dv("성장 " + EN_DASH + " 2024년에는") == [EN_DASH]   # 띄운 삽입구는 차단
    assert dv("서울" + EN_DASH + "부산") == [EN_DASH]       # 단어 연결은 보수 차단
    assert dv("2020" + EN_DASH + "현재", strict=True) == [EN_DASH]
    # 통합: 덱 픽스처로도 고정
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 6, 0.5, "매출 5%" + EN_DASH + "10% 구간", font="Wanted Sans", size=14)
    s = add_slide(p)
    tb(s, 1, 1, 6, 0.5, "성장 " + EN_DASH + " 2024년에는 확대", font="Wanted Sans", size=14)
    errors, _w = lint_full(save(p, tmp_path, "fx.pptx"))
    assert [si for (si, m, d) in by_code(errors, "E2")] == [2], errors


def test_skip_rejects_error_codes(tmp_path):
    """--skip은 WARN 전용: E코드는 exit 2 거부, 적용된 skip은 JSON에 기록(풋건 교정)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    path = save(p, tmp_path, "fx.pptx")
    r = run_cli([path, "--skip", "E1"])
    assert r.returncode == 2 and "WARN" in r.stderr
    doc = json.loads(run_cli([path, "--json", "--profile", "full", "--skip", "W14"]).stdout)
    assert doc["summary"]["skipped_codes"] == ["W14"]
    assert any(e["code"] == "E1" for e in doc["errors"])   # E1은 여전히 살아있음


def test_e1_master_lststyle_font_inheritance(tmp_path):
    """프로브6 실측 반영: 마스터 ph lstStyle의 a:ea가 실효 렌더 폰트다.
    라틴 전용이면 E1, 한글 폰트면 통과('Malgun 폴백 확정' 오탐의 교정)."""
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
    """프로브6 Q1/Q2 실측 반영: 제목 ph는 테마 majorFont ea, 본문은 minorFont ea를 탄다."""
    p = Presentation()
    s = p.slides.add_slide(p.slide_layouts[1])
    s.shapes.title.text_frame.text = "제목 한글"
    s.placeholders[1].text_frame.text = "본문 한글"
    path = save(p, tmp_path, "fx.pptx")
    patch_theme_fonts(path, major_ea="Consolas", minor_ea="맑은 고딕")
    errors, _w = lint_full(path)
    e1 = by_code(errors, "E1")
    assert any("제목" in d for (_si, m, d) in e1), errors     # major=Consolas -> E1
    assert not any("본문" in d for (_si, m, d) in e1), errors  # minor=맑은 고딕 -> 통과


def test_vertical_and_complex_script_geometry_skip(tmp_path):
    """세로쓰기(bodyPr@vert)·복잡 조판 스크립트는 기하 추정 불가: 스킵하되 W18로 표면화."""
    p = new_prs()
    s = add_slide(p)   # p1: 세로쓰기 프레임이 화면 밖까지 -> W16 FP 없어야
    box = tb(s, 12.0, 1.0, 4.0, 0.6, "vertical long text overflowing edge",
             font="Wanted Sans", size=18)
    bodyPr = box.text_frame._txBody.find(qn("a:bodyPr"))
    bodyPr.set("vert", "eaVert")
    s = add_slide(p)   # p2: 아랍어 텍스트 -> 기하 스킵 + W18
    tb(s, 11.0, 1.0, 4.0, 0.6, "مرحبا بالعالم",
       font="Arial", size=18)
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    w16_text = [d for (_si, m, d) in by_code(warns, "W16") if "텍스트" in d]
    assert not w16_text, warns
    w18 = by_code(warns, "W18")
    assert any(si == 1 and "vertical_text" in d for (si, m, d) in w18), warns
    assert any(si == 2 and "complex_script" in d for (si, m, d) in w18), warns


# ---------------------------------------------------------------- 0.3.0 i18n + profiles
def test_lang_english_output(tmp_path):
    """0.3.0 i18n: 영어 환경에선 메시지가 영어, JSON에 lang 기록. 코드는 언어 무관."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    path = save(p, tmp_path, "fx.pptx")
    doc = json.loads(run_cli([path, "--json"], lang="en").stdout)
    assert doc["lang"] == "en"
    e1 = [e for e in doc["errors"] if e["code"] == "E1"][0]
    assert "Hangul" in e1["message"] and "Malgun" in e1["message"]
    doc_ko = json.loads(run_cli([path, "--json"], lang="ko").stdout)
    assert doc_ko["lang"] == "ko" and "한글" in [e for e in doc_ko["errors"] if e["code"] == "E1"][0]["message"]
    # --lang 플래그가 환경변수를 이긴다
    doc_flag = json.loads(run_cli([path, "--json", "--lang", "en"], lang="ko").stdout)
    assert doc_flag["lang"] == "en"


def test_lang_catalog_consistency():
    """카탈로그 규율: 모든 항목에 ko/en이 있고 % 포맷 지시자 순서가 동일하다."""
    fmt = re.compile(r"%[-#0-9.]*[sdfr%]")
    for mid, entry in jmsg.MESSAGES.items():
        assert set(entry) == {"ko", "en"}, mid
        assert fmt.findall(entry["ko"]) == fmt.findall(entry["en"]), mid


def test_profile_core_drops_style_rules(tmp_path):
    """core 프로파일 = 객관 결함만: E2(스타일성 ERROR)까지 제외되지만 선택이 JSON에 남는다."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "이건" + EM_DASH + "차단", font="Wanted Sans", size=12)   # E2만
    path = save(p, tmp_path, "fx.pptx")
    doc = json.loads(run_cli([path, "--json", "--profile", "full"]).stdout)
    assert any(e["code"] == "E2" for e in doc["errors"])           # full 명시 시 차단(0.4.0: 기본은 core)
    doc = json.loads(run_cli([path, "--json", "--profile", "core"]).stdout)
    assert not doc["errors"] and doc["summary"]["pass"]
    assert doc["summary"]["profile"] == "core"
    assert "E2" in doc["summary"]["skipped_codes"]                  # 조용한 우회 아님

    # core에서도 객관 결함(E1)은 그대로 차단
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    doc = json.loads(run_cli([save(p, tmp_path, "fx2.pptx"), "--json", "--profile", "core"]).stdout)
    assert any(e["code"] == "E1" for e in doc["errors"])


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


# ---------------------------------------------------------------- 0.3.1 (3차 외부 리뷰 P0/P1)
def test_e1_para_defrpr_inheritance(tmp_path):
    """P0-1(프로브7 실측): 문단 pPr/defRPr 폰트는 run 다음 순위로 상속되고 lstStyle을 이긴다.
    케이스A: 문단 ea=한글폰트 -> E1 오탐 없어야. 케이스B: 마스터 lstStyle ea=한글폰트인데
    문단 ea=Consolas -> 문단이 이기므로 E1 미탐 없어야."""
    def add_para_ea(para, ea):
        pPr = para._p.get_or_add_pPr()
        defR = pPr.makeelement(qn("a:defRPr"), {})
        defR.append(defR.makeelement(qn("a:ea"), {"typeface": ea}))
        pPr.append(defR)

    # 케이스A: 오탐 검증
    p = new_prs()
    s = add_slide(p)
    box = s.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(0.6))
    para = box.text_frame.paragraphs[0]
    add_para_ea(para, "맑은 고딕")
    r = para.add_run(); r.text = "문단 상속 한글"; r.font.size = Pt(14)
    errors, _w = lint_full(save(p, tmp_path, "a.pptx"))
    assert not by_code(errors, "E1"), errors

    # 케이스B: 미탐 검증(마스터 lstStyle 한글폰트 + 문단 Consolas)
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


def test_geometry_uses_style_resolver(tmp_path):
    """P0-2: 기하 검사가 E3와 동일한 상속 크기를 쓴다. 레이아웃 lstStyle 40pt 텍스트가
    우측을 뚫으면 W16(구버전은 기본 12pt로 계산해 미탐)."""
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
    body.text_frame.text = "Revenue ABC"   # 12pt면 0.95in(통과), 40pt면 3.2in(넘침)
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    assert any("텍스트" in d for (_si, m, d) in by_code(warns, "W16")), warns


def test_table_cell_geometry(tmp_path):
    """P0-3: 네이티브 표 셀 텍스트가 기하 검사에 포함된다(우측 이탈 표 -> W16)."""
    p = new_prs()
    s = add_slide(p)
    gf = s.shapes.add_table(1, 2, Inches(11.0), Inches(2.0), Inches(4.0), Inches(1.0))
    cell = gf.table.cell(0, 1)   # 오른쪽 열은 x=13in 부근에서 시작 -> 캔버스 밖
    cell.text = "overflowing table cell text"
    r = cell.text_frame.paragraphs[0].runs[0]
    r.font.size = Pt(24)
    cell.text_frame.word_wrap = False
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    assert any("텍스트" in d for (_si, m, d) in by_code(warns, "W16")), warns


def test_render_dir_contract(tmp_path):
    """P0-4: --render 폴더 부재·페이지 렌더 누락은 incomplete로 표면화, 그림 없는 덱은 무관."""
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
    assert doc["summary"]["incomplete"] is True, doc["summary"]   # 그림 있는 페이지 렌더 누락
    # 그림 없는 덱은 렌더가 없어도 완전(검사할 W7 대상 자체가 없음)
    p2 = new_prs()
    s2 = add_slide(p2)
    tb(s2, 1, 1, 5, 0.5, "text only", font="Wanted Sans", size=20)
    doc = json.loads(run_cli([save(p2, tmp_path, "t.pptx"), "--json", "--render", empty]).stdout)
    assert doc["summary"]["incomplete"] is False, doc["summary"]


def test_profile_is_engine_policy(tmp_path, monkeypatch):
    """P0-5: 프로파일은 실행 정책이다. 라이브러리 lint(profile=)로 쓸 수 있고,
    제외 규칙은 실행 자체가 안 되므로 그 내부 실패가 W18로 누출되지 않는다."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "이건" + EM_DASH + "차단", font="Wanted Sans", size=12)
    path = save(p, tmp_path, "fx.pptx")
    errors, _w = lint_full(path, profile="core")
    assert not by_code(errors, "E2"), errors            # 라이브러리 API에서 프로파일 동작
    errors, _w = lint_full(path)
    assert by_code(errors, "E2"), errors                # 기본 full은 차단 유지

    def boom(*a, **kw):
        raise RuntimeError("w9 fail")
    monkeypatch.setattr(jl, "accent_vbars_check", boom)
    _e, warns = lint_full(path, profile="core")           # W9 제외 -> 실행 안 함 -> W18 누출 없음
    assert not any("w9" in d for (_si, m, d) in by_code(warns, "W18")), warns
    _e, warns = lint_full(path)                           # full -> 실행 -> 가드 -> W18
    assert any("w9" in d for (_si, m, d) in by_code(warns, "W18")), warns


def test_skip_validation_strengthened(tmp_path):
    """P1: 존재하지 않는 W코드와 W18 억제는 exit 2 거부."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "clean", font="Wanted Sans", size=20)
    path = save(p, tmp_path, "fx.pptx")
    assert run_cli([path, "--skip", "W51"]).returncode == 2
    assert run_cli([path, "--skip", "W18"]).returncode == 2
    assert run_cli([path, "--skip", "W14"]).returncode == 0


def test_e4_hanja_message(tmp_path):
    """P1: E4 메시지가 실제 판정 범위(한글·한자)와 일치한다."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "大韓民國", font="맑은 고딕", size=12, spc=100)
    errors, _w = lint_full(save(p, tmp_path, "fx.pptx"))
    e4 = by_code(errors, "E4")
    assert e4 and "한자" in e4[0][1], e4


def test_w8_extended_hangul(tmp_path):
    """P1: is_cjk 통합으로 반각 한글도 W8 소형 CJK 판정에 잡힌다."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 2, 0.4, chr(0xFFA1) * 6, font="Wanted Sans", size=6)   # 반각 ㄱ x6
    _e, warns = lint_full(save(p, tmp_path, "fx.pptx"))
    assert by_code(warns, "W8"), warns


def test_ghost_prefers_title_placeholder(tmp_path):
    """제목 placeholder가 있으면 더 큰 KPI 빅넘버가 ghost 제목을 밀어내지 않는다(3차 리뷰)."""
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


def test_json_schema_contract(tmp_path):
    """JSON 버전 계약 시작(3차 리뷰): schema_version·tool·target_renderer."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "clean", font="Wanted Sans", size=20)
    doc = json.loads(run_cli([save(p, tmp_path, "fx.pptx"), "--json"]).stdout)
    assert doc["schema_version"] == "1.0"
    assert doc["tool"]["name"] == "archforge" and doc["tool"]["version"]
    assert doc["target_renderer"] == "powerpoint-windows"


def test_w6_detail_english(tmp_path):
    """P1: detail까지 i18n(영어 모드에서 W6 detail이 e.g.로 시작)."""
    p = new_prs()
    for i in range(6):
        s = add_slide(p)
        tb(s, 1, 0.8, 8, 0.6, "Title block %d" % i, font="Wanted Sans", size=24)
        tb(s, 1, 2.0, 6, 2.5, "Body block", font="Wanted Sans", size=12)
        tb(s, 8, 2.0, 4, 2.5, "Side block", font="Wanted Sans", size=12)
    doc = json.loads(run_cli([save(p, tmp_path, "fx.pptx"), "--json", "--profile", "full"], lang="en").stdout)
    w6 = [w for w in doc["warnings"] if w["code"] == "W6"]
    assert w6 and w6[0]["detail"].startswith("e.g."), w6


# ---------------------------------------------------------------- 0.4.0 구조 개편
def test_default_profile_is_core(tmp_path):
    """0.4.0 파괴적 변경: 무옵션 기본은 core(객관 결함만). 스타일 규칙 E2는 옵트인,
    객관 결함 E1은 기본에서도 차단."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "이건" + EM_DASH + "정상 통과", font="Wanted Sans", size=12)
    dash_deck = save(p, tmp_path, "dash.pptx")
    errors, _w = jl.lint(dash_deck)              # 라이브러리 기본값
    assert not by_code(errors, "E2"), errors
    r = run_cli([dash_deck, "--json"])           # CLI 기본값
    doc = json.loads(r.stdout)
    assert r.returncode == 0 and doc["summary"]["profile"] == "core"
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    assert run_cli([save(p, tmp_path, "e1.pptx")]).returncode == 1   # 객관 결함은 차단 유지


def test_config_file(tmp_path):
    """.archforge.json: 덱 폴더에서 자동 발견, CLI 플래그가 설정을 이긴다."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "이건" + EM_DASH + "차단", font="Wanted Sans", size=12)
    deck = save(p, tmp_path, "fx.pptx")
    with open(os.path.join(str(tmp_path), ".archforge.json"), "w", encoding="utf-8") as f:
        json.dump({"profile": "full"}, f)
    doc = json.loads(run_cli([deck, "--json"]).stdout)
    assert any(e["code"] == "E2" for e in doc["errors"])    # 설정의 full 적용
    doc = json.loads(run_cli([deck, "--json", "--profile", "core"]).stdout)
    assert not doc["errors"]                                 # CLI가 설정을 이김
    # fail-safe(4차 리뷰): 알 수 없는 키(오타)는 무시가 아니라 exit 2. 'profle: full'이
    # 조용히 기본 core로 실행되는 사고 방지
    with open(os.path.join(str(tmp_path), ".archforge.json"), "w", encoding="utf-8") as f:
        json.dump({"profile": "full", "no_such_key": 1}, f)
    r = run_cli([deck, "--json"])
    assert r.returncode == 2 and "no_such_key" in r.stderr
    # --no-config: 신뢰 불가 덱 폴더의 설정을 무시(신뢰 경계)
    with open(os.path.join(str(tmp_path), ".archforge.json"), "w", encoding="utf-8") as f:
        json.dump({"profile": "full"}, f)
    doc = json.loads(run_cli([deck, "--json", "--no-config"]).stdout)
    assert not doc["errors"] and doc["summary"]["config"] is None
    doc = json.loads(run_cli([deck, "--json"]).stdout)
    assert doc["summary"]["config"] and doc["summary"]["config"].endswith(".archforge.json")


def test_config_value_validation(tmp_path):
    """설정값 타입·범위는 traceback이 아니라 정돈된 exit 2(4차 리뷰). CLI 범위도 동일
    (--hard-min 0으로 E3를 조용히 끄던 구 X1 우회 봉합)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "clean", font="Wanted Sans", size=20)
    deck = save(p, tmp_path, "fx.pptx")
    with open(os.path.join(str(tmp_path), ".archforge.json"), "w", encoding="utf-8") as f:
        json.dump({"hard_min": "abc"}, f)
    r = run_cli([deck, "--json"])
    assert r.returncode == 2 and "hard_min" in r.stderr and "Traceback" not in r.stderr
    with open(os.path.join(str(tmp_path), ".archforge.json"), "w", encoding="utf-8") as f:
        json.dump({"lang": "fr"}, f)
    assert run_cli([deck, "--json"]).returncode == 2
    with open(os.path.join(str(tmp_path), ".archforge.json"), "w", encoding="utf-8") as f:
        json.dump({}, f)
    assert run_cli([deck, "--hard-min", "0"]).returncode == 2   # CLI 범위 검증


def test_baseline_v2_language_and_count(tmp_path):
    """지문 v2(4차 리뷰 HIGH 교정): ko로 만든 baseline이 en 실행에서도 유효하고,
    발생 수(count) 의미를 지키며, 페이지 무관이라 슬라이드 삽입에도 생존한다."""
    p = new_prs()
    for i in range(6):   # W6 클론 덱(detail이 로케일 문자열이던 대표 규칙)
        s = add_slide(p)
        tb(s, 1, 0.8, 8, 0.6, "Title block %d" % i, font="Wanted Sans", size=24)
        tb(s, 1, 2.0, 6, 2.5, "Body block", font="Wanted Sans", size=12)
        tb(s, 8, 2.0, 4, 2.5, "Side block", font="Wanted Sans", size=12)
    deck = save(p, tmp_path, "w6.pptx")
    bl = os.path.join(str(tmp_path), "bl.json")
    run_cli([deck, "--profile", "full", "--write-baseline", bl], lang="ko")
    doc = json.loads(run_cli([deck, "--profile", "full", "--json", "--baseline", bl],
                             lang="en").stdout)
    assert not any(w["code"] == "W6" for w in doc["warnings"]), doc["warnings"]

    # count 의미: 같은 지문 2건 발생, baseline엔 1건 -> 1건만 억제
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    one = save(p, tmp_path, "one.pptx")
    run_cli([one, "--write-baseline", bl])
    s = add_slide(p)   # 같은 텍스트·같은 폰트 = 같은 지문이 두 페이지에
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    two = save(p, tmp_path, "two.pptx")
    doc = json.loads(run_cli([two, "--json", "--baseline", bl]).stdout)
    assert doc["summary"]["baseline_suppressed"] == 1
    assert doc["summary"]["error_count"] == 1
    # 페이지 무관: 앞에 슬라이드를 끼운 덱에서도 억제 유지
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "새로 삽입된 표지", font="Wanted Sans", size=20)
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    shifted = save(p, tmp_path, "shifted.pptx")
    doc = json.loads(run_cli([shifted, "--json", "--baseline", bl]).stdout)
    assert doc["summary"]["baseline_suppressed"] == 1 and doc["summary"]["error_count"] == 0
    # 텍스트 모드 가시성: 억제가 각주로 표시(불가시 clean 오독 교정)
    r = run_cli([shifted, "--baseline", bl])
    assert "baseline" in r.stdout


def test_lint_rejects_unknown_profile(tmp_path):
    """라이브러리 API: 오타 프로파일이 조용히 full로 동작하던 것 교정(4차 리뷰)."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "clean", font="Wanted Sans", size=20)
    deck = save(p, tmp_path, "fx.pptx")
    with pytest.raises(ValueError):
        jl.lint(deck, profile="ful")


def test_baseline_flow(tmp_path):
    """baseline: 기존 위반 수용 후 신규만 보고. 억제 수는 summary에 기록."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    deck = save(p, tmp_path, "fx.pptx")
    bl = os.path.join(str(tmp_path), "baseline.json")
    r = run_cli([deck, "--write-baseline", bl])
    assert r.returncode == 0 and os.path.exists(bl)
    doc = json.loads(run_cli([deck, "--json", "--baseline", bl]).stdout)
    assert doc["summary"]["pass"] and doc["summary"]["baseline_suppressed"] == 1
    # 신규 결함 추가 -> 신규만 보고
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "새 결함 한글", font="Consolas", size=12)
    deck2 = save(p, tmp_path, "fx2.pptx")
    doc = json.loads(run_cli([deck2, "--json", "--baseline", bl]).stdout)
    assert doc["summary"]["error_count"] == 1
    assert "Consolas" in doc["errors"][0]["detail"]


def test_sarif_output(tmp_path):
    """SARIF 2.1.0 최소 계약: version·rules·results·ruleId·level."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    deck = save(p, tmp_path, "fx.pptx")
    out = os.path.join(str(tmp_path), "out.sarif")
    run_cli([deck, "--json", "--sarif", out])
    with open(out, encoding="utf-8") as f:
        doc = json.load(f)
    assert doc["version"] == "2.1.0"
    run0 = doc["runs"][0]
    assert run0["tool"]["driver"]["name"] == "archforge"
    res = run0["results"]
    assert res and res[0]["ruleId"] == "E1" and res[0]["level"] == "error"
    assert any(r["id"] == "E1" for r in run0["tool"]["driver"]["rules"])


def test_finding_location_payload(tmp_path):
    """구조화 위치(3차 리뷰): shape_id·bbox·part·paragraph·run이 JSON location에 실린다."""
    p = new_prs()
    s = add_slide(p)
    tb(s, 1, 1, 5, 0.5, "모노 폴백 한글", font="IBM Plex Mono", size=12)
    doc = json.loads(run_cli([save(p, tmp_path, "fx.pptx"), "--json"]).stdout)
    e1 = [e for e in doc["errors"] if e["code"] == "E1"][0]
    loc = e1["location"]
    assert isinstance(loc["shape_id"], int)
    assert loc["part"].endswith("slide1.xml")
    assert len(loc["bbox"]) == 4 and abs(loc["bbox"][0] - 1.0) < 0.01
    assert loc["paragraph"] == 0 and loc["run"] == 0


def test_finding_tuple_compat_and_lazy_message():
    """Finding의 4튜플 하위호환과 로케일 중립성(같은 finding을 두 언어로 렌더)."""
    from archforge.findings import Finding
    f = Finding(3, "E2", "e2", (), "cp=U+2014")
    page, code, msg, detail = f
    assert (page, code, detail) == (3, "E2", "cp=U+2014") and f[1] == "E2" and len(f) == 4
    jmsg.set_lang("ko")
    ko = f.message
    jmsg.set_lang("en")
    en = f.message
    jmsg.set_lang("ko")
    assert ko != en and "Dash" in en


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
