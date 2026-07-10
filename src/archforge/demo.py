# -*- coding: utf-8 -*-
"""데모 덱 생성기(0.5.0): 설치 30초 안에 린터가 무엇을 잡는지 보여주는 온보딩 자산.

`archforge demo`가 이 모듈로 broken.pptx(대표 결함 6종을 의도적으로 심은 덱)와
fixed.pptx(같은 내용의 교정본, full 프로파일 클린)를 만들어 즉석 린트한다.
리포 examples/ 의 커밋본도 scripts/make_examples.py 가 이 모듈로 생성한다.

심는 결함은 실제 생성 덱에서 가장 흔한 축만 고른다:
- E1 조용한 한글 폰트 폴백(Arial a:latin에 실린 한글, 테마 ea 빈 슬롯)
- E2 문장 부호로 쓴 em dash(AI 생성 덱 1번 티, full 프로파일)
- E3 판독 불가 크기(4pt 출처 표기)
- E4 연속 한글 양수 자간
- W15 텍스트 프레임 겹침 / W16 캔버스 밖 넘침
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

# 소스에 금지문자를 리터럴로 두지 않는다(tests와 같은 관례): 검사 대상 pptx 안에만 존재
EM_DASH = chr(0x2014)


def _prs():
    p = Presentation()
    p.slide_width = Inches(13.333)
    p.slide_height = Inches(7.5)
    return p


def _slide(p):
    return p.slides.add_slide(p.slide_layouts[6])


def _tb(slide, x, y, w, h, text, size=14, font=None, ea=None, spc=None,
        color="1A1A1A", bold=False):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    r = tf.paragraphs[0].add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = RGBColor.from_string(color)
    if font:
        r.font.name = font
    if ea:
        rPr = r._r.get_or_add_rPr()
        rPr.append(rPr.makeelement(qn("a:ea"), {"typeface": ea}))
    if spc is not None:
        rPr = r._r.get_or_add_rPr()
        rPr.set("spc", str(spc))
    return box


def build_broken(path):
    """대표 결함 6종을 심은 덱. full 프로파일에서 ERROR 4 + WARN(W15/W16)이 나온다."""
    p = _prs()
    s = _slide(p)   # p1: 타이포그래피 결함
    _tb(s, 0.8, 0.6, 8.0, 0.9, "3분기 실적 요약", size=26, font="Arial", bold=True)   # E1
    # E1은 타이틀 하나로 족하다: 나머지 run은 ea를 명시해 결함 하나당 코드 하나로 유지
    # (기본 Office 테마는 ea 빈 슬롯이라 ea 없는 한글 run은 전부 E1이 맞게 발화한다: 실측)
    _tb(s, 0.8, 1.8, 10.0, 0.6,
        "핵심 지표는 전 분기 대비 개선" + EM_DASH + "특히 구독 매출이 견인했습니다",
        size=14, ea="맑은 고딕")                                                      # E2
    _tb(s, 0.8, 2.8, 6.0, 0.5, "자간이 벌어진 한글 강조", size=16, spc=300,
        ea="맑은 고딕")                                                               # E4
    _tb(s, 0.8, 6.9, 5.0, 0.3, "출처: 사내 관리회계, 2026-06", size=4,
        ea="맑은 고딕")                                                               # E3
    s = _slide(p)   # p2: 기하 결함
    _tb(s, 0.8, 0.6, 8.0, 0.9, "분기 핵심 지표", size=26, ea="맑은 고딕", bold=True)
    _tb(s, 1.0, 2.4, 5.0, 1.0, "매출 성장률 +18%", size=24, ea="맑은 고딕")            # W15 쌍
    _tb(s, 1.2, 2.5, 5.0, 1.0, "영업이익률 12.4%", size=24, ea="맑은 고딕")
    _tb(s, 12.0, 4.5, 3.0, 0.6, "다음 분기 가이던스", size=18, ea="맑은 고딕")         # W16
    p.save(path)
    return path


def build_fixed(path):
    """broken과 같은 내용의 교정본. full 프로파일에서 ERROR 0, WARN 0."""
    p = _prs()
    s = _slide(p)
    # E1 교정: a:latin은 그대로 두고 a:ea에 한글 폰트를 명시(가장 견고한 수정)
    _tb(s, 0.8, 0.6, 8.0, 0.9, "3분기 실적 요약", size=26, font="Arial",
        ea="맑은 고딕", bold=True)
    # E2 교정: 산문 대시는 콜론·쉼표·괄호로
    _tb(s, 0.8, 1.8, 10.0, 0.6,
        "핵심 지표는 전 분기 대비 개선: 특히 구독 매출이 견인했습니다",
        size=14, ea="맑은 고딕")
    # E4 교정: 한글 run 자간 0
    _tb(s, 0.8, 2.8, 6.0, 0.5, "자간을 되돌린 한글 강조", size=16, ea="맑은 고딕")
    # E3 교정: 출처·캡션도 최소 9pt
    _tb(s, 0.8, 6.9, 5.0, 0.3, "출처: 사내 관리회계, 2026-06", size=9, ea="맑은 고딕")
    s = _slide(p)
    _tb(s, 0.8, 0.6, 8.0, 0.9, "분기 핵심 지표", size=26, ea="맑은 고딕", bold=True)
    _tb(s, 1.0, 2.4, 5.0, 1.0, "매출 성장률 +18%", size=24, ea="맑은 고딕")
    _tb(s, 6.8, 2.4, 5.0, 1.0, "영업이익률 12.4%", size=24, ea="맑은 고딕")
    _tb(s, 9.2, 4.5, 3.4, 0.6, "다음 분기 가이던스", size=18, ea="맑은 고딕")
    p.save(path)
    return path


def build_warnings(path):
    """core 프로파일은 클린, full 프로파일에서만 스타일 WARN이 나오는 덱(examples/용):
    W13 네이티브 그림자, W14 명사구 타이틀 나열. 프로파일 분리를 가르치는 교보재."""
    p = _prs()
    for i, title in enumerate(("시장 현황", "경쟁 구도 분석", "성장 전략 로드맵")):
        s = _slide(p)
        _tb(s, 0.8, 0.6, 8.0, 0.9, title, size=24, ea="맑은 고딕", bold=True)
        _tb(s, 0.8, 2.0, 10.0, 0.6,
            "본문 요지는 페이지마다 다르게 구성했습니다 (%d)" % (i + 1),
            size=14, ea="맑은 고딕")
        from pptx.enum.shapes import MSO_SHAPE
        # W13은 같은 페이지에 실효 효과 2개+일 때 집계된다(빈 effectLst 오탐 방지 설계)
        for j in range(2):
            sp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                    Inches(0.8 + i * 0.7 + j * 3.2), Inches(4.0 + i * 0.4),
                                    Inches(2.5), Inches(1.0))
            sp.fill.solid()
            sp.fill.fore_color.rgb = RGBColor.from_string("DDE3EA")
            sp.line.fill.background()
            spPr = sp._element.spPr
            eff = spPr.makeelement(qn("a:effectLst"), {})
            sh = spPr.makeelement(qn("a:outerShdw"),
                                  {"blurRad": "40000", "dist": "20000"})
            eff.append(sh)
            spPr.append(eff)
    p.save(path)
    return path
