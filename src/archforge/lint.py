# -*- coding: utf-8 -*-
"""
Archforge(아치포지): 빌드된 .pptx를 받아 반복 결함을 기계로 차단하는 한글 특화 품질 린터.

존재 이유: 규칙을 문서에 적어둬도 만들 때마다 같은 실수가 반복된다. 기계로 검사 가능한 결함은
사람 눈이 아니라 이 린터가 막는다. 한글 덱이 조용히 깨지는 지점(라틴 전용 폰트로의 폴백, CJK 자간,
소형 글자)과 AI 생성 덱 특유의 티(긴 대시, 상투 카피, 레이아웃 재탕)를 빌드 산출물에서 잡는다.
임의 pptx를 받는 최후 방어선이라 상속·슬롯·autofit·표·그룹 같은 우회 경로까지 덮는다(적대 감사 반영).

ERROR = 배포 차단:
  E1 한글 텍스트의 실효 렌더 폰트가 라틴 전용(한글 글리프 없음) = 조용한 Malgun 폴백.
     실효 폰트는 COM 렌더 실측 모델(2026-07-10, docs/CALIBRATION.md)로 해석한다:
     run a:ea > lstStyle 상속 체인(도형→레이아웃→마스터, 프로브6 실측) > 테마 ea
     (제목 패밀리는 majorFont, 그 외 minorFont; 비어있지 않으면 run a:latin보다 우선)
     > (테마 ea 빈 슬롯일 때만) run a:latin > OS 폴백(Malgun).
     트리거는 한글 한정: 가나·한자 전용 런은 한글 커버리지 지식으로 판정하지 않는다(0.2.1).
  E2 대시류 문자가 렌더 텍스트에 포함(em/en/figure dash + 수학 마이너스 U+2212·전각 하이픈 U+FF0D·박스 罫線 U+2500 등).
     판별 축은 기능: en dash는 양옆 토큰이 숫자성이면(범위) 통과, 한쪽만 숫자성이면
     붙었을 때만 통과, 띄운 삽입구·단어 연결은 차단. U+2212는 숫자 앞 통과.
     문맥은 run이 아니라 문단 전체로 본다(run 분할 오탐 방지). --strict 는 전부 차단.
  E3 실효 폰트(autofit fontScale·문단 상속·placeholder 상속 체인 반영)가 HARD_MIN(기본 5.0pt) 미만 = 판독 불가
  E4 연속 한글(2자 이상)에 유의미 양수 트래킹(spc>50 = 0.5pt) = 자간 벌어짐.
     한글 한정: 가나 자간은 일본어 디자인의 정상 관행이라 결함 단정 불가(0.2.1).
WARN = 참고(배포는 통과):
  W1 본문급 프레임(넓고 글자수 많음)의 실효 폰트가 BODY_MIN(기본 9.0pt) 미만 = 본문이 작다
  W8 좁은 프레임의 소형 CJK(실효 [HARD_MIN, SMALL_MIN=7.5)pt) = 기기 목업·카드 내부 텍스트 판독 위험
     (E3 판독불가 하한 위, W1 본문급 아래의 회색지대. 폰 목업 안 서브텍스트를 프리플라이트에 가시화)
  W6 레이아웃 골격 재탕: 도형 bbox 시그니처 코사인>0.90 클러스터 4장+(렌더 불요, 항상 검사)
  W7 이미지 위 텍스트 저대비 ratio<2.5 (--render <pages> 지정 시에만, 렌더 PNG 픽셀 대조)
  W9 accent 세로바 3개+를 리스트 마커로 반복 = 색으로 구조 잡기(클로드 티). 커넥터·제로폭 커버
  W10 직접 그린 도식(단면도 등 장식 텍스처)이 여러 페이지에서 거의 동일 반복 = 재탕인지 의도인지 눈 확정
  W11 AI 티 카피: 버즈워드(좁은 사전, 전 페이지)·뻔한 오프닝 상투구(p1~3만)
  W12 푸터 baseline 어긋남: 지배 baseline(표지 제외, 0.05in 버킷 중앙값) 대비 0.03~0.25in 벗어난 페이지
  W13 PPT 자체 그림자·글로·3D 효과 2개+(올드 티. 자식 없는 빈 effectLst 는 안 셈)
  W14 타이틀 다수(3+, 절반+)가 서술형 명사구 = 액션타이틀 아님. --ghost 로 고스트덱(타이틀 나열)도 출력
  W15 텍스트끼리 겹침 추정(실효 글리프 bbox 근사, 교집>작은쪽 45%) = 가림·충돌, 페이지당 최대 2건
  W16 화면 밖 넘침: 텍스트=실효 글리프 box 경계 밖 0.15in+, 그림=잉크 bbox(알파 트림) 0.12in+
      (풀블리드 70%+ 제외). 장식 도형의 모서리 블리드는 표준 기법이라 검사 대상 아님(렌더 실측 기각)
  W17 텍스트가 비배경 그림의 잉크 경계에 걸침(글리프의 25~75%만 안) = 잘려 보임. 완전 위는 W7 소관
  W5 폰트 크기를 run·문단·상속 체인 어디에서도 못 찾음(체인 전체 침묵 시에만)
  W18 손상·비정형 속성으로 일부 구간 검사 불능: 결과 불완전 가능(가드가 삼킨 구간을
      stderr만이 아니라 출력 계약에 표면화, --strict에선 exit 1로 승격)

W15~W17 기하 강건성(적대 검증 12건 실측 재현 후 수정, 2026-07-03): 그룹 off/chOff 아핀 변환,
wrap=none(word_wrap=False, python-pptx add_textbox 기본) 한 줄 실폭, autofit 퍼센트 문자열
('62.5%')·lnSpcReduction, 문단별 정렬, 빈 문단 endParaRPr 크기, 회전 도형(축정렬 확장 bbox,
회전 텍스트는 스킵), 그림 srcRect 크롭·flipH/V·P모드 tRNS 투명, W17은 사진-캡션 사이 솔리드
카드(z순서)를 억제. 남은 한계: 플레이스홀더의 레이아웃 lstStyle 정렬 상속은 좌정렬로 후퇴.

사용: archforge <built.pptx> [--hard-min 5.0] [--body-min 9.0] [--strict] [--render <pages>]
              [--ghost] [--json] [--skip CODES] [--w6-sim 0.90] [--w6-cluster 3]
  --strict: WARN도 exit 1 + E2 숫자 맥락 예외 해제. --render: W7 활성화.
  --ghost: 타이틀 나열(수평 논리 눈검수). --skip: 지정 코드 억제(예 --skip W14,W6: 에디토리얼 덱).
반환: ERROR가 하나라도 있으면 exit 1, 아니면 0.

서브커맨드: archforge skill [--install [DIR]] [--path]
  동봉된 에이전트 스킬팩(SKILL.md)을 출력하거나 DIR(기본 ./.claude/skills)에 설치한다.
"""
import os
import re
import sys
import glob
import argparse
import colorsys
from collections import Counter, namedtuple
from typing import Dict, List, Optional, Tuple

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

try:
    from .messages import M, set_lang, get_lang
    from .findings import Finding, shape_loc
    from .rules import RULES, ALL_CODES, PROFILES, DEFAULT_PROFILE
    from . import config as _config
    from . import reporters as _reporters
except ImportError:   # 파일 단독 실행(python lint.py) 폴백
    from messages import M, set_lang, get_lang
    from findings import Finding, shape_loc
    from rules import RULES, ALL_CODES, PROFILES, DEFAULT_PROFILE
    import config as _config
    import reporters as _reporters

EMU_PER_IN = 914400
NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
NS_P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

# 한글 글리프가 없어 한글에 지정돼도 적용되지 못하는 라틴 전용 폰트(접두 매칭이라 굵기 변형도 커버).
# 블록리스트 방식의 한계(여기 없는 폰트는 미탐)를 좁히려고 덱에서 흔한 라틴 패밀리를 폭넓게 등재
# (2026-07-10 외부 리뷰의 E1 미탐 지적 반영: Inter·Arial·Calibri 류가 빠져 있었다).
LATIN_ONLY_FONTS = (
    # 모노스페이스
    "ibm plex mono", "courier", "consolas", "cascadia mono", "cascadia code",
    "roboto mono", "jetbrains mono", "fira code", "source code pro", "pt mono",
    "dejavu sans mono", "dejavu sans", "menlo", "monaco", "sf mono", "space mono",
    # 산세리프(웹·디자인 덱 단골)
    "switzer", "inter", "arial", "helvetica", "calibri", "aptos", "segoe ui", "roboto",
    "lato", "montserrat", "poppins", "open sans", "source sans", "raleway",
    "nunito", "work sans", "dm sans", "manrope", "karla", "figtree", "sora",
    "outfit", "plus jakarta", "barlow", "oswald", "rubik", "mulish",
    "tahoma", "verdana", "trebuchet", "century gothic", "franklin gothic",
    "gill sans", "futura", "candara", "corbel", "ibm plex sans", "noto sans",
    "pt sans", "pt serif",
    # 세리프
    "times new roman", "georgia", "garamond", "palatino", "cambria", "constantia",
    "playfair", "merriweather", "lora", "libre baskerville", "crimson", "noto serif",
)

# 위 접두에 걸리지만 실제로는 한글 글리프를 갖춘 예외(먼저 검사해 오탐 방지).
# 예: "IBM Plex Sans KR"은 "ibm plex sans" 접두에 걸리지만 한글 완비 폰트다.
KOREAN_CAPABLE_EXCEPTIONS = (
    "arial unicode", "ibm plex sans kr", "noto sans kr", "noto sans cjk",
    "noto serif kr", "noto serif cjk",
)

# 블록리스트 접두에 걸리지만 가나·한자는 갖춘 JP/SC/TC 서브셋 폰트: 한글 판정에선 라틴
# 전용 취급(한글 글리프 없음), 가나·한자 판정에선 통과. 이 두 층을 구분해야 한자 텍스트의
# 진짜 폴백(Inter·모노 계열=CJK 전무)은 잡으면서 JP 폰트 오탐이 없다(3차 적대 패널).
JP_CN_CAPABLE_PREFIXES = (
    "noto sans jp", "noto sans sc", "noto sans tc", "noto sans hk",
    "noto serif jp", "noto serif sc", "noto serif tc", "ibm plex sans jp",
)

# 긴 대시류 + 전각 하이픈마이너스 + 2·3배 em + 박스 수평선(코드포인트로 구성: 소스에 대시 문자 없음)
LONG_DASHES = {chr(c) for c in (0x2012, 0x2013, 0x2014, 0x2015, 0x2212, 0xFF0D, 0x2E3A, 0x2E3B, 0x2500)}
_EN_DASH = chr(0x2013)
_MINUS = chr(0x2212)

# 규칙 메타데이터·프로파일·전체 코드 목록은 rules.py 레지스트리로 이관(0.4.0 구조 분해).
# PROFILES/ALL_CODES는 위 import로 이 모듈에서 계속 노출된다(하위호환).


def is_cjk(ch):
    """한글(확장 포함)·가나·한자. 스크립트 판정 단일화(3차 리뷰 P1: is_hangul만 확장해
    W8 등 has_cjk 소비자가 확장 한글을 놓치던 불일치의 교정)."""
    return is_hangul(ch) or is_kana(ch) or is_hanja(ch)


def has_cjk(text):
    return any(is_cjk(c) for c in text)


def is_hangul(ch):
    o = ord(ch)
    return (0xAC00 <= o <= 0xD7A3      # 음절
            or 0x1100 <= o <= 0x11FF   # 자모
            or 0x3130 <= o <= 0x318F   # 호환 자모
            or 0xA960 <= o <= 0xA97F   # 자모 확장 A(적대 패널: 옛한글 조합 미탐)
            or 0xD7B0 <= o <= 0xD7FF   # 자모 확장 B
            or 0xFFA0 <= o <= 0xFFDC)  # 반각 한글(레거시 EDI·OCR 산출물)


def has_hangul(text):
    return any(is_hangul(c) for c in text)


def is_kana(ch):
    return 0x3040 <= ord(ch) <= 0x30FF


def is_hanja(ch):
    o = ord(ch)
    return 0x4E00 <= o <= 0x9FFF or 0x3400 <= o <= 0x4DBF


def _geometry_unsupported(text):
    """RTL(아랍·히브리)·복잡 조판 스크립트(인도계·티베트·미얀마·태국·라오·크메르) 포함 여부.
    글자폭 근사표가 라틴/CJK 이분법이라 이 스크립트들의 기하 추정은 무의미하다(0.2.1
    스크립트 레이어): 추정하지 말고 스킵 후 W18로 정직하게 알린다.
    범위는 0x0900~0x109F(인도계~미얀마)+0x1780~0x17FF(크메르)까지(적대 패널 보강)."""
    for c in text:
        o = ord(c)
        if 0x0590 <= o <= 0x08FF or 0x0900 <= o <= 0x109F or 0x1780 <= o <= 0x17FF:
            return True
    return False


def run_fonts(run):
    """run rPr의 a:latin/a:ea/a:cs typeface. run.font.name은 a:latin만 보므로 XML을 직접 읽는다."""
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
    """ST_TextPoint 유니언: '150'(1/100pt 정수)과 '1.5pt' 같은 universal measure 둘 다
    스키마 유효(ECMA-376). int() 단독은 후자에서 ValueError로 죽는다: _pct_attr과 같은
    유니언 함정(적대 검증 실측 2026-07-03)의 spc 판. 파싱 불가한 쓰레기값은 None."""
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
    """문단 pPr/defRPr의 latin/ea typeface. OOXML에서 문단 기본 런 속성으로, run rPr
    바로 다음 순위다(COM 프로브7 실측 2026-07-10: 문단 defRPr ea가 실제로 렌더되고
    lstStyle 체인을 이긴다. 3차 외부 리뷰 P0-1: 이 단계 누락이 E1 오탐·미탐 동시 유발)."""
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
    """run 의 트래킹(spc, 1/100 pt 단위). 없거나 파싱 불가면 None."""
    rPr = run._r.find(NS + "rPr")
    if rPr is None:
        return None
    return _text_point_attr(rPr.get("spc"))


class _FldRun:
    """a:fld(슬라이드 번호·날짜 자동 필드) 어댑터. 스키마상 CT_TextField도 rPr+t 구조라
    PowerPoint는 일반 run과 같은 규칙으로 렌더하는데, python-pptx para.runs가 a:r만
    돌려줘 필드 텍스트가 E1/E3/E4의 사각지대였다(4차 리뷰 이월분, 0.5.0). 검사 코드가
    쓰는 최소 인터페이스만 노출한다: run_fonts/run_track용 ._r, 크기용 .font.size."""
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
        return self   # .size 접근만 쓰인다

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
    """해당 스크립트 글리프가 없는 폰트인가(접두 매칭, 완비 예외 우선).
    script="hangul": 한글 글리프 기준(기본). script="cjk_other": 가나·한자 기준으로,
    JP/SC 서브셋 폰트(가나·한자 보유)는 통과시킨다. 블록리스트 본대(Inter·Arial·모노
    계열)는 CJK 전체가 없어 두 기준 모두 유효하다(3차 적대 패널: 한자 회귀 교정).
    이름에 cjk가 들어간 폰트(Noto/Source Han 계열 슈퍼폰트)는 무조건 통과."""
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
    """E1 판정: CJK 텍스트의 실효 렌더 폰트가 라틴 전용이거나 결정 불능(Malgun 폴백)인가.

    실측 렌더 모델(PowerPoint COM 프로브 2026-07-10, docs/CALIBRATION.md):
      1) run a:ea가 있으면 그 폰트가 한글을 그린다.
      2) 없으면 테마 minorFont a:ea. 비어있지 않으면 run a:latin보다 우선한다.
      3) 테마 ea가 빈 슬롯일 때만 run a:latin이 한글을 그린다(한글 글리프가 있으면).
      4) 그마저 없거나 라틴 전용이면 OS 폴백(Windows Malgun).
    반환: (message, detail) 또는 None. 외부 리뷰(2026-07-10)의 두 지적을 함께 반영:
    ea-or-latin 대용이 만들던 미탐과, ea 단독 판정이 만들 뻔한 합법 패턴
    (font.name=한글폰트, 빈 테마 ea에서 실제로 렌더됨) 오탐을 실측 모델로 동시에 푼다."""
    if script != "hangul":
        # 가나·한자 모드(3차 적대 패널: 한자 회귀 교정): 같은 슬롯 해석을 쓰되,
        # 명시된 폰트가 해당 글리프 없음일 때만 발화한다. 빈 슬롯·미지정은 침묵
        # (비한글 스크립트의 OS 폴백 지형은 미실측이라 단정하지 않는다).
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
        return None   # 비어있지 않은 한글 테마 ea가 렌더를 받는다(run latin은 못 이김: 실측)
    run_latin = fonts.get("latin")
    if run_latin:
        if is_latin_only_font(run_latin):
            return ("e1_latin_empty_theme", "font=%r text=%r" % (run_latin, text[:24]))
        return None   # 빈 테마 ea에선 한글 지원 latin이 실제로 한글을 그린다(실측)
    return ("e1_nofont", "text=%r" % text[:24])


def _is_digit_ch(c: str) -> bool:
    """숫자 판정은 ASCII·전각 숫자만: isdigit()은 위첨자 각주(U+00B9)·원문자(U+2460)까지
    True라 '매출¹' 같은 각주 달린 단어가 숫자성으로 오인돼 예외를 뚫었다(적대 패널 실측)."""
    return "0" <= c <= "9" or 0xFF10 <= ord(c) <= 0xFF19


def _dash_neighbor(text: str, i: int, step: int) -> Tuple[str, bool]:
    """i 위치 대시의 이웃 토큰(최대 12자)과 공백 유무. step=-1 왼쪽, +1 오른쪽."""
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
    # 왼쪽 스캔(step=-1)은 역순으로 쌓이므로 뒤집어 원문 순서로: 숫자-선두 판정이
    # 대시 인접 문자가 아니라 토큰의 실제 첫 글자를 보게 한다(적대 패널 실측 교정)
    tok = "".join(reversed(buf)) if step < 0 else "".join(buf)
    return tok, spaced


def dash_violations(text: str, strict: bool = False,
                    span: Optional[Tuple[int, int]] = None) -> List[str]:
    """E2 위반 문자 추출. 판별 축은 문자가 아니라 기능이다: 범위 연결(정당 타이포)과
    문장 부호(AI 삽입구 티)를 이웃 토큰의 숫자성과 붙임 여부로 가른다(0.2.1 v2,
    2차 외부 재점검의 잔존 오탐 4형태 교정).

    en dash(U+2013):
      - 양옆 토큰이 둘 다 숫자성(숫자 포함: 2020, Q1, 5%, FY24) -> 통과(띄어도 범위)
      - 한쪽만 숫자성 + 대시가 붙어 있음 -> 통과(2020~현재식 범위)
      - 한쪽만 숫자성 + 띄어 있음 -> 차단("성장 - 2024년에는"식 AI 삽입구)
      - 단어~단어 -> 차단(삽입구와 기계 구분 불가, 보수 선택. Seoul~Busan류가 잔여 오탐)
    수학 마이너스(U+2212): 바로 뒤 숫자면 통과(음수·산식).
    em dash 등 나머지 대시류: 전 맥락 차단(간판 기능). strict=True면 예외 없이 전부 차단.

    span=(start,end)를 주면 보고는 그 구간의 문자만 하되 이웃 문맥은 text 전체에서 본다.
    호출자가 text=문단 전체, span=run 구간으로 부르면 run 경계로 쪼개진 범위가 오탐되지
    않는다(적대 패널 확정 2026-07-10)."""
    lo, hi = span if span is not None else (0, len(text))
    bad = []
    for i in range(lo, min(hi, len(text))):
        c = text[i]
        if c not in LONG_DASHES:
            continue
        if not strict:
            # 연속 대시(전후 인접 문자가 또 대시)는 어떤 표기 관습에도 없다: 무조건 차단.
            # '2020--2024'처럼 정확히 2연속일 때 양쪽 대시가 서로를 이웃으로 만나 토큰이
            # 비면서 한쪽-숫자 예외를 뚫던 구멍의 봉합(적대 패널 실측).
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
                # 한쪽-숫자 예외는 숫자로 시작하는 토큰만: '결론2024' 같은 단어+숫자
                # 혼합 토큰이 붙은 삽입구를 통과시키던 우회의 봉합(적대 패널 실측)
                l_lead = bool(lt) and _is_digit_ch(lt[0])
                r_lead = bool(rt) and _is_digit_ch(rt[0])
                if (l_lead or r_lead) and not lsp and not rsp:
                    continue
            elif c == _MINUS:
                # 공백 허용: '− 3.2%'처럼 부호가 띄어진 재무 표기 오탐 교정(적대 패널)
                j = i + 1
                while j < len(text) and text[j].isspace():
                    j += 1
                if j < len(text) and _is_digit_ch(text[j]):
                    continue
        bad.append(c)
    return bad


def _pct_attr(v, default):
    """OOXML 퍼센트 유니언 타입: '62500'(1/1000 %)과 '62.5%'(문자열형) 둘 다 유효
    (ST_TextFontScalePercentOrPercentString). int() 단독은 후자에서 ValueError로 죽어
    blanket except가 1.0으로 삼켰다(적대 검증 실측 2026-07-03)."""
    if not v:
        return default
    v = v.strip()
    if v.endswith("%"):
        return float(v[:-1]) / 100.0
    return int(v) / 100000.0


def frame_autofit(tf):
    """(fontScale, lnSpcReduction) 비율 쌍. normAutofit 없으면 (1.0, 0.0)."""
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
    """텍스트프레임 autofit fontScale(비율). E3 등 기존 소비자 호환 유지."""
    return frame_autofit(tf)[0]


def _group_xf(sp, xf):
    """그룹 도형의 off/ext vs chOff/chExt 아핀을 부모 xf에 합성. 계수 (ax,bx,ay,by):
    abs = a*raw + b (EMU). 파싱 실패 시 부모 xf 그대로(항등 후퇴).

    grpSpPr는 슬라이드에선 p: 네임스페이스다(내부 xfrm·off류만 a:). 종전 코드가
    a:grpSpPr로 find해 항상 None → 항등 후퇴였던 잠복 버그의 교정(0.5.0 실측:
    이동 desync 그룹의 loc bbox가 raw로 나오는 재현으로 발견)."""
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
    """(text_frame, width_emu, owner_shape, cell_rc, xf) 리스트. 그룹 재귀 + 네이티브 표
    셀 포함. cell_rc는 표 셀이면 (행,열) 0기반, 아니면 None. xf는 그룹 절대좌표 아핀
    (iter_shapes_geo와 동일 계수)로, run 단위 finding의 loc bbox를 그룹 chOff 좌표계가
    아닌 슬라이드 실좌표로 만드는 데 쓴다(4차 리뷰 이월분, 0.5.0)."""
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
            for ri, row in enumerate(tbl.rows):
                for ci, cell in enumerate(row.cells):
                    out.append((cell.text_frame, (sp.width or 0) // ncol, sp, (ri, ci), xf))
            continue
        if sp.has_text_frame:
            out.append((sp.text_frame, sp.width or 0, sp, None, xf))
    return out


def iter_shapes(shapes):
    """그룹 재귀로 모든 도형을 평탄하게 순회."""
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


# W15~W17 기하 소비자용 절대좌표 순회. 그룹 자식의 raw left/top은 그룹 chOff 좌표계라
# 그룹이 이동·리사이즈된 pptx(off!=chOff desync, PowerPoint 드래그 표준 동작)에서
# 슬라이드 좌표와 어긋난다(적대 검증 실측 2026-07-03). off/ext vs chOff/chExt 아핀을
# 합성해 (도형, z순서, 절대 xfrm 함수 계수)로 순회한다. xf=(ax,bx,ay,by): abs = a*raw + b (EMU).
def iter_shapes_geo(shapes, xf=(1.0, 0.0, 1.0, 0.0), _z=None):
    if _z is None:
        _z = [0]
    for sp in shapes:
        try:
            is_grp = sp.shape_type == MSO_SHAPE_TYPE.GROUP
        except Exception:
            is_grp = False
        if is_grp:
            # 그룹 내부 좌표 g: abs_g = ox + (g - cx)*sx 를 부모 xf에 합성(_group_xf)
            for inner in iter_shapes_geo(sp.shapes, _group_xf(sp, xf), _z):
                yield inner
            continue
        z = _z[0]
        _z[0] += 1
        yield sp, z, xf


def _geo_rect(sp, xf):
    """xf 적용한 절대 bbox(in). 회전 도형은 중심 고정 축정렬 bbox로 확장. 실패 시 None."""
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
    """단색 채움 hex(6자리 대문자) 또는 None. python-pptx 접근자라 네임스페이스·커넥터 내부를 알아서 처리."""
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
    """선/윤곽 색 hex 또는 None. 커넥터(cxnSp)의 line 색까지 포함(실덱의 커넥터 세로바가 이 경로로 통과했던 실측)."""
    try:
        c = sp.line.color
        if c.type is not None and c.rgb is not None:
            return str(c.rgb).upper()
    except Exception:
        pass
    return None


def _is_accent(hexc):
    """의미 강조색 판별: HSV 채도>=0.55 그리고 명도 0.18~0.78. 배경·괘선·본문·저채도 보조색은 배제."""
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
    """W9: accent 세로바를 리스트 마커로 반복해 항목 구조를 색으로 잡는 AI 생성물 티(실덱 실측).
    적대 감사(2026-07-02) 반영: 색은 sp.line/sp.fill 접근자로 읽어 네임스페이스·커넥터를 커버,
    세로바는 제로폭 커넥터(w~0)를 명시 포함, 오른쪽 인접 텍스트로 리스트 마커임을 확증."""
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
        if 0.2 <= h <= 1.0 and (w < 0.05 or h > 3 * w):   # 세로 막대(제로폭 커넥터 포함)
            bars.append((x, y, w, h, hexc))
    if len(bars) < 3:
        return
    hues = {b[4] for b in bars}
    if len(hues) != 1:                       # 다색은 데이터 인코딩(범례)이라 정당
        return
    xs = [b[0] for b in bars]
    if max(xs) - min(xs) > 0.15:             # 세로 정렬 스택만(가로 분산=차트·구분선 제외)
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
    """슬라이드의 단색 채움 도형을 (색, 24분할 위치·크기) 토큰 멀티셋으로. 풀블리드 배경은 제외."""
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
    """테마 XML의 major/minor 폰트 슬롯 4종: {"mn-ea","mn-lt","mj-ea","mj-lt"}.
    XML 파싱이라 따옴표 직렬화·속성 순서에 무관(byte 정규식이던 시절의 취약점 교정,
    외부 리뷰 2026-07-10). run rPr의 "+mn-lt" 류 테마 토큰 해석에도 쓴다(적대 패널 확정:
    토큰을 실폰트명으로 취급하면 E1 미탐). 파싱 실패는 None."""
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
    """슬라이드마스터별 테마 폰트 슬롯 맵. 키=마스터 partname 문자열, 값=None이면 파싱 실패.
    멀티마스터 덱에서 iter_parts 첫 테마를 잡아 콘텐츠와 무관한 마스터의 빈 슬롯으로
    E1을 오발화하던 것의 교정: 마스터→테마 관계(rels)로 해석한다(외부 리뷰 2026-07-10)."""
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
    """하위호환: 마스터별 minorFont a:ea만 뽑은 맵(값 None=파싱 실패, ""=빈 슬롯)."""
    return {k: (v.get("mn-ea", "") if v is not None else None)
            for k, v in theme_fonts_by_master(prs).items()}


def _theme_colors_from_blob(blob: bytes) -> Optional[Dict[str, str]]:
    """테마 clrScheme의 색 이름 -> RRGGBB(sysClr는 lastClr). W7의 schemeClr 해석용
    (3차 외부 리뷰 P1: 직접 RGB만 읽어 테마색 텍스트가 검사에서 빠지던 것 보강)."""
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
        # schemeClr 참조명 매핑(표준 clrMap 기본): tx1->dk1, tx2->dk2, bg1->lt1, bg2->lt2
        for ref, base in (("tx1", "dk1"), ("tx2", "dk2"), ("bg1", "lt1"), ("bg2", "lt2")):
            if base in out:
                out.setdefault(ref, out[base])
        return out
    except Exception:
        return None


def theme_colors_by_master(prs) -> Dict[str, Optional[Dict[str, str]]]:
    """슬라이드마스터별 테마 색 맵. 키=마스터 partname 문자열."""
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
    """run 슬롯 값의 OOXML 테마 토큰("+mn-lt"/"+mj-ea" 류)을 실폰트명으로 치환.
    토큰이 해석 불가(테마 파싱 실패·빈 값)면 그 슬롯을 비운다(= 상속 체인에 맡김).
    적대 패널 확정(2026-07-10): 토큰을 문자 그대로 블록리스트에 대조하면 절대 매칭되지
    않아 라틴 전용 테마 폰트로 폴백되는 런이 E1을 조용히 통과했다."""
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
    """하위호환 진입점: 첫 마스터의 테마 a:ea(관계 기반 해석).
    반환: 폰트명 / ""(빈 슬롯 = Windows에서 Malgun 폴백) / None(테마 파싱 실패)."""
    ea_map = theme_ea_by_master(prs)
    for v in ea_map.values():
        if v is not None:
            return v
    return None


def _sz_from_defrpr(d) -> Optional[float]:
    """defRPr@sz(1/100pt 정수)를 pt로. 없거나 쓰레기값이면 None."""
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
    """lstStyle 형 컨테이너(a:lvlXpPr 자식들)에서 해당 레벨의 defRPr 요소."""
    if lst_el is None:
        return None
    lvl = min(max(int(lvl or 0), 0), 8)
    p = lst_el.find(NS + "lvl%dpPr" % (lvl + 1))
    if p is None:
        return None
    return p.find(NS + "defRPr")


def _lst_sz_pt(lst_el, lvl: int) -> Optional[float]:
    """하위호환 셔: lstStyle에서 해당 레벨의 defRPr sz(pt)."""
    return _sz_from_defrpr(_lst_defrpr(lst_el, lvl))


class StyleResolver:
    """run·문단에 명시 속성이 없을 때 OOXML 상속 체인으로 실효 스타일을 해석한다.
    크기와 폰트(a:ea/a:latin)를 같은 체인으로 해석하되 속성별로 독립 탐색한다.

    체인(ECMA-376 텍스트 스타일 계층): 도형 txBody lstStyle → (placeholder면)
    레이아웃 동일 idx placeholder lstStyle → 마스터 placeholder lstStyle →
    마스터 txStyles(title/body/other) → 프레젠테이션 defaultTextStyle.

    크기: 전부 없으면 None(=W5). 0.1.0은 run·문단만 보고 전부 W5로 흘려 placeholder
    덱에서 E3/W1/W8이 사실상 죽어 있었다(외부 리뷰 2026-07-10).
    폰트: 마스터 lstStyle의 a:ea가 실제 렌더에 상속됨을 COM 프로브로 실측 확인
    (2026-07-10 프로브6, docs/CALIBRATION.md). 0.2.0은 run rPr만 봐서 표준 기업
    템플릿에서 'Malgun 폴백 확정' 오탐을 냈다(2차 외부 재점검 확정)."""

    def __init__(self, prs):
        self._prs = prs
        self._default_el = None
        self._default_loaded = False
        # 상속 의존 덱은 run 수천 개가 같은 layout/master를 반복 조회한다: defRPr 요소를
        # partname 키로 메모이제이션(적대 패널 지적: 매 run 재순회하던 비대칭의 교정)
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
        """placeholder면 title/body/other, 아니면 None. E1의 majorFont 분기용
        (제목 패밀리는 테마 majorFont ea를 탄다: 프로브6 Q1 실측)."""
        try:
            if getattr(sp, "is_placeholder", False):
                return self._ph_family(sp.placeholder_format.type)
        except Exception:
            pass
        return None

    def _layout_ph_defrpr(self, slide, idx: int, lvl: int):
        # 가드는 placeholder 단위: 루프 전체를 한 try로 감싸면 손상된 placeholder 하나가
        # 검색을 통째로 중단시켜 뒤 순번의 진짜 스타일을 놓친다(적대 패널 확정 2026-07-10)
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
        """상속 체인의 (defRPr 요소, 출처) 순회. 요소가 없는 노드는 건너뜀."""
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
        """(실효 크기 pt, 출처) 또는 (None, None). 출처 구분은 제목 수집기 범람 방지용:
        defaultTextStyle 폴백 18pt는 게이트엔 유효하지만 '이 덱이 의도한 제목 크기'가
        아니므로 제목 후보에서 제외한다(2차 외부 재점검 확정 회귀의 교정)."""
        for el, src in self._chain(tf, sp, slide, lvl):
            v = _sz_from_defrpr(el)
            if v is not None:
                return v, src
        return None, None

    def resolve_font(self, tf, sp, slide, lvl: int, slot: str) -> Optional[str]:
        """상속 체인에서 defRPr의 a:{slot}(ea/latin) typeface. 테마 토큰일 수 있음."""
        for el, _src in self._chain(tf, sp, slide, lvl):
            f = el.find(NS + slot)
            if f is not None:
                name = f.get("typeface")
                if name:
                    return name
        return None

    def resolve(self, tf, sp, slide, lvl: int) -> Optional[float]:
        """하위호환 진입점: 실효 크기(pt)만. tf=검사 중 텍스트프레임(표 셀 포함), sp=소유 도형."""
        return self.resolve_size(tf, sp, slide, lvl)[0]


SizeResolver = StyleResolver   # 하위호환 별칭(0.2.0 공개명)


# AI 티 카피: 명백한 클리셰만(좁은 사전 = 오탐 억제). 맥락상 정당할 수 있는 일반어는 안 넣는다.
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
    """W11: AI 티 카피. 버즈워드는 전 페이지, 뻔한 오프닝은 도입부(p1~3)만."""
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
    """슬라이드 하단 밴드(y>0.88H)의 최하단 텍스트 top(in). 푸터 없으면 None."""
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
    """W12: 페이지간 푸터 baseline 어긋남. 50덱 실측(2026-07-02)에서 절대 편차 방식은 표지
    크레딧·하단 캡션을 푸터로 오인해 17덱 오탐 → 표지(p1) 제외 + 지배 baseline(0.05in 양자화
    최빈, 3페이지+)을 하우스 푸터로 보고, 거기서 '살짝'(0.03~0.25in) 어긋난 페이지만 지목.
    0.25in 넘게 다른 건 캡션·디바이더 등 다른 요소로 보고 무시(존재 여부도 안 본다)."""
    tops = [(si, t) for si, t in foot_tops.items() if t is not None and si > 1]
    if len(tops) < 4:
        return
    q = Counter(round(t / 0.05) for _si, t in tops)
    qv, cnt = q.most_common(1)[0]
    if cnt < 3:
        return
    # 하우스 기준 = 지배 버킷의 실제 값 중앙값(버킷 중심을 쓰면 7.08 다수가 7.10과 딱 0.02
    # 차이로 부동소수점 경계 오탐, 50덱 재스캔서 확인)
    bvals = sorted(t for _si, t in tops if round(t / 0.05) == qv)
    base = bvals[len(bvals) // 2]
    off = [(si, t) for si, t in tops if 0.03 < abs(t - base) <= 0.25]
    if off:
        ex = " ".join("p%d=%.2f" % (si, t) for si, t in off[:4])
        warns.append(Finding(0, "W12", "w12", (base, len(off)), ex))


_EFFECT_TAGS = tuple(NS + t for t in ("outerShdw", "innerShdw", "glow", "reflection"))
_3D_TAGS = tuple(NS + t for t in ("sp3d", "scene3d"))


def effects_count(slide):
    """슬라이드의 실효 PPT 효과(그림자·글로·3D) 수와 종류. 일부 생성기는 상속 차단용으로
    자식 없는 빈 effectLst만 남기므로 실제 효과 자식이 있어야만 센다(빈 요소 오탐 방지)."""
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
    """W13: 덱 단위 1회 집계(페이지별 반복 발화는 소음: 50덱 코퍼스 실측).
    의도적 네온·글로 스타일일 수 있어 WARN, 판단은 눈."""
    hits = [(si, n, kinds) for si, (n, kinds) in per_page.items() if n >= 2]
    if not hits:
        return
    total = sum(n for _si, n, _k in hits)
    kinds = sorted(set().union(*[k for _si, _n, k in hits]))
    pages = ",".join("p%d" % si for si, _n, _k in hits[:6])
    warns.append(Finding(0, "W13", "w13", (total, len(hits)),
                         "%s | %s" % (pages, ",".join(kinds))))


# W15 텍스트 겹침: 생성 덱에서 가장 흔한 결함 축(수정 라운드마다 요소가 포개진다)인데
# 프레임 bbox 는 넉넉히 잡는 관례라 못 쓰고, 실효 글리프 폭을 근사해 잡는다.
_W_CJK, _W_LAT, _W_SP = 0.96, 0.52, 0.28   # 글자폭/폰트크기 비 근사(보수적: 오탐 억제)

# 문단 단위 실효 글리프 bbox(in). 매직 인덱스 튜플이던 것을 명명 필드로(외부 리뷰 2026-07-10).
# sp(소유 도형)는 W15~W17 finding의 loc 페이로드용(0.5.0); 좌표는 이미 그룹 절대좌표다.
GlyphBox = namedtuple("GlyphBox", "x0 y0 x1 y1 rep max_pt frame_id sp")
GlyphBox.__new__.__defaults__ = (None,)


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
    """빈 문단(스페이서)의 실효 크기: endParaRPr/defRPr sz 우선(4pt 스페이서를 12pt로
    통산해 팬텀 높이를 만들던 것의 교정, 적대 검증 실측)."""
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
    """문단 단위 실효 글리프 bbox(in) 근사. 반환 [(x0,y0,x1,y1,대표텍스트,max_pt,frame_id)].
    폭은 run별 실제 크기로 합산, 행간은 line_spacing 실값(없으면 1.2)에 autofit
    lnSpcReduction 반영. 문단별 정렬(명시 없으면 프레임 첫 명시값, 그것도 없으면 좌)로
    x를 잡아 혼합 정렬 프레임의 오배치를 줄인다. wrap=none(word_wrap=False, python-pptx
    add_textbox 기본)은 한 줄로 프레임 밖까지 뻗으므로 랩 접기 없이 실폭 그대로.
    회전 프레임은 추정 무효라 스킵. 그룹은 iter_shapes_geo로 절대좌표 변환.
    실덱 렌더 대조 + 적대 검증 실측 재현으로 캘리브레이션.
    스크립트 레이어(0.2.1): 세로쓰기(bodyPr@vert)와 RTL·복잡 조판 스크립트가 든 프레임은
    글자폭 근사가 무의미해 스킵하고, skipped Counter가 오면 집계해 W18로 표면화한다.
    0.3.1(3차 외부 리뷰 P0): styler(StyleResolver)를 받으면 크기 없는 run을 E3와 동일한
    상속 체인으로 해석한다(한 문서에 두 실효 스타일 모델이 존재하던 불일치의 교정).
    네이티브 표는 열 폭·행 높이 누적으로 셀 사각형을 계산해 셀 텍스트도 포함한다.
    알려진 한계: 플레이스홀더가 레이아웃 lstStyle로 정렬을 상속하는 경우는 좌정렬로 후퇴,
    회전 프레임 스킵은 장식 관행이라 W18 집계에서 제외."""
    import math
    out = []

    def emit_frame(tframe, fx, fy, fw, fh, fid, owner_sp):
        try:
            bodyPr = tframe._txBody.find(NS + "bodyPr")
            vert = bodyPr.get("vert") if bodyPr is not None else None
            if vert not in (None, "horz"):
                if skipped is not None:
                    skipped["vertical_text"] += 1
                return
        except Exception:
            pass
        try:
            frame_text = "".join(r.text for para in tframe.paragraphs for r in para.runs)
            if _geometry_unsupported(frame_text):
                if skipped is not None:
                    skipped["complex_script"] += 1
                return
        except Exception:
            pass
        fw2 = max(fw, 0.05)
        scale, lnred = frame_autofit(tframe)
        wrap = tframe.word_wrap is not False   # None(속성 없음)=OOXML 기본 square=랩
        frame_align = None
        for para in tframe.paragraphs:
            if para.alignment is not None:
                frame_align = para.alignment
                break
        paras = []   # (pw, pmx, ptxt, factor, n, align)
        for para in tframe.paragraphs:
            pw, pmx, ptxt = 0.0, 0.0, ""
            for r in para.runs:
                t = r.text
                if not t:
                    continue
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
                pw += _glyph_w(t, sz)
                ptxt += t
                if sz > pmx:
                    pmx = sz
            if not ptxt.strip():
                paras.append((0.0, _empty_para_pt(para, default_pt) * scale, "", 1.2, 1, None))
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
            n = max(1, math.ceil(pw / (fw2 * 1.04))) if wrap else 1
            al = para.alignment if para.alignment is not None else frame_align
            paras.append((pw, pmx, ptxt, factor, n, al))
        gh_total = sum(n * pmx * max(f, 0.95) * (1.0 - lnred) / 72.0
                       for (pw, pmx, ptxt, f, n, al) in paras)
        if gh_total <= 0:
            return
        va = str(tframe.vertical_anchor) if tframe.vertical_anchor is not None else ""
        if "MIDDLE" in va:
            cy = fy + max(0.0, (fh - gh_total) / 2)
        elif "BOTTOM" in va:
            cy = fy + max(0.0, fh - gh_total)
        else:
            cy = fy
        for (pw, pmx, ptxt, factor, n, al) in paras:
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
                out.append(GlyphBox(x0, cy, x0 + gw, cy + ph, ptxt[:24], pmx, fid, owner_sp))
            cy += ph

    for sp, _z, xf in iter_shapes_geo(slide.shapes):
        if getattr(sp, "has_table", False):
            # 네이티브 표 셀(3차 외부 리뷰 P0: 자동 생성 덱의 표가 기하 사각지대였다).
            # 셀 사각형 = 표 원점 + 열 폭·행 높이 누적(EMU에 xf 스케일 적용).
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
                        ctf = cell.text_frame
                    except Exception:
                        cx_off += cw
                        continue
                    emit_frame(ctf,
                               tx + ax * cx_off / EMU_PER_IN,
                               ty + ay * cy_off / EMU_PER_IN,
                               ax * cw / EMU_PER_IN,
                               ay * rh / EMU_PER_IN,
                               id(cell._tc), sp)
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
        emit_frame(sp.text_frame, fx, fy, fw, fh, id(sp), sp)
    return out


def text_overlap_check(slide, si, warns, boxes: Optional[List[GlyphBox]] = None):
    """W15: 서로 다른 텍스트 프레임의 실효 글리프 영역이 유의미하게 포개짐(가림·충돌).
    근사 기반이라 WARN. 교집 면적이 작은 쪽의 45% 초과일 때만, 페이지당 최대 2건.
    임계 45%는 렌더 대조 실측: 30~35%대는 전부 오탐(빅넘버 아래 타이틀·2단 사이 침범 추정),
    60%+는 전부 실겹침(차트가 엑서빗 라벨 침범, 캡션 칩이 범례 위 등).
    의도적 레이어는 제외: 동일 텍스트 에코(잔상 타이포 연출), 1~2자 대형 글자(드롭캡·장 번호 I·II 류).
    boxes 를 주면 재계산 없이 그대로 쓴다(슬라이드당 1회 계산 캐시)."""
    if boxes is None:
        boxes = _text_glyph_boxes(slide)
    hits = []
    for i in range(len(boxes)):
        for j in range(i + 1, len(boxes)):
            a, b = boxes[i], boxes[j]
            if a.frame_id == b.frame_id:      # 같은 프레임의 문단끼리는 스택이라 제외
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
    # 정렬 키는 frac만: GlyphBox의 sp 필드는 비교 불가라 튜플 통비교가 죽는다(0.5.0)
    for frac, a, b in sorted(hits, key=lambda h: h[0], reverse=True)[:2]:
        # loc: 프레임 raw bbox가 아니라 겹침을 판정한 실효 글리프 bbox(절대 in) 그대로.
        # related에 상대 프레임을 실어 에이전트가 어느 쌍을 움직일지 특정하게 한다(0.5.0).
        loc = shape_loc(a.sp, bbox=[a.x0, a.y0, a.x1 - a.x0, a.y1 - a.y0]) or {}
        rel = shape_loc(b.sp, bbox=[b.x0, b.y0, b.x1 - b.x0, b.y1 - b.y0])
        if rel:
            loc["related"] = rel
        warns.append(Finding(si, "W15", "w15", (frac * 100,), "%r ~ %r" % (a.rep, b.rep),
                             loc=loc or None))


def _pic_boxes(slide, sw_in, sh_in, skipped=None):
    """비배경 그림의 실효 잉크 bbox(in)와 z순서. 반환 [(x0,y0,x1,y1,z)].
    슬라이드의 70%+를 덮는 풀블리드·mesh 배경은 제외. 투명 PNG(matplotlib 차트 등)는
    프레임 bbox가 잉크보다 훨씬 커서 오탐 원천 → 알파 불투명 bbox로 트림.
    성능 예산(0.4.0, 3차 리뷰): 25MP 초과 이미지는 알파 트림을 생략하고 프레임 bbox를
    그대로 쓰며 skipped 카운터로 공개한다(대량 배치에서 디코드 폭주 방지).
    적대 검증 실측 반영(2026-07-03): P모드+tRNS는 RGBA 변환 후 트림, srcRect 크롭은
    보이는 소스 창으로 좁혀 매핑, flipH/flipV는 창 안에서 미러, 회전 그림은 축정렬
    확장 bbox를 쓰되 잉크 트림은 무효라 건너뜀, 그룹 자식은 절대좌표 변환."""
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
                    raise RuntimeError("image decode budget")   # 아래 except가 프레임 bbox 유지
                if im.mode == "P" and "transparency" in im.info:
                    im = im.convert("RGBA")
                if "A" in im.getbands():
                    bb = im.getchannel("A").point(lambda a: 255 if a > 16 else 0).getbbox()
                    if bb is None:
                        continue   # 전면 투명 = 잉크 없음
                    iw, ih = im.size
                    wl, wt = iw * float(sp.crop_left or 0), ih * float(sp.crop_top or 0)
                    wr = iw * (1.0 - float(sp.crop_right or 0))
                    wb = ih * (1.0 - float(sp.crop_bottom or 0))
                    l, t, r, b = bb
                    l, r = max(l, wl), min(r, wr)
                    t, b = max(t, wt), min(b, wb)
                    if l >= r or t >= b:
                        continue   # 크롭 창 안에 잉크 없음
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
            except Exception:
                pass
        out.append((x, y, x + w, y + h, z, sp))
    return out


def overflow_check(slide, si, sw_in, sh_in, warns,
                   boxes: Optional[List[GlyphBox]] = None, pics: Optional[list] = None):
    """W16: 화면 밖 넘침. 프레임 bbox 기준은 넉넉한 프레임 관례 때문에 오탐이 커서 기각했었지만
    W15의 실효 글리프 bbox가 그 반론을 해소했다(2026-07-03): 텍스트는 실제 글자 영역이 경계를
    뚫을 때만. 비텍스트는 그림(잉크 bbox 트림)만 본다: 장식 도형(글로우 원 등)을 모서리로
    흘리는 블리드는 표준 기법이라 결함이 아님(코퍼스 렌더 실측으로 도형 검사는 기각)."""
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
                         shape_loc(gb.sp, bbox=[gb.x0, gb.y0, gb.x1 - gb.x0, gb.y1 - gb.y0])))
    for (px0, py0, px1, py1, _z, psp) in pics:
        over = max(-px0, -py0, px1 - sw_in, py1 - sh_in)
        if over > TOL_S:
            hits.append((over, M("w16_pic") % (px1 - px0, py1 - py0),
                         "p|%.1fx%.1f" % (px1 - px0, py1 - py0),
                         shape_loc(psp, bbox=[px0, py0, px1 - px0, py1 - py0])))
    # 정렬 키는 over만: loc dict는 비교 불가라 튜플 통비교가 죽는다(0.5.0)
    for over, what, fpk, loc in sorted(hits, key=lambda h: h[0], reverse=True)[:2]:
        # fp_key: detail(what)은 로케일 문자열이라 baseline 지문에서 제외(4차 리뷰)
        warns.append(Finding(si, "W16", "w16", (over,), what, fp_key=fpk, loc=loc))


def _occluder_boxes(slide, sw_in, sh_in):
    """그림 위에 얹힌 솔리드 채움 도형(카드·패널) bbox와 z. 사진 위 카드에 캡션이 앉은
    합법 레이아웃이 W17에 걸리던 것의 억제용(적대 검증 실측)."""
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
    """W17: 텍스트가 비배경 그림의 잉크 경계에 걸침(글리프의 25~75%만 이미지 안) = 반은 위·
    반은 밖이라 잘리거나 배경이 갈라져 보임. 완전 위(오버레이 캡션)는 W7 대비 게이트 소관,
    완전 밖은 무관. 1제곱인치 미만(아이콘·로고급) 그림 무시. 사진과 텍스트 사이 z에 솔리드
    카드가 끼어 텍스트 영역의 90%+를 받치고 있으면 걸침이 아니라 카드 위 캡션이라 제외.
    페이지당 최대 2건."""
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
    # 정렬 키는 frac만(sp 필드 비교 불가, 0.5.0)
    for frac, gb, pic in sorted(hits, key=lambda h: h[0], reverse=True)[:2]:
        loc = shape_loc(gb.sp, bbox=[gb.x0, gb.y0, gb.x1 - gb.x0, gb.y1 - gb.y0]) or {}
        rel = shape_loc(pic[4], bbox=[pic[0], pic[1], pic[2] - pic[0], pic[3] - pic[1]])
        if rel:
            loc["related"] = rel
        warns.append(Finding(si, "W17", "w17", (frac * 100,), "%r" % gb.rep, loc=loc or None))


def action_title_check(titles, warns):
    """W14: 타이틀 다수가 서술형 명사구(시장 현황·경쟁 분석류) = 액션타이틀 아님(MBB: 제목만
    읽어도 주장이 흐르게). 종결어미 휴리스틱이라 주장으로 오판하는 미탐은 있어도(바다 등 '다'
    끝 명사) 오탐은 좁게: 한글 타이틀 3개+가 명사구이고 전체의 절반 이상일 때만 덱 단위 1회."""
    entries = []
    for si in sorted(titles):
        txt = " ".join(titles[si][1]).strip()
        chars = [c for c in txt if not c.isspace()]
        cjk_n = sum(1 for c in chars if is_cjk(c))
        # 한글 3자 미만·한글 비중 30% 미만은 제외: 빅스탯 숫자("300만+ 18.3%")·짧은 브랜드명을
        # 타이틀로 오인하던 것(50덱 스캔 실측)과 영문 타이틀을 걸러낸다.
        if len(chars) < 4 or cjk_n < 3 or cjk_n < 0.3 * len(chars):
            continue
        core = txt.rstrip(" ?!.…”’")
        # 숫자+단위가 박힌 타이틀("매출 3배 성장")은 명사로 끝나도 주장형 헤드라인이다:
        # 종결어미만 보던 판별이 이런 타이틀을 명사구로 오분류하던 것의 교정(외부 리뷰 2026-07-10).
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
    """두 페이지 공유 채움도형 멀티셋에서 '장식 텍스처 클론'(작은 점·절리 등, 24분할 1x1 이하) 수를 센다.
    카드형(중간 블록)·표(전폭 밴드)는 W6 소관이라 세지 않아 3열 비교카드 오탐을 피한다(적대 감사 반영)."""
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
    """슬라이드 요소 배치를 gw x gh 그리드 점유 벡터로. 풀블리드 배경은 제외.
    텍스트=1 도형=0.5 이미지=2 가중. 어디에 무엇이 있는지(골격) 비교용."""
    sig = [0.0] * (gw * gh)
    n = 0
    for sp in iter_shapes(slide.shapes):
        try:
            L, T, Wd, Ht = sp.left, sp.top, sp.width, sp.height
        except Exception:
            continue
        if None in (L, T, Wd, Ht) or not Wd or not Ht:
            continue
        if Wd > 0.9 * sw and Ht > 0.9 * sh:      # 풀블리드 배경/이미지 제외
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
    # 동일 벡터가 부동소수점 오차로 1.0을 살짝 넘어 w6_sim=1.0 상한까지 뚫는 것 방지
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


def _resolve_run_rgb(run, para, tframe, sp, slide, styler=None, thm_colors=None):
    """텍스트색 해석: run rPr 직접 RGB → 문단 defRPr → lstStyle 상속 체인 → schemeClr를
    테마 clrScheme으로 해석(3차 리뷰 P1: 직접 RGB만 읽어 테마색 텍스트가 W7에서 빠지던 것)."""
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
        return None

    direct = _run_rgb(run)
    if direct:
        return direct
    try:
        c = from_el(run._r.find(NS + "rPr"))
        if c:
            return c
    except Exception:
        pass
    try:
        pPr = para._p.find(NS + "pPr")
        c = from_el(pPr.find(NS + "defRPr") if pPr is not None else None)
        if c:
            return c
    except Exception:
        pass
    if styler is not None:
        try:
            for el, _src in styler._chain(tframe, sp, slide, getattr(para, "level", 0)):
                c = from_el(el)
                if c:
                    return c
        except Exception:
            pass
    return None


def contrast_check(slide, si, sw, sh, render_dir, warns, styler=None, thm_colors=None):
    """이미지 위 텍스트 저대비 감지(렌더 PNG 근사). picture와 겹치는 텍스트 프레임만, 슬라이드당 1회.
    반환: "no_pics"(검사 대상 없음) / "no_png"(그림은 있는데 규약 렌더 없음=불완전) / "ok"(검사함).
    좌표는 그룹 변환 포함 절대좌표(3차 리뷰 P1: raw 좌표라 그룹 내부 그림·텍스트가 어긋나던 것)."""
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
        # 배경 휘도는 텍스트색 반대 극단의 분위수로: 최악-국소 대비를 재고(WCAG는 최악값 기준)
        # 텍스트 잉크가 섞인 평균 오염을 피한다(밝은 글자면 어두운 15퍼센타일, 어두운 글자면 밝은 85퍼센타일).
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
    return "ok"   # 이 페이지의 렌더 PNG를 찾아 검사했음(W7 발화 여부와 무관)


def lint(path, hard_min=5.0, body_min=9.0, small_min=7.5, render_dir=None, ghost=None,
         strict=False, w6_sim=0.90, w6_min_cluster=3, profile=DEFAULT_PROFILE):
    """profile은 실행 정책이다(3차 외부 리뷰 P0: CLI 사후 필터가 아니라 엔진 단계 적용).
    제외된 규칙은 아예 실행하지 않으므로 O(S^2) 비교 비용도 안 들고, 제외 규칙의 내부
    실패가 W18로 누출되지도 않으며, 라이브러리 호출에서도 프로파일을 쓸 수 있다.

    0.4.0 파괴적 변경: 기본 프로파일이 core(객관 결함만)다. AI 티·하우스 스타일 규칙
    (E2, W6, W9~W14)까지 원하면 profile="full"을 명시한다. 첫 사용자 무옵션 실행이
    정상 문장부호로 exit 1을 받는 첫인상 문제의 교정(외부 전략 리뷰, 사용자 확정)."""
    if profile not in PROFILES:
        # 오타 프로파일이 조용히 full(빈 제외집합)로 동작하던 것 교정(4차 리뷰)
        raise ValueError("unknown profile %r (choices: %s)" % (profile, ", ".join(sorted(PROFILES))))
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
    deck_skipped = Counter()   # 덱 레벨 검사 불능(W18 p00 발화용)
    if render_dir and not os.path.isdir(render_dir):
        # --render 폴더가 없으면 W7이 조용히 0건 수행되던 경로(3차 리뷰 P0): 불완전으로 표면화
        deck_skipped["render_dir_missing"] += 1
        print(M("note_render_dir_missing") % render_dir, file=sys.stderr)
    theme_fails = sum(1 for v in fonts_map.values() if v is None)
    if theme_fails:
        # 파싱 실패(None)와 확인된 빈 슬롯("")은 E1 분기에선 같은 폴백 가정으로 후퇴한다.
        # 구분이 사라지는 지점이라 W18로 표면화(적대 패널·2차 재점검: 침묵 붕괴 방지).
        deck_skipped["theme_parse"] = theme_fails
        print(M("note_theme_parse"), file=sys.stderr)
    styler = StyleResolver(prs)

    render_png_hits = 0
    for si, slide in enumerate(prs.slides, 1):
        # 슬라이드가 쓰는 마스터의 테마 폰트 슬롯(멀티마스터 덱에서 E1 오발화 방지)
        thm_fonts = thm_default
        thm_colors = colors_default
        try:
            pn = str(slide.slide_layout.slide_master.part.partname)
            thm_fonts = fonts_map.get(pn, thm_default)
            thm_colors = colors_map.get(pn, colors_default)
        except Exception:
            pass
        skipped = Counter()   # W18: 검사 불능 구간 집계(조용한 저하를 JSON에 표면화)
        try:
            slide_part = str(slide.part.partname)   # finding location의 part 필드용
        except Exception:
            slide_part = None
        # sig와 toks는 각자 가드: 한 try에 묶으면 sig 성공 후 toks 실패 시 except의
        # 재append로 sigs가 밀려 W6 페이지 번호가 어긋난다(적대 패널 실측 재현).
        # 프로파일에서 제외된 규칙의 수집·검사는 아예 실행하지 않는다(3차 리뷰 P0).
        if "W6" not in excl:
            try:
                sig = slide_layout_sig(slide, sw, sh)
            except Exception as e:
                sig = ([0.0] * 24, 0)   # 위치 보존(sigs는 페이지 순서 인덱스)
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
                                    styler=styler, thm_colors=thm_colors)
                if r7 == "ok":
                    render_png_hits += 1
                elif r7 == "no_png":
                    # 그림이 있는 페이지인데 규약 렌더가 없음 = 이 페이지 W7 미수행(불완전)
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
        # 기하 기초자료는 슬라이드당 1회만 계산(이전엔 W15~W17이 각자 재계산: 텍스트박스 3회,
        # 그림 PIL 디코드 2회. 대형 덱 성능 지적의 교정, 외부 리뷰 2026-07-10)
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
        # 실패한 축만 빈 리스트로 후퇴시키고 살아있는 축은 계속 검사한다: tboxes 실패가
        # 무관한 그림 W16까지 침묵시키던 공유 게이트의 교정(적대 패널 실측 재현 2건).
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
        # 코어 라인 게이트(E1~E4·W1/W5/W8). 가드는 run 단위: 프레임 단위 가드는 한 run의
        # 쓰레기 속성이 같은 프레임의 진짜 위반(이웃 run)까지 삼켜 거짓 통과를 만들었다
        # (적대 패널 실측 재현 2026-07-10). 삼켜진 구간은 W18로 JSON에 표면화한다.
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
                    # 문서 순서 아이템 (run_like, para.runs 인덱스 or None, is_fld):
                    # a:fld(자동 필드)는 일반 run과 같은 rPr로 렌더되므로 같은 게이트를
                    # 통과해야 하고, a:br은 E2 문맥·오프셋에서 줄바꿈 한 글자로 취급한다
                    # (4차 리뷰 이월분, 0.5.0). a:r 개수가 안 맞으면 종전 runs 경로로 후퇴.
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
                    p_fonts = para_fonts(para)   # 문단 defRPr: run rPr 다음 순위(프로브7)
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
                        if run is None:   # a:br: 문맥 기여만, 검사 대상 아님
                            continue
                        t = run.text
                        if not t:
                            continue
                        if not is_fld:
                            page_texts.setdefault(si, []).append(t)
                        lvl = getattr(para, "level", 0)

                        # E1: 실효 렌더 폰트 판정(실측 렌더 모델, e1_violation 참조).
                        # 스크립트별 판정: 한글은 실측 모델 전체, 가나·한자는 CJK 전무 폰트
                        # (Inter·모노 계열)만 발화하고 JP/SC 서브셋 폰트는 통과(3차 패널).
                        # run rPr에 없는 슬롯은 lstStyle 상속 체인에서 보충(프로브6 실측: 마스터
                        # lstStyle의 a:ea가 실제 렌더에 상속됨), 제목 패밀리는 테마 majorFont ea.
                        script = None
                        if has_hangul(t):
                            script = "hangul"
                        elif any(is_kana(c) or is_hanja(c) for c in t):
                            script = "cjk_other"
                        if script:
                            fonts = run_fonts(run)
                            for slot in ("ea", "latin"):
                                if slot not in fonts and slot in p_fonts:
                                    fonts[slot] = p_fonts[slot]   # 문단 defRPr(프로브7: lstStyle을 이김)
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

                        # E2: 긴 대시류. 문맥은 문단 전체(ptext)에서 보고 보고는 이 run 구간만:
                        # run 경계로 쪼개진 '2020'/'-2021' 범위 오탐 방지(적대 패널 확정)
                        bad = [] if "E2" in excl else \
                            dash_violations(ptext, strict=strict,
                                            span=(run_offs[ii], run_offs[ii] + len(t)))
                        if bad:
                            errors.append(Finding(si, "E2", "e2", (),
                                                  "cp=%s text=%r" % (",".join("U+%04X" % ord(c) for c in sorted(set(bad))), t[:24]),
                                                  loc=shape_loc(owner_sp, paragraph=pi, run=ri, part=slide_part, cell=cell_rc, xf=sp_xf, field=is_fld)))

                        # E3 / W1 / W5: 실효 폰트 크기(run → 문단 → placeholder 상속 체인, autofit 반영).
                        # size_src는 제목 수집 자격 판정용: defaultTextStyle 폴백(18pt)은 게이트엔
                        # 유효하지만 의도된 제목 크기가 아니라 ghost/W14를 범람시킨다(회귀 교정).
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
                            # 제목 자격: 명시 크기 또는 제목 패밀리 placeholder만. lstStyle·
                            # 마스터 bodyStyle 상속 크기가 18pt를 넘으면 본문 산문이 ghost/W14에
                            # 쓸려 들어간다(3차 적대 패널: own lstStyle 20pt 본문 실측 재현)
                            # 자동 필드(슬라이드 번호 등)는 크기가 커도 제목이 아니다(0.5.0)
                            if eff >= 18 and t.strip() and not is_fld \
                                    and (size_src == "explicit" or frame_fam == "title"):
                                # 제목 placeholder가 있으면 크기 무관 그것이 제목이다: 60pt
                                # KPI 빅넘버가 26pt 실제 제목을 밀어내던 것 교정(3차 리뷰)
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
                                # 좁은 프레임(<=4in)만: 넓은 프레임의 소형 한글은 목업·카드 내부가 아니라
                                # 캡션·주석일 여지가 커 메시지(목업 추정)와 어긋난다(공개 위생 감사 반영).
                                warns.append(Finding(si, "W8", "w8", (eff, small_min),
                                                     "w=%.1fin text=%r" % (fw_in, t[:24]),
                                                     loc=shape_loc(owner_sp, paragraph=pi, run=ri, part=slide_part, cell=cell_rc, xf=sp_xf, field=is_fld)))
                        else:
                            warns.append(Finding(si, "W5", "w5", (), "text=%r" % t[:24],
                                                 loc=shape_loc(owner_sp, paragraph=pi, run=ri, part=slide_part, cell=cell_rc, xf=sp_xf, field=is_fld)))

                        # E4: 연속 한글·한자 2자 이상 + 유의미 양수 트래킹. 가나가 섞인 런은
                        # 일본어라 제외(가나 자간 벌리기는 정상 디자인 관행). 한자 전용 런은
                        # 한국 덱의 인명·법률용어가 흔해 대상 유지(3차 패널: 한자 회귀 교정)
                        if not any(is_kana(c) for c in t) and \
                                sum(1 for c in t if is_hangul(c) or is_hanja(c)) >= 2:
                            tr = run_track(run)
                            if tr is not None and tr > 50:
                                errors.append(Finding(si, "E4", "e4", (tr,), "text=%r" % t[:24],
                                                      loc=shape_loc(owner_sp, paragraph=pi, run=ri, part=slide_part, cell=cell_rc, xf=sp_xf, field=is_fld)))
                    except Exception as e:
                        skipped["run"] += 1
                        print("E1-E4 skipped p%02d run: %s" % (si, e), file=sys.stderr)

        # W18: 이 페이지에서 가드가 삼킨 구간을 출력 계약(JSON·텍스트)에 표면화.
        # stderr에만 남으면 exit code·JSON summary만 보는 CI가 불완전 검사를 통과로 오독한다
        # (적대 패널 확정 2026-07-10). --strict에선 exit 1로 승격된다.
        if skipped:
            det = ", ".join("%s=%d" % (k, v) for k, v in sorted(skipped.items()))
            warns.append(Finding(si, "W18", "w18_page", (), det))

    # 규약(p01.png) 매치가 0인데 다른 이름의 png가 있으면 파일명 규약을 안내.
    # 불완전성 자체는 페이지 단위 w7_no_render 카운터가 W18/incomplete로 표면화한다(0.3.1).
    if render_dir and os.path.isdir(render_dir) and render_png_hits == 0:
        anypng = glob.glob(os.path.join(render_dir, "*.png"))
        if anypng:
            print(M("note_render_naming") % (render_dir, os.path.basename(anypng[0])), file=sys.stderr)

    # W6: 레이아웃 골격 재탕. 성긴 슬라이드(디바이더·커버 등 비영 셀<4)는 제외하고, 콘텐츠
    # 슬라이드 중 한 장이 유사(>w6_sim)한 다른 슬라이드가 w6_min_cluster장 이상이면 경고.
    # 덱 길이에 불변(전체 쌍 비율은 큰 덱에서 국소 재탕을 희석해 놓치므로 최대 클러스터로 판정).
    # 임계는 CLI 튜너블(--w6-sim/--w6-cluster): 의도적 템플릿 일관성이 강한 하우스는 조여서
    # 억제 가능(장르 무지 지적의 교정, 외부 리뷰 2026-07-10. 통째 억제는 --skip W6).
    try:
        content = [] if "W6" in excl else \
            [(i + 1, sig) for i, (sig, n) in enumerate(sigs) if n >= 3]
        if len(content) > 200:
            # 성능 예산(0.4.0): O(S^2) 쌍 비교 상한. 잘림은 공개한다.
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
                warns.append(Finding(0, "W6", "w6", (len(worst) + 1,), M("w6_detail") % ex,
                                     fp_key=ex))
    except Exception as e:
        deck_skipped["w6"] += 1
        print("W6 skipped: %s" % e, file=sys.stderr)

    # W11 카피 클리셰 · W12 푸터 정렬 · W13 효과 · W14 액션타이틀 (전부 덱 단위, 프로파일 게이트)
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
            # ghost 는 W14 한글 필터와 무관하게 수집한 전 타이틀(18pt+): 영문·숫자 덱에서도
            # 제목만 읽어 수평 논리를 검토하는 용도라, 필터 후 entries 가 아니라 raw titles 에서 뽑는다.
            ghost.extend((si, " ".join(titles[si][1]).strip()) for si in sorted(titles))
    except Exception as e:
        deck_skipped["w11_w14"] += 1
        print("W11/W12/W13/W14 skipped: %s" % e, file=sys.stderr)

    # W10: 직접 그린 도식(단면도 등) 클론 재탕. 장식 텍스처(marks) 경로로 특화해 3열 카드형은
    # W6에 맡기고 여기선 세지 않는다(적대 감사 2026-07-02: 카드형 subblk 오탐 회피). W6가 놓친
    # 0.90 미달·단일 쌍(단면 도식 반복 실측)을 pptx 도형만으로 메운다.
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
                warns.append(Finding(0, "W10", "w10", (len(grp),),
                                     M("w10_detail") % pages_key, fp_key=pages_key))
    except Exception as e:
        deck_skipped["w10"] += 1
        print("W10 skipped: %s" % e, file=sys.stderr)

    # 덱 레벨 W18: 덱 단위 검사(W6·W10·W11~W14)와 테마 파싱의 검사 불능도 출력 계약에
    # 표면화한다. 0.2.0은 기하·코어 게이트만 집계해 문서의 '전부 표면화' 주장과 어긋났다
    # (2차 외부 재점검 확정: W18 부분 구현의 교정).
    if deck_skipped:
        det = ", ".join("%s=%d" % (k, v) for k, v in sorted(deck_skipped.items()))
        warns.append(Finding(0, "W18", "w18_deck", (), det))

    return errors, warns


def _skill_res():
    """패키지에 동봉된 스킬팩 SKILL.md(importlib.resources, py3.9+).
    다인자 joinpath는 3.11+라 3.9의 zip 로더까지 안전한 / 체인을 쓴다."""
    from importlib import resources
    return resources.files("archforge") / "skills" / "archforge-pptx-lint" / "SKILL.md"


def skill_main(argv=None):
    """`archforge skill`: 동봉 스킬팩을 출력하거나 에이전트 스킬 폴더에 설치한다.
    pip 사용자에게 스킬팩이 안 딸려가던 배포 구멍의 교정(외부 리뷰 2026-07-10):
    이제 wheel에 SKILL.md가 동봉되고 이 서브커맨드로 바로 받는다."""
    ap = argparse.ArgumentParser(prog="archforge skill", description=M("skill_desc"))
    ap.add_argument("--install", nargs="?", const="", metavar="DIR", help=M("help_skill_install"))
    ap.add_argument("--path", action="store_true", help=M("help_skill_path"))
    # 프로그램 전역 플래그: 값은 main()의 프리스캔이 이미 반영했고 여기선 수용만
    # (없으면 `archforge skill --lang ko`가 unrecognized로 죽는다: 3차 적대 패널 실측)
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


def main():
    # 한국어 메시지·argparse 도움말이 non-UTF-8 stdout(파이프·cp949/cp1252)에서 UnicodeEncodeError
    # 로 죽지 않게 파서 생성보다 먼저 재설정한다(--help 는 parse_args 중에 출력되므로 순서가 중요).
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass
    argv = sys.argv[1:]
    # --lang은 --help 문자열보다 먼저 확정돼야 해서 파서 생성 전에 프리스캔한다.
    # 반복 지정 시 argparse 관례대로 마지막 값이 이긴다(3차 적대 패널: 첫 매치 고정 교정).
    lang_arg = None
    for i, tok in enumerate(argv):
        if tok == "--lang" and i + 1 < len(argv):
            lang_arg = argv[i + 1]
        elif tok.startswith("--lang="):
            lang_arg = tok.split("=", 1)[1]
    if lang_arg:
        set_lang(lang_arg)
    # skill 서브커맨드 판별은 --lang류 선행 플래그를 건너뛰고 본다:
    # `archforge --lang ko skill`이 "skill 파일 린트"로 오해석되던 것 교정(3차 패널)
    rest = list(argv)
    while rest and (rest[0] == "--lang" or rest[0].startswith("--lang=")):
        rest = rest[2:] if rest[0] == "--lang" else rest[1:]
    if rest and rest[0] == "skill":
        # 파일명이 정확히 "skill"인 대상을 린트하려는 경우와의 충돌 안내(적대 패널 지적:
        # 조용한 오동작 방지). 그 파일을 린트하려면 `archforge ./skill`처럼 경로로 부른다.
        if os.path.exists("skill"):
            print(M("skill_conflict"), file=sys.stderr)
        sys.exit(skill_main(rest[1:]))
    if rest and rest[0] in ("scan", "demo"):
        if os.path.exists(rest[0]) and os.path.isfile(rest[0]):
            print(M("subcmd_conflict") % (rest[0], rest[0]), file=sys.stderr)
        sys.exit(scan_main(rest[1:]) if rest[0] == "scan" else demo_main(rest[1:]))
    ap = argparse.ArgumentParser(prog="archforge", description=M("prog_desc"))
    ap.add_argument("pptx")
    ap.add_argument("--render", default=None, help=M("help_render"))
    ap.add_argument("--write-baseline", default=None, metavar="PATH", help=M("help_write_baseline"))
    _add_common_flags(ap)
    a = ap.parse_args(argv)

    res = _lint_one(a.pptx, a)
    if res is None:   # --write-baseline: 기록 후 종료
        sys.exit(0)

    if a.sarif:
        import json
        sarif_doc = _reporters.build_sarif(a.pptx, res["errors"], res["warns"])
        with open(a.sarif, "w", encoding="utf-8", newline="\n") as f:
            json.dump(sarif_doc, f, ensure_ascii=False, indent=2)

    if a.json:
        import json
        doc = _reporters.build_json_doc(a.pptx, res["errors"], res["warns"],
                                        res["ghost"], res["summary"])
        print(json.dumps(doc, ensure_ascii=False, indent=2))
        sys.exit(1 if res["fail"] else 0)

    for line in _reporters.render_text(a.pptx, res["errors"], res["warns"], res["ghost"],
                                       res["profile"], res["profile_excl"], res["skip"],
                                       config_path=res["cfg_path"],
                                       baseline_suppressed=res["baseline_suppressed"],
                                       baseline_path=res["baseline_path"]):
        print(line)

    sys.exit(1 if res["fail"] else 0)


def _add_common_flags(ap):
    """단일 파일 모드와 scan 모드가 공유하는 플래그(중복 정의 드리프트 방지, 0.5.0)."""
    ap.add_argument("--hard-min", type=float, default=None, help=M("help_hard_min"))
    ap.add_argument("--body-min", type=float, default=None, help=M("help_body_min"))
    ap.add_argument("--strict", action="store_true", help=M("help_strict"))
    ap.add_argument("--small-min", type=float, default=None, help=M("help_small_min"))
    ap.add_argument("--ghost", action="store_true", help=M("help_ghost"))
    ap.add_argument("--json", action="store_true", help=M("help_json"))
    ap.add_argument("--skip", default=None, metavar="CODES", help=M("help_skip"))
    ap.add_argument("--profile", default=None, choices=sorted(PROFILES), help=M("help_profile"))
    ap.add_argument("--lang", default=None, choices=("ko", "en"), help=M("help_lang"))
    ap.add_argument("--w6-sim", type=float, default=None, help=M("help_w6_sim"))
    ap.add_argument("--w6-cluster", type=int, default=None, help=M("help_w6_cluster"))
    ap.add_argument("--config", default=None, metavar="PATH", help=M("help_config"))
    ap.add_argument("--no-config", action="store_true", help=M("help_no_config"))
    ap.add_argument("--sarif", default=None, metavar="PATH", help=M("help_sarif"))
    ap.add_argument("--baseline", default=None, metavar="PATH", help=M("help_baseline"))


def _lint_one(path, a):
    """한 파일의 설정 해석→검사→필터→summary. main(단일)과 scan_main이 공유한다(0.5.0).
    사용 오류(설정·플래그)는 종전 계약대로 즉시 exit 2. --write-baseline이면 기록 후
    None을 반환한다. 반환 dict: errors/warns/ghost/summary/fail/profile/profile_excl/
    skip/cfg_path/baseline_suppressed/baseline_path."""
    if not os.path.exists(path):
        print(M("err_notfound") % path, file=sys.stderr)
        sys.exit(2)

    # 설정 파일(.archforge.json/.yml): CLI 플래그가 항상 이긴다(0.4.0).
    # 신뢰 경계(4차 리뷰): 덱 폴더의 설정이 게이트를 약화시킬 수 있으므로, 적용된 설정
    # 경로를 출력 계약(JSON summary.config, 텍스트 각주)에 남기고 --no-config로 끌 수 있다.
    cfg = {}
    cfg_path = None if a.no_config else _config.find_config(path, a.config)
    if a.config and not a.no_config and cfg_path is None:
        print(M("err_config") % a.config, file=sys.stderr)
        sys.exit(2)
    if cfg_path:
        try:
            cfg, cfg_warns = _config.load_config(cfg_path)
            for wmsg in cfg_warns:
                print("archforge: %s (%s)" % (wmsg, cfg_path), file=sys.stderr)
        except Exception as e:
            print(M("err_config") % ("%s (%s)" % (cfg_path, e)), file=sys.stderr)
            sys.exit(2)

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
        print(M("err_config") % ("threshold: %s" % e), file=sys.stderr)
        sys.exit(2)
    # 범위 검증: --hard-min 0 이 E3(판독 불가 차단)를 조용히 끄던 우회의 봉합(구 X1)
    if hard_min <= 0 or body_min <= 0 or small_min <= 0 or not (0 < w6_sim <= 1) or w6_cluster < 1:
        print(M("err_config") % ("threshold out of range (hard_min/body_min/small_min > 0, "
                                 "0 < w6_sim <= 1, w6_cluster >= 1)"), file=sys.stderr)
        sys.exit(2)
    profile = pick(a.profile, "profile", DEFAULT_PROFILE)
    if profile not in PROFILES:
        print(M("err_config") % ("profile=%r" % profile), file=sys.stderr)
        sys.exit(2)
    lang_final = pick(a.lang, "lang", None)
    if lang_final:
        set_lang(lang_final)
    baseline_path = pick(a.baseline, "baseline", None)
    skip_raw = pick(a.skip, "skip", "")
    if isinstance(skip_raw, list):
        skip_raw = ",".join(str(c) for c in skip_raw)

    ghost = [] if (a.ghost or a.json) else None
    try:
        errors, warns = lint(path, hard_min, body_min, small_min, render_dir=a.render,
                             ghost=ghost, strict=a.strict, w6_sim=w6_sim,
                             w6_min_cluster=w6_cluster, profile=profile)
    except Exception as e:
        print(M("err_open") % (path, type(e).__name__), file=sys.stderr)
        sys.exit(2)

    # baseline 기록 모드: 현재 위반(불완전성 신호 W18 제외)을 지문으로 저장하고 종료
    if getattr(a, "write_baseline", None):
        n = _config.write_baseline(a.write_baseline,
                                   [f for f in list(errors) + list(warns) if f.code != "W18"],
                                   profile=profile, lang=get_lang())
        print(M("baseline_written") % (n, a.write_baseline))
        return None

    # 검사 불완전 여부는 필터 전에 확정: W18을 --skip해도 기계 판독 신호는 남는다
    has_w18 = any(w[1] == "W18" for w in warns)
    baseline_suppressed = 0
    if baseline_path:
        try:
            known = _config.load_baseline(baseline_path)
        except Exception as e:
            print(M("err_config") % ("baseline %s (%s)" % (baseline_path, e)), file=sys.stderr)
            sys.exit(2)
        errors, s1 = _config.apply_baseline(errors, known)
        warns, s2 = _config.apply_baseline(warns, known)
        baseline_suppressed = s1 + s2
    skip = {c.strip().upper() for c in skip_raw.split(",") if c.strip()}
    # --skip은 WARN 전용: E코드까지 조용히 삼키면 배포 차단 게이트가 흔적 없이 꺼지는
    # 풋건이 된다(2차 외부 재점검 확정). 적용된 skip은 JSON summary에 기록해 흔적을 남긴다.
    bad_skip = sorted(c for c in skip if not c.startswith("W"))
    if bad_skip:
        print(M("err_skip_e") % ",".join(bad_skip), file=sys.stderr)
        sys.exit(2)
    # 존재하지 않는 코드 거부(오타가 조용히 통과하면 CI가 정상처럼 보임: 3차 리뷰 P1)
    unknown_skip = sorted(c for c in skip if c not in ALL_CODES)
    if unknown_skip:
        print(M("err_skip_unknown") % ",".join(unknown_skip), file=sys.stderr)
        sys.exit(2)
    # W18은 검사 불완전성 신호라 억제 대상이 아니다(3차 리뷰 P1)
    if "W18" in skip:
        print(M("err_skip_w18"), file=sys.stderr)
        sys.exit(2)
    # 프로파일 제외는 엔진 단계에서 이미 실행되지 않았다(0.3.1). 여기선 --skip만 필터.
    profile_excl = PROFILES[profile]
    excluded = skip | profile_excl
    if skip:
        warns = [w for w in warns if w[1] not in skip]

    summary = {"error_count": len(errors), "warn_count": len(warns),
               "pass": not errors and not (a.strict and warns),
               "incomplete": has_w18,
               "profile": profile,
               "skipped_codes": sorted(excluded),
               "baseline_suppressed": baseline_suppressed,
               "config": cfg_path}   # 어떤 설정이 게이트를 조정했는지 항상 가시화(신뢰 경계)

    return {"errors": errors, "warns": warns, "ghost": ghost, "summary": summary,
            "fail": bool(errors or (a.strict and warns)),
            "profile": profile, "profile_excl": profile_excl, "skip": skip,
            "cfg_path": cfg_path, "baseline_suppressed": baseline_suppressed,
            "baseline_path": baseline_path}


def _expand_scan_paths(patterns):
    """scan 인자 확장: 디렉터리는 재귀 .pptx, 글롭 패턴은 glob, 그 외는 파일 경로 그대로.
    PowerPoint 잠금 파일(~$*.pptx)은 제외, 중복은 순서 유지로 제거."""
    out = []
    for pat in patterns:
        if os.path.isdir(pat):
            for root, _dirs, files in os.walk(pat):
                for fn in sorted(files):
                    if fn.lower().endswith(".pptx") and not fn.startswith("~$"):
                        out.append(os.path.join(root, fn))
        elif any(ch in pat for ch in "*?["):
            for p in sorted(glob.glob(pat, recursive=True)):
                if p.lower().endswith(".pptx") and not os.path.basename(p).startswith("~$"):
                    out.append(p)
        else:
            out.append(pat)
    seen, uniq = set(), []
    for p in out:
        k = os.path.normcase(os.path.abspath(p))
        if k not in seen:
            seen.add(k)
            uniq.append(p)
    return uniq


def scan_main(argv=None):
    """`archforge scan PATHS...`: 여러 파일·디렉터리·글롭을 한 번에 린트(0.5.0, CI·
    pre-commit용). 파일별 판정은 단일 모드와 동일 경로(_lint_one)를 지나며, exit는
    하나라도 실패면 1. 매치 0건은 조용한 통과가 아니라 exit 2다(CI 풋건 방지)."""
    ap = argparse.ArgumentParser(prog="archforge scan", description=M("scan_desc"))
    ap.add_argument("paths", nargs="+", help=M("help_scan_paths"))
    _add_common_flags(ap)
    # 단일 모드 전용 플래그는 미지원: _lint_one 호환용 기본값만 심는다
    # (--render는 페이지 렌더 폴더가 덱마다 달라 scan에선 의미가 없다)
    ap.set_defaults(render=None, write_baseline=None)
    a = ap.parse_args(argv)

    files = _expand_scan_paths(a.paths)
    if not files:
        print(M("err_scan_none") % " ".join(a.paths), file=sys.stderr)
        return 2

    results = []   # (path, res)
    for path in files:
        results.append((path, _lint_one(path, a)))

    failed = sum(1 for (_p, r) in results if r["fail"])

    if a.sarif:
        import json
        sarif_doc = _reporters.build_sarif_multi(
            [(p, r["errors"], r["warns"]) for (p, r) in results])
        with open(a.sarif, "w", encoding="utf-8", newline="\n") as f:
            json.dump(sarif_doc, f, ensure_ascii=False, indent=2)

    if a.json:
        import json
        docs = [_reporters.build_json_doc(p, r["errors"], r["warns"], r["ghost"], r["summary"])
                for (p, r) in results]
        agg = {"schema_version": "1.0",
               "tool": {"name": "archforge", "version": _reporters._tool_version()},
               "lang": get_lang(),
               "scan": {"inputs": list(a.paths)},
               "files": docs,
               "summary": {"file_count": len(results), "failed_files": failed,
                           "error_count": sum(len(r["errors"]) for (_p, r) in results),
                           "warn_count": sum(len(r["warns"]) for (_p, r) in results),
                           "pass": failed == 0,
                           "incomplete": any(r["summary"]["incomplete"] for (_p, r) in results)}}
        print(json.dumps(agg, ensure_ascii=False, indent=2))
        return 1 if failed else 0

    for p, r in results:
        for line in _reporters.render_text(p, r["errors"], r["warns"], r["ghost"],
                                           r["profile"], r["profile_excl"], r["skip"],
                                           config_path=r["cfg_path"],
                                           baseline_suppressed=r["baseline_suppressed"],
                                           baseline_path=r["baseline_path"]):
            print(line)
        print()
    print(M("scan_summary") % (len(results), failed))
    return 1 if failed else 0


def demo_main(argv=None):
    """`archforge demo`: 결함 심은 덱과 교정본을 생성해 즉석에서 린트(0.5.0 온보딩).
    설치 30초 안에 무엇을 잡는지 직접 보여주는 첫 실행 경험용이다."""
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
    _demo.build_broken(broken)
    _demo.build_fixed(fixed)
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
            rc = 1   # 교정본은 항상 클린이어야 한다(테스트로 고정된 계약)
    print(M("demo_next") % broken)
    return rc


if __name__ == "__main__":
    main()
