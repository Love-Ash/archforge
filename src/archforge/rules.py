# -*- coding: utf-8 -*-
"""규칙 레지스트리(0.4.0, 3차 외부 리뷰 구조 개편의 1단계).

규칙의 메타데이터(심각도·분류·메시지 id)와 프로파일 정의를 검사 구현에서 분리한다.
CLI 검증(--skip 오타), 프로파일, SARIF rules[] 생성이 전부 이 표 하나를 소비한다.
검사 구현 자체의 물리적 분해(rules/typography 등 패키지화)는 다음 단계다.
"""

# code -> (severity, category, 대표 메시지 id)
RULES = {
    "E1": ("error",   "typography", "e1_nofont"),
    "E2": ("error",   "style",      "e2"),
    "E3": ("error",   "typography", "e3"),
    "E4": ("error",   "typography", "e4"),
    "W1": ("warning", "typography", "w1"),
    "W5": ("warning", "typography", "w5"),
    "W6": ("warning", "structure",  "w6"),
    "W7": ("warning", "render",     "w7"),
    "W8": ("warning", "typography", "w8"),
    "W9": ("warning", "style",      "w9"),
    "W10": ("warning", "structure", "w10"),
    "W11": ("warning", "style",     "w11_buzz"),
    "W12": ("warning", "structure", "w12"),
    "W13": ("warning", "style",     "w13"),
    "W14": ("warning", "style",     "w14"),
    "W15": ("warning", "geometry",  "w15"),
    "W16": ("warning", "geometry",  "w16"),
    "W17": ("warning", "geometry",  "w17"),
    "W18": ("warning", "meta",      "w18_page"),
}

ALL_CODES = frozenset(RULES)

# 프로파일 = 엔진 실행 정책(0.3.1부터 제외 규칙은 실행 자체를 안 함).
# 0.4.0: 기본 프로파일이 core로 바뀜(파괴적 변경). 객관 결함만 기본이고,
# AI 티·하우스 스타일 규칙(E2, W6, W9~W14)은 full 옵트인이다.
PROFILES = {
    "full": frozenset(),
    "core": frozenset({"E2", "W6", "W9", "W10", "W11", "W12", "W13", "W14"}),
    "editorial": frozenset({"W6", "W14"}),
}

DEFAULT_PROFILE = "core"


def severity(code: str) -> str:
    return RULES.get(code, ("warning",))[0]


def category(code: str) -> str:
    return RULES.get(code, ("", "unknown"))[1]
