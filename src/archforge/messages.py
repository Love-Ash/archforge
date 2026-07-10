# -*- coding: utf-8 -*-
"""사용자 대면 메시지 카탈로그(en/ko). 0.3.0 세계화 레이어.

원칙: 게이트 코드(E1, W15)와 JSON 키는 언어 무관 안정 계약이고, 사람 읽는 메시지만
언어를 탄다. 리포트 언어는 덱이 아니라 사용자를 따른다(한국 사용자가 영어 덱을
검사해도 리포트는 한국어): 우선순위 --lang > ARCHFORGE_LANG > 시스템 로케일 > en.
포맷 인자(%s, %d)는 두 언어 템플릿이 순서까지 동일해야 한다.
"""
import contextvars
import os
from typing import Optional

# 전역 아닌 컨텍스트 변수: 스레드·async로 임베드된 라이브러리 사용에서 언어가 테넌트 간
# 새지 않는다(적대 패널 경쟁 조건 실측 재현의 교정). CLI 동작은 동일.
_LANG: "contextvars.ContextVar[Optional[str]]" = contextvars.ContextVar("archforge_lang", default=None)

MESSAGES = {
    # ---- E gates
    "e1_run_ea": {
        "ko": "한글 런의 a:ea가 라틴 전용 폰트(한글 글리프 없음, Malgun 폴백)",
        "en": "Hangul run's a:ea is a Latin-only font (no Hangul glyphs; silent Malgun fallback)",
    },
    "e1_theme_ea": {
        "ko": "한글 런에 ea 미지정 + 테마 a:ea가 라틴 전용(Malgun 폴백)",
        "en": "No ea on Hangul run and the theme a:ea is Latin-only (Malgun fallback)",
    },
    "e1_latin_empty_theme": {
        "ko": "한글 텍스트가 라틴 전용 폰트 지정 + 테마 a:ea 빈 슬롯(Malgun 폴백)",
        "en": "Hangul text set in a Latin-only font with an empty theme a:ea slot (Malgun fallback)",
    },
    "e1_nofont": {
        "ko": "한글 런에 폰트 미지정 + 테마 a:ea 빈 슬롯(Malgun 폴백 확정)",
        "en": "No font anywhere on Hangul run and empty theme a:ea slot (guaranteed Malgun fallback)",
    },
    "e1_cjk_other": {
        "ko": "가나·한자 텍스트가 해당 글리프 없는 폰트로 지정됨(OS 폴백)",
        "en": "Kana/Hanzi text set in a font without those glyphs (OS fallback)",
    },
    "e2": {
        "ko": "긴 대시류 문자 포함",
        "en": "Dash-family character in rendered text",
    },
    "e3": {
        "ko": "실효 폰트 %.1fpt < 하한 %.1fpt(판독 불가)%s",
        "en": "Effective font %.1fpt below hard floor %.1fpt (unreadable)%s",
    },
    "e3_note": {
        "ko": " (명목 %.1f * autofit %.2f)",
        "en": " (nominal %.1f * autofit %.2f)",
    },
    "e4": {
        "ko": "한글·한자 run 에 양수 트래킹 %d(0.5pt 초과, 자간 벌어짐)",
        "en": "Positive tracking %d on a Hangul/Hanja run (over 0.5pt; letter-spacing damage)",
    },
    # ---- W gates
    "w1": {
        "ko": "본문급 프레임 글자 %.1fpt < 권장 %.1fpt(출처·캡션이면 무시)",
        "en": "Body-class frame text %.1fpt below recommended %.1fpt (ignore for sources/captions)",
    },
    "w5": {
        "ko": "폰트 크기를 run·문단·상속 체인 어디에서도 못 찾음",
        "en": "Font size not found on run, paragraph, or anywhere in the inheritance chain",
    },
    "w6": {
        "ko": "레이아웃 골격이 %d개 페이지에서 반복됨(클로드 티 의심, 성긴 디바이더 제외, 의도된 템플릿 시스템이면 무시)",
        "en": "Layout skeleton repeated across %d pages (AI-deck tell; sparse dividers excluded; ignore for intentional template systems)",
    },
    "w7": {
        "ko": "이미지 위 텍스트 대비 낮음 %.1f:1 (스크림·색 보강 필요)",
        "en": "Low text-over-image contrast %.1f:1 (add a scrim or adjust colors)",
    },
    "w8": {
        "ko": "소형 CJK %.1fpt < %.1fpt(목업·카드 내부 추정, 판독 위험, 캡션이면 무시)",
        "en": "Small CJK %.1fpt below %.1fpt in a narrow frame (likely mockup/card text; readability risk; ignore for captions)",
    },
    "w9": {
        "ko": "accent 세로바 %d개로 항목 구조 반복(구조는 괘선·여백·활자로, accent는 점 하나로)",
        "en": "%d accent vertical bars repeated as list markers (structure with rules/whitespace/type; keep accent to a single dot)",
    },
    "w10": {
        "ko": "직접 그린 도식이 %d개 페이지에서 거의 동일 반복(의도된 교육 시퀀스인지 게으른 재탕인지 눈으로 확정)",
        "en": "Hand-drawn diagram repeated nearly identically on %d pages (confirm by eye: intentional sequence or lazy reuse)",
    },
    "w11_buzz": {
        "ko": "AI 티 버즈워드 %d종(카피 재작성 검토)",
        "en": "%d AI-tell buzzword type(s) (consider rewriting the copy)",
    },
    "w11_open": {
        "ko": "뻔한 오프닝 상투구(표지·도입은 자기 목소리로)",
        "en": "Stock opening cliche (open the deck in your own voice)",
    },
    "w12": {
        "ko": "푸터 baseline 어긋남: 하우스 %.2fin 대비 살짝 다른 페이지 %d개(정렬 실수 의심)",
        "en": "Footer baseline drift: house baseline %.2fin, %d page(s) slightly off (suspected alignment slip)",
    },
    "w13": {
        "ko": "PPT 자체 효과 %d개 / %d개 페이지(그림자·글로·3D = 올드 티 의심, 의도한 스타일이면 무시)",
        "en": "%d native PowerPoint effects across %d pages (shadow/glow/3D read dated; ignore if intentional)",
    },
    "w14": {
        "ko": "타이틀 %d/%d개가 서술형 명사구(read 덱은 액션타이틀로, 에디토리얼·작품 소개 헤드라인은 무시)",
        "en": "%d/%d titles are nominal phrases (use action titles for read decks; ignore for editorial headlines)",
    },
    "w15": {
        "ko": "텍스트끼리 겹침 추정 %.0f%%(가림·충돌, 렌더 확인)",
        "en": "Estimated text-on-text overlap %.0f%% (occlusion/collision; verify on render)",
    },
    "w16": {
        "ko": "화면 밖 넘침 %.2fin(잘림, 렌더 확인)",
        "en": "Off-canvas overflow %.2fin (clipping; verify on render)",
    },
    "w16_text": {"ko": "텍스트 %r", "en": "text %r"},
    "w16_pic": {"ko": "그림 %.1fx%.1fin", "en": "picture %.1fx%.1fin"},
    "w17": {
        "ko": "텍스트가 이미지 경계에 걸침(안 %.0f%%, 잘려 보임 의심, 렌더 확인)",
        "en": "Text straddles an image ink edge (%.0f%% inside; may look clipped; verify on render)",
    },
    "w18_page": {
        "ko": "일부 구간을 검사하지 못함(손상·비정형 속성): 이 페이지 결과는 불완전할 수 있음",
        "en": "Some spans on this page could not be checked (malformed/atypical attributes); results may be incomplete",
    },
    "w18_deck": {
        "ko": "덱 단위 일부 검사를 수행하지 못함(손상·비정형 구조): 결과가 불완전할 수 있음",
        "en": "Some deck-level checks could not run (malformed/atypical structure); results may be incomplete",
    },
    "w6_detail": {"ko": "예 %s", "en": "e.g. %s"},
    "w10_detail": {"ko": "페이지 %s", "en": "pages %s"},
    # ---- 진단(stderr)
    "note_theme_parse": {
        "ko": "theme parse 실패 마스터 있음: E1 테마 판정이 빈 슬롯 가정으로 후퇴",
        "en": "a master's theme failed to parse: E1 theme judgment falls back to the empty-slot assumption",
    },
    "note_render_dir_missing": {
        "ko": "--render 폴더가 없음: W7 이미지 대비 검사를 수행하지 못함(%s)",
        "en": "--render folder not found: the W7 on-image contrast check could not run (%s)",
    },
    "note_render_naming": {
        "ko": "W7 참고: %s 에 p01.png·p02.png 형식 렌더가 없어 이미지 대비 검사를 건너뜀(현재 파일: %s ...)",
        "en": "W7 note: no p01.png/p02.png-named renders in %s; on-image contrast check skipped (found: %s ...)",
    },
    # ---- 리포트 스캐폴딩
    "ghost_header": {
        "ko": "--- ghost deck (제목만 읽기: 주장이 이야기로 흐르는가) ---",
        "en": "--- ghost deck (read titles top to bottom: does the argument flow?) ---",
    },
    "skip_applied": {
        "ko": "  (--skip 적용: %s)",
        "en": "  (--skip applied: %s)",
    },
    "profile_applied": {
        "ko": "  (--profile %s 적용, 제외 코드: %s)",
        "en": "  (--profile %s applied; excluded codes: %s)",
    },
    # ---- CLI 오류
    "err_notfound": {
        "ko": "archforge: 파일을 찾을 수 없습니다: %s",
        "en": "archforge: file not found: %s",
    },
    "err_open": {
        "ko": "archforge: pptx 를 열 수 없습니다(유효한 .pptx 인지 확인): %s (%s)",
        "en": "archforge: cannot open pptx (check that it is a valid .pptx): %s (%s)",
    },
    "err_skip_e": {
        "ko": "archforge: --skip 은 WARN 코드 전용입니다(배포 차단 ERROR는 억제 불가): %s",
        "en": "archforge: --skip accepts WARN codes only (deploy-blocking ERRORs cannot be suppressed): %s",
    },
    "err_skip_unknown": {
        "ko": "archforge: --skip 에 존재하지 않는 코드가 있습니다(오타 확인): %s",
        "en": "archforge: --skip contains unknown rule codes (check for typos): %s",
    },
    "err_skip_w18": {
        "ko": "archforge: W18은 검사 불완전성 신호라 --skip 으로 억제할 수 없습니다",
        "en": "archforge: W18 signals incomplete checking and cannot be suppressed with --skip",
    },
    "skill_installed": {
        "ko": "archforge: 스킬 설치 완료 -> %s",
        "en": "archforge: skill installed -> %s",
    },
    "skill_conflict": {
        "ko": "archforge: 참고: 현재 폴더에 'skill' 파일이 있지만 서브커맨드를 실행합니다. 그 파일을 린트하려면 `archforge ./skill`",
        "en": "archforge: note: a file named 'skill' exists here, but the subcommand runs. To lint that file use `archforge ./skill`",
    },
    # ---- CLI 도움말
    "prog_desc": {
        "ko": "빌드된 .pptx를 배포 전에 기계로 검사하는 한글 특화 품질 린터 (서브커맨드: archforge skill = 에이전트 스킬팩 출력/설치)",
        "en": "Korean-typography-aware quality linter for built .pptx files (subcommand: archforge skill = print/install the agent skill pack)",
    },
    "help_hard_min": {"ko": "E3 판독 불가 하한(pt, 기본 5.0)", "en": "E3 unreadable hard floor in pt (default 5.0)"},
    "help_body_min": {"ko": "W1 본문급 권장 하한(pt, 기본 9.0)", "en": "W1 body-class recommended floor in pt (default 9.0)"},
    "help_strict": {
        "ko": "WARN도 exit 1 + E2 숫자 맥락 예외(범위 en dash·음수 부호) 해제",
        "en": "WARNs also exit 1, and E2 numeric-context exemptions (range en dash, minus sign) are disabled",
    },
    "help_small_min": {"ko": "W8 좁은 프레임 소형 CJK 상한(pt)", "en": "W8 small-CJK ceiling in narrow frames (pt)"},
    "help_render": {
        "ko": "렌더 PNG 폴더(p01.png·p02.png 형식) 지정 시 이미지 위 텍스트 대비(W7) 검사 활성화",
        "en": "folder of rendered PNGs (p01.png, p02.png naming) to enable the on-image contrast check (W7)",
    },
    "help_ghost": {
        "ko": "고스트덱(페이지별 타이틀만 나열) 출력: 제목만 읽어 주장이 흐르는지 수평 논리 눈검수",
        "en": "print the ghost deck (per-page titles only) to eyeball the horizontal logic",
    },
    "help_json": {"ko": "기계 판독용 JSON 출력(에이전트·CI 연동)", "en": "machine-readable JSON output (agents/CI)"},
    "help_skip": {
        "ko": "억제할 WARN 코드 콤마 목록(예 --skip W14,W6). ERROR 코드는 불가, 적용 내역은 JSON summary.skipped_codes에 기록",
        "en": "comma list of WARN codes to suppress (e.g. --skip W14,W6). ERROR codes are refused; applied skips are recorded in JSON summary.skipped_codes",
    },
    "help_w6_sim": {"ko": "W6 골격 유사도 임계(기본 0.90)", "en": "W6 skeleton similarity threshold (default 0.90)"},
    "help_w6_cluster": {"ko": "W6 클러스터 최소 이웃 수(기본 3 = 같은 골격 4장+)", "en": "W6 minimum cluster neighbors (default 3 = 4+ pages sharing a skeleton)"},
    "help_profile": {
        "ko": "규칙 프리셋: full(기본, 전부) / core(객관 결함만: 스타일·관행 규칙 제외) / editorial(에디토리얼 덱: W6·W14 제외). 제외 내역은 JSON에 기록",
        "en": "rule preset: full (default, everything) / core (objective defects only; style/convention rules off) / editorial (editorial decks; W6/W14 off). Exclusions are recorded in JSON",
    },
    "help_lang": {
        "ko": "리포트 언어(ko/en). 기본은 ARCHFORGE_LANG 환경변수, 없으면 시스템 로케일",
        "en": "report language (ko/en). Defaults to ARCHFORGE_LANG env var, then system locale",
    },
    "skill_desc": {
        "ko": "동봉된 에이전트 스킬팩(SKILL.md)을 stdout으로 출력하거나 --install 로 설치",
        "en": "print the bundled agent skill pack (SKILL.md) to stdout, or install it with --install",
    },
    "help_skill_install": {
        "ko": "스킬을 DIR/archforge-pptx-lint/SKILL.md 로 설치(기본 DIR=./.claude/skills)",
        "en": "install the skill to DIR/archforge-pptx-lint/SKILL.md (default DIR=./.claude/skills)",
    },
    "help_skill_path": {"ko": "동봉 SKILL.md 의 패키지 내 경로만 출력", "en": "print only the in-package path of the bundled SKILL.md"},
}


def detect_lang() -> str:
    """ARCHFORGE_LANG > LANG/LC_ALL > 시스템 로케일 > en."""
    v = os.environ.get("ARCHFORGE_LANG", "").strip().lower()
    if v.startswith("ko"):
        return "ko"
    if v:
        return "en"
    for var in ("LC_ALL", "LANG"):
        lv = os.environ.get(var, "").lower()
        if lv.startswith("ko"):
            return "ko"
        if lv:
            return "en"
    try:
        import locale
        loc = (locale.getlocale()[0] or "").lower()
        if "ko" in loc or "korean" in loc:
            return "ko"
    except Exception:
        pass
    return "en"


def set_lang(lang: Optional[str]) -> str:
    """명시 언어 설정(ko/en). None이면 자동 감지로 되돌림. 확정 언어를 반환."""
    _LANG.set(lang if lang in ("ko", "en") else None)
    return get_lang()


def get_lang() -> str:
    v = _LANG.get()
    return v if v is not None else detect_lang()


def M(msg_id: str) -> str:
    """현재 언어의 메시지 템플릿. 포맷 인자는 호출부에서 % 로 채운다."""
    entry = MESSAGES[msg_id]
    return entry.get(get_lang(), entry["en"])
