# -*- coding: utf-8 -*-
"""User-facing message catalog (en/ko). 0.3.0 globalization layer.

Principle: gate codes (E1, W15) and JSON keys are a language-independent stable
contract; only the human-read messages carry a language. Report language follows the
user, not the deck (a Korean user checking an English deck still gets a Korean report):
priority is --lang > ARCHFORGE_LANG > system locale > en. Format args (%s, %d) must
match in order, not just count, across both language templates.
"""
import contextvars
import os
from typing import Optional

# A context variable, not a global: in threaded/async embedded-library usage, language
# does not leak across tenants (fixes a bug from a measured reproduction of an
# adversarial-panel race condition). CLI behavior is unchanged.
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
    # ---- Diagnostics (stderr)
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
    # ---- Report scaffolding
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
    # ---- CLI errors
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
    # ---- CLI help
    "prog_desc": {
        "ko": "빌드된 .pptx를 배포 전에 기계로 검사하는 한글 특화 품질 린터 (서브커맨드: scan = 다중 파일/디렉터리, demo = 첫 실행 투어, skill = 에이전트 스킬팩)",
        "en": "Preflight quality linter for built .pptx files, deep CJK font coverage (subcommands: scan = many files/dirs, demo = first-run tour, skill = agent skill pack)",
    },
    "help_hard_min": {"ko": "E3 판독 불가 하한(pt, 기본 5.0)", "en": "E3 unreadable hard floor in pt (default 5.0)"},
    "help_body_min": {"ko": "W1 본문급 권장 하한(pt, 기본 9.0)", "en": "W1 body-class recommended floor in pt (default 9.0)"},
    "help_strict": {
        "ko": "--fail-on-warning + --fail-incomplete + --e2-no-exemptions 세 정책의 합집합(호환 별칭)",
        "en": "union of --fail-on-warning, --fail-incomplete, and --e2-no-exemptions (compatibility alias)",
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
        "ko": "규칙 프리셋: core(기본, 객관 결함만) / full(전부: AI 티·스타일 규칙 포함) / editorial(에디토리얼 덱: W6·W14 제외). 제외 내역은 JSON에 기록",
        "en": "rule preset: core (default, objective defects only) / full (everything incl. AI-tell/style rules) / editorial (editorial decks; W6/W14 off). Exclusions are recorded in JSON",
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
    "help_config": {
        "ko": "설정 파일 경로(기본: 덱 폴더/현재 폴더의 .archforge.json|.yml). CLI 플래그가 설정을 이김",
        "en": "config file path (default: .archforge.json|.yml in the deck dir or cwd). CLI flags override config",
    },
    "help_sarif": {
        "ko": "SARIF 2.1.0 결과를 이 경로에 기록(GitHub code scanning 연동)",
        "en": "write SARIF 2.1.0 results to this path (GitHub code scanning)",
    },
    "help_baseline": {
        "ko": "baseline 파일의 기존 위반을 억제하고 신규만 보고(W18 제외, 억제 수는 summary에 기록)",
        "en": "suppress findings recorded in this baseline file and report only new ones (W18 exempt; count recorded in summary)",
    },
    "help_write_baseline": {
        "ko": "현재 위반을 baseline 파일로 기록하고 종료(기존 덱을 있는 그대로 수용하는 도입 경로)",
        "en": "record current findings to a baseline file and exit (adoption path for existing decks)",
    },
    "err_config": {
        "ko": "archforge: 설정을 읽을 수 없습니다: %s",
        "en": "archforge: cannot read config: %s",
    },
    "baseline_written": {
        "ko": "archforge: baseline 기록 완료(%d건) -> %s",
        "en": "archforge: baseline written (%d finding(s)) -> %s",
    },
    "help_no_config": {
        "ko": "설정 파일 자동 탐색을 끔(신뢰할 수 없는 출처의 덱을 린트할 때)",
        "en": "disable config file discovery (when linting decks from untrusted sources)",
    },
    "config_applied": {
        "ko": "  (설정 적용: %s)",
        "en": "  (config applied: %s)",
    },
    "baseline_applied": {
        "ko": "  (baseline 억제: %d건, %s)",
        "en": "  (baseline suppressed: %d finding(s), %s)",
    },
    # ---- scan/demo subcommands (0.5.0)
    "scan_desc": {
        "ko": "여러 파일·디렉터리·글롭을 한 번에 린트(CI·pre-commit용). 하나라도 실패면 exit 1",
        "en": "lint multiple files, directories, or globs in one run (CI/pre-commit). Exits 1 if any file fails",
    },
    "help_scan_paths": {
        "ko": "pptx 파일, 디렉터리(재귀), 글롭 패턴(예: decks/**/*.pptx)의 나열",
        "en": "any mix of .pptx files, directories (recursive), and glob patterns (e.g. decks/**/*.pptx)",
    },
    "err_scan_none": {
        "ko": "archforge: 매치되는 .pptx 가 없습니다(조용한 통과 방지, exit 2): %s",
        "en": "archforge: no .pptx files matched (refusing to silently pass, exit 2): %s",
    },
    "scan_summary": {
        "ko": "=== 스캔 요약: 파일 %d개, 실패 %d개 ===",
        "en": "=== scan summary: %d file(s), %d failed ===",
    },
    "err_scan_baseline": {
        "ko": "archforge: scan에서 --baseline은 대상이 1개 파일일 때만 허용됩니다(지문에 파일 정체성이 없어 덱 A의 수용이 덱 B의 동일 결함을 숨김). 덱별 baseline은 각 덱 폴더의 설정 파일로 지정하세요",
        "en": "archforge: --baseline under scan is allowed only when exactly one file matched (fingerprints carry no file identity, so deck A's acceptance would suppress the same finding in deck B). Point each deck at its own baseline via its folder config",
    },
    "scan_file_error": {
        "ko": "== 오류: %s == %s (이 파일은 건너뛰고 스캔을 계속했습니다)",
        "en": "== ERROR: %s == %s (file skipped; scan continued)",
    },
    "note_baseline_meta": {
        "ko": "archforge: 경고: baseline의 기록 조건과 현재 실행이 다릅니다(%s: 기록 %r, 현재 %r). 억제 결과가 기대와 다를 수 있습니다",
        "en": "archforge: warning: baseline was recorded under different conditions (%s: recorded %r, current %r); suppression may not match expectations",
    },
    "help_fail_on_warning": {
        "ko": "WARN이 하나라도 있으면 exit 1 (권고를 차단으로 승격)",
        "en": "exit 1 if any WARN is present (promotes advisories to blockers)",
    },
    "help_fail_incomplete": {
        "ko": "검사 불완전(W18/summary.incomplete)이면 exit 1. CI 게이트에 권장",
        "en": "exit 1 when checking was incomplete (W18/summary.incomplete). Recommended for CI gates",
    },
    "help_e2_no_exemptions": {
        "ko": "E2 숫자 맥락 예외(범위·음수) 해제. E2가 실행되는 프로파일(full)에서만 의미",
        "en": "lift E2's numeric-context exemptions (ranges, minus). Only meaningful in profiles that run E2 (full)",
    },
    "help_junit": {
        "ko": "JUnit XML 결과를 이 경로에 기록(Jenkins·GitLab 등 테스트 리포트 UI 연동)",
        "en": "write JUnit XML results to this path (Jenkins/GitLab test-report UIs)",
    },
    "help_allow_empty_pattern": {
        "ko": "매치 0건인 입력 패턴을 허용(기본은 exit 2: 오타·빌드 실패가 다른 패턴 뒤에 숨는 것 방지)",
        "en": "allow an input pattern that matched nothing (default exits 2 so a typo cannot hide behind another pattern)",
    },
    "err_scan_pattern_empty": {
        "ko": "archforge: 다음 입력이 아무 .pptx도 매치하지 못했습니다(--allow-empty-pattern으로 허용 가능): %s",
        "en": "archforge: these inputs matched no .pptx (use --allow-empty-pattern to permit): %s",
    },
    "rules_desc": {
        "ko": "규칙 한 줄 요약 목록(코드·심각도·카테고리·프로파일)",
        "en": "one-line summary of every rule (code, severity, category, profiles)",
    },
    "explain_desc": {
        "ko": "규칙 하나의 의미·발화 조건·수정법 설명",
        "en": "what one rule means, when it fires, and how to fix it",
    },
    # ---- per-code fix guidance (explain subcommand + docs/rules generation)
    "fix_e1": {
        "ko": "한글 run의 a:ea에 한글 폰트를 명시하세요. font.name(a:latin)만으로는 테마 ea가 비어있을 때만 통합니다",
        "en": "Set a CJK-capable font on the run's a:ea slot. font.name (a:latin) alone only works when the theme ea slot is empty",
    },
    "fix_e2": {
        "ko": "산문 대시는 콜론·쉼표·괄호·줄바꿈으로 바꾸세요. 숫자 범위와 음수는 기본 통과합니다",
        "en": "Replace prose dashes with a colon, comma, parentheses, or a line break. Numeric ranges and minus signs pass by default",
    },
    "fix_e3": {
        "ko": "숫자만 키우지 말고 재설계하세요: 항목을 줄이고 대표 요소 하나를 키우는 쪽이 맞습니다",
        "en": "Redesign instead of bumping the number: fewer items, one representative element bigger",
    },
    "fix_e4": {
        "ko": "한글 run의 자간(spc)을 0으로. 트래킹은 ASCII 전용 라벨에만 쓰세요",
        "en": "Set tracking (spc) to 0 on Hangul runs; track ASCII-only labels only",
    },
    "fix_w1": {"ko": "출처·캡션이면 무시, 본문이면 9pt 이상으로", "en": "Ignore for sources/captions; raise to 9pt+ for body text"},
    "fix_w5": {"ko": "게이트가 측정할 수 있게 크기를 명시하세요", "en": "Set sizes explicitly so the gates can measure"},
    "fix_w6": {"ko": "페이지마다 그리드를 다르게. 의도된 템플릿 시스템이면 --w6-sim/--w6-cluster 조정 또는 --skip W6", "en": "Vary the grid per page; for an intentional template system tune --w6-sim/--w6-cluster or --skip W6"},
    "fix_w7": {"ko": "스크림을 깔거나 이미지 쪽 밝기를 조정하세요", "en": "Add a scrim or darken/lighten the image side"},
    "fix_w8": {"ko": "목업 안 라벨은 밖의 콜아웃으로 빼세요", "en": "Move labels out of the mockup into callouts"},
    "fix_w9": {"ko": "구조는 괘선·여백·활자로, 색 세로바 대신 점 하나로", "en": "Structure with rules/whitespace/type; use a dot instead of colored bars"},
    "fix_w10": {"ko": "재탕인지 의도된 반복인지 눈으로 확정 후 재설계", "en": "Confirm by eye whether it is reuse or intent, then redesign"},
    "fix_w11": {"ko": "덱의 자기 목소리로 카피를 다시 쓰세요", "en": "Rewrite the copy in the deck's own voice"},
    "fix_w12": {"ko": "푸터를 한 baseline에 정렬하세요", "en": "Align footers to one baseline"},
    "fix_w13": {"ko": "네이티브 그림자·글로·3D를 제거하세요(올드 티)", "en": "Remove native shadow/glow/3D effects; they read dated"},
    "fix_w14": {"ko": "액션 타이틀로 다시 쓰세요(--ghost 목록이 이야기로 읽혀야). 에디토리얼 덱은 --skip W14", "en": "Rewrite as action titles (the --ghost list should read as a story). Editorial decks: --skip W14"},
    "fix_w15": {"ko": "렌더에서 확인 후 한쪽을 이동·축소하세요", "en": "Check the render, then move or shrink one frame"},
    "fix_w16": {"ko": "내용을 캔버스 안으로. 장식 도형 블리드는 자동 제외됩니다", "en": "Pull content inside the canvas; decorative shape bleed is auto-excluded"},
    "fix_w17": {"ko": "캡션을 이미지 위나 밖으로 완전히 옮기세요", "en": "Move the caption fully on or off the image"},
    "fix_w18": {"ko": "검사 못 한 구간이 있습니다. stderr에서 원인을 보고 원본을 고친 뒤 재린트하세요. CI는 --fail-incomplete 권장", "en": "Part of the deck went unchecked; see stderr for why, fix the malformed source, re-lint. Use --fail-incomplete in CI"},
    "subcmd_conflict": {
        "ko": "archforge: 참고: 현재 폴더에 %r 파일이 있지만 서브커맨드를 실행합니다. 그 파일을 린트하려면 `archforge ./%s`",
        "en": "archforge: note: a file named %r exists here, but the subcommand runs. To lint that file use `archforge ./%s`",
    },
    "demo_desc": {
        "ko": "결함을 심은 데모 덱(broken.pptx)과 교정본(fixed.pptx)을 생성해 즉석에서 린트(첫 실행 경험)",
        "en": "generate a defect-seeded demo deck (broken.pptx) and its fix (fixed.pptx), then lint both (first-run tour)",
    },
    "help_demo_dir": {
        "ko": "데모 덱을 생성할 폴더(기본 ./archforge-demo)",
        "en": "directory to write the demo decks to (default ./archforge-demo)",
    },
    "demo_built": {
        "ko": "archforge: 데모 덱 생성 완료 -> %s (broken.pptx = 결함 6종, fixed.pptx = 교정본)",
        "en": "archforge: demo decks written -> %s (broken.pptx = seeded defects, fixed.pptx = the corrected version)",
    },
    "demo_next": {
        "ko": "다음: 직접 돌려보세요. archforge %s --profile full --json (기계 생성 덱 검사는 full 프로파일)",
        "en": "next: run it yourself. archforge %s --profile full --json (machine-made decks want the full profile)",
    },
}


def detect_lang() -> str:
    """ARCHFORGE_LANG > LANG/LC_ALL > system locale > en."""
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
    """Explicitly set the language (ko/en). None reverts to auto-detection. Returns the
    resolved language."""
    _LANG.set(lang if lang in ("ko", "en") else None)
    return get_lang()


def get_lang() -> str:
    v = _LANG.get()
    return v if v is not None else detect_lang()


def M(msg_id: str) -> str:
    """Message template for the current language. The caller fills format args with %."""
    entry = MESSAGES[msg_id]
    return entry.get(get_lang(), entry["en"])
