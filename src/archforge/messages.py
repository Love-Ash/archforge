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
        "ko": "한글에 지정된 한글용 폰트에 정작 한글 글자가 없습니다. 보는 사람 환경의 기본 폰트로 바뀌어 표시됩니다",
        "en": "Hangul run's a:ea is a Latin-only font (no Hangul glyphs; silent Malgun fallback)",
    },
    "e1_theme_ea": {
        "ko": "한글용 폰트가 지정되지 않았고, 테마의 기본 한글 폰트에도 한글 글자가 없습니다. 환경에 따라 다른 폰트로 표시됩니다",
        "en": "No ea on Hangul run and the theme a:ea is Latin-only (Malgun fallback)",
    },
    "e1_latin_empty_theme": {
        "ko": "한글에 라틴 전용 폰트만 지정되어 있습니다. 한글 글자가 없어 보는 사람 환경의 기본 폰트로 대체됩니다",
        "en": "Hangul text set in a Latin-only font with an empty theme a:ea slot (Malgun fallback)",
    },
    "e1_nofont": {
        "ko": "이 한글에는 폰트가 전혀 지정되어 있지 않습니다. 어떤 폰트로 보일지는 여는 환경이 정하게 됩니다",
        "en": "No font anywhere on Hangul run and empty theme a:ea slot (guaranteed Malgun fallback)",
    },
    "e1_cjk_other": {
        "ko": "가나·한자 텍스트에 해당 글자가 없는 폰트가 지정되어 있어 환경에 따라 다르게 표시됩니다",
        "en": "Kana/Hanzi text set in a font without those glyphs (OS fallback)",
    },
    "e2": {
        "ko": "문장 부호로 쓰인 긴 대시(—)가 있습니다. 쉼표나 쌍점으로 바꾸는 편이 자연스럽습니다",
        "en": "Dash-family character in rendered text",
    },
    "e3": {
        "ko": "글자가 %.1fpt로, 읽기 하한 %.1fpt보다 작아 사실상 읽을 수 없습니다%s",
        "en": "Effective font %.1fpt below hard floor %.1fpt (unreadable)%s",
    },
    "e3_note": {
        "ko": " (지정 크기 %.1fpt가 자동 맞춤으로 %.2f배 줄어든 결과)",
        "en": " (nominal %.1f * autofit %.2f)",
    },
    "e4": {
        "ko": "한글·한자 자간이 %d만큼 벌어져 있습니다. 한글은 자간을 띄우면 어색해 보이므로 되돌리는 편이 좋습니다",
        "en": "Positive tracking %d on a Hangul/Hanja run (over 0.5pt; letter-spacing damage)",
    },
    # ---- W gates
    "w1": {
        "ko": "본문 크기 글상자의 글자가 %.1fpt로, 권장 하한 %.1fpt보다 작습니다. 출처 표기나 캡션이라면 무시해도 됩니다",
        "en": "Body-class frame text %.1fpt below recommended %.1fpt (ignore for sources/captions)",
    },
    "w5": {
        "ko": "이 텍스트의 글자 크기가 파일 어디에도 지정되어 있지 않습니다. 여는 프로그램에 따라 다르게 표시될 수 있습니다",
        "en": "Font size not found on run, paragraph, or anywhere in the inheritance chain",
    },
    "w6": {
        "ko": "같은 레이아웃 골격이 %d개 페이지에서 반복됩니다. 의도한 템플릿이라면 무시해도 되지만, 자동 생성 덱에서 흔한 흔적이기도 합니다",
        "en": "Layout skeleton repeated across %d pages (AI-deck tell; sparse dividers excluded; ignore for intentional template systems)",
    },
    "w7": {
        "ko": "이미지 위 글자의 명암 대비가 %.1f:1로 낮아 읽기 어렵습니다. 글자 색을 바꾸거나 이미지 위에 어두운 반투명 층을 깔아 보세요",
        "en": "Low text-over-image contrast %.1f:1 (add a scrim or adjust colors)",
    },
    "w8": {
        "ko": "좁은 틀 안의 글자가 %.1fpt로, %.1fpt보다 작습니다. 기기 목업이나 카드 속 작은 글씨로 보이는데, 읽혀야 하는 내용이라면 키워야 합니다",
        "en": "Small CJK %.1fpt below %.1fpt in a narrow frame (likely mockup/card text; readability risk; ignore for captions)",
    },
    "w9": {
        "ko": "강조색 세로 막대 %d개가 목록 구분에 반복 사용됐습니다. 자동 생성 덱의 흔한 장식이라, 괘선이나 여백으로 바꾸면 인상이 정돈됩니다",
        "en": "%d accent vertical bars repeated as list markers (structure with rules/whitespace/type; keep accent to a single dot)",
    },
    "w10": {
        "ko": "직접 그린 도식이 %d개 페이지에서 거의 그대로 반복됩니다. 의도한 반복인지 눈으로 확인해 보세요",
        "en": "Hand-drawn diagram repeated nearly identically on %d pages (confirm by eye: intentional sequence or lazy reuse)",
    },
    "w11_buzz": {
        "ko": "AI가 쓴 글에서 흔한 상투어가 %d종 발견됐습니다. 문구를 직접 다듬어 보세요",
        "en": "%d AI-tell buzzword type(s) (consider rewriting the copy)",
    },
    "w11_open": {
        "ko": "표지나 도입부가 흔한 상투구로 시작합니다. 첫 문장은 직접 쓰는 편이 좋습니다",
        "en": "Stock opening cliche (open the deck in your own voice)",
    },
    "w12": {
        "ko": "이 덱의 다른 페이지들은 하단 요소를 위에서 %.2fin 위치에 맞춰 두었는데, 거기서 살짝 벗어난 페이지가 %d개 있습니다. 정렬 실수로 보입니다",
        "en": "Footer baseline drift: house baseline %.2fin, %d page(s) slightly off (suspected alignment slip)",
    },
    "w13": {
        "ko": "파워포인트 기본 효과(그림자·광선·입체)가 %d곳, %d개 페이지에 쓰였습니다. 요즘 덱에서는 낡아 보이기 쉬우니 의도한 스타일인지 확인하세요",
        "en": "%d native PowerPoint effects across %d pages (shadow/glow/3D read dated; ignore if intentional)",
    },
    "w14": {
        "ko": "제목 %d개(전체 %d개 중)가 내용을 설명만 하는 명사구입니다. 그 장의 주장을 담은 문장형 제목으로 바꾸면 전달력이 좋아집니다. 표지·목차 같은 구조 슬라이드는 세지 않았습니다",
        "en": "%d/%d titles are nominal phrases (use action titles for read decks; ignore for editorial headlines)",
    },
    "w15": {
        "ko": "텍스트 상자 두 개가 겹쳐 보입니다(추정 겹침 %.0f%%). 실제 화면에서 확인해 하나를 옮기거나 줄이세요",
        "en": "Estimated text-on-text overlap %.0f%% (occlusion/collision; verify on render)",
    },
    "w16": {
        "ko": "내용이 슬라이드 밖으로 %.2fin 나가 있습니다. 잘려 보일 수 있으니 화면에서 확인하세요",
        "en": "Off-canvas overflow %.2fin (clipping; verify on render)",
    },
    "w16_text": {"ko": "텍스트 %r", "en": "text %r"},
    "w16_pic": {"ko": "그림 %.1fx%.1fin", "en": "picture %.1fx%.1fin"},
    "w17": {
        "ko": "글자가 이미지 가장자리에 걸쳐 있습니다(이미지 안쪽 %.0f%%). 잘린 것처럼 보일 수 있으니 화면에서 확인하세요",
        "en": "Text straddles an image ink edge (%.0f%% inside; may look clipped; verify on render)",
    },
    "w19": {
        "ko": "글자 색과 상자 배경색의 명암 대비가 %.1f:1로, 거의 안 보이는 수준입니다. 글자 색이나 배경색을 바꿔 주세요",
        "en": "Text-on-fill contrast %.1f:1, close to invisible (change the text or fill color)",
    },
    "w20": {
        "ko": "글자가 아래 도형 위에 %.0f%% 걸쳐 있고 대비가 %.1f:1이라 묻혀서 읽기 어렵습니다. 글자를 도형 밖으로 옮기거나 색을 바꿔 주세요",
        "en": "Text sits on the shape beneath it (%.0f%% covered) at %.1f:1 contrast and is buried (move it off the shape or change the color)",
    },
    "w20_bg": {
        "ko": "글자가 슬라이드 배경 위에 그대로 놓여 있는데 배경과의 대비가 %.1f:1이라 사실상 안 보입니다. 글자 색이나 배경색을 바꿔 주세요",
        "en": "Text sits directly on the slide background at %.1f:1 contrast, close to invisible (change the text or background color)",
    },
    "w18_page": {
        "ko": "이 페이지의 일부 구간은 손상되거나 비정형인 속성 때문에 검사하지 못했습니다. 결과가 불완전할 수 있습니다",
        "en": "Some spans on this page could not be checked (malformed/atypical attributes); results may be incomplete",
    },
    "w18_deck": {
        "ko": "덱 전체 대상 검사 중 일부를 수행하지 못했습니다. 결과가 불완전할 수 있습니다",
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
    # explain used to borrow err_skip_unknown, so a typo'd code was reported as a problem
    # with --skip, a flag the subcommand does not even accept.
    "err_unknown_code": {
        "ko": "archforge: 존재하지 않는 규칙 코드입니다(오타 확인): %s. 전체 목록은 archforge rules",
        "en": "archforge: unknown rule code (check for typos): %s. See archforge rules for the full list",
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
    # Subcommands are dispatched by hand rather than by an argparse subparser, so argparse
    # cannot list them and this string is the only place they are advertised. It named
    # three of seven; fix, explain, rules and baseline all worked but appeared nowhere in
    # --help, including the auto-remediation path.
    "prog_desc": {
        "ko": "빌드된 .pptx를 배포 전에 기계로 검사하는 한글 특화 품질 린터\n"
              "서브커맨드: scan = 다중 파일/디렉터리, fix = 안전한 결함 자동 수정, "
              "demo = 첫 실행 투어, rules = 규칙 목록, explain = 규칙 하나 설명, "
              "baseline = 기존 위반 수용 파일 관리, skill = 에이전트 스킬팩",
        "en": "Preflight quality linter for built .pptx files, deep CJK font coverage\n"
              "subcommands: scan = many files/dirs, fix = auto-correct the safe defects, "
              "demo = first-run tour, rules = list the rules, explain = describe one rule, "
              "baseline = manage the accepted-violation file, skill = agent skill pack",
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
    "help_schema": {
        "ko": "JSON 스키마 버전(1.0 기본=errors/warnings 분리, 2.0=findings[] 단일배열+severity+data+capabilities+abstentions)",
        "en": "JSON schema version (1.0 default: split errors/warnings; 2.0: single findings[] with severity+data, plus capabilities and abstentions)",
    },
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
    "err_out_path": {
        "ko": "archforge: 출력 경로의 상위 폴더가 없습니다: %s",
        "en": "archforge: the parent directory of the output path does not exist: %s",
    },
    "help_timeout": {
        "ko": "전체 실행에 초 단위 벽시계 제한. 자식 프로세스로 격리해 시간 초과 시 exit 124(악성·병적 덱이 CI를 멈추는 것 방지)",
        "en": "wall-clock limit in seconds for the whole run, isolated in a child process; exit 124 on timeout (keeps a hostile/pathological deck from hanging CI)",
    },
    "err_timeout": {
        "ko": "archforge: 제한 시간 %.1f초를 초과해 중단했습니다",
        "en": "archforge: aborted after exceeding the %.1fs timeout",
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
    "help_html": {
        "ko": "주석 시각 리포트를 이 경로에 기록(슬라이드 와이어프레임+판정 박스, 단일 정적 HTML)",
        "en": "write the annotated visual report to this path (slide wireframes + finding boxes, one static HTML file)",
    },
    "fix_desc": {
        "ko": "기계적으로 안전한 세 규칙(E1 폰트·E2 대시·E4 자간)만 결정적으로 자동 수정. 수정 후 재린트 권장",
        "en": "deterministic auto-fixes for the three mechanically safe rules (E1 font, E2 dash, E4 tracking); re-lint afterwards",
    },
    "help_fix_output": {
        "ko": "수정본을 저장할 경로(원본은 건드리지 않음)",
        "en": "where to write the fixed copy (the original is never touched)",
    },
    "help_fix_rules": {
        "ko": "적용할 규칙(기본 E1,E2,E4. 그 외 규칙은 레이아웃 판단이 필요해 자동수정 제외)",
        "en": "which rules to apply (default E1,E2,E4; everything else needs layout judgment and stays find-only)",
    },
    "help_fix_ea": {
        "ko": "E1 수정에 쓸 한글 폰트(기본: 맑은 고딕)",
        "en": "the Hangul-capable font for E1 fixes (default: Malgun Gothic)",
    },
    "err_fix_rules": {
        "ko": "archforge: 자동수정 불가 규칙: %s (가능: %s)",
        "en": "archforge: not auto-fixable: %s (fixable: %s)",
    },
    "fix_summary": {
        "ko": "%d건 수정 → %s (재린트로 확인하세요)",
        "en": "%d fix(es) applied -> %s (re-lint to verify)",
    },
    "baseline_desc": {
        "ko": "baseline 파일 점검: 무엇을 억제하는지, 어떤 기록 조건인지",
        "en": "inspect a baseline file: what it suppresses, under which recorded conditions",
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
    "fix_w19": {"ko": "글자 색이나 상자 배경색을 바꿔 대비를 2:1 이상으로. 같은 색 유령 텍스트는 삭제", "en": "Change the text or fill color to reach at least 2:1; delete same-color ghost text"},
    "fix_w20": {"ko": "글자를 도형 밖으로 옮기거나 글자·도형 색을 바꿔 대비를 2:1 이상으로", "en": "Move the text off the shape, or change either color to reach at least 2:1"},
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
    "rollup_header": {
        "ko": "-- 같은 원인 묶음 (상세는 아래 개별 행) --",
        "en": "-- grouped by shared cause (details below) --",
    },
    "rollup_line": {
        "ko": "%s x%d",
        "en": "%s x%d",
    },
    "rollup_line_cause": {
        "ko": "%s x%d, 최다 원인 %s (%d건)",
        "en": "%s x%d, dominant cause %s (%d hits)",
    },
    "demo_built": {
        "ko": "archforge: 데모 덱 생성 완료 -> %s (broken.pptx = 결함 6종, fixed.pptx = 교정본, 검사는 --profile full 기준)",
        "en": "archforge: demo decks written -> %s (broken.pptx = seeded defects, fixed.pptx = the corrected version; linted with --profile full)",
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
