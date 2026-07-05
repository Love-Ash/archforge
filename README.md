<div align="center">

# Archforge

**한글 덱이 조용히 깨지는 곳을 배포 전에 잡아내는 PPTX 품질 린터**

아치포지 · arch(구조) + forge(대장간)

![python](https://img.shields.io/badge/python-3.9%2B-3776AB)
![license](https://img.shields.io/badge/license-MIT-green)
![status](https://img.shields.io/badge/status-beta-orange)

</div>

---

빌드된 `.pptx`를 배포 직전에 기계로 검사합니다. 파워포인트 없이 파일만으로 동작하고, 한글
타이포그래피를 일급으로 다룹니다. 이름은 덱의 구조와 한글 타이포를 배포 전에 벼려 다듬는
대장간이라는 뜻입니다.

## 왜

한글 덱이 깨지는 지점은 대부분 조용합니다.

- 라틴 전용 폰트에 실린 한글은 에러 없이 Malgun으로 폴백됩니다.
- 자간(tracking)은 한글 낱자 사이를 소리 없이 벌려 놓습니다.
- autofit은 글자를 판독 불가 크기까지 줄여 놓습니다.

코드 리뷰로는 안 보이고 렌더를 눈으로 봐야만 잡히던 이 결함들을, Archforge는 pptx 내부
(XML, 폰트 슬롯, 좌표, 이미지 알파)를 직접 읽어 빌드 직후에 잡습니다.

## 설치

```bash
pip install archforge          # 배포 전이면
pip install -e .               # 리포에서 직접
```

## 사용

```bash
archforge deck.pptx                 # 사람용 리포트, ERROR 있으면 exit 1
archforge deck.pptx --json          # 기계 판독용 JSON (에이전트·CI)
archforge deck.pptx --strict        # WARN도 실패로 취급
archforge deck.pptx --ghost         # 페이지별 타이틀 나열 (수평 논리 검토)
archforge deck.pptx --render pages/ # p01.png·p02.png 형식 렌더로 이미지 위 대비(W7)까지
```

`--json` 출력:

```json
{
  "file": "deck.pptx",
  "errors":   [{ "page": 3, "code": "E1", "message": "...", "detail": "..." }],
  "warnings": [{ "page": 5, "code": "W15", "message": "...", "detail": "..." }],
  "summary":  { "error_count": 1, "warn_count": 2, "pass": false }
}
```

## 무엇을 잡나

**ERROR** (배포 차단, exit 1)

| 코드 | 내용 | 고치는 법 |
|:----:|------|-----------|
| `E1` | 한글이 라틴 전용 폰트 슬롯이나 빈 테마 `a:ea` 슬롯으로 렌더됨 (조용한 Malgun 폴백) | 한글 런에 CJK 폰트를, 모노는 영문·숫자에만 |
| `E2` | 대시류 문자 (em·en·figure dash, 수학 마이너스 U+2212, 전각 하이픈 U+FF0D 등) | 콜론·쉼표·괄호로, 마이너스는 ASCII 하이픈으로 |
| `E3` | 실효 크기(autofit·상속 반영) 5pt 미만: 판독 불가 | 요소를 줄이고 대표 하나를 크게 |
| `E4` | 연속 한글에 양수 자간: 낱자가 벌어짐 | CJK 런은 자간 0, ASCII 라벨만 트래킹 |

**WARN** (권고)

| 코드 | 내용 |
|:----:|------|
| `W1` | 본문급 프레임이 9pt 미만 |
| `W5` | 폰트 크기를 못 찾음 (상속 미해석) |
| `W6` | 같은 레이아웃 골격이 4장 이상 반복 |
| `W7` | 이미지 위 텍스트 대비 낮음 (`--render` 필요) |
| `W8` | 좁은 프레임(≤4in)의 소형 한글 (목업·카드 내부) |
| `W9` | 색 세로바를 리스트 마커로 반복 |
| `W10` | 직접 그린 도식이 여러 페이지에서 반복 |
| `W11` | AI 티 카피 (버즈워드·뻔한 오프닝) |
| `W12` | 푸터 baseline 어긋남 |
| `W13` | PPT 자체 그림자·글로·3D 효과 |
| `W14` | 서술형 명사구 타이틀 (액션타이틀 아님) |
| `W15` | 텍스트끼리 겹침 |
| `W16` | 화면 밖 넘침 |
| `W17` | 텍스트가 이미지 잉크 경계에 걸침 |

## 작동 방식

`W15`~`W17`은 프레임 박스가 아니라 실효 글리프·잉크 영역을 근사해서 봅니다. run별 크기,
실제 행간, autofit(퍼센트 문자열 포함), wrap 모드, 그룹 좌표 변환, 정렬, 이미지 알파
트림·크롭·flip까지 반영하고, 드롭캡·잔상 타이포·장식 블리드·카드 위 캡션 같은 의도적
연출은 제외합니다.

임계값은 취향이 아니라 캘리브레이션 결과입니다. 실덱 코퍼스 50여 개를 전수 스캔해 플래그된
페이지를 실제 렌더와 대조하고, 진짜 결함과 오탐이 갈리는 지점에 임계를 놓았습니다. 이후
적대적 검증(재현 pptx로 오탐을 공격)으로 그룹 이동, wrap=none, 팔레트 투명 PNG, 크롭·회전·flip
같은 외부 pptx의 다양성까지 회귀 픽스처로 고정했습니다.

> 알려진 한계: 레이아웃 lstStyle로 정렬을 상속하는 플레이스홀더는 좌정렬로 후퇴 판정합니다.
> 템플릿 덱에서 `W15`/`W16`이 이상해 보이면 렌더로 확정하세요.

## 에이전트 연동

LLM 에이전트가 python-pptx 류로 덱을 만드는 워크플로를 일차 사용자로 설계했습니다.

```
빌드 → archforge --json → error_count 0 될 때까지 수정 → WARN은 렌더 보고 판단
```

`skills/pptx-lint/`의 Agent Skills 스킬팩이 이 루프와 코드별 수정 가이드를 에이전트에게
가르칩니다. SKILL.md + YAML frontmatter 표준이라 지원하는 어느 에이전트(Claude Code, Codex 등)든
그대로 설치됩니다.

린트 통과가 완성이라는 뜻은 아닙니다. 이 린터는 기계로 잡히는 결함군을 담당하고, 페이지 구성과
서사의 품질은 여전히 렌더를 보는 눈의 몫입니다.

## English

**Archforge** is a CLI quality linter for built `.pptx` files with first-class Korean
typography awareness. It catches the silent failure modes of Korean decks (Latin-only
font fallback, CJK tracking damage, sub-5pt autofit shrinkage) plus the general mechanical
defect class (text collisions, off-canvas bleed, text straddling image edges, AI-generated
deck tells) by reading the pptx internals directly, no PowerPoint required.

Thresholds are calibrated against rendered output of a real deck corpus and hardened with
adversarial reproduction fixtures (group transforms, `wrap="none"`, palette-transparency
PNGs, crop/flip/rotation). Built for LLM-agent build-lint-fix loops: stable codes, `--json`
output, and an Agent Skills pack in `skills/pptx-lint/`.

## License

MIT © Minjae Kwon (Ash)
