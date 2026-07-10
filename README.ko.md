<div align="center">

# Archforge

**한글 덱이 조용히 깨지는 곳을 배포 전에 잡아내는 PPTX 품질 린터**

아치포지 · arch(구조) + forge(대장간)

![python](https://img.shields.io/badge/python-3.9%2B-3776AB)
![license](https://img.shields.io/badge/license-MIT-green)
![status](https://img.shields.io/badge/status-beta-orange)
[![ci](https://github.com/Love-Ash/archforge/actions/workflows/ci.yml/badge.svg)](https://github.com/Love-Ash/archforge/actions/workflows/ci.yml)

[English README](README.md)

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
pip install archforge          # PyPI
pip install -e .               # 리포에서 직접(개발용)
```

파이썬 3.9 이상이면 되고, python-pptx와 Pillow만 딸려 옵니다. 파워포인트는 필요 없습니다.

에이전트 스킬팩도 wheel에 동봉돼 있어 설치 후 바로 받을 수 있습니다.

```bash
archforge skill                  # 스킬팩(SKILL.md) 출력
archforge skill --install        # ./.claude/skills/archforge-pptx-lint/ 에 설치
archforge skill --install DIR    # 원하는 스킬 폴더에 설치
```

## 사용

```bash
archforge deck.pptx                 # 사람용 리포트, ERROR 있으면 exit 1
archforge deck.pptx --json          # 기계 판독용 JSON (에이전트·CI)
archforge deck.pptx --strict        # WARN도 실패 + E2 숫자 맥락 예외 해제
archforge deck.pptx --ghost         # 페이지별 타이틀 나열 (수평 논리 검토)
archforge deck.pptx --render pages/ # p01.png·p02.png 형식 렌더로 이미지 위 대비(W7)까지
archforge deck.pptx --profile core  # 객관 결함만(스타일·관행 규칙 제외)
archforge deck.pptx --skip W14,W6   # 특정 WARN 억제(JSON에 기록됨)
archforge deck.pptx --lang ko       # 리포트 언어(기본: ARCHFORGE_LANG, 그다음 OS 로케일)
archforge deck.pptx --hard-min 5 --body-min 9 --small-min 7.5   # 크기 게이트 임계 조정
archforge deck.pptx --w6-sim 0.95 --w6-cluster 5                # W6 골격 반복 임계 조정
```

`--json` 출력:

```json
{
  "file": "deck.pptx",
  "lang": "ko",
  "errors":   [{ "page": 3, "code": "E1", "message": "...", "detail": "..." }],
  "warnings": [{ "page": 5, "code": "W15", "message": "...", "detail": "..." }],
  "ghost":    [{ "page": 1, "title": "..." }],
  "summary":  { "error_count": 1, "warn_count": 2, "pass": false,
                "incomplete": false, "profile": "full", "skipped_codes": [] }
}
```

코드(E1, W15)는 언어 무관 안정 식별자이고, 메시지 언어는 덱이 아니라 사용자를 따릅니다
(`--lang` > `ARCHFORGE_LANG` > OS 로케일).

## 무엇을 잡나

**ERROR** (배포 차단, exit 1)

| 코드 | 내용 |
|:----:|------|
| `E1` | 한글을 실제로 렌더할 폰트가 라틴 전용(한글 글리프 없음): 조용한 Malgun 폴백. 실효 폰트는 실측 렌더 모델로 해석(아래) |
| `E2` | 대시류 문자를 문장 부호로 사용(AI 생성 덱 1번 티). en dash로 쓴 숫자 범위(2020~2024, Q1~Q3, 5%~10% 류)와 음수 부호는 기본 통과, `--strict`는 전부 차단 |
| `E3` | 실효 크기(autofit·문단·placeholder 상속 체인 반영) 5pt 미만: 판독 불가 |
| `E4` | 연속 한글에 양수 자간: 낱자가 벌어짐(한글 한정: 가나 자간은 일본어의 정상 관행) |

**WARN** (권고)

| 코드 | 내용 |
|:----:|------|
| `W1` | 본문급 프레임이 9pt 미만 |
| `W5` | 상속 체인 어디에도 크기 없음 |
| `W6` | 같은 레이아웃 골격이 4장 이상 반복 (`--w6-sim`/`--w6-cluster`로 조정) |
| `W7` | 이미지 위 텍스트 대비 낮음 (`--render` 필요) |
| `W8` | 좁은 프레임(≤4in)의 소형 CJK (목업·카드 내부) |
| `W9` | 색 세로바를 리스트 마커로 반복 |
| `W10` | 직접 그린 도식이 여러 페이지에서 반복 |
| `W11` | AI 티 카피 (버즈워드·뻔한 오프닝) |
| `W12` | 푸터 baseline 어긋남 |
| `W13` | PPT 자체 그림자·글로·3D 효과 |
| `W14` | 서술형 명사구 타이틀 (숫자+단위 타이틀은 주장으로 인정) |
| `W15` | 텍스트끼리 겹침 |
| `W16` | 화면 밖 넘침 |
| `W17` | 텍스트가 이미지 잉크 경계에 걸침 |
| `W18` | 손상·비정형 속성으로 일부 구간 검사 불능 (결과 불완전 가능, `--strict`면 실패) |

프로파일이 객관 결함과 스타일 정책을 분리합니다. `core`는 기계적 게이트만(E1/E3/E4,
W1/W5/W7/W8, W15~W18), `editorial`은 에디토리얼·포트폴리오 덱을 위해 W6/W14 제외,
`full`(기본)은 전부. 프로파일 선택은 JSON summary에 기록되므로 조용한 우회가 아닙니다.

## 작동 방식

`E1`의 폰트 해석은 규격 추정이 아니라 실측입니다. PowerPoint COM으로 프로브 덱을 렌더해
확정한 우선순위: run `a:ea` > lstStyle 상속 체인(도형 > 레이아웃 ph > 마스터 ph > 마스터
txStyles > defaultTextStyle) > 테마 ea(제목 placeholder는 majorFont, 그 외 minorFont;
비어있지 않으면 run `a:latin`보다 우선) > 빈 테마일 때만 run `a:latin` > OS 폴백.
기록은 [docs/CALIBRATION.md](docs/CALIBRATION.md)에 있고, 테마는 슬라이드가 실제로 쓰는
마스터의 관계로 해석하므로 멀티마스터 덱에서도 엉뚱한 테마로 판정하지 않습니다.

실효 크기도 같은 상속 체인을 해석해서, 명시 크기 없는 템플릿·placeholder 덱에서도
`E3`/`W1`/`W8`이 실제로 돕니다.

`W15`~`W17`은 프레임 박스가 아니라 실효 글리프·잉크 영역을 근사해서 봅니다. run별 크기,
실제 행간, autofit(퍼센트 문자열 포함), wrap 모드, 그룹 좌표 변환, 정렬, 이미지 알파
트림·크롭·flip까지 반영하고, 드롭캡·잔상 타이포·장식 블리드·카드 위 캡션 같은 의도적
연출은 제외합니다.

임계값은 취향이 아니라 캘리브레이션 결과입니다. 실덱 코퍼스 50여 개를 전수 스캔해 렌더와
대조하고, 적대적 검증(오탐을 공격하는 재현 pptx)으로 회귀 픽스처를 고정했습니다. 게이트별
임계와 근거, 실측 기록은 [docs/CALIBRATION.md](docs/CALIBRATION.md)에 있습니다.

임의 pptx가 들어와도 리포트는 살아남습니다. run·슬라이드 단위 가드가 쓰레기 속성을
흡수하고, 가드가 삼킨 구간은 `W18`과 기계 판독용 `summary.incomplete` 플래그로
표면화됩니다. CI는 `pass`와 `incomplete`를 함께 보거나 `--strict`(W18도 실패)로
돌리세요. `pass` 단독은 ERROR만 반영합니다.

> 범위에 대한 정직한 주석. 폰트 커버리지 지식(E1/E4)은 현재 한글 심층입니다. 다른
> 스크립트로 쓰인 런은 절대 오탐하지 않고(런 단위 유니코드 스크립트 판별), 일본어·중국어
> 심층 지원은 스크립트별 커버리지 표를 추가하는 방식으로 확장합니다. 세로쓰기와
> RTL·복잡 조판 스크립트의 기하 추정은 추측하지 않고 스킵 후 W18로 알립니다. 렌더 모델의
> 타겟은 PowerPoint for Windows입니다(Mac·웹·LibreOffice는 폰트 해석이 다를 수 있음).

## 에이전트 연동

LLM 에이전트가 python-pptx 류로 덱을 만드는 워크플로를 일차 사용자로 설계했습니다.

```
빌드 → archforge --json → error_count 0 될 때까지 수정 → WARN은 렌더 보고 판단
```

Agent Skills 스킬팩(SKILL.md + YAML frontmatter 표준)이 이 루프와 코드별 수정 가이드를
에이전트에게 가르칩니다. wheel에 동봉되므로 `pip install archforge` 후
`archforge skill --install`이면 끝이고, 리포를 클론했다면 `skills/archforge-pptx-lint/`를
그대로 써도 됩니다.

린트 통과가 완성이라는 뜻은 아닙니다. 이 린터는 기계로 잡히는 결함군을 담당하고, 페이지 구성과
서사의 품질은 여전히 렌더를 보는 눈의 몫입니다.

## License

MIT © Minjae Kwon (Ash)
