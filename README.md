# jangpyo (장표)

빌드된 `.pptx`를 배포 전에 기계로 검사하는 품질 린터입니다. 파워포인트 없이 파일만으로
동작하고, 한글 타이포그래피를 일급으로 다룹니다.

한글 덱이 깨지는 지점은 대부분 조용합니다. 라틴 전용 폰트에 실린 한글은 에러 없이 Malgun으로
폴백되고, 자간(tracking)은 한글 낱자 사이를 소리 없이 벌려 놓고, autofit은 글자를 판독 불가
크기까지 줄여 놓습니다. 코드 리뷰로는 보이지 않고 렌더를 눈으로 봐야만 잡히던 이 결함들을,
jangpyo는 pptx 내부(XML·폰트 슬롯·좌표·이미지 알파)를 직접 읽어 빌드 직후에 잡습니다.

LLM 에이전트가 python-pptx 류로 덱을 생성하는 워크플로를 일차 사용자로 설계했습니다.
`--json` 출력과 안정된 코드 체계로 어느 에이전트(Claude Code, Codex 등)든 빌드-검사-수정
루프를 돌 수 있고, `skills/pptx-lint/`의 Agent Skills 스킬팩을 설치하면 에이전트가 코드별
수정 방법까지 알고 움직입니다.

## 설치와 사용

```bash
pip install jangpyo        # (배포 전이면) pip install -e <repo>

jangpyo deck.pptx                 # 사람용 리포트, ERROR 있으면 exit 1
jangpyo deck.pptx --json          # 기계 판독용(에이전트·CI)
jangpyo deck.pptx --strict        # WARN도 실패로
jangpyo deck.pptx --ghost         # 페이지별 타이틀만 나열(수평 논리 검토)
jangpyo deck.pptx --render pages/ # p01.png·p02.png 형식 렌더가 있으면 이미지 위 대비(W7)까지
```

`--render` 는 페이지별 렌더 PNG를 `p01.png`, `p02.png` 형식(페이지 번호 2자리)으로 찾습니다.
다른 이름이면 W7이 조용히 건너뛰어지고, 폴더에 png가 있는데 규약과 다르면 stderr로 알립니다.

## 무엇을 잡나

ERROR(배포 차단) 4종:

| 코드 | 내용 |
|------|------|
| E1 | 한글이 라틴 전용 폰트 슬롯(모노 계열 등)이나 빈 테마 `a:ea` 슬롯으로 렌더됨: 조용한 Malgun 폴백 |
| E2 | 대시류 문자가 본문에 포함(em·en·figure dash와 함께 수학 마이너스 U+2212, 전각 하이픈 U+FF0D, 박스 罫線 U+2500 등도 차단). 마이너스 기호는 ASCII 하이픈으로 |
| E3 | 실효 크기(autofit·상속 반영) 5pt 미만: 판독 불가 |
| E4 | 연속 한글에 양수 자간: 낱자가 벌어짐 |

WARN(권고) 14종: 본문 소형(W1), 크기 미상(W5), 레이아웃 재탕(W6), 이미지 위 저대비(W7),
좁은 프레임 소형 한글(W8), 색 세로바 마커 반복(W9), 도식 클론(W10), AI 티 카피(W11),
푸터 어긋남(W12), 그림자·글로·3D(W13), 명사구 타이틀(W14), 텍스트 겹침(W15),
화면 밖 넘침(W16), 이미지 경계 걸침(W17).

W15~W17은 프레임 박스가 아니라 실효 글리프·잉크 영역을 근사해서 봅니다. run별 크기,
실제 행간, autofit(퍼센트 문자열 포함), wrap 모드, 그룹 좌표 변환, 정렬, 이미지 알파
트림·크롭·flip까지 반영하고, 드롭캡·잔상 타이포·장식 블리드·카드 위 캡션 같은 의도적
연출은 제외합니다.

## 임계값은 실측으로 정했습니다

겹침 45%, 넘침 0.15in 같은 숫자는 취향이 아니라 캘리브레이션 결과입니다. 실덱 코퍼스
50여 개를 전수 스캔하고, 플래그된 페이지를 실제 렌더와 대조해 진짜 결함과 오탐이 갈리는
지점에 임계를 놓았습니다. 이후 적대적 검증(재현 pptx를 만들어 오탐을 공격)으로 그룹 이동,
wrap=none, 팔레트 투명 PNG, 크롭·회전·flip 같은 외부 pptx의 다양성까지 회귀 픽스처로
고정했습니다.

알려진 한계: 레이아웃 lstStyle로 정렬을 상속하는 플레이스홀더는 좌정렬로 후퇴 판정합니다.
템플릿 덱에서 W15/W16이 이상해 보이면 렌더로 확정하세요.

## 에이전트 연동

```
빌드 → jangpyo --json → error_count 0 될 때까지 수정 루프 → WARN은 렌더 보고 판단
```

`skills/pptx-lint/SKILL.md`가 이 루프와 코드별 수정 가이드를 에이전트에게 가르칩니다.
Agent Skills 표준(SKILL.md + YAML frontmatter)이라 지원하는 어느 에이전트에든 그대로
설치됩니다.

린트 통과가 완성이라는 뜻은 아닙니다. 이 린터는 기계로 잡히는 결함군을 담당하고,
페이지 구성과 서사의 품질은 여전히 렌더를 보는 눈의 몫입니다.

## English summary

jangpyo is a CLI quality linter for built `.pptx` files with first-class Korean
typography awareness. It catches the silent failure modes of Korean decks (Latin-only
font fallback, CJK tracking damage, sub-5pt autofit shrinkage) plus the general
mechanical defect class (text collisions, off-canvas bleed, text straddling image
edges, AI-generated deck tells) by reading the pptx internals directly, no PowerPoint
required. Thresholds are calibrated against rendered output of a real deck corpus and
hardened with adversarial reproduction fixtures (group transforms, wrap="none",
palette-transparency PNGs, crop/flip/rotation). Built for LLM-agent build-lint-fix
loops: stable codes, `--json` output, and an Agent Skills pack in `skills/pptx-lint/`.

## License

MIT
