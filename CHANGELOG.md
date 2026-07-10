# Changelog

## 0.3.0 (2026-07-10)

세계화 릴리스: 리포트 언어와 규칙 정책을 분리 가능한 레이어로 만들었습니다.

- 메시지 카탈로그 i18n(en/ko): 모든 게이트 메시지·CLI 도움말·오류가 영어/한국어 카탈로그로
  이동. 리포트 언어는 덱이 아니라 사용자를 따른다: `--lang` > `ARCHFORGE_LANG` >
  시스템 로케일 > en. 게이트 코드(E1, W15)와 JSON 키는 언어 무관 안정 계약이며 JSON에
  `lang` 필드 추가. 두 언어 템플릿의 포맷 인자 순서 일치는 테스트로 고정.
- 프로파일: `--profile full|core|editorial`. core=객관 결함만(E2 등 스타일·관행 규칙
  제외), editorial=W6·W14 제외. `--skip`(WARN 전용)과 달리 이름 붙은 공개 정책이라
  스타일성 ERROR 제외를 허용하되, 선택이 JSON `summary.profile`과 `skipped_codes`에
  기록된다(조용한 우회 아님). 외부 전략 리뷰의 "객관 결함과 하우스 스타일 분리" 반영.
- README 영어 전면화 + README.ko.md 분리(PyPI 페이지도 영어). 렌더 모델의 타겟
  렌더러(PowerPoint for Windows)를 문서에 명시.
- SKILL.md를 0.3.0 계약(lang·profile·skipped_codes, E2 v2, 한글 스코프)으로 갱신.
- 자체 3차 적대 패널(4관점) 확정분 반영:
  - 스크립트 레이어 2층 구조: 블록리스트 본대(Inter·Arial·모노 계열)는 CJK 전체가 없어
    가나·한자 판정에도 유효하고, JP/SC 서브셋(Noto Sans JP 류)만 한글 판정 전용.
    한자 전용 런의 E1·E4가 0.2.1에서 통째로 빠졌던 회귀 교정(한국 덱 인명·법률용어),
    가나 섞인 런만 E4 제외. 한글 유니코드 범위에 반각·자모 확장 A/B 추가,
    기하 스킵 스크립트에 티베트·미얀마·크메르 추가.
  - E2 봉합 4건: 2연속 대시 미탐, 단어+숫자 혼합 토큰(결론2024류)의 붙은 삽입구 우회,
    위첨자 각주 숫자 오인(숫자 판정을 ASCII·전각으로 한정), 띄어진 음수 부호 오탐.
  - 제목 자격을 명시 크기 또는 제목 패밀리 placeholder로 한정(도형 lstStyle·마스터
    bodyStyle 상속 크기가 본문 산문을 ghost/W14에 쓸어 담던 잔여 경로 봉합).
  - --lang 반복 지정 시 마지막 값 적용, `archforge skill --lang` 수용,
    `--lang ko skill` 순서 오해석 교정. 언어 상태를 contextvars로(스레드 임베드 안전).
  - `summary.incomplete` 신설: W18 여부의 기계 판독 신호. pass 단독은 ERROR만
    반영하므로 CI는 pass와 incomplete를 함께 보거나 --strict를 쓴다(문서 과장 교정).
- 테스트 56개에서 65개로.

수용한 한계(수정 안 함): argparse 자체 골격 문구(usage/error)는 영어 고정,
레거시 cp949 콘솔에서 UTF-8 출력은 깨져 보일 수 있음(크래시 방지가 우선).

## 0.2.1 (2026-07-10)

0.2.0에 대한 2차 외부 재점검(17에이전트, 확정 7건)을 전량 반영한 안전 릴리스입니다.
어떤 언어의 덱이 들어와도 오탐하지 않는 것(스크립트 레이어)이 주제입니다.

- 제목 수집 범람 회귀 수정: defaultTextStyle 폴백 18pt는 게이트엔 쓰되 제목 후보에서
  제외(크기 출처 구분). 크기 미지정 덱에서 --ghost가 본문 전문을 뱉고 W14가 오발화하던
  0.2.0 신규 회귀의 교정.
- 스크립트 레이어: E1·E4 트리거를 한글 포함 텍스트로 한정. 가나·한자 전용 런은 한글
  커버리지 지식으로 판정하지 않는다(Noto Sans JP/SC 등 일본어·중국어 덱 오차단 교정.
  일본어 가나 자간은 정상 관행이라 E4 대상 아님). 세로쓰기(bodyPr@vert)와 RTL·복잡
  조판 스크립트(아랍·히브리·인도계·태국)는 기하 추정을 스킵하고 W18로 표면화.
- E1 폰트 상속 체인: run rPr에 없는 폰트 슬롯을 lstStyle 체인(도형 → 레이아웃 ph →
  마스터 ph → 마스터 txStyles → defaultTextStyle)에서 해석. 마스터 lstStyle의 a:ea가
  실제 렌더에 상속됨과 제목 placeholder가 테마 majorFont ea를 탄다는 것을 COM 프로브로
  실측 확인 후 반영(docs/CALIBRATION.md 프로브6). 표준 기업 템플릿에서 나던
  "Malgun 폴백 확정" 오탐과 mj-ea 미사용 문제의 교정.
- E2 v2: 판별 축을 문자에서 기능으로. en dash는 양옆 토큰이 숫자성(숫자 포함: 2020,
  Q1, 5%, FY24)이면 띄어도 통과, 한쪽만 숫자성이면 붙었을 때만 통과(범위), 띄운
  삽입구와 단어 연결은 차단. 공백 범위·Q1 범위·% 범위·숫자+한글 연결의 잔존 오탐 4형태
  교정.
- (파괴적 변경) --skip을 WARN 코드 전용으로 제한(E코드는 exit 2 거부), 적용 내역을
  JSON summary.skipped_codes에 기록. 배포 차단 게이트가 흔적 없이 꺼지던 풋건의 교정.
  0.2.0에서 `--skip E2`를 쓰던 파이프라인은 `--profile core`로 이전.
- W18 배선 전체 확장: W6/W7/W9~W14 가드, 테마 파싱 실패, 덱 단위 검사(p00)까지.
  "가드가 삼킨 건 전부 출력에 표면화"라는 문서 주장이 이제 사실.
- 블록리스트: cjk가 이름에 든 폰트는 무조건 통과(Noto Sans Mono CJK KR 오탐 교정),
  Aptos(현 Office 기본)·Courier·PT Sans/Serif 추가.
- 테스트 47개에서 56개로.

## 0.2.0 (2026-07-10)

외부 심층 리뷰(6관점 + 적대 재검증)의 지적 전량을 반영하고, 반영분 자체를 다시
자체 적대 패널(8관점 리뷰 + 발견별 반박 검증, 22에이전트)에 통과시켜 확정 11건을
추가 수정한 릴리스입니다.

### 자체 적대 패널 라운드에서 추가된 것

- W18 신설: 가드가 삼킨 검사 불능 구간(손상·비정형 속성)을 stderr만이 아니라
  JSON·텍스트 출력에 경고로 표면화. `summary.pass`만 보는 CI가 불완전 검사를 완전
  통과로 오독하던 침묵 저하의 교정. `--strict`에선 exit 1로 승격.
- 코어 라인 게이트 가드를 프레임 단위에서 run 단위로: 한 run의 쓰레기 속성이 같은
  프레임 이웃 run의 진짜 E1 위반까지 삼켜 거짓 clean을 만들던 것 교정.
- 기하 캐시 디커플링: 글리프 박스 계산 실패가 무관한 그림 W16까지 침묵시키던 공유
  게이트 제거(실패한 축만 후퇴, 살아있는 축은 계속 검사).
- E2 숫자 맥락 예외를 문단 컨텍스트로 판정: run 경계로 쪼개진 연도 범위("2020" +
  "(U+2013)2024")가 오탐되던 것 교정(PowerPoint는 철자검사·서식 경계로 run을 흔히 쪼갬).
- OOXML 테마 토큰("+mn-lt"/"+mj-ea" 류)을 실폰트로 해석 후 E1 판정: 토큰을 문자
  그대로 블록리스트에 대조해 라틴 전용 테마 폰트 폴백이 미탐되던 것 교정.
- W6용 슬라이드 시그니처의 이중 append 경로 제거(sig와 tokens를 각자 가드): 토큰
  수집 실패 시 W6 페이지 번호가 밀리던 것 교정.
- SizeResolver: placeholder 순회 가드를 루프 전체에서 항목 단위로(손상 placeholder
  하나가 검색을 통째로 중단시키던 것), layout/master 조회 메모이제이션 추가.
- `archforge skill` 실행 시 현재 폴더에 "skill" 파일이 있으면 stderr 안내(그 파일을
  린트하려면 `archforge ./skill`).
- 테마 파싱 실패 마스터가 있으면 stderr로 후퇴 사실 고지(파스 실패와 빈 슬롯 구분
  소실 지점의 신호 보존).

### E1: 폰트 판정 재설계 (실측 렌더 모델)

- PowerPoint COM 프로브로 CJK 폰트 해석 우선순위를 실측 확정: run `a:ea` > 비어있지 않은
  테마 minorFont `a:ea`(run `a:latin`보다 우선) > 빈 테마일 때만 run `a:latin` > OS 폴백.
  기록은 `docs/CALIBRATION.md`.
- 구버전 `ea or latin` 대용 판정 폐기. 이로써 (1) 빈 테마에서 라틴 전용 latin 폰트가
  만들던 미탐과 (2) 비어있지 않은 한글 테마 ea가 렌더를 받는데도 latin만 보고 내던 오탐이
  함께 사라짐.
- 블록리스트를 덱 상용 라틴 패밀리 60여 종으로 확장(Inter, Arial, Calibri, Segoe UI,
  Roboto, IBM Plex Sans, Noto Sans 등). 한글 완비 변형 접두 오탐은 예외 목록으로 차단
  (Arial Unicode, IBM Plex Sans KR, Noto Sans/Serif KR·CJK).
- NanumGothicCoding을 블록리스트에서 제거(한글 완비 고정폭 폰트를 라틴 전용으로 오분류).
- 테마 `a:ea`를 슬라이드 → 마스터 → 테마 관계(rels)로 마스터별 해석. 멀티마스터 덱에서
  무관한 첫 테마 파트로 E1을 오발화하던 것 교정. byte 정규식 대신 XML 파싱.

### E2: 정당 타이포 예외

- 숫자 사이 en dash(U+2013 범위 표기)와 숫자 앞 수학 마이너스(U+2212 음수)는 기본 모드
  통과. `--strict`는 예외 없이 전부 차단(기존 전면 차단 규율 유지용).

### 크기 게이트: 상속 체인 해석

- 실효 크기가 run > 문단 > 도형 lstStyle > 레이아웃 placeholder > 마스터
  placeholder·txStyles > `defaultTextStyle`을 해석. placeholder·템플릿 덱에서 E3/W1/W8이
  실제로 돌게 됨. W5는 체인 전체가 침묵할 때만 발화.

### 견고성

- 코어 라인 게이트(E1~E4)와 W6/W10 클러스터링에 프레임·블록 단위 가드. 외부 생성기의
  쓰레기 속성 하나가 lint 전체를 죽여 exit 2("못 엶")로 오라벨되던 것 교정.
- `spc` 트래킹이 OOXML universal measure("1.5pt")와 쓰레기값을 흡수(autofit 퍼센트
  유니언과 같은 함정의 spc 판).
- 완전 클론 슬라이드의 코사인이 부동소수점상 1.0을 넘던 것 클램프.

### 장르·튜닝

- `--skip CODES`: 장르에 안 맞는 경고 선택 억제(에디토리얼 덱의 W14 등).
- `--w6-sim` / `--w6-cluster`: W6 골격 반복 임계 튜너블.
- W14가 숫자+단위 타이틀("매출 3배 성장")을 주장형으로 인정.

### 패키징·배포

- 스킬팩을 wheel에 동봉(`src/archforge/skills/`, package-data). 0.1.0은 pip 사용자가
  SKILL.md를 받을 방법이 없었음.
- `archforge skill` 서브커맨드 신설: 출력 / `--install [DIR]` 설치 / `--path`.
- 스킬 디렉터리명을 frontmatter name과 일치하도록 `archforge-pptx-lint/`로 변경
  (Agent Skills 규격). 루트 `skills/` 사본과 패키지 정본의 동일성은 테스트로 고정.
- GitHub Actions CI 신설: ubuntu 3.9/3.12/3.13 + windows 3.12 매트릭스, wheel 동봉·설치
  스모크 잡. `[project.optional-dependencies] test` 선언.

### 성능·아키텍처

- 슬라이드당 글리프·그림 잉크 bbox를 1회만 계산해 W15~W17에 주입(기존 텍스트 3회,
  그림 PIL 디코드 2회 재계산).
- E1 판정을 `e1_violation()`으로, E2를 `dash_violations()`으로 추출(단위 테스트 가능).
  기하 박스는 매직 인덱스 튜플 대신 `GlyphBox` NamedTuple. 주요 공개 함수에 타입 힌트.

### 문서·테스트

- `docs/CALIBRATION.md` 신설: 코퍼스 구성·방법, E1 실측 표, 게이트별 임계 근거, 재보정
  손잡이.
- README·SKILL.md를 0.2.0 동작으로 갱신(JSON `ghost` 필드, 튜너블 플래그, E2 예외,
  E1 모델, 스킬 설치 경로).
- 테스트 18 → 42개: E1 모델 매트릭스, 테마 관계, E2 맥락·strict, spc 방어, 부분 생존,
  placeholder 상속(E3/W1), defaultTextStyle 제거 시 W5, 표 셀 E1, exit 2 경로, 텍스트
  출력, `--ghost`, `--skip`, W6 튜너블, W14 숫자 주장, 스킬 동기화·frontmatter·서브커맨드.

## 0.1.0 (2026-07-10)

최초 공개. E1~E4 / W1~W17 게이트, `--json`, Agent Skills 스킬팩(리포 내), pytest 18개.
