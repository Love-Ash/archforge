# Changelog

## 0.5.0 (2026-07-10)

배포·온보딩 릴리스입니다. 4차 리뷰 3건이 공통으로 짚은 병목(엔진이 아니라 전달)을
다룹니다: 30초 온보딩(demo), CI 통합(scan·GitHub Action·pre-commit), 에이전트 자동
수정용 location 계약 완성, 그리고 README 전면 개편.

### 신설

- `archforge demo`: 결함 6종을 심은 broken.pptx와 교정본 fixed.pptx를 생성해 즉석
  린트(첫 실행 경험). 리포 `examples/`에 같은 소스로 만든 커밋본 3종(broken/fixed/
  style_warnings)과 기대 출력 문서를 추가했습니다.
- `archforge scan PATHS...`: 파일·디렉터리(재귀)·글롭 혼합 다중 린트. 파일별 판정은
  단일 모드와 같은 경로를 지나고, 집계 JSON(`files[]` + aggregate summary)과 다중
  파일 SARIF(한 run, 파일별 artifactLocation)를 냅니다. 하나라도 실패면 exit 1,
  매치 0건은 조용한 통과가 아니라 exit 2(CI 풋건 방지). PowerPoint 잠금 파일
  (~$*.pptx)은 제외.
- GitHub Action(composite `action.yml`): `uses: Love-Ash/archforge@v0.5.0`에
  files/profile/strict/sarif/version 입력. pre-commit 훅(`.pre-commit-hooks.yaml`)도
  추가.
- location 계약 완성(에이전트 자동 수정 타깃): 그룹 변환을 합성한 절대 bbox, 표 셀
  `cell`=[행,열], 자동 필드 `field: true`, W15~W17의 실효 글리프 bbox + 쌍 판정
  `related`(상대 도형), W7 절대 bbox.
- a:fld(슬라이드 번호·날짜 필드)가 일반 run과 같은 게이트(E1/E3/E4)를 지납니다.
  스키마상 fld도 rPr+t라 같은 규칙으로 렌더되는데 python-pptx run 순회가 건너뛰어
  사각지대였습니다. 필드는 ghost/W14 제목 수집에서 제외(큰 페이지 번호 오염 방지).
  a:br은 E2 문맥·오프셋에서 줄바꿈 한 글자로 취급됩니다.
- 커뮤니티 문서: CONTRIBUTING(증거 기준: 게이트는 렌더 대조로만 조정),
  SECURITY(위협 모델), 이슈 템플릿 3종(오탐 신고 템플릿 포함), PR 템플릿.

### 수정(정확성)

- 그룹 아핀 잠복 버그: 슬라이드의 grpSpPr는 p: 네임스페이스인데 a:로 find해 그룹
  변환이 항상 항등으로 후퇴했습니다. 이동 desync 그룹에서 W15~W17 기하와 loc bbox가
  raw 좌표로 판정되던 것을 로컬네임 매칭으로 교정하고 절대좌표 테스트로 고정했습니다
  (0.5.0 loc 작업 중 실측으로 발견).
- W15/W16/W17 상위 2건 선별의 정렬 키를 초과량·비율 단일 키로 명시(동률일 때
  텍스트 사전순이 개입하던 것 제거).

### 4차 리뷰 정확성 반영분(main에 누적돼 있던 미출시 배치)

4차 외부 리뷰(전략·확산 관점 2건 + 자체 검증 세션 CONFIRMED 9건)의 정확성 결함
반영분. 릴리스 규율 지적을 수용해 당시 버전을 올리지 않고 이 릴리스에 합류했습니다.

- (HIGH) 설정 파일 신뢰 경계: `--no-config` 신설, 적용된 설정 경로를 JSON
  `summary.config`와 텍스트 각주에 항상 표시. 덱 폴더의 공격자 제어 설정이 게이트를
  조용히 약화시키던 경로의 가시화·차단 수단.
- (HIGH) baseline 지문 v2(스키마 2, 재생성 필요): 로케일 중립 키(fp_key)로 언어 의존
  해소, 페이지 번호 제거로 슬라이드 삽입 생존, 발생 수(count) 기반 multiset 의미,
  실행 조건 메타데이터(tool_version·profile·lang) 기록. v1 파일은 명확한 오류로 거부.
  텍스트 출력에 억제 각주 추가(불가시 clean 오독 교정). baseline은 beta로 표기.
- `--help`의 기본 프로파일 안내가 구버전(full)이던 것 교정(파괴적 변경 릴리스에서
  --help가 틀려 있던 문서 드리프트).
- 설정 fail-safe: 알 수 없는 키(오타)는 exit 2, 값 타입·범위 검증(traceback 제거),
  lang 검증, baseline 경로는 설정 파일 기준 해석. CLI 임계값도 범위 검증
  (`--hard-min 0`으로 E3를 조용히 끄던 우회 봉합).
- 라이브러리 `lint(profile=...)`가 오타 프로파일을 조용히 full로 처리하던 것을
  ValueError로 교정.
- 문서 드리프트: README·SKILL의 E4 서술을 실제 범위(한글·한자, 가나 제외)로, JSON
  예시의 하드코딩 버전 제거, `--strict`의 E2 해제가 full에서만 의미임을 명시,
  CHANGELOG의 내부 의사결정 표현 제거.

수용 보류·기각(판단 기록): core 기본값 자체는 유지(3차에서 사용자 확정, 가시성으로
보완). a:fld/a:br 텍스트 미검사와 그룹 내 loc bbox의 raw 좌표는 알려진 한계로
공개했다가 이 릴리스의 신설·수정 항목에서 해소했습니다.

### 문서

- README(en/ko) 전면 개편: 실렌더 before/after 대조 이미지와 데모 GIF(둘 다
  examples 덱의 PowerPoint 실렌더로 제작), 30초 온보딩 섹션, CI 섹션, location
  계약 문서화. SKILL.md에 scan/demo와 location 키 계약 추가.
- docs/assets/social-preview.jpg(1280x640) 동봉: 리포 Settings의 Social preview에
  업로드하는 용도.

### 테스트

- 85개에서 93개로: fld 게이트·br 오프셋·표 셀 loc·그룹 절대 bbox·W15/W16 loc·
  scan/demo CLI·다중 SARIF·examples 계약.

## 0.4.0 (2026-07-10)

3차 외부 리뷰의 구조 개편 로드맵과 기본 프로파일 결정을 반영한 릴리스입니다.
기본값 변경의 목적은 첫 실행 오탐 제거와 온보딩 개선입니다(스타일 규칙은 옵트인).

### 파괴적 변경

- 기본 프로파일이 full에서 core로 바뀌었습니다. 무옵션 실행은 객관 결함만
  (E1/E3/E4, W1/W5/W7/W8, W15~W18) 검사하고, AI 티·하우스 스타일 규칙(E2 대시,
  W6 반복, W9~W14)은 `--profile full` 옵트인입니다. 첫 사용자가 정상 문장부호로
  exit 1을 받던 첫인상 문제의 교정. 기계 생성 덱을 검사하는 에이전트 루프는 full을
  명시하세요(SKILL.md의 기본 명령이 그렇게 바뀌었습니다).

### 구조 개편

- Finding 모델(findings.py): 검사 엔진은 로케일 중립 finding(코드+메시지 id+인자)만
  만들고 언어는 리포터 단계에서 결정. 기존 4튜플 언패킹·인덱싱 하위호환 유지.
  finding에 구조화 위치(location: shape_id·shape_name·bbox·part·paragraph·run)가
  실려 에이전트가 detail 문자열 재파싱 없이 수정 대상을 특정할 수 있습니다.
- 규칙 레지스트리(rules.py)와 리포터(reporters.py: text/json/sarif) 분리.
- 설정 파일: 덱 폴더/작업 폴더의 .archforge.json(PyYAML 설치 시 .yml). CLI가 설정을
  이기고, 알 수 없는 키는 경고 후 무시.
- baseline: `--write-baseline`으로 기존 덱의 위반을 수용하고 이후 신규 위반만 보고
  (W18은 대상 외, 억제 수는 summary.baseline_suppressed에 기록).
- SARIF 2.1.0 출력(`--sarif PATH`): GitHub code scanning 연동.
- 성능 예산: 25MP 초과 이미지는 알파 트림 생략(공개), W6/W10 쌍 비교 200장 상한(공개).
- python-pptx 상한(<2): 내부 API 의존을 검증 안 된 메이저 업그레이드로부터 보호.
- benchmarks/: 결함 매트릭스에서 픽스처를 생성해 규칙별 precision/recall을 재현
  가능하게 채점하는 공개 하네스(CI 게이트 겸용). 생성기 커버리지는 python-pptx부터.

테스트 76개에서 82개로.

## 0.3.1 (2026-07-10)

0.3.0에 대한 3차 외부 심층 리뷰의 P0 5건과 P1 7건을 전량 반영한 정합성 릴리스입니다.
주제는 기능 추가가 아니라 엔진 내부의 일관성입니다: 하나의 실효 문서 모델.

### P0 (배포 차단급)

- E1이 문단 pPr/defRPr 폰트 상속을 읽는다. COM 프로브7로 실측 확정: 문단 defRPr의
  a:ea가 실제로 렌더되고 lstStyle 체인을 이긴다(우선순위 run rPr > 문단 defRPr >
  lstStyle 체인 > 테마). 이 단계 누락이 오탐(문단에 한글 폰트를 준 덱 차단)과
  미탐(문단이 라틴 전용으로 덮는데 상위 체인만 보고 통과)을 동시에 내고 있었다.
- W15~W17 기하 검사가 E3와 동일한 StyleResolver로 크기 없는 run을 해석한다.
  같은 문서에 두 개의 실효 스타일 모델(E3=40pt, 기하=기본 12pt)이 존재하던 구조적
  불일치의 교정. 상속 크기 큰 제목의 화면 이탈·충돌 미탐이 사라진다.
- 네이티브 표 셀 텍스트가 기하 검사에 포함된다(열 폭·행 높이 누적으로 셀 사각형 계산).
  자동 생성 덱의 표가 W15~W17 사각지대였다.
- --render 계약: 폴더 부재, 그림 있는 페이지의 렌더 누락이 W18/incomplete로 표면화된다.
  조용한 0건 검사 경로 제거(그림 없는 덱은 렌더가 없어도 완전으로 판정).
- 프로파일이 CLI 사후 필터가 아니라 엔진 실행 정책이 됐다: lint(profile=...)로
  라이브러리에서도 사용 가능, 제외 규칙은 실행 자체를 안 하므로 O(S^2) 비교 비용이
  안 들고 제외 규칙의 내부 실패가 W18로 누출되지 않는다.

### P1

- --skip 검증 강화: 존재하지 않는 코드(오타) exit 2 거부, W18(불완전성 신호) 억제 불가.
- W7 텍스트색 해석 확장: 직접 RGB만이 아니라 문단 defRPr, lstStyle 상속 체인,
  schemeClr(테마 clrScheme 해석)까지. 그림·텍스트 좌표도 그룹 변환 포함 절대좌표로.
- E4 메시지가 판정 범위와 일치(한글·한자). 한자 전용 런 E4는 유지하되 향후 프로파일 후보.
- detail·stderr까지 i18n(W6 detail의 예/e.g., W10 페이지/pages, 진단 노트 3종).
- is_cjk를 is_hangul+is_kana+is_hanja 합성으로 단일화(확장 한글이 W8에서 빠지던 불일치).
- ghost 제목 수집이 제목 placeholder를 크기보다 우선(60pt KPI 빅넘버가 26pt 실제 제목을
  밀어내던 것).
- JSON 버전 계약 시작: schema_version("1.0"), tool{name,version}, target_renderer
  ("powerpoint-windows"). SKILL의 에이전트 종료 조건을 error_count==0 AND
  incomplete==false로 갱신.

테스트 65개에서 76개로.

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
