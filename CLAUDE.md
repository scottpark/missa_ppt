# CLAUDE.md — missa_ppt 프로젝트

매일미사 웹데이터로 미사 PPT를 자동 생성하는 도구. 성인미사(exe 배포)와 청년미사(Python
소스 배포, OneDrive 연동) 두 갈래를 지원한다.

**파이프라인:** `missa_to_json.py`(missa.cbck.or.kr 크롤링 → `missa_YYYYMMDD.json`) →
`missa_to_ppt.py`(JSON + 성가번호로 참조/템플릿 PPT를 수정해 결과 PPT 생성).

## 모듈 구성 (한 줄 요약, 상세는 `docs/missa_to_ppt-module-notes.md`)

- `missa_to_ppt.py` — 진입점. CLI 인자·파일 탐색·JSON 로드·`main()`/`_run_gui_mode()` 분기.
- `missa_ooxml_utils.py` — 슬라이드/도형 복사·rId 매핑·배경 상속 등 OOXML 저수준 프리미티브.
- `missa_gui.py` — Tkinter 팝업 전부 + config.json 로드/저장 + OneDrive 온디맨드 조회·캐싱.
- `missa_reading_layout.py` — 독서·복음 절 파싱 → 슬라이드 분배 → Pillow/COM 실측 → 재조정.
- `missa_sections.py` — 섹션 탐색(`find_sections`)과 검증(`validate*`).
- `missa_content_updaters.py` — 화답송·성가·입당송 등 섹션별 콘텐츠 갱신(`update_*`).
- `missa_psalm_score_image.py` — 화답송 악보 원본 이미지 → 슬라이드 변환(leaf 모듈).
- `missa_youth_gospel.py` — 청년미사 영문 독서/복음 조회(universalis.com).
- `missa_youth_hymn_pdf.py` — PDF 성가집(나주노/야훼이레) → 성가 슬라이드 PPT 생성.
- `missa_onedrive.py` — Microsoft Graph API 접근(성당 공용 계정, Device Code Flow, leaf 모듈).
- `missa_updater.py` — GitHub 기반 자동 업데이트(화이트리스트 파일만 덮어씀).
- `ppt_com_verify.py` — PowerPoint COM으로 실제 줄 수 실측 검증.
- `tests/` — 테스트 전부(`test_missa_progression.py`=신규 동작 명세, `test_missa_regression.py`
  =회귀, 나머지는 OneDrive/GUI/COM/업데이터별 분리). 저장소 루트 `conftest.py`가 `sys.path`에
  루트를 추가해 `tests/`에서 루트의 런타임 모듈을 import할 수 있게 한다.
- `dist/` — 빌드 산출물(exe/zip) 전용 폴더, git 추적 안 함.
- `docs/` — 요구사항·구현 계획(성인+청년 통합, 공통/성인전용/청년전용 3섹션 구조),
  `archive/`(과거 버전 원문), `ooxml-pitfalls-log.md`(아래 규칙들의 상세 발견 경위),
  `missa_to_ppt-module-notes.md`(모듈별 구현 상세).

## OOXML·GUI·캐싱 함정 규칙 (필수 적용)

각 규칙의 전체 설명·코드 예제·발견 경위 전문은 `docs/ooxml-pitfalls-log.md`의 동일 제목
절에 있다. 새 함정을 발견하면 여기에 규칙 한 줄 + 날짜만 추가하고, 전문은 그 문서에 쓴다.

1. **`pPr.append()` 금지** — `<a:pPr>` 자식은 `lnSpc→spcBef→spcAft→...→extLst` 순서가
   강제된다. 추가/교체는 `insert(0, ...)` 또는 스키마 위치에 맞게. append는 항상 맨 끝에
   붙어 순서를 깬다.
2. **순서 스키마는 pPr 외 다른 요소에도 적용** — 예: `<a:p>`는 `pPr?→(run|br)*→endParaRPr?`.
   `endParaRPr`가 남아있는 단락에 새 run을 append하면 순서 위반.
3. **슬라이드 복사 후 배경 재설정 금지** — `copy_slide_from_prs()`가 이미 배경을 정확히
   복사하므로 직후 `_set_slide_bg_black()` 호출 금지. 같은 프레젠테이션 내부 복제
   (`duplicate_slide`)에는 걸어도 안전.
4. **슬라이드 복제 시 `showMasterSp` 속성도 원본과 맞춘다** — `p:bg`와 별개 속성, 마스터
   상속 요소 노출 여부를 결정. `add_slide()` 기본값엔 이 속성이 없다.
5. **같은 프레젠테이션 내부 복제는 원본 레이아웃을 그대로 재사용** — `_blank_layout()`류
   "이름으로 검색"은 다른 프레젠테이션에서 복사할 때만의 최후 폴백.
6. **fit-보정 함수의 안전 마진은 이미 실측된 보정치를 재사용** — 새 임의 값을 발명하지
   말고 `_FIT_HEIGHT_SAFETY_MARGIN=1.15`(Pillow가 실제 줄 높이보다 14~18% 작다는 실측)를
   재사용한다.
7. **빈 문자열은 falsy** — 도형 분류에서 `if key=='content' and txt:` 패턴은 텍스트 없는
   배경 도형(`BlackBg`)을 걸러내지 못한다. 이름으로 명시적으로 먼저 제외한다.
8. **텍스트 키워드로 도형 식별 시 부분 문자열 충돌 주의** — 키워드가 다른 도형에도 우연히
   포함되면 순회 순서에만 의존하는 암묵적 안전성이 된다. 더 구체적인 패턴으로 좁힌다.
9. **post-write 재조정 이후엔 계획값이 아니라 실제 물리 상태를 재측정** — 앞 단계가 슬라이드를
   추가/재배치했을 수 있으므로 `units_pages` 같은 계획값 참조 금지.
10. **1줄 높이를 "박스 height ÷ 고정 줄 수"로 역산하지 않는다** — 실제 폰트 메트릭
    (`_content_line_height_emu()`)으로 계산.
11. **한글 줄 수 계산: TTC 인덱스는 실측, Pillow 폭은 보정** — `batang.ttc` 등 인덱스를
    추측하지 말고 `getname()`으로 확인. `libraqm` 미지원으로 폭이 넓게 나와
    `_RENDER_WIDTH_CALIBRATION=1.03` 보정.
12. **페이지네이션과 검증은 반드시 같은 박스 폭을 써야 한다** — 한쪽에만 안전 마진을 곱하면
    실제로 들어가는 단어가 다음 슬라이드로 밀려나는 반대 방향 버그가 생긴다(`_PIL_WRAP_
    SAFETY` 사례).
13. **단락 분리·병합 함수는 run 개수를 2개로 가정하지 않는다** — 절 병합 시 run이 3개 이상
    될 수 있다. 분리 지점을 실제로 계산해 앞뒤 run의 rPr을 보존한다.
14. **run 재작성은 통짜 재배치 대신 의미 단위 경계로 나눈다** — `_update_prefix_in_runs()`의
    `segments=[(old,new),...]`처럼, 길이 변화가 다른 의미 구간의 run을 침범하면 안 된다.
15. **검증은 텍스트 존재가 아니라 절 번호별 오렌지색 렌더링을 확인** — `_missing_orange_
    verse_numbers()`.
16. **이미지에서 여러 그룹을 크롭할 때 전역 경계가 아니라 그룹별 경계를 쓴다** —
    `_group_content_bbox(gray, group)`처럼 그 그룹의 행 범위 안에서만 잉크 경계 재계산.
17. **재시도 루프는 "정확히 못 맞추면 포기"가 아니라 best-so-far를 보존** — 불변조건(오버
    플로 금지)을 지키는 가장 근접한 결과를 버리지 않는다. 이웃을 통째로 흡수해 빈 이웃이
    남으면 삭제해서 병합하는 것이 "아무것도 안 하는 것"보다 낫다.
18. **GUI 위젯의 "펼침" 등 다중 트리거 상태 변화는 가상 이벤트 하나로 감지** — Treeview는
    더블클릭과 화살표클릭이 다른 이벤트다. `<<TreeviewOpen>>` 하나에만 로직을 건다.
19. **표시 문자열과 실제 처리값은 반드시 분리** — `textvariable`을 보여주기 좋게 다듬으면
    그 값을 읽는 실제 처리 로직이 존재하지 않는 경로를 받는다. 표시용/처리용 변수를 분리하고
    단일 헬퍼로 우선순위를 강제한다.
20. **계속 확장되는 함수에 새 부작용을 추가할 때는 그 함수를 모킹 없이 실행하는 기존 테스트를
    찾아 갱신한다** — 안 하면 헤드리스 테스트가 실제 모달 창 앞에서 무한 대기한다.
21. **Tkinter 합성 클릭(`event_generate`)은 좌표+`update_idletasks()` 없이 무시될 수 있다** —
    좌표 없이는 트리거 안 됨, 좌표만으론 깊이 중첩된 위젯에서 부족할 수 있음.
22. **`Application.Visible=True`는 창을 보이게 하지만 포커스를 가져가지 않는다**(실측
    확인) — 사용자 안내 문구는 이 사실대로 정정.
23. **캐시 계층이 여러 개 겹칠 때, 한 계층의 스냅샷을 다른 계층의 최종 판단에 재사용하지
    않는다** — "UI 표시용"과 "최종 재사용 판단용"은 다른 질문이다. 후자는 항상 실시간
    재확인 경로를 거친다.

## 회귀 테스트

`tests/test_missa_regression.py`에 주일·평일 통합 테스트와 핵심 함수 단위 테스트가 있다.

**선택 실행 우선, 전체 실행은 최종 게이트 1회**: 기능별 클래스(`TestIsSundayMass`/
`TestWrapLineCount`/`TestSplitAndAdjustViaCom`/`TestComVerify` 등)와 별도 파일(OneDrive/
GUI-OneDrive/COM검증/업데이터)로 이미 나뉘어 있다. 작업 중엔 관련 클래스/파일만 `-k`나
`::클래스명`으로 좁혀 돌리고, 전체 스위트(`pytest tests/ -v`, 총 300개+·3~4분)는 리뷰
통과 후 최종 게이트로 한 번만 돌린다.

**알려진 환경 함정**: `ppt_com_verify.py`가 띄우는 PowerPoint COM `Application` 핸들에
확실한 정리 경로(`atexit`)가 없어, pytest 프로세스가 비정상 종료하면 `POWERPNT.EXE`가
좀비로 남을 수 있다. COM 관련 테스트(`TestComVerify` 등) 실행 전 `tasklist`로 잔여
프로세스가 없는지 확인하는 습관이 필요하다(수정 미착수).

## 하네스: missa_ppt 개발

**목표:** 성인미사 전용으로 짜인 코드를 청년미사·어린이미사 등 새 미사 유형/성당으로
안전하게 확장한다. 새 기능은 실패하는 프로그레션 테스트로 먼저 명세하고, 안정화되면
회귀 테스트로 승격한다. 설계·구현·리뷰·검증을 분리된 전문 에이전트가 담당한다.

**트리거:** 새 미사 유형/성당 지원 추가, PPT 템플릿 구조 확장, 섹션 처리·OOXML 로직
구현/버그 수정 요청 시 `mass-ppt-dev` 스킬을 사용하라. 단순 질문은 직접 응답 가능.

**구성:** `.claude/agents/`(mass-template-architect·ppt-ooxml-specialist·
ooxml-code-reviewer·regression-qa), `.claude/skills/`(mass-ppt-dev 오케스트레이터 +
mass-template-analysis·tdd-progression-testing·ooxml-pitfalls·boundary-verification·
docs-sync). 변경 이력은 `docs/missa_to_ppt 구현 계획 변경이력.md` 참고.
