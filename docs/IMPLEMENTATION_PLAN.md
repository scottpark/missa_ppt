# IMPLEMENTATION_PLAN — 현재 진행 작업

> 임시/단기 문서. 작업이 끝나면 내용을 비우고(또는 `archive/`로 이동) 다음 작업으로 교체한다.
> 작성 순서: PRD → SPEC → ARCHITECTURE → 이 문서, **각 문서마다 사용자 승인**. 이 문서 승인 후 코딩 시작.

## 진행 중: 성인미사 화답송 악보 폐지 (2026-11-17, 사용자 승인: SPEC 승인 + ARCH/PLAN 면제·즉시 개발)

명세: `SPECS/SPEC_성인미사.md` §2.1·2.2·2.3·2.4. 설계: `ARCHITECTURE.md` §2.1. 경량 경로(TDD→독립 리뷰→검증→문서동기화).

### 단계별 체크리스트
- [x] 1. 프로그레션 테스트 선작성(`tests/test_missa_progression.py`, 접두 `test_nopsalm_`) — 아래 행동 목록, red 확인
- [x] 2. `missa_content_updaters.py` `update_화답송` 단순화(주일 분기·악보 인자 제거)
- [x] 3. `missa_to_ppt.py` `--화답송` 인자·악보 PPT 탐색·사진 입력·경고 출력 제거
- [x] 4. `missa_gui.py` 화답송 악보 행·`화답송_pptx` browse 키·14.화답송 악보 기본폴더·미선택 확인 팝업 제거, 옛 키 무시
- [x] 5. 기존 테스트(악보 의존) 갱신/폐지, `tools/build_dist_zip.py` 화이트리스트는 보존 모듈이라 유지
- [x] 6. 독립 리뷰 → 관련 테스트 → 전체 pytest 1회 → 문서 동기화(운영자 README·module-notes·CLAUDE.md) → 하네스 변경 이력 해당 없음

### 테스트 대상 행동
1. 주일(is_calendar_sunday) 미사에서도 화답송 슬라이드는 절 수만큼 텍스트 슬라이드이며 악보 슬라이드가 없다(후렴 ◎ 표시).
2. `update_화답송` 시그니처에 악보 경로/`is_sunday` 인자가 없다.
3. CLI 파서에 `--화답송`이 없고, 날짜 폴더의 `화답송 악보*.pptx`가 있어도 사용되지 않는다.
4. 첫 팝업(성인)에 화답송 악보 행이 없고 미선택 확인 팝업이 뜨지 않는다. 청년 팝업은 변화 없음.
5. config에 옛 `화답송_pptx`/`last_psalm_dir`가 있어도 오류 없이 무시된다.
6. 청년 성가 PDF 모듈(`missa_psalm_score_image` import)은 정상 동작.

## 추가 작업: 화답송 서식 보존·넘침 보정 (2026-11-18, 11월 주일 3개 PPT 육안 검수 결과)

명세: `SPECS/SPEC_성인미사.md` §화답송 "서식 원칙". 사용자 결정: ○ 템플릿 구성 그대로(라벨 30+본문 32) / 줄간격 1.0 먼저 후 ○ 본문 축소(하한 16pt) / COM 실측 검증 상시 / CLAUDE.md 함정 규칙+pitfalls 로그.
- [x] 1. 재현·원인 측정(11/8 슬라이드 50: `_adjust_fit_if_needed`가 ◎·○ 전 run을 일괄 36pt+줄간격 1.0으로 덮어씀, `_set_화답송_content_text`가 ○ 단락을 첫 run 서식으로 통일해 본문 32→30pt) + 평일/청년에서 안 보이던 이유 확인
- [x] 2. 프로그레션 테스트 선작성(`test_psalmfmt_`): 템플릿 run 크기 보존, 조정 시 ◎ 불변·○ 라벨 불변·본문만 축소, 줄간격 순서
- [x] 3. `_set_화답송_content_text`(라벨/본문 run 구분 복제)·화답송 전용 넘침 보정 구현(실측 COM 기반)
- [x] 4. 생성 직후 COM 실측 검증(화답송 슬라이드 하단이 슬라이드 높이 초과 시 경고) 추가
- [x] 5. 독립 리뷰 → 관련 테스트 → 11/8·11/15·11/22 재생성 → PowerPoint 열어 실측·이미지 확인 → CLAUDE.md 규칙/pitfalls 로그/ module-notes 동기화
