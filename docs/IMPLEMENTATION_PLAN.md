# IMPLEMENTATION_PLAN — 현재 진행 작업

> 임시/단기 문서. 작업이 끝나면 내용을 비우고(또는 `archive/`로 이동) 다음 작업으로 교체한다.
> 작성 순서: PRD → SPEC → ARCHITECTURE → 이 문서, **각 문서마다 사용자 승인**. 이 문서 승인 후 코딩 시작.

## 진행 중: OneDrive 폴더 구조 개편 대응 (2026-10-18, 사용자 지시로 승인 면제)

목표·명세: `SPECS/SPEC_OneDrive_폴더구조_개편.md`(§4 확정 결정 포함). 설계: `ARCHITECTURE.md` §3.3·§3.4.

### 단계별 체크리스트
- [x] 1. 이 계획 + ARCHITECTURE 갱신
- [x] 2. 프로그레션 테스트 선작성(`tests/test_missa_progression.py`, 접두 `test_fs_`) — 아래 행동 목록
- [x] 3. `missa_gui.py` 루트 3종 관리·경로 규칙·찾아보기 기억 구현, Graph 폴백·원격 브라우저·프리페치·캐시 삭제
- [x] 4. `missa_to_ppt.py`·`missa_content_updaters.py` 호출부를 새 API로 교체(`ensure_onedrive_login`·모드 분기 제거)
- [x] 5. `missa_onedrive.py`·`msal` 의존 제거(참조 grep 후), `tools/build_dist_zip.py`·`requirements.txt`·운영자 안내서 문구 갱신
- [x] 6. 기존 테스트 갱신/폐지(`test_missa_youth_local_mode.py` 재작성, `test_missa_gui_onedrive.py`·`test_missa_onedrive.py` 삭제, regression/progression의 Graph 의존 테스트 정리)
- [x] 7. SPEC·README·CLAUDE.md 문구 갱신, 하네스 변경 이력은 해당 없음(코드/문서만)
- [ ] 8. `pytest tests/ -q` 전체 1회

### 테스트 대상 행동 (인수 기준 §3 매핑)
1. 루트 유효성: 공용=`09.가톨릭 성가`, 성인=`10.기타 PPT문서`, 청년=`2.성가` 포함 시 유효, 공백 차이 무시.
2. 필요 루트만 묻는다: 청년 실행은 공용+청년, 성인 실행은 공용+성인. 불필요 루트는 묻지 않음.
3. 청년 PC(두 루트가 서로 다른 위치): 선택창 2회 후 config에 저장, 재실행 시 묻지 않음.
4. 마스터 편의: 'PPT 문서'(상위)를 고르면 세 루트가 모두 채워진다.
5. 옛 구조 config(`onedrive_ppt_folder`가 옛 'PPT 문서'/`onedrive_hymn_folder`)는 무효로 보고 선택창으로 간다. 새 구조를 품은 `onedrive_ppt_folder`는 세 루트로 자동 이관.
6. 무효 선택 시 안내 후 재선택, 취소 시 False 반환(호출부는 안내 후 종료).
7. 경로 규칙 표 전 행: 가톨릭 성가 검색, 청년 성가 검색/신규 저장, 결과 복사(성인/청년), 찾아보기 기본 폴더(시작기도·미사후기도·공지·참조·화답송).
8. 찾아보기 기억: 행별 저장, 참조 PPT는 성인/청년 구분, 화답송 `last_psalm_dir` 흡수, 저장 폴더 삭제 시 기본 폴더로.
9. 폐지 확인: `missa_gui`에 Graph 관련 심볼(`ensure_onedrive_login`·`_mirror_onedrive_folder` 등)이 없고 `missa_onedrive` 모듈 파일이 없다.
10. 모달 방지: 선택창은 `ask_if_missing=True`인 GUI 진입점에서만, 헤드리스 테스트는 monkeypatch.
