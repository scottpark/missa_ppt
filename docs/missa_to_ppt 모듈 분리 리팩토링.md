# missa_to_ppt.py 모듈 분리 리팩토링 — 분석·설계·구현계획

**범위**: 순수 파일 분리 리팩토링. 로직/이름/동작은 바꾸지 않고, 책임별로 파일만 나눈다.
**실행 주체**: 메인 에이전트(오케스트레이터 없이 직접, 단계별 진행 + 매 단계 대화로 확인).
**목표 모듈 수**: 6개 (신규 5개 + 슬림화된 진입점 1개).

---

## 1. 분석

### 1.1 현황

- `missa_to_ppt.py`: 7,342줄, 함수 약 100개. 단일 파일에 GUI, OOXML 저수준 조작, 섹션 탐색,
  독서/복음 레이아웃 엔진, 화답송·성가 등 콘텐츠 갱신, CLI 진입점이 모두 섞여 있다.
- 관련 파일: `missa_to_json.py`(231줄), `missa_psalm_score_image.py`(427줄)는 이미 분리돼 있고
  `missa_to_ppt.spec`이 `hiddenimports`로 정상 빌드 중 — 다중 모듈 구조 자체는 PyInstaller에서
  이미 검증된 패턴.

### 1.2 테스트 커버리지로 본 리팩토링 안전망 (사용자 질의에 대한 답)

**결론: 기존 회귀 스위트만으로는 부족하다.**

| 구분 | 현재 상태 |
|---|---|
| 통합 테스트 커버 케이스 | `MASS_CASES`에 주일(`20260712`)·평일(`20260624`) 2건만. `20260705`(성수축복)는 픽스처로 추적되지만 미사용 — 단, 사용자 확인상 성수축복은 주일 템플릿의 고정 문구 일부만 다른 구조적 동일 케이스라 리팩토링 검증 목적상 별도 편입은 불필요 |
| 검증 방식 | 절 번호 오렌지색, 줄 수 초과, BlackBg 크기 등 **이미 알려진 버그 패턴에 대한 속성 검증**. "체크리스트에 없는" 서식·위치 변경은 통과된 것처럼 보임 |
| GUI 팝업(`_ask_*_popup`, `_run_with_progress_window`, `_show_result_window`, 약 1,500줄) | 자동 테스트 커버리지 0% (Tkinter, subprocess 경로는 CLI 인자로 GUI를 우회) |
| 커버리지 측정 | `pytest-cov` 미사용 — 사각지대가 수치로 확인된 적 없음 |
| 출력 결정성 | `uuid4()`는 COM probe 임시파일명에만 쓰이고 출력 pptx 내용에는 개입하지 않음(확인됨) → 동일 입력에 동일 출력이 기대되는 결정적 파이프라인. **골든 마스터 diff가 가능하고, 가장 강력한 안전망** |

**추가 검증 방법(구현계획 §3.2에 반영)**:
1. 골든 마스터 XML diff — 리팩토링 각 단계마다 리팩토링 전/후 출력 pptx가 바이트/XML 단위로
   동일한지 비교. "체크리스트에 없는" 회귀까지 전부 잡는 유일한 방법.
2. 기존 회귀(52개)·프로그레션(56개) 스위트를 각 단계마다 실행.
3. GUI 팝업 함수는 자동 검증 불가 — 모듈 이동 후 수동 스모크 테스트 체크리스트로 별도 확인
   (§3.4).

### 1.3 함수 인벤토리 → 호출 그래프 조사 결과 요약

전체 함수를 6개 카테고리로 배정 가능함을 확인했다(상세 근거는 §2). 경계를 넘나드는 함수 5종과
잠재적 순환 임포트 1건을 발견했고, 순환 임포트는 사용자와 협의해 해소 방안을 확정했다(§2.3).

부수적으로 `_josa()`, `_set_두_줄_text()`가 파일 내 어디서도 호출되지 않는 **죽은 코드**임을
확인했다. 이번 리팩토링(순수 이동)의 범위 밖이므로 삭제하지 않고 그대로 이동만 하되, 별도
정리 작업으로 제안한다(§4).

---

## 2. 설계

### 2.1 목표 모듈 구조

```
missa_ooxml_utils.py      (leaf, 프로젝트 내부 의존성 없음)
        ↑
missa_gui.py               (leaf, Tkinter + 날짜판단/config/OneDrive 폴더 설정)
        ↑
missa_sections.py          (find_sections + validate — ooxml_utils, reading_layout 의존)
        ↑
missa_reading_layout.py    (독서/복음 레이아웃 엔진 — ooxml_utils 의존)
        ↑
missa_content_updaters.py  (화답송·성가·입당송 등 섹션별 갱신 — ooxml_utils, sections 의존)
        ↑
missa_to_ppt.py             (슬림화된 진입점 — 위 5개 전부 의존)
```

의존 방향은 전부 단방향이며 순환 없음(아래 화살표 = "가 사용하는"):

- `missa_ooxml_utils` → (없음, leaf)
- `missa_gui` → (없음, leaf)
- `missa_sections` → `missa_ooxml_utils`, `missa_reading_layout`(validate가 parse_into_verse_units 호출)
- `missa_reading_layout` → `missa_ooxml_utils`
- `missa_content_updaters` → `missa_ooxml_utils`, `missa_sections`(replace_성가가 find_sections 직접 호출)
- `missa_to_ppt`(entry) → 위 5개 전부

### 2.2 모듈별 함수 배정

**① `missa_ooxml_utils.py`** — 슬라이드/도형 범용 프리미티브 + 텍스트 런 조작 저수준 유틸
`_slide_text`, `all_slide_texts`, `delete_slide`, `move_slide`, `_blank_layout`, `_copy_spTree`,
`_copy_image_rels`, `_update_rId_in_spTree`, `_effective_bg`, `_copy_bg_image_rels`,
`duplicate_slide`, `insert_slide_copy`, `copy_slide_from_prs`, `_com_probe_path`,
`_build_com_probe_pptx`, `find_slide_with_text`, `find_shape_exact_text`, `_set_slide_bg_black`,
`_find_content_shape`, `_has_ending_text`, `_clear_text_frame`, `_safe_next_slide_partname`,
`_para_append_run`, `_replace_para_text_clone`, `_set_두_줄_text`, `_set_화답송_content_text`,
`_set_single_para_text`, `_josa`, `_update_book_name_after_br`
전역: `_IMG_ID_COUNTER`

> 뒤 5개(`_para_append_run` 등)와 `_josa`/`_update_book_name_after_br`는 원래
> reading_layout/content_updaters 후보였지만, 실제로는 여러 상위 모듈이 공통으로 쓰는 범용
> 텍스트/런 헬퍼라 여기로 옮겼다 — 그렇지 않으면 content_updaters가 이 5개 함수만 쓰려고
> 훨씬 큰 reading_layout 모듈 전체를 임포트해야 한다.

**② `missa_gui.py`** — Tkinter 팝업 전부 + 날짜판단/설정(사용자 결정: 순환 임포트 해소를 위해 흡수)
`_set_window_icon`, `_apply_theme`, `_center_window`, `_ask_numbers_popup`, `_ask_date_popup`,
`_ask_input_files_popup`, `_ask_combined_input_popup`, `_report_progress`,
`_run_with_progress_window`, `_show_result_window`, `_ask_onedrive_path_popup`,
`is_sunday_mass`, `_load_config`, `_save_config`, `get_onedrive_hymn_folder`,
`_com_verification_enabled`
전역: `_UI`, `_ICON_PATH`, `_icon_image`, `_progress_callback`, `_first_dialog_pos`,
`_last_output_path`, `_last_date_str`, `_preloaded_inputs`, `CONFIG_FILE`, `_SCRIPT_DIR`,
`HYMN_TYPES`, `_TEST_HYMN_DEFAULTS`, `_COM_VERIFY_ENABLED_CACHE`

> `_last_output_path`/`_last_date_str`는 `main()`(entry)이 쓰고 `_show_result_window()`(gui)가
> 읽는 채널이다. entry→gui 단방향만 유지하려면 이 전역들도 gui 쪽에 두고, entry는
> `missa_gui._last_output_path[0] = ...`처럼 모듈 참조로 값을 채워 넣는다(리스트 셀이라
> import 방식과 무관하게 같은 객체를 공유 — 순수 이동으로 동작 100% 보존).

**③ `missa_sections.py`** — 섹션 탐색 + 검증
`find_content_range`, `find_복음_content_range`, `find_sections`, `_is_hymn_divider`,
`validate_pptx_structure`, `_orange_verse_numbers_in_range`, `_missing_orange_verse_numbers`,
`validate`, `_strip_slide_xml`, `strip_ppt2007_incompatible`

**④ `missa_reading_layout.py`** — 독서/복음 절 파싱·줄 수 계산·레이아웃·분리/재조정
`_ends_sentence`, `parse_into_verse_units`, `_visual_lines`, `_wrap_line_count`,
`_page_visual_lines`, `layout_units_on_slides`, `_verify_and_rebalance_pages`, `_set_reading_text`,
`_count_slide_lines`, `_rendered_wrap_count`, `_get_slide_render_params`,
`_count_slide_lines_rendered`, `_count_slide_lines_verified`, `_split_para_at_lines`,
`_restore_para_from_backup`, `_split_and_adjust_via_com`, `_rebalance_reading_slides_post_write`,
`replace_reading_slides`, `_align_ending_slides_to_제2독서`, `_reposition_merged_ending_shapes`
전역: `CHARS_PER_LINE`, `LINES_PER_SLIDE`, `_RENDER_WIDTH_CALIBRATION`, `_PILLOW_FONT_CACHE`,
`ORANGE`, `_SENTENCE_ENDERS`, `_COM_DISABLED`, `_COM_MISMATCH_COUNT`, `_COM_ATEXIT_REGISTERED`

**⑤ `missa_content_updaters.py`** — 화답송·성가·입당송·복음환호송·영성체송·기도문 등 갱신
`update_title_slide`, `update_입당송`, `update_reading_title_slide`, `update_복음_title_slide`,
`_update_화답송_title_in_slide`, `update_화답송`, `update_복음환호송`, `_find_last_row_top`,
`_shape_first_run_font_size_emu`, `_shape_first_para_line_spacing_pct`,
`_set_shape_all_para_line_spacing`, `_set_shape_all_run_font_size`, `_estimate_text_lines`,
`_adjust_fit_if_needed`, `update_영성체송`, `replace_시작기도문`, `replace_미사후기도`,
`_update_성가_divider_number`, `_update_prefix_in_runs`, `_update_성가_header`, `replace_성가`

**⑥ `missa_to_ppt.py`(슬림화된 진입점)**
`_infer_hymn_numbers`, `parse_args`, `find_files`, `apply_화답송_override`, `get_json_data`,
`main`, `__main__` 블록
전역: `OUTPUT_ROOT`(2번째 세션에서 신설)

### 2.3 순환 임포트 해소 (확정)

`is_sunday_mass`/config류 함수는 GUI 팝업(`_ask_combined_input_popup`)과 진입점(`main`,
`parse_args`, `find_files`) 양쪽에서 호출된다. 그대로 entry에 두면 entry↔gui 순환 임포트가
발생한다. **사용자 결정: 전부 `missa_gui.py`로 흡수, entry→gui 단방향만 유지.** (§2.2 ②)

### 2.4 미확정/재검증 필요 사항

- `HYMN_TYPES` 등 일부 상수는 이번 조사에서 gui 소속으로 잠정 배정했으나, `content_updaters`
  쪽(`replace_성가`)에서도 참조하는지 Phase 0에서 실제 grep으로 재확인 후 확정한다. 이 문서의
  모듈 배정표를 "복사된 기대값"으로 그대로 신뢰하지 않고, 각 파일 추출 직전에 반드시
  `grep -n '이름'`으로 실제 참조처를 재확인한다(프로젝트에 이미 기록된 "검증 우선순위 0" 원칙).
- 죽은 코드(`_josa`, `_set_두_줄_text`)는 이번 스코프에서 삭제하지 않고 이동만 한다. 삭제 여부는
  별도로 여쭤본다(§4).

---

## 3. 구현 계획

### 3.1 단계 순서 (leaf → root, 매 단계 후 커밋 가능한 단위)

1. **Phase 0 — 안전망 구축**: 골든 마스터 스냅샷 생성(§3.2), `pytest-cov`로 현재 커버리지
   스냅샷 확보, 회귀(52)·프로그레션(56) 스위트 베이스라인 그린 확인(이미 완료).
2. **Phase 1 — `missa_ooxml_utils.py` 추출** (leaf, 위험 최소)
3. **Phase 2 — `missa_gui.py` 추출** (leaf, Tkinter 코드라 자동 검증 불가 → 수동 스모크 필요)
4. **Phase 3 — `missa_reading_layout.py` 추출** (ooxml_utils 의존)
5. **Phase 4 — `missa_sections.py` 추출** (ooxml_utils, reading_layout 의존)
6. **Phase 5 — `missa_content_updaters.py` 추출** (ooxml_utils, sections 의존)
7. **Phase 6 — `missa_to_ppt.py` 슬림화**: 남은 진입점 코드만 정리, 5개 모듈 import로 배선
8. **Phase 7 — 빌드 검증**: `missa_to_ppt.spec`의 `hiddenimports`에 신규 모듈 5개 추가 후
   PyInstaller 빌드 + exe 실행 스모크 테스트

각 Phase는 "함수/전역을 새 파일로 옮기고, 원래 자리에는 `from missa_X import *`로 재노출"
방식으로 진행해 **각 단계마다 기존 회귀/프로그레션 스위트가 수정 없이 그대로 통과**하도록 한다
(테스트가 `import missa_to_ppt as m; m.함수명(...)` 형태로 접근하므로, 재노출 없이 옮기면 모든
테스트를 동시에 고쳐야 해서 안전망이 일시적으로 무력화된다 — 단계별 재노출로 이를 피한다).
마지막 Phase 6에서 재노출 래퍼를 걷어내고, 이때 테스트 파일들의 import 방식을 함께 정리한다.

### 3.2 골든 마스터 diff 스크립트 (신규 도입)

`tools/golden_diff.py`(가칭) — 리팩토링 전/후 출력이 100% 동일한지 확인하는 전용 도구.

```
python tools/golden_diff.py snapshot   # 현재 코드로 3개 케이스 출력 생성 → _golden/ 에 저장
python tools/golden_diff.py check      # 다시 생성 후 _golden/ 과 XML 레벨로 비교, 차이나면 실패
```

- 대상 케이스: `20260712`(주일), `20260624`(평일), `20260705`(성수축복 — 구조는 주일과 동일하나
  문구 차이가 실제로 반영되는지 함께 확인).
- 비교 방식: 출력 pptx를 zip으로 열어 `ppt/slides/slide*.xml` 등 각 파트를 유니코드 텍스트로
  읽어 파일 단위 diff. 완전 바이트 동일이 이상적이나, 한 번 실측해 실제로 결정적인지 먼저
  확인(§1.2에서 `uuid4()`가 출력에 개입하지 않음은 이미 코드 레벨로 확인함 — 실행 결과로 한 번
  더 검증).
- Phase 1~6 각 단계 완료 후 `check` 실행 → 차이가 하나라도 있으면 그 단계에서 멈추고 원인 파악
  (순수 이동이므로 차이가 나면 100% 리팩토링 버그).

### 3.3 각 Phase 공통 절차

1. 옮길 함수/전역 목록을 grep으로 실제 참조처 재확인(§2.4)
2. 새 모듈 파일 생성, 함수/전역 이동, 필요한 import 추가
3. 원래 위치에 `from missa_X import *` (또는 명시적 이름 나열) 재노출
4. `pytest test_missa_regression.py test_missa_progression.py -v` 전체 그린 확인
5. `python tools/golden_diff.py check` 통과 확인
6. 커밋(Phase 단위 1커밋)

### 3.4 GUI 팝업 수동 스모크 체크리스트 (Phase 2 전용, 자동화 불가 영역)

- 인수 없이 실행 → 날짜 입력 팝업 → 파일 선택 팝업 → 성가번호 팝업 순서로 뜨는지
- OneDrive 성가 폴더 최초 설정 팝업(설정 없을 때) 동작
- 진행률 창이 단계별로 갱신되는지
- 성공/실패 결과 팝업 + 로그 파일(`output/YYYYMMDD/log/`) 저장 확인
- EXE(GUI) 모드 — 빌드 후 실제 더블클릭 실행으로 1회 확인

### 3.5 리스크 및 완화

| 리스크 | 완화 |
|---|---|
| 전역 상태(list-cell) 이동 시 read/write 순서가 깨짐 | §2.2 ②처럼 모듈 참조(`missa_gui.X[0] = ...`)로 접근, `from X import Y`가 아닌 방식 우선 검토 |
| 모듈 배정표가 실제 코드와 어긋남(§2.4) | Phase별 grep 재확인을 스킵하지 않음 |
| PyInstaller 빌드 누락 | Phase 7에서 `hiddenimports`에 신규 모듈 추가 후 실제 exe 빌드+실행까지 확인 |
| GUI 코드 자동 검증 불가 | §3.4 체크리스트로 수동 보완 |
| 재노출 래퍼가 Phase 6까지 남아 있어 순환처럼 보일 수 있음 | 각 Phase 커밋 메시지에 "임시 재노출, Phase 6에서 제거 예정" 명시 |

---

## 4. 사용자 확인 필요 (남은 질문)

1. **죽은 코드 삭제**: `_josa()`, `_set_두_줄_text()`는 호출부가 없다. 이번엔 이동만 하고
   삭제는 별도 커밋으로 진행할지, 이번 리팩토링에 포함해 같이 삭제할지?
2. **재노출 방식**: Phase 1~5에서 `missa_to_ppt.py`에 남기는 임시 재노출을
   `from missa_X import *`(간단하지만 어떤 이름이 실제로 노출되는지 안 보임) vs 명시적 이름
   나열(장황하지만 명확) 중 선호가 있는지?
3. **착수 시점**: 지금 바로 Phase 0(골든 마스터 스냅샷 도구 작성)부터 시작할지, 이 문서를 먼저
   검토할 시간을 가질지?
