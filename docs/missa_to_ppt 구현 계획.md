# missa_to_ppt 구현 계획

- 최종 수정일: 2026-10-01
- 성인미사 구현 계획과 청년미사 2단계 구현 계획을 한 문서로 통합했다. 구성: **1부 공통**
  (두 미사 유형이 공유하는 파이프라인·엔진·유틸) · **2부 성인미사 전용** · **3부 청년미사
  전용**. 날짜순 상세 변경 이력(버그 발견 경위·리뷰 왕복 포함)은
  `docs/missa_to_ppt 구현 계획 변경이력.md`로 분리했다 — 이 문서는 "지금 어떻게 동작하는가"만
  담는다. OOXML 함정 재발 방지 규칙은 `CLAUDE.md`, 모듈별 책임 요약도 `CLAUDE.md` "핵심 파일"
  절 참고(이 문서 §1.9 함수 목록과 상호 보완).
- 요구사항은 `docs/missa_to_ppt 요구사항.md`. 청년미사 1·2단계 원문, 청년미사 2단계 구현
  계획 원문은 `docs/archive/`에 보관.

---

# 1부. 공통 구현

## 1.1 처리 파이프라인 (9단계 + 청년 전용 단계)

`missa_to_ppt.py`의 `main(date_str, hymn_numbers, files, mass_type='adult', ...)`가 아래
순서로 호출한다. 괄호 안은 성인/청년 분기 여부.

1. `get_json_data()` — 미사 JSON 로드/크롤링(청년: `input_date_str`로 출력 폴더를 입력 날짜
   기준으로 계산, §3.7)
2. (청년만) `missa_youth_gospel.get_youth_mass()` 조회 + `merge_youth_gospel_content()`로
   복음 content 교체
3. `resolve_mass_flags(mass_type, date_str)` — §1.2의 4축 계산
4. `find_sections(prs, mass_type)` — 섹션 인덱스 탐지(§1.3)
5. 전례 텍스트 교체(제목/입당송/1독서/2독서/복음환호송/복음/영성체송) — 복음환호송·복음은
   `mass_type` 전달
6. 시작기도문 교체(지정 시)
7. 성가 처리 — 성인 `replace_성가()` vs 청년 `replace_성가_youth()`(§3.5)
8. [6.7] `find_sections(prs)` 재호출 후 공지사항 삽입(성인만 해당 슬롯 존재, 함수 자체는 공통)
9. 검증(`validate()`) → 저장 → PPT2007 호환성 정리(OneDrive 복사는 결과창 버튼, §3.4.2)

## 1.2 `resolve_mass_flags(mass_type, date_str) -> dict`

기존 `is_sunday`(불리언 1개)가 실제로 4개 독립 축(①달력상 일요일 여부 ②표시 형식(평일식
여부) ③2차봉헌 포함 여부 ④성가 악보 실제 복사 여부)을 겸하던 것을 분리했다.

```python
is_calendar_sunday = is_sunday_mass(date_str)          # ①
if mass_type == 'adult':
    use_weekday_display = not is_calendar_sunday        # ② 기존과 동일
    include_2차봉헌      = is_calendar_sunday            # ③ 기존과 동일
    copy_hymn_scores     = is_calendar_sunday            # ④ 기존과 동일
else:  # 'youth'
    use_weekday_display = True                          # ② 고정
    include_2차봉헌      = True                          # ③ 고정
    copy_hymn_scores     = True                          # ④ 고정
```

adult 경로는 4개 값 전부 `is_calendar_sunday`(또는 부정)로만 정의되므로 기존 동작과 100%
동일하다. 호출부 치환: `find_files(..., copy_hymn_scores)`(매개변수명 `is_sunday`→
`copy_hymn_scores`로 리네임), `update_화답송(..., is_sunday=not use_weekday_display)`(**부호
반전 — `is_sunday`는 "평일식 표시인가"가 아니라 "주일 포맷(악보 포함)인가"를 뜻하므로 그대로
대입하면 안 됨**), `replace_성가(..., copy_scores=copy_hymn_scores)`, 미사후기도 처리는
`use_weekday_display`가 참일 때만, `validate(prs2, json_data)`는 `is_sunday` 매개변수 자체가
제거됨(화답송 라벨 체크를 항상 수행).

## 1.3 `find_sections(prs, mass_type='adult') -> dict`

섹션 탐색의 뼈대(입당~영성체송의 골격, 공지사항 슬롯 등)는 공통이다. `mass_type='youth'`일 때
추가로 동작하는 3곳(§3.2)을 제외하면 완전히 동일한 탐지 로직을 쓴다. 성가 5종은 성인 경로에서
`{htype}_divider`/`_content_start`/`_content_end`(flat 키)만 채우고, 청년 경로는 추가로
`{htype}_songs`(곡 리스트) 키를 채운다 — `2차봉헌_content_end` 같은 flat 키는 청년 경로에서도
`songs[0]['content_end']`로 alias돼 하위 호환된다(`insert_공지사항()`이 그대로 참조 가능).

## 1.4 독서·복음 레이아웃 엔진 (`missa_reading_layout.py`)

요구사항 §1.5(공통 알고리즘)를 구현하는 핵심 모듈. 한글(전각 문자 수 기반)과 영문(Pillow
실측 word-wrap 기반)이 페이지 분배 함수만 다르고, 나머지 파이프라인은 완전히 공유한다.

- **파싱**: `parse_into_verse_units()`(절 번호 파싱) → `_merge_continuation_units()`(continuation 병합).
  이 units를 `replace_reading_slides(prs, s, e, units, template_idx, ...)`가 그대로 받는다(2026-10-03:
  사전 분할 결과 `units_pages`가 아니라 units 입력).
- **기록**: `_set_reading_text(tf, units, line_spacing, align=None, normalize_page_size=False)`
  — 절 번호 오렌지 run + 본문 run 생성. `align`(기본 `None`=기존 좌측, 청년 복음만 `'ctr'`)과
  `normalize_page_size`(기본 `False`, 청년 복음만 `True` — 재사용되는 모든 페이지 박스 크기를
  `template_idx` 기준으로 정규화)는 청년 전용 opt-in 파라미터이며 성인 호출부는 이 값을 넘기지
  않아 완전히 무변경으로 동작한다.
- **슬라이드 단위 채우기(2026-10-03, 옛 "추정 일괄 분할→사후 보정" 구조 교체)**:
  `replace_reading_slides()`가 슬라이드마다 `_measure_slide_cut()`으로 정확히 9줄까지를 확정한다.
  1) 남은 본문(`_reading_paras_from_units()`: 문단 텍스트 + 오렌지 절번호 구간 오프셋, 절 번호는 run이
  아니라 오프셋으로 들고 다녀 run 개수를 가정하지 않음)의 앞부분(약 1200자)을 `_write_reading_paras()`/
  `_set_reading_text`로 그 슬라이드의 실제 본문 상자에 임시 기록 → 2) COM 한 번(`_com_measure_line_starts()`
  → `ppt_com_verify.measure_line_starts()`: 줄 수와 앞 10줄의 시작 오프셋)으로 10번째 줄 시작을 얻어
  `_split_paras_at()`으로 자름(10번째 줄이 안 보이면 프리픽스를 2배로 늘려 재측정, 응답이 기록한 텍스트와
  어긋나면 그 슬라이드만 Pillow 폴백) → 3) 앞 9줄 확정, 나머지로 다음 슬라이드(템플릿 복제로 증가)를
  반복, 마지막 슬라이드는 남은 만큼.
  COM 불가/`com_verification_enabled=false`/COM 실패 시 같은 루프를 Pillow 줄바꿈
  (`_pil_line_starts()`, 폰트도 없으면 `_char_line_starts()` 27자 근사)로 수행. 슬라이드 수가 모자라면
  복제, 남으면 정리하며 반환값은 종전과 같은 순삽입 슬라이드 수. 복제한 슬라이드에서 본문 상자를 못
  찾으면 `RuntimeError`.
  **제거된 코드**: `layout_units_on_slides`/`layout_units_on_slides_pil`/`_verify_and_rebalance_pages`/
  `_rebalance_reading_slides_post_write`(+`_try_absorb_underfull`)/`_split_and_adjust_via_com`/
  `_split_para_at_lines`/`_prevent_widow_tails`/`_count_slide_lines_verified`/`_count_slide_lines`/
  `_rendered_wrap_count`/`_PIL_WRAP_SAFETY` 등(경위는 `docs/ooxml-pitfalls-log.md`와
  `docs/missa_to_ppt 구현 계획 변경이력.md`).
- **종료 텍스트 병합 판정**: 마지막 슬라이드를 확정할 때 얻은 줄 수(COM 실측 / Pillow 줄 수 / 폰트 없을
  때만 27자 근사)로 `merge_threshold`(5) 이하이면 통합한다.
- **종료 슬라이드**: `_align_ending_slides_to_제2독서()`(위치·크기를 제2독서 기준으로 복사 —
  빈 텍스트 도형(`BlackBg`)은 falsy 함정을 피하려 이름으로 먼저 제외), `_reposition_merged_
  ending_shapes()`(1줄 높이를 `_content_line_height_emu()`로 폰트 실측, 역산 금지),
  `_move_ending_shape_to_next_slide()`(한 슬라이드에 안 들어가면 통째로 이동).
- 안전 마진 상수: `_RENDER_WIDTH_CALIBRATION=1.03`(한글, Pillow 커닝 미지원 보정),
  `_FIT_HEIGHT_SAFETY_MARGIN=1.15`(화답송/복음환호송/영성체송 공용 `_adjust_fit_if_needed()`가
  쓰는 값, 아래 "Pillow 줄 높이 14~18% 과소" 실측치 재사용). 검증된 계수만 재사용하고 새 임의
  값을 발명하지 않는다(CLAUDE.md "여러 곳에서 재사용하는 fit-보정 함수..." 원칙).

## 1.5 슬라이드 복사·서식 보존 유틸 (`missa_ooxml_utils.py`)

- `duplicate_slide(prs, src_idx)`: 같은 프레젠테이션 내부 복제(shape 트리 + 이미지 rel + 배경 +
  `showMasterSp` 속성 모두 복사). **레이아웃은 `src_slide.slide_layout`을 그대로 재사용**한다
  (이름으로 "빈" 레이아웃을 검색하는 `_blank_layout()` 방식은 같은 프레젠테이션 내부 복제에는
  쓰지 않음 — 이름이 다른 "빈" 레이아웃이 여러 개 공존하면 원본과 다른 레이아웃이 골라질 수
  있음).
- `insert_slide_copy(prs, position, src_idx)`: `duplicate_slide` 후 이동.
- `copy_slide_from_prs(target_prs, position, source_prs, source_idx)`: **다른** 프레젠테이션
  에서 원본 서식(배경 포함) 그대로 복사. 소스/타겟 슬라이드 크기가 다르면 `scale_x`/`scale_y`
  비율로 도형 위치·크기와 폰트 크기·단락 여백(`text_scale = min(scale_x, scale_y)`)을 함께
  보정한다. `_build_com_probe_pptx()`가 크기를 원본과 동일하게 맞춰 호출하므로 COM 실측 경로
  에서는 이 보정이 항상 스킵된다(scale=1).
- `_set_slide_bg_black(slide, prs=None)`: `p:bg` 삽입 + 전체 크기 `BlackBg` 사각형 삽입.
  **같은 프레젠테이션 내부 템플릿 복제(`insert_slide_copy`)에만 걸어도 안전**하고, **다른
  프레젠테이션에서 복사한 슬라이드(`copy_slide_from_prs`)에는 걸면 안 된다**(이미 원본 배경을
  정확히 복사했으므로 덮어쓰면 서식 파괴).
- 텍스트 서식 보존 유틸(모두 기존 run/pPr을 deepcopy 후 텍스트만 교체): `_set_single_para_text`,
  `_replace_para_text_clone`, `_set_reading_text`, `_update_prefix_in_runs(para, segments)`
  (`segments=[(old_text, new_text), ...]` — 세그먼트별 독립 재배치, 옛 run 길이로 통짜
  재배치하지 않음), `_restore_para_from_backup()`.

## 1.6 PPT2007 호환성 처리 (`strip_ppt2007_incompatible`)

저장 직후 호출. `ppt/changesInfos/`(공동 작성 추적) 제거, `<a14:imgProps>`(고화질 이미지
레이어) 제거, 슬라이드 내 외부 비디오 링크 제거, 고아 image/hdphoto Relationship 제거,
`docProps/app.xml` 카운트 수정.

## 1.7 검증 (`missa_sections.py`)

- `validate_pptx_structure(pptx_path)`: 저장된 OOXML이 손상 없이 열리는지.
- `validate(prs, json_data)`: liturgy 텍스트 존재(오류) → 입당송·영성체송·화답송 슬라이드
  존재(경고, 화답송 라벨은 `'화답송'`과 평일 템플릿의 `'화 답 송'` 둘 다 인식) →
  `_missing_orange_verse_numbers()`로 절 번호별 오렌지색 렌더링 대조(경고). `is_sunday`
  매개변수는 제거됐다(화답송 체크를 항상 수행).

## 1.8 회귀 테스트 스위트

```
pytest tests/test_missa_regression.py tests/test_ppt_com_verify.py -v
```

`tests/` 폴더로 이동(2026-10-01 저장소 정리)했으며, 루트 `conftest.py`가
`sys.path`에 저장소 루트를 추가해 `tests/` 안에서 애플리케이션 모듈을 바로 import할 수 있게
한다. **선택 실행 우선, 전체 실행은 최종 게이트 1회** — 개발 중에는 건드린 클래스/파일만
`-k`나 `::클래스명`으로 좁혀 돌리고, 전체 스위트(`pytest tests/ -v`, 총 300개 이상)는 리뷰
통과 후 최종 게이트로 한 번만 돌린다. 알려진 환경 함정(MSAL 행 유발 테스트 3개 deselect 필요,
PowerPoint COM 좀비 프로세스 정리 필요)은 CLAUDE.md "회귀 테스트" 절 참고.

테스트 정책: 새 동작은 `tests/test_missa_progression.py`에 먼저 실패하는 테스트로 명세(TDD
red→green)하고, 안정화되면 `regression-qa`가 `tests/test_missa_regression.py`로 승격한다.

## 1.9 함수 목록 (모듈별 — 2026-10-01 기준)

CLAUDE.md "핵심 파일" 절의 모듈 책임 요약과 함께 본다. 함수 동작 설명은 이 문서의 해당 절을
참고하고, 여기서는 "지금 어느 파일에 있는지"만 정리한다.

**`missa_to_ppt.py`**(진입점): `_build_arg_parser`/`parse_args`/`find_files`/`get_json_data`
(§3.7 `mass_type`/`input_date_str` 파라미터)/`apply_화답송_override`/`_infer_hymn_numbers`,
`_dispatch_cli_or_gui`/`_run_gui_mode`/`_main_should_use_preloaded_or_fallback`/
`_show_children_mass_not_supported_message`(§2.1),
`resolve_mass_flags`/`merge_youth_gospel_content`/`_resolve_mass_type`/
`_youth_content_date_str`(§3.7), `_get_youth_english_mass`(§3.8), 청년 CLI 성가 인자 파싱(`_parse_youth_hymn_cli_entry`/`_args`), `main`.

**`missa_ooxml_utils.py`**(leaf, 프로젝트 내부 의존성 없음): §1.5의 슬라이드 복사 유틸 전체 +
`find_slide_with_text`/`find_shape_exact_text`/`_slide_text`/`all_slide_texts`/
`_find_content_shape`/`_has_ending_text`/`_clear_text_frame` + `HYMN_TYPES` 상수(entry·
content_updaters 순환 임포트 방지용) + `_safe_next_slide_partname`(python-pptx 몽키패치).

**`missa_gui.py`**(leaf, Tkinter 팝업 + 설정): `_load_config`/`_save_config`/
`_com_verification_enabled`/`is_sunday_mass`/`output_folder_key`/`_log_dir_for`, 팝업군
(`_ask_date_popup`/`_ask_input_files_popup`/`_ask_combined_input_popup`/`_ask_numbers_popup`/
`_ask_youth_hymn_popup`/`_report_progress`/`_run_with_progress_window`/`_show_result_window`),
`_get_app_version`(§3.1), OneDrive 관련 함수 전체(§3.3·§3.4).

**`missa_sections.py`**: §1.3·§1.7 + 청년 전용 탐지 함수(§3.2).

**`missa_reading_layout.py`**: §1.4 전체(한글+영문 공유).

**`missa_content_updaters.py`**: 섹션별 업데이트(`update_title_slide`/`update_입당송`/
`update_reading_title_slide`/`update_복음_title_slide`/`update_화답송`/
`update_복음환호송(..., mass_type='adult')`/`update_영성체송`), 본문 자동 맞춤
(`_adjust_fit_if_needed` 등 §1.4), `replace_시작기도문`/`replace_미사후기도`/
`insert_공지사항`/`replace_성가`(성인) + 청년 전용(§3.5) `resolve_youth_hymn_pptx`/
`_update_youth_title_slide`/`replace_성가_youth`/`_replace_one_youth_song`.

**`missa_psalm_score_image.py`**(leaf, `missa_ppt` 계열을 import하지 않는 단방향 의존):
화답송 악보 사진 → 슬라이드. 진입점 `render_화답송_score_slide()`. §2.4(요구사항 문서) 참고.

**`missa_youth_gospel.py`**(leaf): universalis.com 영문 미사 조회. `get_youth_mass(date)`
(§3.6), 하위호환 래퍼 `get_youth_gospel`/`parse_gospel`.

**`missa_youth_hymn_pdf.py`**(leaf): PDF 성가집 → PPT. `build_hymn_pptx()`(§3.5). 자산
`assets/청년미사_성가_template.pptx`, `assets/나주노_copyright_bbox_cache.json`(개발자 1회
OCR 생성, 런타임은 캐시만 읽음).

**`missa_onedrive.py`**(leaf): §3.3. `get_access_token`/`list_children`/`download_file`/
`upload_file`/`ensure_folder`/`get_item_metadata`/`_resolve_remote_base`/`_ensure_base_resolved`.

**`missa_updater.py`**(leaf): §3.1. `check_for_update`/`apply_update`/`get_latest_commit_sha`.

**`ppt_com_verify.py`**(별도 모듈): `is_available`/`count_slide_lines`/`shutdown`.

## 1.10 의존성 / 빌드

python-pptx, lxml, Pillow, numpy(화답송 악보 사진 배열 연산), pywin32(PowerPoint COM, 미설치
시 Pillow 전용 폴백), requests/beautifulsoup4(크롤링), msal(OneDrive 인증), tkinter(표준
라이브러리), pytest.

빌드는 **exe(PyInstaller) 방식을 폐기하고 Python 인터프리터 + `.bat` 런처**로 전환했다(§3.1)
— `run_missa.bat`가 `pythonw.exe missa_to_ppt.py`를 실행한다. `missa_to_ppt.spec`/`dist/`는
과거 exe 빌드 산출물 경로로 여전히 존재하지만(`PyInstaller.config.CONF['distpath']`를
`dist/`로 지정), 현재 배포 주력 경로는 zip 소스 배포다(`tools/build_dist_zip.py`,
`tools/build_youth_dist_zip.py`).

---

# 2부. 성인미사 전용 구현

성인미사는 위 1부의 공유 함수들이 `mass_type='adult'`(기본값)로 호출될 때의 동작 전체이며,
`resolve_mass_flags()`가 4개 값을 전부 `is_calendar_sunday`에서 파생시키므로 이 문서 이전
버전(v1.0~2026-09 초)의 성인 전용 동작과 100% 동일하다. 아래는 성인 경로에서만 실행되는
분기·함수를 모아 정리한다.

## 2.1 화답송 — `update_화답송(prs, json_data, sections, 화답송_pptx_path,
화답송_img_path=None, is_sunday=True)`

- **주일**(`is_sunday=True`): 필요 슬라이드 `2n+1`(악보 n+1장 + 텍스트 n장), 짝수 인덱스가
  악보. 악보 슬라이드는 `copy_slide_from_prs()`로 화답송 악보 PPT에서 복사(§1.5). PPT가
  없고 사진만 있으면 `missa_psalm_score_image.render_화답송_score_slide()`를 내부에서
  지연 import해 대신 호출(모듈 최상단 import 없음 — 상시 하드 의존 아님).
- **평일**(`is_sunday=False`): 텍스트 슬라이드만(`needed = len(verses)`), 참조 PPT의 화답송
  템플릿을 `insert_slide_copy()`로 복제. 후렴(◎)도 표시.

## 2.2 복음환호송 — `update_복음환호송(prs, json_data, sections, mass_type='adult')`

`mass_type='adult'`(기본값): 참조 PPT의 기존 3단락(`para[0]`=◎/`para[1]`=○/`para[2]`=◎)을
템플릿으로 재사용해 각 줄을 deepcopy 후 텍스트만 교체. 청년 분기는 §3.6 참고.

## 2.3 성가 교체 — `replace_성가(prs, 성가_map, hymn_numbers, copy_scores=True)`

5종(입당→봉헌→성체→2차봉헌→파견) 순서대로, 슬라이드 수가 바뀔 때마다 `find_sections()`를
재호출. `copy_scores=True`(주일): 기존 content 삭제 후 성가 PPT 슬라이드를 `copy_slide_
from_prs()`로 삽입. `copy_scores=False`(평일): 삽입 생략, 삭제만(악보 슬라이드 제거,
divider만 남음). 구분 라벨 교정은 `_update_성가_header(slide, expected_type, new_number,
slide_no=None, n_slides=None, label_override=None)`(`CANONICAL_LABEL` 매핑, `label_override`는
청년 가톨릭성가 분기용 — §3.5).

## 2.4 미사 후 기도 — `replace_미사후기도(prs, path, sections) -> int` (평일미사 전용)

파일 미지정 시 아무 것도 하지 않음(오류 아님). 있으면 기존 슬라이드를 삭제하고 새 PPT로
교체. 위치: 파견 콘텐츠 → blank → 미사 후 기도 → blank → (다음 섹션 또는 끝).

## 2.5 공지사항 — `insert_공지사항(prs, path, sections) -> int`

삭제/교체가 아닌 순수 삽입. `sections.get('2차봉헌_content_end')`가 `None`이면(슬롯 없음)
경고 후 `return 0`. 앞쪽 blank를 그대로 재사용하고 뒤쪽 구분선 1장만 `insert_slide_copy()`로
추가. 반환값 = 공지사항 장수 + 1.

---

# 3부. 청년미사 전용 구현

청년미사(토요일 저녁) 지원은 2단계로 진행됐다: **1단계**(영문 복음 조회 + PDF 성가집→슬라이드,
독립 leaf 모듈 2개) → **2단계**(공유 파이프라인 함수에 `mass_type` 분기 추가 + `main()` 배선 +
배포/OneDrive 인증/전체 통합). 둘 다 완료됐다. 청년 전용 참조 PPT 127슬라이드(`Template_토요일
저녁 청년 주일미사_20260822_연중 제21주일.pptx`) 기준 슬라이드 인덱스 표는 §3.2 참고.

## 3.1 배포 전환

- **업데이트 메커니즘**: `missa_updater.py`(leaf) — `get_latest_commit_sha()`(GitHub API,
  인증 불필요·public repo) → `check_for_update()`(config.json의 `last_update_commit`과 비교)
  → `apply_update(sha)`(zip 다운로드 → 압축 해제 → 화이트리스트(`*.py`, `*.spec`, `VERSION`,
  `assets/`)만 복사). `config.json`/`output/`/`reference/`/`cache/`/`.git/`은 절대 덮어쓰지
  않는다. GUI의 "🔄 업데이트 확인" 링크에서 트리거.
- **배포 형태**: PyInstaller exe를 폐기하고 Python 인터프리터 + `.bat` 런처로 전환(§1.10) —
  `.py` 교체만으로 갱신되지 않는 exe 방식으로는 업데이트 메커니즘이 성립하지 않기 때문.

## 3.2 `find_sections()` 청년 전용 3곳

전체 슬라이드 골격은 성인 로직과 공통이고(§1.3), 신규 로직이 필요한 곳은 정확히 3곳이다.

| 문제 | 실측 근거 | 해결 |
|---|---|---|
| 성가 title 탐지 | 진짜 title 6장은 `AUTO_SHAPE`, 우연히 텍스트가 같은 114번("파 견" 인사말)은 `PLACEHOLDER` | `_is_hymn_title_youth()` — shape 타입으로 구분(텍스트 "첫 매칭" 의존 금지) |
| 복음환호송 "가운데" 탐지 | 48/49/50번 전부 라벨 포함, 가운데(49)만 `○`로 시작, 앞뒤는 `©`만 | `_find_복음환호송_middle()` — `○` 포함 여부로 판별. 성인 템플릿에서도 참이라 어댑터 없이 대체 가능(회귀 없음) |
| 성체 2곡 콘텐츠 분리 | 구분 슬라이드 없이 title2가 content1 뒤에 바로 이어짐 | 콘텐츠 스캔 중 `_is_hymn_title_youth()`가 다시 참이 되는 지점을 곡 경계로 처리, `{htype}_songs`를 1~2곡 리스트로 반환 |

참조 PPT 슬라이드 인덱스(0-based, 127장): 12=입당title, 13-16=입당content, 49=복음환호송
가운데, 52=복음title(한글 인트로 유지), 53-55=복음content(영문), 56=복음종료(한글 클로징
자동 보존), 65=봉헌title, 66-69=봉헌content, 103=성체1title, 104-105=성체1content,
106=성체2title(기타), 107=성체2content, 109=2차봉헌title, 110-111=2차봉헌content,
112=2차봉헌 뒤 구분(`insert_공지사항()` 참조 지점), 114=파견 인사(PLACEHOLDER, title
아님), 116=파견title(AUTO_SHAPE), 117-119=파견content.

## 3.3 OneDrive 인증 (`missa_onedrive.py`)

- Device Code Flow(MSAL Python), `AUTHORITY="https://login.microsoftonline.com/consumers"`,
  `SCOPES=["Files.ReadWrite"]`. App ID `21b7d1a4-87bb-4148-8ae8-8b57231c125b`(퍼블릭 클라이언트,
  client secret 불필요 — public repo 노출 안전). 토큰 캐시 `msal_token_cache.bin`
  (`MSAL SerializableTokenCache`, `has_state_changed`일 때만 쓰기, `.gitignore` 등록).
- `_APP_CACHE = [app, cache]`(모듈 레벨) + `_get_app()`로 `PublicClientApplication` 인스턴스를
  프로세스 수명 동안 1회만 생성(매번 생성하면 ~1초 걸림 — instance discovery 재수행 추정).
  `_headers()`를 거치는 모든 공개 함수가 이 캐시로 자동 이득.
- **remoteItem 함정**: 공용 계정 OneDrive의 최상위 `PPT 문서`는 진짜 폴더가 아니라 다른
  드라이브를 가리키는 **공유 바로가기**다. `_resolve_remote_base(first_segment)`가 실제 GET으로
  remoteItem 여부를 확인해 `(driveId, itemId)`를 `_REMOTE_ITEM_CACHE`(프로세스 수명)에 저장하고,
  `_ensure_base_resolved(path)`를 4개 공개 함수 모두 URL 생성 전에 호출한다. `_item_url()`
  자체는 순수 함수로 유지(캐시를 읽기만 함) — 네트워크 GET과 URL 조립을 분리해 기존 순수
  문자열 변환 테스트를 보존한다. 404는 여기서 raise하지 않고 `None` 처리(진짜 "경로 없음"은
  호출부의 기존 404 처리가 담당).

## 3.4 `config.json` 스키마

로컬 절대경로 키(1순위, 기존)와 원격 상대경로 키(2순위, 신규)가 공존한다
(`missa_gui._resolve_onedrive_folder()`):

| 로컬 절대경로 키 | 원격 상대경로 키(기본값) |
|---|---|
| `onedrive_hymn_folder` | `onedrive_hymn_path` (`PPT 문서/09.가톨릭 성가/성가-악보버전`) |
| `onedrive_youth_hymn_folder` | `onedrive_youth_hymn_path` (`PPT 문서/20.청년 미사/2.성가`) |

우선순위: (1) 로컬 키가 유효한 폴더면 그대로(Graph API 전혀 안 탐) → (2) 없으면 원격 경로로
미러링/조회 → (3) 그래도 실패하면 수동 선택 팝업.

> **청년미사는 §3.4.1로 대체 예정**(2026-10-02 요구사항 확정, 구현 대기): 위 표의
> `onedrive_youth_hymn_folder`는 청년미사에서 더 이상 읽지 않고 신규 키 `onedrive_ppt_folder` 하나로
> 판정한다. 성인미사(`onedrive_hymn_folder`)는 그대로.

## 3.4.1 청년미사 접속 모드(로컬 폴더 우선 / 로그인 폴백) — 설계·구현 완료(2026-10-02)

요구사항은 `docs/missa_to_ppt 요구사항.md` §3.3.1. 아래는 코드 전수 조사(2026-10-02) 결과로
확정해 구현한 변경 지점이다. 테스트는 `tests/test_missa_youth_local_mode.py`(14개), 테스트 간 모드
상태 격리는 루트 `conftest.py`의 autouse 픽스처(`_youth_ppt_mode_isolated`)가 맡는다.
구현 시 `find_youth_onedrive_hymn_file()`에 `local_kind` 인자를 추가했고, 구 청년 성가 폴더 키
(`onedrive_youth_hymn_folder`)를 읽던 `_youth_onedrive_local_sync_dir()`는 청년 경로에서 더 이상
호출되지 않는다(함수 자체는 기존 테스트 호환을 위해 남김).

**1) 모드 판정 단일 지점 — `missa_gui.py`**
- `get_youth_ppt_mode(ask_if_missing: bool) -> str`(`'local'`/`'fallback'`)과 모듈 상태
  `_YOUTH_PPT_FOLDER = [None]`(Path|None). 호출 순서: `config['onedrive_ppt_folder']`가 유효한
  디렉터리면 `'local'` → 아니면 `ask_if_missing`일 때 `_ask_youth_ppt_folder_popup()`(기존
  `_ask_onedrive_path_popup()`의 안내문만 바꿔 재사용) → 선택한 폴더가 디렉터리면 `_save_config`
  후 `'local'`, 취소/무효면 `'fallback'`(저장 안 함 — 다음 실행에 다시 묻는다).
- `ask_if_missing=False`는 GUI 없는 CLI 실행(`--y입당` 등 인자 직접 지정)용 — 팝업 없이 설정이
  있으면 local, 없으면 fallback(로그인).
- 시작 시 1회만 호출하고 이후 모든 단계는 `_YOUTH_PPT_FOLDER[0]`/`get_youth_ppt_mode()`의 **저장된
  결과**를 읽는다(`get_onedrive_youth_hymn_remote_path()`가 겪은 "두 곳이 따로 판정해 어긋남"
  버그 재발 방지, CLAUDE.md 규칙 23·19). 경로 조합은 순수 헬퍼 `youth_local_subdir(kind)`
  (`'catholic_hymn'|'naju'|'yahweh'|'output'|'start_dir'`)가 담당하고 하드코딩 문자열은 이곳에만
  둔다(`'09.가톨릭 성가/성가-악보버전'`, `'20.청년 미사/2.성가/{나주노 성가|야훼이레 성가}'`,
  `'20.청년 미사/{YYYY}'`).

**2) 변경 대상 코드(전수 조사 결과)**

| 파일·위치 | 현재 | 변경 |
|---|---|---|
| `missa_to_ppt.py` `_run_gui_mode()` 1454~1474 | 청년이면 항상 `ensure_onedrive_login()` | 청년이면 먼저 `get_youth_ppt_mode(True)`; `'fallback'`일 때만 `ensure_onedrive_login()` |
| `missa_gui.py` `_ask_combined_input_popup()` ~1754~1800, ~1839, ~1996 | 청년이면 OneDrive 브라우저·다운로드 대기·프리페치 스레드 | `'local'`이면 3곳 모두 건너뛰고 일반 `askopenfilename(initialdir=PPT폴더)` 사용 |
| `missa_gui.py` `find_youth_onedrive_hymn_file()` 335~367 | 로컬 캐시 → Graph 재귀 검색 | `'local'`이면 `youth_local_subdir()` 폴더를 `rglob`으로 직접 검색(캐시/Graph 생략) |
| `missa_content_updaters.py` `resolve_youth_hymn_pptx()` 1644~1723 | `_youth_onedrive_local_sync_dir('onedrive_youth_hymn_folder')` | 새 키·모드 기준으로 교체(`'local'`: 새로 만든 성가를 `나주노 성가`/`야훼이레 성가` 폴더에 저장, 업로드 없음) |
| `missa_to_ppt.py` `_upload_youth_output_to_onedrive()` | (2026-10-02) 모드별 복사/업로드 | **2026-10-03 삭제** — 결과창 버튼 `copy_result_to_onedrive()`로 대체(§3.4.2) |
| `missa_gui.py` `_resolve_onedrive_folder()`/`get_onedrive_youth_hymn_root()`/`get_onedrive_youth_hymn_remote_path()` | 구 청년 성가 폴더 해석 | 청년 경로에서는 더 이상 호출하지 않음(폴백 모드 전용으로 남기거나 데드코드 정리 — 구현 중 grep으로 호출부 0건 확인 후 결정) |
| 성인미사 경로 | `get_onedrive_hymn_folder()` | **변경 없음** |

**3) 지켜야 할 제약**
- 로컬 모드에서는 `missa_onedrive`/`msal`을 import조차 하지 않는다(msal 미설치 PC 보호, 기존
  `_mirror_onedrive_folder` 주석과 같은 원칙) — 모든 Graph 호출은 폴백 분기 안에서 지연 import.
- 새 폴더 선택 팝업은 Tk 모달이라, `_run_gui_mode()`·`resolve_youth_hymn_pptx()`를 모킹 없이
  실행하는 기존 테스트(`test_y2m_*`, 청년 OneDrive/GUI 관련 테스트)가 실제 창 앞에서 무한 대기하지
  않도록 구현 시 해당 테스트를 먼저 찾아 모드를 `'fallback'`/`'local'`로 고정하는 monkeypatch를
  추가한다(CLAUDE.md 규칙 20).
- 폴더 검증 `_is_valid_youth_ppt_folder(path)`(= 디렉터리이며 `20.청년 미사` 하위 폴더 존재)를
  **선택 직후와 매 실행 시작 시 모두** 호출한다. 선택 직후 실패하면 경고 후 재선택, 재선택에서
  취소하면 폴백. 저장된 값이 검증에 실패하면 "설정 없음"과 동일하게 선택창을 다시 띄운다.
- 로컬 파일 선택창의 `filetypes`·공지사항 `.pptx` 검증·필수/선택 규칙은 성인 행과 동일 코드를
  재사용한다.

**4) 테스트 계획(프로그레션 → 안정화 후 회귀 승격)**
- 모드 판정: 유효 경로→local / 키 없음+선택→저장 후 local / 키 없음+취소→fallback(미저장) /
  저장 경로가 사라짐→선택창 재표시.
- `youth_local_subdir()` 경로 조합(3종 + 연도 폴더).
- 로컬 모드에서 성가 조회·생성 저장 위치, `missa_onedrive` 미호출(import 감시).
- 결과 PPT 복사: 연도 폴더 생성, 기존 `output/` 저장 유지, 복사 실패 시 경고만.
- 폴백 모드 회귀: 기존 K/L·`test_y2m_` 동작 불변.

## 3.5 OneDrive 파일 획득 — 두 메커니즘

**(A) 개별 파일 브라우저** (참조 PPT·화답송·시작기도·공지사항·미사후기도,
`_ask_onedrive_file_browser_popup()`): Treeview 탐색, 최상위 3개 허용 루트(`_ONEDRIVE_ROOT_
ALLOWED_FOLDERS`)만 노출, 하위는 무제한. 폴더 확장은 더블클릭·확장 화살표 모두
`<<TreeviewOpen>>` 가상 이벤트로 통일 처리(`_ensure_folder_children_loaded(iid)`) — 물리
이벤트별로 따로 걸면 확장 화살표 단일 클릭 경로를 놓친다. `_start_onedrive_prefetch_thread()`
가 입력창이 뜨는 즉시(청년미사만, 논블로킹) 3개 루트의 모든 하위 폴더/파일을 재귀
프리페치(`_prefetch_onedrive_children()`)해 캐시를 채운다. 선택한 파일은 공유 캐시 폴더
(`_ONEDRIVE_CACHE_ROOT = 'onedrive_cache'`, 날짜 폴더와 별개)에 원본 1부 보관 — 매니페스트
(`onedrive_cache/manifest.json`)의 `lastModifiedDateTime`이 **문자열 완전일치**할 때만
재다운로드 스킵(`_resolve_onedrive_download()`). **이 신선도 판단은 항상
`missa_onedrive.get_item_metadata()`를 실시간 호출한다 — 프리페치 캐시(트리 표시용 스냅샷)를
절대 재사용하지 않는다**(두 캐시 계층은 "무엇에 답하는 캐시인지"가 다름, CLAUDE.md "서로 다른
목적의 캐시 계층이 겹칠 때" 참고). 파일 선택(`_on_choose()`)은 다운로드를 기다리지 않고
목적지 경로만 계산해 즉시 반환하고, `_download_worker`(generation 추적 백그라운드 스레드)가
실제 다운로드를 맡는다 — 같은 파일명 재선택 시 구세대의 느린 다운로드가 최신 선택을 덮어쓰지
않도록 **커밋(`_commit_onedrive_blob_to_dest()`)은 `_download_lock` 안에서 generation을
재확인한 경우에만** 수행한다("확인" 클릭 시 `_wait_for_pending_downloads_or_report_failure()`가
미완료 다운로드를 전부 대기).

**(B) 성가 PPT 온디맨드 조회** (나주노/야훼이레/가톨릭성가, `find_youth_onedrive_hymn_file
(remote_path_key, default_remote_path, subfolder, pattern, cache_subdir)`, 2026-09-27
도입 — §3.5의 §6.2 이전에는 두 함수 모두 `get_onedrive_hymn_folder()`/
`get_onedrive_youth_hymn_root()`로 **폴더 전체**를 미러링했었다): 우선순위 (1) 로컬 캐시
(`cache/{cache_subdir}/`, 성가 PPT는 불변이므로 매번 재확인 안 함) → (2) `_find_onedrive_
file(remote_path, pattern)`으로 OneDrive를 재귀 검색(다운로드 없이 `list_children()`만,
개별 폴더 실패는 그 아래만 건너뛰고 나머지는 계속 탐색)해 매칭 파일 하나만 다운로드. 성인
미사의 `get_onedrive_hymn_folder()`(가톨릭성가 공유)는 이 전환과 무관하게 계속 전체 미러링
그대로다 — `mass_type == 'adult'`로 좁혀진 `missa_to_ppt.py` 호출부에서만 쓰인다.
`get_onedrive_youth_hymn_remote_path(root)`(캐시 역업로드 판단용, §3.9)만 여전히
`get_onedrive_youth_hymn_root()`의 반환값과 비교하는 용도로 남아 있다.

## 3.6 성가 5종 × 4출처 — `replace_성가_youth(prs, 성가_선택)`

`resolve_youth_hymn_pptx(entry, htype) -> (Presentation|None, resolved_title)`가 출처별로
분기(`entry = {'출처', '번호', '제목'}`):

- **나주노/야훼이레**: §3.5(B)로 기존 PPT 탐색, 없으면 `missa_youth_hymn_pdf.build_hymn_
  pptx()`로 생성 후 같은 캐시에 저장. 캐시 적중 여부와 무관하게 매번 `build_header_runs()`를
  다시 호출해 헤더를 이번 주 값(htype/제목)으로 강제 갱신.
- **가톨릭성가**: §3.5(B)로 `'성가 {번호}(?!\\d)'` 패턴 검색. 제목 미입력 시 파일명
  (`성가 {번호} {제목}.pptx` 관례)에서 파싱해 폴백(2026-09-24 수정 — PDF 인덱스가 없어
  자동조회 불가능하므로 빈 문자열 방지). 헤더는 `_update_성가_header(..., label_override=
  _YOUTH_HEADER_LABEL.get(htype))`로 처리(청년 템플릿의 '2차봉헌' 콘텐츠 헤더가 공백 없는
  표기라 성인 `CANONICAL_LABEL`의 `'2차 봉헌'`과 다름).
- **기타**: `(None, entry['제목'])` — PPT 없음. 타이틀 슬라이드 텍스트만
  `_update_youth_title_slide(slide, htype, 출처, number, title)`(기존 단락 전부 제거 후
  `sz` 문자열(`'5200'`=라벨/`'3500'`=제목/`'2000'`=출처)로 run 템플릿을 찾아 재조립 — 문단
  순서가 아니라 크기로 식별)로 갱신하고, 가사 placeholder는 채우지 않는다. **콘텐츠 슬라이드는
  "손대지 않는다"가 아니라 "성체2 슬롯과 동일한 빈 placeholder로 강제 교체"한다**
  (`_capture_youth_blank_content_scratch_slide()`로 성체2 슬롯을 scratch로 복제해두고, 각
  '기타' 슬롯을 그 scratch로 교체 — 템플릿 기본 콘텐츠가 실제 악보 그림인 htype(예: 2차봉헌)에
  '기타'를 고르면 "손대지 않는다"가 곧 "지난주 그림이 남는다"가 되는 버그를 막기 위함).

`replace_성가_youth()`는 5개 htype을 처리한 뒤 여분의 곡 슬롯을 삭제하며, **곡 하나 처리할
때마다** `find_sections(mass_type='youth')`를 재호출한다(htype 단위가 아니라 곡 단위 —
슬라이드 삭제/복사로 인덱스가 계속 밀리므로).

## 3.7 영문 복음 레이아웃

- `missa_youth_gospel.get_youth_mass(content_date_str)`(리다이렉트 차단 — 3xx면
  `GospelFetchError`, 최종 경로가 요청 날짜와 다르면 에러) → `merge_youth_gospel_content()`가
  한글 JSON의 `복음.content` 자리에 영문 본문을 끼워 넣음 → 기존 `replace_reading_slides()`
  파이프라인에 그대로 태움.
- 영문 페이지 분배도 §1.4의 슬라이드 단위 채우기(COM 실측, 폴백 Pillow `_pil_line_starts()`)를 그대로
  쓴다. 분할과 검증은 **항상 같은 `box_px`**를 써야 한다 — 과거 페이지네이션에만 적용되던
  `_PIL_WRAP_SAFETY` 안전 마진(0.97)이 실제로 들어가는 단어까지 다음 슬라이드로 밀어내는 반대 방향
  버그를 만들어(20260926 슬라이드 54 "father's" 실사용 사례) 무력화됐고, 슬라이드 단위 구조 전환 뒤
  이 계수 자체가 없어졌다(CLAUDE.md "텍스트 배치를 결정하는 단계와 검증하는 단계는 반드시 같은 박스 폭을
  써야 한다" 원칙; 회귀 `test_y2_G1b_20260926_father_fits_first_page`).
- `get_json_data(content_date_str, mass_type='adult', input_date_str=None)` — 청년 호출부는
  `input_date_str`(사용자가 입력한 토요일)을 넘겨, JSON 캐시 폴더를 `output_folder_key(
  input_date_str or content_date_str, mass_type)`로 계산한다(파일명은 여전히
  `missa_{content_date_str}.json`). `_get_youth_english_mass(content_date_str, output_key)`가
  영문 JSON을 `missa_en_{content_date_str}.json`으로 같은 폴더에 캐시-우선 저장한다(과거엔
  메모리에서만 쓰이고 파일로 저장되지 않았다). 성인 호출부는 `input_date_str=None`이라 100%
  하위호환.

## 3.4.2 'PPT 문서' 루트·찾아보기 시작 폴더·결과창 복사 버튼 — 구현 완료(2026-10-03)

요구사항은 `docs/missa_to_ppt 요구사항.md` §1.9(공통)·§2.1(C)(성인)·§3.3.1/§3.8(청년). 코드는 전부
`missa_gui.py`의 한 절('PPT 문서' 루트…)에 모았고 하위 폴더명 하드코딩은 그 절에만 있다.

- **루트 상태**: 청년은 기존 `_YOUTH_PPT_STATE`, 성인은 신규 `_ADULT_PPT_STATE['folder']`. 공용 조회 헬퍼
  `get_ppt_root(mass_type)`(청년은 `mode=='local'`일 때만 폴더, 아니면 None). 성인 선택·검증은
  `get_adult_ppt_folder(ask_if_missing)`/`_is_valid_adult_ppt_folder`(`09.가톨릭 성가` 포함 여부)/
  `_select_adult_ppt_folder_interactively`. 우선순위: 캐시 → `onedrive_ppt_folder` →
  `onedrive_hymn_folder`의 두 단계 위(기존 설정 역산, 저장은 안 함) → 선택창(선택 시
  `onedrive_ppt_folder`+`onedrive_hymn_folder` 저장). 선택창은 `_run_gui_mode()`(성인)만
  `ask_if_missing=True`로 호출한다(CLAUDE.md 규칙 20 — 다른 호출이 모달 앞에서 멈추지 않게).
- **시작 폴더**: `_BROWSE_START_SUBPATHS`(공통/성인/청년 표) → `browse_start_subpath(mass_type, key)` →
  `resolve_browse_start_dir(root, subpath)`(`_find_child_dir`가 공백 무시 비교, 없으면 있는 곳까지) →
  `browse_initial_dir(mass_type, key)`가 `_browse()`의 `askopenfilename(initialdir=…)`에 쓰인다. 화답송(성인)은
  `psalm_browse_start_dir()`(`last_psalm_dir` 우선) + 선택 후 `remember_psalm_dir()`. 청년 폴백 브라우저는
  `_ask_onedrive_file_browser_popup(start_subpath=…)`가 `_expand_start_path()`로 해당 폴더까지 펼친다.
  Graph 허용 폴더 상수는 실제 이름 `13.기도문`으로 정정.
- **결과 복사**: `result_copy_dir`/`result_copy_exists`/`copy_result_to_onedrive`(순수 로직, 로컬이면
  `shutil.copy2`, 청년 폴백이면 `ensure_folder`+`upload_file`) + 버튼 핸들러 `_copy_result_with_ui`(성인 루트
  미설정 시 선택창 → 덮어쓰기 확인 → 복사 → 저장 폴더 안내/오류 안내). `_show_result_window()`가 `btn_frame`에서
  '파일 열기' 바로 옆에 같은 글꼴(10pt bold)·폭 16의 버튼을 둔다(2026-10-04: 처음엔 3버튼 아래 작은 버튼이었음). 버튼 색은 `_RESULT_BTN_COLORS`(목업 C안: 파일 열기 #59001d, 복사 #871b24, 폴더 열기 #3f5a73, 닫기 흰 배경+#59001d 테두리). `main()`이
  `_last_output_mass[0]`에 미사 유형을 기록한다. `main()`의 `[8.5] OneDrive 업로드`와
  `_upload_youth_output_to_onedrive()`는 삭제됐다.
- **PowerPoint 안내**: `_show_powerpoint_background_notice()`(messagebox)와 그 호출을 삭제하고,
  `_run_with_progress_window()`가 진행 창 첫머리에 `_POWERPOINT_NOTICE_TEXT` 라벨을 만들어
  `_POWERPOINT_NOTICE_SECONDS`(5)초 뒤 `after()`로 제거한다(진행 메시지 라벨과 별개라 동시 표시).
- **테스트**: `tests/test_missa_youth_local_mode.py`(복사·성인 루트·시작 폴더·화답송), 진행 테스트 `j7a~c`·
  `e1`(파이프라인이 업로드하지 않음) 갱신. `conftest.py`의 autouse 픽스처가 성인 루트·`_last_output_mass`도 격리.

## 3.8 `main()` 청년 배선

- 미사 유형은 `--미사유형` CLI 플래그(`_dispatch_cli_or_gui()`가 `__main__` 최상단에서 처리)로
  정해지고, `_resolve_mass_type()`(`{'성인':'adult','청년':'youth'}`, 미지값은 `'adult'` 폴백)
  이 한글→영문 키를 변환한다.
- `_youth_content_date_str(date_str)`: 사용자 입력(토요일) +1일 → 콘텐츠 조회일(일요일).
  `resolve_mass_flags('youth', date_str)`에는 **토요일 그대로** 넘긴다(달력 판정 오염 방지).
- **청년 전용 방어 게이트**: adult의 "주일 5종 성가 악보 필수" 검증은 `mass_type == 'adult'`로
  좁혀져 있다(청년은 `files['성가']`를 쓰지 않음). 대신 `성가_선택`에 5종 중 하나라도 비어
  있으면 `sys.exit(1)`로 중단한다(`_ask_youth_hymn_popup()` 취소 시 `RuntimeError` — 무경고로
  빈 dict가 반환되면 성가 5종 전부가 조용히 스킵되는 것을 막는 2차 방어선).
- 결과 PPT는 `output/`에만 저장한다. 이전의 `_upload_youth_output_to_onedrive()`(자동 복사·Graph 업로드)는
  2026-10-03 삭제됐고, OneDrive 저장은 결과창 '원드라이브로 복사' 버튼(§3.4.2)이 맡는다.

## 3.9 범위 밖 — Graph 미러링 캐시 역방향 업로드

`resolve_youth_hymn_pptx()`가 나주노/야훼이레 PDF로 새 성가를 생성할 때, 지금 쓰는 `root`가
Graph 미러링 캐시 폴더(§3.4의 구(舊) 전체 미러링 경로)와 정확히 같으면
`missa_onedrive.upload_file()`로 되올리는 기능이 있다(`get_onedrive_youth_hymn_remote_path
(root)`가 판단). §3.5(B)의 온디맨드 전환 이후 나주노/야훼이레가 더 이상 이 미러링 경로를
기본으로 타지 않으므로, 실무에서 이 역업로드가 트리거되는 빈도는 줄었다 — 로컬 동기화 폴더가
있는 개발 PC에서는 OneDrive 앱이 자동 업로드해주고, Graph 미러링 캐시로만 동작하는 PC에서는
로컬 캐시에만 남아 다른 PC와 공유되지 않는다(회귀가 아니라 원래 "아직 완성하지 않은 개선").

## 3.10 청년 전용 테스트

`tests/test_missa_regression.py`에 `test_y2_`(25개, F/S/H/G/C군) +
`test_y2m_`(9개, main() 통합) 접두 테스트가 승격되어 있다. K/L그룹(OneDrive 프리페치·캐시)과
E~J그룹(실사용자 검수 후속) 다수는 청년미사 전용 경로라 관련 스코프 테스트로만 검증되고
(사용자 지시로 전체 회귀 실행이 매번 생략된 라운드가 많음), 최종적으로 B/A/C/D 그룹 승격
시점(2026-09-24)에 전체 스위트 322개 중 319 passed·3 deselected로 무회귀가 확인됐다. 상세
목록·리뷰 왕복은 `docs/missa_to_ppt 구현 계획 변경이력.md` 참고.

---

*OOXML 함정 재발 방지 규칙 전체는 `CLAUDE.md` 참고. 날짜순 상세 변경 이력(버그 발견 경위·
리뷰 라운드)은 `docs/missa_to_ppt 구현 계획 변경이력.md`. 원문(구 버전)은 `docs/archive/`.*
