# 구조·기술 사양 (ARCHITECTURE)

> 2026-10-06 새 문서 체계로 이관(원문 `missa_to_ppt 구현 계획.md`, 현재는 `archive/legacy-2026-10-06/`). "지금 어떻게 동작하는가"만 담는다. 모듈별 요약은 `missa_to_ppt-module-notes.md`, OOXML 함정은 `ooxml-pitfalls-log.md`·루트 `CLAUDE.md`. 요구사항은 `PRD.md`·`SPECS/`. 본문의 `§번호`는 이관 전 원문 절 번호다.

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
  때만 27자 근사)로 `merge_threshold`(5) 이하이면 통합한다. 단 `allow_ending_merge=False`(청년 영문 복음, `main()`이
  `mass_type != 'youth'`로 전달)이면 줄 수와 무관하게 통합하지 않는다(2026-10-05).
- **종료 슬라이드**: `_align_ending_slides_to_제2독서(prs, sections, align_gospel)`(위치·크기를 제2독서 기준으로 복사 —
  청년은 `align_gospel=False`라 복음 종료 슬라이드가 템플릿 위치를 유지 —
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

**`missa_updater.py`**(leaf): §3.1. `check_for_update`/`apply_update`/`get_latest_commit_sha`.

**`ppt_com_verify.py`**(별도 모듈): `is_available`/`count_slide_lines`/`shutdown`.

## 1.10 의존성 / 빌드

python-pptx, lxml, Pillow, numpy(화답송 악보 사진 배열 연산), pywin32(PowerPoint COM, 미설치
시 Pillow 전용 폴백), requests/beautifulsoup4(크롤링), tkinter(표준
라이브러리), pytest.

빌드는 **exe(PyInstaller) 방식을 폐기하고 Python 인터프리터 + `.bat` 런처**로 전환했다(§3.1; 청년은 2026-10-01,
성인은 2026-10-05) — `run_missa_성인.bat`/`run_missa_청년.bat`가 `pythonw.exe missa_to_ppt.py --미사유형 …`을
실행한다. 배포는 `tools/build_dist_zip.py [성인|청년|all]` 하나가 유형별 zip(`dist/missa_ppt_성인미사.zip`·
`dist/missa_ppt_청년미사.zip`)을 만든다(`KINDS` 표: zip 이름·실행 .bat·`docs/README_*미사_운영자.md`·추가 파일 — 청년만
성가집 PDF 2개). 공용 `install.bat`이 `pip install -r requirements.txt` 후 `create_shortcut.py`(폴더에 있는
`run_missa_*.bat`으로 유형을 판단해 `성인미사 PPT.lnk`/`청년미사 PPT.lnk` 생성)를 호출한다. 삭제된 exe 전용
도구(2026-10-05): `missa_to_ppt.spec`, `tools/build_shortcuts.py`, `tools/build_youth_dist_zip.py`, `run_missa.bat`.
코드의 `sys.frozen` 분기(`_SCRIPT_DIR` 등)는 무해해서 남겨 두었다.

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
  `assets/`, 그리고 `docs/README_*_운영자.md` → 설치 폴더 최상위(이미 있는 이름만, 2026-10-05))만 복사). `config.json`/`output/`/`reference/`/`cache/`/`.git/`은 절대 덮어쓰지
  않는다. GUI의 "🔄 업데이트 확인" 링크에서 트리거.
- **배포 형태**: PyInstaller exe를 폐기하고 Python 인터프리터 + `.bat` 런처로 전환(§1.10) —
  `.py` 교체만으로 갱신되지 않는 exe 방식으로는 업데이트 메커니즘이 성립하지 않기 때문. 청년이 먼저 전환했고(2026-10-01)
  성인도 같은 방식으로 통일했다(2026-10-05, 요구사항 §3.1). 업데이터 화이트리스트(`*.py`, `*.spec`, `VERSION`, `assets/`)는
  성인·청년 공통이다.

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

## 3.3 OneDrive 폴더 루트 3종 (2026-10-18 폴더 구조 개편)

OneDrive 접근은 **로컬 동기화 폴더 전용**이다. Graph API(`missa_onedrive.py`, MSAL 기기 코드 로그인)·원격 파일
브라우저·프리페치·다운로드 캐시·로그인 폴백은 모두 폐지·삭제됐다(`msal` 의존도 제거). 명세는
`SPECS/SPEC_OneDrive_폴더구조_개편.md`. 코드는 `missa_gui.py`의 한 절('OneDrive 폴더 루트 3종')에 모았고 하위 폴더
이름 상수는 이 절에만 있다.

- **루트 3종**: 공용(`00.공용 PPT`, 표지 하위 폴더 `09.가톨릭 성가`) · 성인(`10.성인 미사`, `10.기타 PPT문서`) ·
  청년(`20.청년 미사`, `2.성가`). `_ROOT_SPECS`(role → config 키·폴더명·표지·안내명), `_MASS_ROOT_ROLES`
  (성인=공용+성인, 청년=공용+청년), 프로세스 캐시 `_ROOTS`. 새 미사 유형은 두 표에 한 줄씩 추가한다.
- **조회·확정**: `get_root(role)`은 선택창 없이 캐시 → config(유효할 때만)를 반환한다. 선택창은 GUI 진입점
  `_run_gui_mode()`만 `ensure_ppt_roots(mass_type, ask_if_missing=True)`로 띄운다(CLAUDE.md 규칙 20). 필요한 루트만
  묻는다. 사용자가 상위 'PPT 문서'를 고르면 `_derive_roots()`가 `00./10./20.`을 찾아 모두 저장한다. 무효 선택은 안내
  후 재선택, 취소는 `False` → `_run_gui_mode()`가 안내(`_show_root_required_message`) 후 종료한다.
- **유효성**: `is_valid_ppt_root(role, path)` = 디렉터리 + 표지 하위 폴더 존재(`_find_child_dir`가 공백 유무 무시).
- **이관**: `_migrate_legacy_ppt_folder()`가 옛 `onedrive_ppt_folder`가 새 구조를 품은 경우에만 세 루트를 채우고 옛
  키를 지운다. 옛 구조('PPT 문서' 아래 `09.`/`10.`가 바로 있는 형태)는 유도 결과가 비어 선택창으로 간다.
  `onedrive_hymn_folder`/`onedrive_youth_hymn_folder`는 더 이상 읽지 않는다.
- **경로 규칙**: `ppt_subdir(kind, subfolder, year)`(`catholic_hymn`=`{공용}/09.가톨릭 성가/성가-악보버전`,
  `youth_hymn`=`{청년}/2.성가/{나주노 성가|야훼이레 성가}`, `adult_output`/`youth_output`=`{성인|청년}/{YYYY}`). 이미
  있는 폴더는 공백 차이를 무시한 실제 이름을 쓰고 없는 곳부터는 그대로 이어 붙인다(신규 폴더 생성 겸용). 루트가 없으면
  `RuntimeError`. `get_onedrive_hymn_folder()`는 이를 감싼 성인 호출부용 이름이다.
- **테스트**: 명세는 `tests/test_missa_progression.py`의 `test_fs_*`, 소비 동작은
  `tests/test_missa_youth_local_mode.py`. 루트 `conftest.py`가 모든 테스트의 `CONFIG_FILE`을 임시 파일로 바꾸고
  `_ROOTS`를 격리한다. 서브프로세스 실행(`tests/_gen_helper.py`)은 환경변수 `MISSA_CONFIG_FILE`로 임시 설정을 쓴다.

## 3.4 `config.json` 스키마 (OneDrive 관련)

| 키 | 의미 |
|---|---|
| `onedrive_common_folder` / `onedrive_adult_folder` / `onedrive_youth_folder` | 세 루트의 로컬 절대경로 |
| `browse_dirs` | `{행 id: 폴더}` — 첫 화면 '찾아보기'로 고른 폴더를 행별로 기억(행 id: `시작기도`·`미사후기도`·`공지사항`·`화답송_pptx`·`adult.ref_pptx`·`youth.ref_pptx`) |

옛 키(`onedrive_ppt_folder`·`onedrive_hymn_folder`·`onedrive_youth_hymn_folder`·`onedrive_hymn_path`·
`onedrive_youth_hymn_path`·`last_psalm_dir`)는 읽지 않는다(`onedrive_ppt_folder`는 이관 대상, `last_psalm_dir`은
화답송 행의 첫 시작 폴더로만 흡수된 뒤 삭제).

## 3.4.2 찾아보기 시작 폴더·행별 기억·결과창 복사 버튼

- **시작 폴더**: `browse_initial_dir(mass_type, key)` 우선순위 = `browse_dirs`(존재할 때만) → `_BROWSE_DEFAULTS`
  (행 id → 루트 role + 하위 경로: 시작기도·미사후기도 `{공용}/13.기도문`, 공지사항 `{공용}/11.공지사항 PPT문서`, 참조 PPT
  성인 `{성인}/10.기타 PPT문서/PPT 자동화/1.Template`·청년 `{청년}/1.Template`, 화답송 `{성인}/14.화답송 악보`; 없는
  하위 폴더는 `resolve_browse_start_dir`가 있는 곳까지) → 루트 → `None`(OS 기본). `_browse()`가 모든 행에서
  `askopenfilename(initialdir=…)` 후 `remember_browse_dir()`로 선택 파일의 폴더를 저장한다.
- **결과 복사**: `result_copy_dir`/`result_copy_exists`/`copy_result_to_onedrive`(`shutil.copy2`, OneDrive 앱이 동기화) +
  버튼 핸들러 `_copy_result_with_ui`(루트 미설정이면 선택창 → 덮어쓰기 확인 → 복사 → 안내). `_show_result_window()`는
  '파일 열기' 바로 옆에 같은 글꼴의 복사 버튼을 둔다(색은 `_RESULT_BTN_COLORS`). `main()`이 `_last_output_mass[0]`에
  미사 유형을 기록한다.
- **PowerPoint 안내**: `_run_with_progress_window()`가 진행 창 첫머리에 `_POWERPOINT_NOTICE_TEXT` 라벨을 만들어
  `_POWERPOINT_NOTICE_SECONDS`(5)초 뒤 `after()`로 제거한다.

## 3.5 성가 PPT 조회

`find_youth_hymn_file(kind, subfolder, pattern)`가 `ppt_subdir(kind, subfolder)`를 `rglob('*.pptx')`로 검색해 패턴에
맞는 첫 파일을 반환한다(없으면 `None`, 잠금 파일 `~$` 제외). 가톨릭성가는 `catholic_hymn`, 나주노·야훼이레는
`youth_hymn`. 나주노·야훼이레를 새로 만들면 같은 폴더(`{청년}/2.성가/{출처 폴더}`)에 저장한다(업로드 없음). 성인
주일은 `get_onedrive_hymn_folder()`를 `rglob`한다.

## 3.6 성가 5종 × 4출처 — `replace_성가_youth(prs, 성가_선택)`

`resolve_youth_hymn_pptx(entry, htype) -> (Presentation|None, resolved_title)`가 출처별로
분기(`entry = {'출처', '번호', '제목'}`):

- **나주노/야훼이레**: §3.5로 기존 PPT 탐색, 없으면 `missa_youth_hymn_pdf.build_hymn_
  pptx()`로 생성 후 같은 폴더에 저장. 캐시 적중 여부와 무관하게 매번 `build_header_runs()`를
  다시 호출해 헤더를 이번 주 값(htype/제목)으로 강제 갱신.
- **가톨릭성가**: §3.5로 `'성가 {번호}(?!\\d)'` 패턴 검색. 제목 미입력 시 파일명
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

## 3.10 청년 전용 테스트

`tests/test_missa_regression.py`에 `test_y2_`(25개, F/S/H/G/C군) +
`test_y2m_`(9개, main() 통합) 접두 테스트가 승격되어 있다. E~J그룹(실사용자 검수 후속) 다수는 청년미사 전용 경로라 관련 스코프 테스트로만 검증되고
(사용자 지시로 전체 회귀 실행이 매번 생략된 라운드가 많음), 최종적으로 B/A/C/D 그룹 승격
시점(2026-09-24)에 전체 스위트 322개 중 319 passed·3 deselected로 무회귀가 확인됐다. 상세
목록·리뷰 왕복은 `docs/missa_to_ppt 구현 계획 변경이력.md` 참고.

---

*OOXML 함정 재발 방지 규칙 전체는 `CLAUDE.md` 참고. 날짜순 상세 변경 이력(버그 발견 경위·
리뷰 라운드)은 `docs/missa_to_ppt 구현 계획 변경이력.md`. 원문(구 버전)은 `docs/archive/`.*
