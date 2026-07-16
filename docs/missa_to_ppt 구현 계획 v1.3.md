# missa_to_ppt.py 구현 계획 v1.3

- 대상 파일: `missa_to_ppt.py` (~6,600줄), `missa_to_json.py`, `test_missa_regression.py`
- 최종 수정일: 2026-07-16
- 이 문서는 현재 구현 전체를 처음부터 끝까지 담은 완결 문서다. 과거 버전(`docs/archive/`)을
  함께 읽을 필요가 없다. 버전별로 무엇이 언제 추가됐는지는 맨 끝 "부록 A: 버전별 변경 이력"만
  참고하면 된다. OOXML 조작 시 지켜야 할 함정·규칙은 `CLAUDE.md`에 별도로 정리되어 있다.

---

## 1. 개요

참조 PPT를 python-pptx로 열어 JSON(미사 본문) + 성가/시작기도/화답송악보/미사후기도 PPT로
내용을 교체한 뒤 새 PPT로 저장한다. 참조 PPT의 폰트·서식·레이아웃·배경은 최대한 그대로
유지하고, 텍스트와 슬라이드 구성만 바꾸는 것이 핵심 원칙이다.

**실행 진입점**
- 인수 없이 실행 → 대화형 팝업 모드 (`_ask_date_popup` → `_ask_input_files_popup` →
  `_ask_numbers_popup`) 또는 EXE GUI 모드(`_ask_combined_input_popup`에서 미리 수집한 값 사용)
- 인수 지정 실행 → CLI 모드: `python missa_to_ppt.py YYYYMMDD [--입당 N ...] [--화답송 PATH] [--미사후기도 PATH] [--test]`

두 경로 모두 `main()`(6011행) 안에서 합쳐져 이후 처리 흐름은 동일하다.

## 2. 처리 파이프라인 (main 흐름)

`_report_progress(pct, label)`로 진행률을 GUI 진행창에 보고하며 아래 순서로 진행된다.

| 단계 | 진행률 | 내용 |
|---|---|---|
| 날짜 확정 | — | `is_sunday_mass(date_str)`로 미사 유형 판단, 로그 출력 |
| [1] 파일 확인 | 5% | 참조/시작기도/화답송악보/미사후기도/성가 파일 존재 여부 출력 |
| [2] JSON 생성 | 10% | `missa_to_json.py` 실행(없을 때만) → `get_json_data()`로 로드 |
| [3] 참조 PPT 열기 | 15% | `Presentation(ref_pptx)` |
| [4] 전례 텍스트 교체 | 20~75% | 제목 → 입당송 → 제1독서 → 화답송 → 제2독서 → 복음환호송/복음 → 종료 슬라이드 정렬 → 영성체송 |
| [5] 시작기도문 교체 | 80% | `replace_시작기도문()` (파일 지정 시) |
| [6] 성가 교체 | 88% | `replace_성가()` × 5종, `copy_scores=is_sunday` |
| [6.5] 미사 후 기도 | 92% | `replace_미사후기도()` (평일 + 파일 지정 시) |
| [7] 저장 | 95% | `prs.save(output_path)`, 파일명 `YYYYMMDD/{liturgy}.pptx` |
| [7.5] PPT 호환성 정리 | 96% | `strip_ppt2007_incompatible()` |
| [8] 검증 | 98% | `validate_pptx_structure()` → `validate()` |

## 3. 설정 및 파일 탐색

### 3.1 미사 유형 판단

```python
def is_sunday_mass(date_str: str) -> bool:
    return datetime.strptime(date_str, '%Y%m%d').weekday() == 6  # 6 = 일요일
```

전역 상태로 두지 않고 `is_sunday` 값을 각 함수에 파라미터로 명시 전달한다. 토요일도 현재는
`False`(평일미사)로 판정된다.

### 3.2 config.json 기반 OneDrive 경로 관리

성가 PPT는 스크립트 실행 위치와 무관하게 OneDrive 동기화 폴더에서 중앙 관리한다.

- `CONFIG_FILE = <script_dir>/config.json`, 형식: `{"onedrive_hymn_folder": "C:/.../성가폴더"}`
- `_load_config() -> dict` / `_save_config(config)`: 존재하지 않거나 파싱 실패 시 빈 dict
- `_ask_onedrive_path_popup() -> str`: `filedialog.askdirectory()`
- `get_onedrive_hymn_folder() -> Path`: config에 유효한 경로가 있으면 반환, 없거나 사라졌으면
  팝업으로 다시 입력받아 저장

### 3.3 `find_files(date_str, hymn_numbers) -> dict` (CLI 모드 파일 탐색)

`YYYYMMDD/` 폴더를 스캔해 다음을 채운 dict를 반환한다: `ref_pptx`, `시작기도`, `화답송_pptx`,
`미사후기도`, `성가`(dict).

- 화답송 악보: 파일명에 "화답송 악보" 포함
- 시작기도: 파일명에 "시작기도" 포함
- 미사후기도: 파일명에 "미사후기도" 포함
- 성가: OneDrive 폴더에서 `re.search(rf'성가 {num}(?!\d)', f.name)`로 우선 탐색, 못 찾으면 날짜
  폴더 fallback
- 참조 PPT: 위 항목에 해당하지 않고, 파일명에 오늘 날짜(`date_str`)가 **포함되지 않은** `.pptx`
  중 선택. `^\d{8}_` 패턴(YYYYMMDD_로 시작)이 있으면 그것을 우선 사용, 없으면 후보 중 첫 번째

### 3.4 `parse_args()` (CLI 인수 파싱)

`argparse`로 `date`, `--입당/--봉헌/--성체/--2차봉헌/--파견`(정수), `--화답송`, `--미사후기도`
(경로 override), `--test`(팝업 없이 폴더 파일로 번호 자동 추론)를 받는다. 성가번호가 하나라도
비어 있으면 `_infer_hymn_numbers()`로 테스트 기본값(`_TEST_HYMN_DEFAULTS`)을 채운다.

### 3.5 대화형 UI 팝업

- `_ask_date_popup() -> str`: YYYYMMDD 형식 검증, Enter/확인으로 제출
- `_ask_input_files_popup()` / `_ask_combined_input_popup()`: 참조 PPT(필수) → 화답송 악보
  PPT(선택, 주일에 비어 있으면 경고) → 시작기도 PPT(선택) → 미사후기도 PPT(선택) 순서의 파일
  선택 화면
- `_ask_numbers_popup(defaults) -> dict`: 성가 5종 번호 입력
- 진행창: `_run_with_progress_window()`가 `_report_progress()` 콜백을 받아 진행바 갱신
- 결과창: `_show_result_window()`

## 4. 섹션 탐색 (find_sections)

`find_sections(prs) -> dict` (1917행)이 슬라이드 전체 텍스트를 스캔해 각 섹션의 슬라이드
인덱스를 키-값으로 반환한다. 이후 거의 모든 처리 함수가 이 dict를 받아 동작하며, 슬라이드
수가 바뀔 때마다(삽입/삭제) **반드시 재호출**해야 인덱스가 어긋나지 않는다.

주요 키: `title`, `시작기도_start/end`, `입당송`, `제1독서_title/start/end`, `화답송_start/end`,
`제2독서_title/start/end`, `복음환호송`, `복음_title/start/end`, `영성체송`,
`{입당|봉헌|성체|2차봉헌|파견}_divider`, `{...}_content_start/end`.

- 독서·복음: title 슬라이드(성서명) 인덱스 + content 범위(start, end)를 `find_content_range()`/
  `find_복음_content_range()`로 계산
- 복음: shape 텍스트가 "복 음" 단독으로 존재하는 슬라이드를 title로 인식
- 성가: `_is_hymn_divider()`로 구분 슬라이드를 판별, 그 뒤 content 범위를 계산

## 5. 독서·복음 처리 (핵심 엔진)

가장 복잡하고 버그가 많이 발견된 부분이다. 처리는 4단계로 나뉜다: **파싱 → 사전 배분 →
슬라이드 기록 → post-write 재조정**.

### 5.1 절 파싱 — `parse_into_verse_units(content) -> list`

- `content`에 절 번호(숫자 + 이어지는 텍스트) 경계에서 regex split
- 각 조각을 `{text, verse_num, is_continuation}` 단위로 변환
- "주님의 말씀입니다"·"◎ 하느님"·"◎ 그리스도님" 등 종료 텍스트는 파싱 단계에서 이미 제외됨
  (JSON 생성 시 `missa_to_json.py`가 제거)

### 5.2 사전 배분 — `layout_units_on_slides(units) -> list[list[dl]]`

- `is_continuation=True`인 unit은 이전 unit과 하나의 논리 단락으로 병합
- 각 논리 단락을 `CHARS_PER_LINE`(27, 32pt 바탕체 24.8cm 텍스트박스 실측 기준) 단위로 word-wrap
  시뮬레이션해 display line(`dl`)으로 분해 — 이 값은 **근사치**이며, 실제 렌더링 줄 수는 §5.4의
  Pillow 계산으로 별도 검증한다
- `LINES_PER_SLIDE`(9)개의 dl씩 슬라이드(`page`)로 묶음. 절 경계와 무관하게 꽉 채우되, 문장이
  끊기면 안 되는 경우는 분리하지 않음
- 절 번호(`extra_verses`)는 dl 텍스트 첫머리에 올 때만 오렌지색(`ORANGE = RGBColor(255,192,0)`)
- `_verify_and_rebalance_pages(pages, label)`: 위 배분 결과를 한 번 더 검증해 사전 단계에서
  잡을 수 있는 불균형을 정리 (post-write 재조정과는 별개 단계)

### 5.3 슬라이드 기록 — `replace_reading_slides(prs, content_start, content_end, units_pages, template_idx, ...) -> int`

1. 기존 content 범위 뒤에서 `_has_ending_text()`로 연속된 종료 슬라이드 수(`n_ending`) 감지
2. `n_usable = 기존 슬라이드 수 - n_ending`. `needed`(계획된 페이지 수) > `n_usable`이면
   `insert_slide_copy()`로 부족분 추가, 적으면 초과분 삭제
3. 각 본문 슬라이드에 `_set_reading_text(text_frame, page_units, line_spacing)`로 텍스트 기록
   (절 번호 run은 오렌지, 본문 run은 원래 색 — pPr/rPr은 clone 방식으로 보존)
4. **post-write 재조정** — `n_content > 1`이면 `_rebalance_reading_slides_post_write()` 호출 (§5.4)
5. 종료 슬라이드들의 본문 텍스트박스를 비움(`_clear_text_frame`) — 참조 PPT 잔여 내용 제거
6. **종료 텍스트 병합 판단**: 마지막 실제 본문 슬라이드(post-write 재조정 이후 위치,
   `content_start + needed + n_rebalance_inserted - 1`)의 **실제 렌더링 줄 수**를
   `_count_slide_lines()`로 측정해 `merge_threshold`(5) 이하이면, 종료 슬라이드의 종료
   텍스트박스만 마지막 본문 슬라이드로 옮기고 종료 슬라이드는 삭제

   > 이 판단은 반드시 재조정 **이후**의 실제 상태를 봐야 한다. post-write 재조정이 슬라이드를
   > 추가/재배치할 수 있어, 재조정 이전 계획값(`units_pages[-1]`)을 쓰면 실제로는 짧은 마지막
   > 슬라이드인데도 병합 판단이 틀어질 수 있다(2026-07-16 발견·수정).

### 5.4 줄 수 계산 정확도

두 가지 계산 방식이 있다.

- `_wrap_line_count(text) -> int` (2646행): `CHARS_PER_LINE`(27자) 근사치 기반, 사전 배분(§5.2)에 사용
- `_rendered_wrap_count(text, pil_font, box_px) -> int` (3248행): Pillow 실제 폰트 메트릭 기반,
  post-write 재조정(§5.5)과 종료 슬라이드 병합 판단(§5.3-6)에 사용. `_count_slide_lines_rendered(slide)`가
  슬라이드의 콘텐츠 shape 전체에 대해 이 값을 합산한다

Pillow 계산이 실제 PowerPoint 렌더링과 어긋나지 않도록 `_get_slide_render_params(slide)`
(3266행)에서 두 가지를 신경 써야 한다.

1. **폰트 로딩**: 슬라이드의 첫 run에서 폰트명(`typeface`)과 크기(`sz`)를 읽어 Windows 폰트
   파일에 매핑한다. `batang.ttc`/`gulim.ttc`처럼 한 파일에 여러 서체가 들어 있는 TTC는 인덱스를
   실제로 확인해야 한다(이 환경은 0=Batang, 1=BatangChe / 0=Gulim, 1=GulimChe — 반대로
   가정하면 폭이 과대평가된다).
2. **커닝 보정**: 이 환경의 Pillow는 `libraqm`(커닝) 미지원 빌드라 글자 폭을 실제보다 넓게
   계산한다. `box_px`에 `_RENDER_WIDTH_CALIBRATION`(현재 1.03)을 곱해 보정한다. 실측 확인된
   슬라이드 기준 1.03~1.04에서만 전체 일치했다(0~1.02는 과소보정, 1.05 이상은 과보정).

### 5.5 Post-write 재조정 — `_rebalance_reading_slides_post_write(prs, content_start, n_content, label) -> int`

실제로 슬라이드에 기록된 텍스트를 다시 읽어(§5.4의 Pillow 기준) `LINES_PER_SLIDE`(9)에 정확히
맞도록 조정한다. 마지막 슬라이드는 별도 처리(아래)하며, 반환값은 초과 처리 중 새로 삽입한
슬라이드 수(`n_inserted`, 종료 슬라이드 병합 판단에 쓰임).

**메인 루프** (`while changed:` — 변화가 없을 때까지 전체를 반복 스캔, 마지막 슬라이드 제외):

- **9줄 초과**: 마지막 단락을 다음 슬라이드로 이동 시도. 다음 슬라이드도 넘치면(그리고
  마지막에서 두 번째가 아니면) 롤백 후 "현재 슬라이드에 `LINES_PER_SLIDE`줄까지만 남기고 나머지는
  다음으로" 단락 분리(`_split_para_at_lines`)를 시도. 그마저 불가능하면(이동해도 여전히 초과분을
  다 해소 못 하는 연쇄 초과) **현재 위치 뒤에 새 슬라이드를 삽입**해 초과분을 옮기고 처음부터
  재스캔한다(2026-07-16 추가 — 이전에는 "분리 불가" 경고만 남기고 방치했음)
- **9줄 미만**: 다음 슬라이드의 첫 단락을 흡수 시도, 초과하면 단락 분리로 필요한 만큼만 가져옴.
  다음 슬라이드가 단락 1개뿐이라도(예: 종료 텍스트 병합 대상 슬라이드) **현재 슬라이드가
  마지막이 아니면** 단락을 분리해 9줄을 채운다 — 다음 슬라이드에는 최소 1줄 이상 남긴다
  (2026-07-16 추가 — 이전에는 단락이 1개면 무조건 건너뛰어 마지막이 아닌 슬라이드가 9줄
  미만으로 남는 경우가 있었음)

**마지막 슬라이드 처리**: 9줄 초과 시에만 새 슬라이드를 삽입해 넘치는 단락을 옮긴다. 단,
이동 후 남는 줄이 `LINES_PER_SLIDE // 2` 이하로 불균형해지면 분리하지 않고 유지한다(과분리 방지).

**연속 단편 병합**: 단락 분리로 문장이 중간에 끊긴 채 인접 슬라이드에 놓인 경우("...하고" /
"불렀다." 등 문장이 안 끝난 단락 + 절 번호로 시작하지 않는 다음 단락), 공백 하나로 이어붙여
자연스러운 한 단락으로 합친다.

**병합 후 재검증**: 위 병합으로 단락 경계가 사라지면서 실제 렌더링 줄 수가 줄어들 수 있다
(단락 구분은 줄 앞에서 강제 개행되므로, 병합하면 남는 폭을 재활용해 원래보다 적은 줄에 들어갈
수 있음). 병합이 모두 끝난 뒤 각 슬라이드를 다시 스캔해 9줄 미만이면 다음 슬라이드에서 다시
흡수한다(2026-07-16 추가).

### 5.6 종료 슬라이드 위치 정렬

- `_align_ending_slides_to_제2독서(prs, sections)` (3964행): 제2독서의 마지막 종료 슬라이드에서
  종료 텍스트박스와 콘텐츠 텍스트박스의 위치·크기를 수집한 뒤, 제1독서·복음의 종료 슬라이드에
  동일하게 적용한다. 도형을 종료 키워드 포함 여부로 `ending`/`content` 두 종류로 분류하는데,
  **빈 텍스트 도형(`BlackBg` 배경 사각형 등)은 반드시 이름으로 먼저 제외**해야 한다 — 빈
  문자열은 falsy라 "content이면서 텍스트가 있을 때만 스킵" 조건에 걸리지 않고 오분류되어 잘못된
  크기로 리사이즈된다(2026-07-16 발견·수정. `CLAUDE.md` 참고)
- `_reposition_merged_ending_shapes(prs, sections)` (4079행): §5.3-6에서 본문+종료 텍스트가
  통합된 슬라이드의 종료 텍스트박스를, 실제 본문 마지막 줄로부터 두 줄 공백 위치로 재조정한다.
  `_align_ending_slides_to_제2독서()` 실행 **이후**에 호출해야 한다

## 6. 화답송 처리 — `update_화답송(prs, json_data, sections, 화답송_pptx_path, is_sunday=True)` (4371행)

`content`를 `\n` 기준으로 분리(`segments`), 첫 항목이 후렴(◎), 나머지가 절(○)이다.

- **주일미사**: 필요 슬라이드 = `2n+1`(악보 n+1장 + 텍스트 n장), 짝수 인덱스가 악보 슬라이드다.
  - 텍스트 슬라이드는 후렴 제외, 절 내용만(`○\t` + 내용)
  - 악보 슬라이드는 화답송 악보 PPT에서 `copy_slide_from_prs()`로 복사한다. **이 함수가 이미
    원본 배경(레이아웃/마스터 상속분까지 해석)을 정확히 복사하므로, 복사 후 `_set_slide_bg_black()`을
    호출해 덮어쓰면 안 된다**(2026-07-16 발견·수정 — 텍스트 슬라이드 분기는 같은 프레젠테이션
    내부 템플릿 복제이므로 `_set_slide_bg_black()`을 걸어도 안전함)
- **평일미사**: 텍스트 슬라이드만(`needed = len(verses)`), 참조 PPT의 화답송 텍스트 템플릿을
  `insert_slide_copy()`로 복제. 후렴(◎)도 각 슬라이드에 함께 표시(`◎\t...` + `○\t...` 2단락)

화답송 제목 갱신(`_update_화답송_title_in_slide`)은 주일·평일 공통.

## 7. 복음환호송 — `update_복음환호송(prs, json_data, sections)` (4607행)

참조 PPT의 기존 단락 3개를 템플릿으로 재사용한다: `para[0]`=첫 ◎ 알렐루야, `para[1]`=○ 구절,
`para[2]`=마지막 ◎ 알렐루야(각각 다른 `lnSpc`/`tabLst`를 가짐). 각 줄은 해당 템플릿을 deepcopy한
뒤 run 텍스트만 교체해, 대한체(◎)·바탕체(○) run 구조가 참조 PPT와 동일하게 유지된다.

## 8. 시작기도문 교체 — `replace_시작기도문(prs, 시작기도_path, sections)` (5224행)

`시작기도_start`~`시작기도_end` 범위 슬라이드만 삭제(직전의 빈 슬라이드는 범위 밖이라 자동
보존됨)하고, 시작기도 PPT의 슬라이드를 `copy_slide_from_prs()`로 그 자리에 순서대로 삽입한다.
파일 미지정 시 아무 것도 하지 않는다.

## 9. 성가 교체 — `replace_성가(prs, 성가_map, hymn_numbers, copy_scores=True)` (5506행)

5종(입당→봉헌→성체→2차봉헌→파견) 순서대로 처리하며, 슬라이드 수가 바뀌므로 매번
`find_sections()`를 다시 호출해 인덱스를 갱신한다.

- 공통(주일·평일 동일): divider 번호 갱신(`_update_성가_divider_number`), 헤더 타입/번호 수정
  (`_update_성가_header`, run 서식 보존)
- `copy_scores=True`(주일): 기존 content 슬라이드 삭제 후 성가 PPT 슬라이드를
  `copy_slide_from_prs()`로 삽입(악보 포함)
- `copy_scores=False`(평일): 삽입을 생략하고 삭제만 수행 — 참조 PPT에 남아 있던 기존 악보
  슬라이드가 제거되고 divider만 남는다

## 10. 미사 후 기도 — `replace_미사후기도(prs, path, sections) -> int` (5272행, 평일미사 전용)

- `path`가 없으면 아무 것도 하지 않음 (기존 미사 후 기도 슬라이드가 있으면 그대로 보존, 오류 아님)
- 있으면 기존 미사 후 기도 슬라이드를 삭제하고 새 PPT의 슬라이드로 교체
- 위치: 파견 성가 콘텐츠 → blank divider → 미사 후 기도 → blank → (다음 섹션 또는 끝)

## 11. 슬라이드 복사·서식 보존 핵심 유틸

**슬라이드 복사**
- `duplicate_slide(prs, src_idx) -> int`: 같은 프레젠테이션 내부에서 슬라이드 복제(shape 트리 +
  이미지 rel + 배경 복사), 맨 끝에 추가
- `insert_slide_copy(prs, position, src_idx)`: `duplicate_slide` 후 원하는 위치로 이동
- `copy_slide_from_prs(target_prs, position, source_prs, source_idx)`: **다른** 프레젠테이션에서
  슬라이드를 원본 서식(배경 포함) 그대로 복사. 소스 레이아웃 이름으로 타겟에서 매칭 레이아웃을
  찾아 사용하며, `_effective_bg()`로 슬라이드→레이아웃→마스터 순으로 실제 배경을 해석해 그대로
  옮긴다 — **이 함수가 배경을 이미 정확히 복원하므로, 호출부에서 다시 배경을 강제하면 안 된다**
  (§6 참고)
- `_copy_spTree` / `_copy_image_rels`(image·hdphoto 타입 모두 복사) / `_update_rId_in_spTree` /
  `_effective_bg` / `_copy_bg_image_rels`

**텍스트 서식 보존** (모두 기존 run/pPr을 deepcopy한 뒤 텍스트만 교체하는 clone 방식)
- `_set_single_para_text(tf, text)`: 단일 단락 텍스트 교체
- `_replace_para_text_clone(para, text)`: 단락 run 재구성
- `_set_reading_text(tf, units, line_spacing)`: 절 번호 오렌지 run + 본문 run 생성
- `_para_append_run(p, new_r)` / `_update_prefix_in_runs(para, old, new)` / `_clear_text_frame(tf)`
- `_has_ending_text(slide)` / `_find_content_shape(slide)`

> **OOXML 자식 요소 순서**: `<a:pPr>`에 요소를 추가할 때 `append()`를 쓰면 안 된다(스키마 순서
> 위반으로 PowerPoint가 파일을 손상됐다고 표시함). 반드시 `insert(0, ...)` 또는 스키마 순서에
> 맞는 위치에 삽입한다. 상세 규칙과 발견 경위는 `CLAUDE.md` 참고.

## 12. 배경 검정 설정 — `_set_slide_bg_black(slide, prs=None)` (5116행)

1. `p:cSld` 아래에 `p:bg`(solidFill 검정) 삽입
2. `spTree` 최하단(첫 번째 도형 위치)에 슬라이드 전체 크기의 검정 사각형(이름: `BlackBg`) 삽입
   — 레이아웃 상속 도형이 위에서 배경을 가리는 경우를 방지

**같은 프레젠테이션 내부 템플릿을 복제**(`insert_slide_copy`)한 새 슬라이드에는 걸어도 안전하다.
**다른 프레젠테이션에서 복사한 슬라이드**(`copy_slide_from_prs`)에는 걸면 안 된다 — 그 함수가
이미 원본 배경을 정확히 복사했으므로 덮어쓰는 것은 서식 파괴다(§6, §11 참고).

## 13. PPT 2007 호환성 처리 — `strip_ppt2007_incompatible(pptx_path)` (5827행)

저장 직후 호출. 제거 대상:
- `ppt/changesInfos/`(PowerPoint 2016 공동 작성 추적 파일) 및 관련 관계/Content_Types 항목
- `<a14:imgProps>` 블록(고화질 이미지 레이어, PPT 2007 미지원) 및 제거 후 남는 빈 `<a:extLst>`
- 슬라이드 내 외부 비디오 링크
- 슬라이드 rels에서 고아 image/hdphoto Relationship (shape는 삭제됐으나 rel만 남은 경우)
- `docProps/app.xml`의 슬라이드/MMClips 카운트를 실제 값으로 수정

## 14. 검증

- `validate_pptx_structure(pptx_path) -> list`: 저장된 OOXML 자체가 손상 없이 열리는지 확인
- `validate(prs, json_data, is_sunday=True) -> bool` (5665행):
  - liturgy 텍스트 존재 확인 → 없으면 오류
  - 입당송·영성체송 슬라이드 존재 확인(주일은 화답송도) → 없으면 경고
  - 제1독서가 있으면 절 번호 오렌지색(`ORANGE`) 사용 여부 샘플 확인 → 없으면 경고
  - 오류가 있으면 `False`(실행 실패로 표시), 경고만 있으면 `True`(로그만 남기고 계속 진행)

## 15. 회귀 테스트 스위트 (`test_missa_regression.py`)

목적: 주일·평일 처리 로직의 회귀 기준선. 향후 토요일미사 등 신규 기능 추가 시 기존 로직이
깨지지 않았는지 확인한다.

```
pip install pytest
pytest test_missa_regression.py -v
```

날짜 폴더에 `~$*.pptx` 잠금 파일이 있으면(PowerPoint에서 열려 있음) 해당 케이스는 자동 skip된다.

**테스트 데이터**: 주일 `20260712`(연중 제15주일), 평일 `20260624`(성 요한 세례자 탄생 대축일,
수요일). 두 폴더 모두 참조 PPT·JSON·성가 파일이 미리 준비되어 있어야 한다.

**단위 테스트**
| 클래스 | 검증 대상 |
|---|---|
| `TestIsSundayMass` | `is_sunday_mass()` 요일 판단. 토요일이 현재 "평일미사"로 처리됨을 명시적으로 고정 |
| `TestWrapLineCount` | `_wrap_line_count()` 줄바꿈 계산 |
| `TestParseIntoVerseUnits` | `parse_into_verse_units()` 절 번호 파싱 |

**통합 테스트** (`generated_case` fixture, `MASS_CASES` 파라미터화 — 실제 CLI 실행으로 PPT 생성 후 검증)
| 테스트 | 검증 내용 | 관련 절 |
|---|---|---|
| `test_output_opens_without_corruption` | 출력 파일이 손상 없이 열리는지 | — |
| `test_liturgy_type_detected_correctly` | 로그에 미사 유형이 올바르게 찍히는지 | — |
| `test_required_sections_present` | 필수 섹션이 모두 있는지 | §4 |
| `test_no_reading_slide_line_overflow` | 독서·복음 본문 슬라이드가 9줄을 넘지 않는지 | §5.4, §5.5 |
| `test_blackbg_covers_full_slide` | `BlackBg` 도형이 항상 슬라이드 전체를 덮는지 | §5.6 |
| `test_화답송_pattern_matches_mass_type` | 화답송이 주일/평일 패턴에 맞는지 | §6 |
| `test_hymn_score_copy_matches_mass_type` | 성가가 주일/평일 패턴에 맞는지 | §9 |

**신규 케이스 추가 방법** (예: 토요일미사): `MASS_CASES`에 케이스 추가 → 토요일 전용 판단
로직이 생기면 `TestIsSundayMass`에 대응 테스트 추가(`test_saturday_currently_treated_as_weekday`는
새 기대값으로 갱신) → 필요한 통합 검증 함수 추가 → 전체 재실행해 기존 케이스가 여전히
통과하는지 확인 후 병합.

## 16. 전체 함수 목록 (2026-07-16 기준, `missa_to_ppt.py`)

**설정 관리**: `_load_config` `_save_config` `_ask_onedrive_path_popup` `get_onedrive_hymn_folder`

**팝업 UI**: `_ask_date_popup` `_ask_input_files_popup` `_ask_combined_input_popup`
`_ask_numbers_popup` `_report_progress` `_run_with_progress_window` `_show_result_window`

**파일/인수**: `parse_args` `_infer_hymn_numbers` `find_files` `get_json_data`

**미사 유형**: `is_sunday_mass`

**슬라이드 복사 유틸**: `delete_slide` `move_slide` `_blank_layout` `duplicate_slide`
`insert_slide_copy` `copy_slide_from_prs` `_copy_spTree` `_copy_image_rels`
`_update_rId_in_spTree` `_effective_bg` `_copy_bg_image_rels` `_set_slide_bg_black`

**섹션 탐색**: `find_sections` `find_content_range` `find_복음_content_range`
`find_slide_with_text` `find_shape_exact_text` `_is_hymn_divider` `_slide_text` `all_slide_texts`

**텍스트 서식 유틸**: `_para_append_run` `_replace_para_text_clone` `_set_두_줄_text`
`_set_화답송_content_text` `_set_single_para_text` `_josa` `_update_book_name_after_br`
`_ends_sentence` `_update_prefix_in_runs`

**독서·복음 파싱/배분/기록**: `parse_into_verse_units` `_visual_lines` `_wrap_line_count`
`_page_visual_lines` `layout_units_on_slides` `_verify_and_rebalance_pages`
`_find_content_shape` `_has_ending_text` `_clear_text_frame` `_set_reading_text`
`_count_slide_lines` `_rendered_wrap_count` `_get_slide_render_params`
`_count_slide_lines_rendered` `_split_para_at_lines` `_rebalance_reading_slides_post_write`
`replace_reading_slides`

**종료 슬라이드**: `_align_ending_slides_to_제2독서` `_reposition_merged_ending_shapes`

**섹션별 업데이트**: `update_title_slide` `update_입당송` `update_reading_title_slide`
`update_복음_title_slide` `_update_화답송_title_in_slide` `update_화답송` `update_복음환호송`
`update_영성체송`

**본문 자동 맞춤**: `_find_last_row_top` `_shape_first_run_font_size_emu`
`_shape_first_para_line_spacing_pct` `_set_shape_all_para_line_spacing`
`_set_shape_all_run_font_size` `_estimate_text_lines` `_adjust_fit_if_needed`

**시작기도·성가·미사후기도**: `replace_시작기도문` `replace_미사후기도`
`_update_성가_divider_number` `_update_성가_header` `replace_성가`

**검증·저장**: `validate_pptx_structure` `validate` `_strip_slide_xml`
`strip_ppt2007_incompatible`

**진입점**: `main`

## 17. 의존성

- python-pptx, lxml, Pillow(폰트 메트릭 측정)
- requests, beautifulsoup4 (missa_to_json.py의 크롤링)
- tkinter (표준 라이브러리, GUI)
- pytest (회귀 테스트 실행용, 프로덕션 실행에는 불필요)
- json, pathlib, re, copy, io, subprocess, zipfile (표준 라이브러리)

---

## 부록 A: 버전별 변경 이력

| 버전 | 주요 변경 | 원문 |
|---|---|---|
| v1.0 | 최초 구현. 주일미사 전용 CLI. 독서/복음 파싱·배분, 슬라이드 복사 유틸, 서식 보존 함수, PPT2007 호환성 처리의 기본 골격 확립 | `docs/archive/missa_to_ppt 구현 계획 v1.0.md` |
| v1.1 | 대화형 팝업 모드, `config.json` 기반 OneDrive 성가 폴더 관리, 종료 슬라이드 통합(`merge_threshold`)과 `_reposition_merged_ending_shapes()` 동적 위치 재조정 추가 | `docs/archive/missa_to_ppt 구현 계획 v1.1.md` |
| v1.2 | 평일미사 지원 추가 — `is_sunday_mass()` 날짜 기반 자동 판단, UI 재배열, 화답송/성가/미사 후 기도를 주일·평일로 분기하는 `is_sunday`/`copy_scores` 파라미터화 | `docs/archive/missa_to_ppt 구현 계획 v1.2 (평일미사 추가).md` |
| v1.3 | v1.2까지 정의됐던 동작(9줄 제한, 화답송 서식 보존, 배경 전체 커버)이 실제로는 지켜지지 않던 버그 6건 수정(§5.3, §5.5, §5.6, §6, §5.4), 회귀 테스트 스위트(`test_missa_regression.py`) 신설. 이 버전부터 구현계획 문서를 "diff" 대신 "현재 전체 구현"을 담은 완결 문서로 관리 | (이 문서) |

## 부록 B: 재발 방지 규칙

OOXML 조작 시 반복적으로 발생했던 함정과 규칙(`pPr` 자식 요소 순서, 슬라이드 복사 후 배경
재설정 금지, 빈 텍스트 도형의 falsy 분류 함정, post-write 단계 간 stale 값 참조 주의, TTC
폰트 인덱스 확인, Pillow 커닝 미지원 보정)은 `CLAUDE.md`에 정리되어 있다. 코드를 수정하기 전에
먼저 읽는다.
</content>
