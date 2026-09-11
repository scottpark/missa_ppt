# CLAUDE.md — missa_ppt 프로젝트

## 프로젝트 구조

매일미사 웹데이터로 미사 PPT를 자동 생성하는 도구.

**파이프라인:** `missa_to_json.py`(missa.cbck.or.kr 크롤링 → `missa_YYYYMMDD.json`)
→ `missa_to_ppt.py`(JSON + 성가번호로 참조/템플릿 PPT를 수정해 결과 PPT 생성).

**핵심 파일:**
- `missa_to_json.py` — 독서·복음·화답송·성가 등 섹션을 크롤링해 JSON 출력.
- `missa_to_ppt.py` — 진입점. CLI 인자 파싱(`parse_args`)·파일 탐색(`find_files`)·JSON 로드
  (`get_json_data`)·`main()`/`__main__` 흐름만 담당. 실제 처리 로직은 아래 5개 모듈에 위치하며
  `missa_to_ppt.py`는 이들을 조합해 호출한다(2026-09 모듈 분리 리팩토링, 순수 이동·동작 무변경).
- `missa_ooxml_utils.py` — 슬라이드/도형 복사·rId 매핑·배경 상속·텍스트런 조작 등 OOXML 저수준
  프리미티브. 프로젝트 내부 의존성이 없는 leaf 모듈. `HYMN_TYPES`도 여기 위치(entry↔content_updaters
  순환 임포트 방지).
- `missa_gui.py` — Tkinter 팝업 전부 + `is_sunday_mass`/config.json 로드·저장/
  `get_onedrive_hymn_folder`. 이 함수들은 GUI 팝업과 진입점 양쪽에서 호출되어 entry에 그대로
  두면 순환 임포트가 생기므로 gui 모듈에 흡수했다(entry→gui 단방향). `OUTPUT_ROOT`도 entry·gui
  양쪽에서 쓰여 여기 위치.
- `missa_reading_layout.py` — 독서/복음 절 파싱(`parse_into_verse_units`)부터 슬라이드 분배
  (`layout_units_on_slides`), Pillow 추정 + PowerPoint COM 실측 줄 수 계산, 기록 후 재조정까지
  독서/복음 처리 파이프라인 전체.
- `missa_sections.py` — 섹션 탐색(`find_sections`)과 검증(`validate`,
  `validate_pptx_structure`, `strip_ppt2007_incompatible`).
- `missa_content_updaters.py` — 화답송·성가·입당송·복음환호송·영성체송·시작기도문·미사후기도 등
  섹션별 콘텐츠 갱신(`update_*`/`replace_*` 함수들).
- `missa_psalm_score_image.py` — 화답송 악보 원본 이미지(PNG/JPG) → 슬라이드 변환. 순수 함수
  묶음, `missa_to_ppt.py`를 import하지 않는 단방향 의존. 진입점
  `render_화답송_score_slide()`. 고정 템플릿 자산은 `assets/화답송_악보_template.pptx`
  (`tools/build_화답송_template.py`로 생성). `missa_to_ppt.py`(`find_files()`/CLI `--화답송`
  오버라이드/입력창/`update_화답송()`)에 배선 완료 — 수작업 PPT가 없을 때 자동 폴백된다.
- `missa_youth_gospel.py` — (청년미사 1단계) universalis.com에서 날짜별 **미사 전 섹션**(First
  reading·Responsorial Psalm·Second reading·Gospel Acclamation·Gospel)을 조회해 JSON으로 반환
  (2026-09-11 v2 확장 — 최초 버전은 Gospel만 조회했다). `fetch_gospel_html`(리다이렉트 차단 —
  3xx면 `GospelFetchError`, 과거 날짜가 조용히 오늘자로 리다이렉트되는 함정 차단), `parse_mass`
  (stdlib `html.parser`, 섹션 마커=`align="left"` th/참조=그 뒤 첫 `align="right"` th, 영어 키만
  사용·한글 키 없음, Psalm은 reference만, Gospel Acclamation의 "Or:" 대체 환호송은 `"or"`
  중첩), `get_youth_mass`(기본 출력 파일명 `missa_en_YYYYMMDD.json`, 한글 미사
  `missa_YYYYMMDD.json`과 구분). `parse_gospel`/`get_youth_gospel`(Gospel 섹션만, 한글 키
  `"복음"`)은 하위호환 래퍼로 유지. **완전 독립 leaf** — `missa_to_ppt.py`/`missa_to_json.py`를
  import하지 않고, 본 파이프라인도 아직 이 모듈을 호출하지 않는다(통합은 후속 단계).
- `missa_youth_hymn_pdf.py` — (청년미사 1단계) PDF 성가집(나주노=스캔 이미지 / 야훼 이레=텍스트
  레이어) → 성가 슬라이드 PPT 생성. 번호 조회 → 이미지 오선 밴드 탐지(`detect_system_bands`,
  콘텐츠폭 기준 row_fill>0.40) → `_group_content_bbox`(그룹 **행 범위 안의 잉크 열**만으로 좌우
  경계를 잡는 per-group 크롭 — 전역 `detect_content_bounds`를 쓰면 그 그룹과 무관한 다른 줄
  (저작권/제목 등)이 폭을 오염시켜 크롭 우측에 불필요한 흰 여백이 생긴다. 2026-09-11 나주노 447
  우측 여백 버그의 원인이자 수정 지점 — **여러 시스템/그룹을 한 이미지에서 크롭하는 코드는 항상
  전역이 아니라 그 그룹의 실제 행 범위로 좌우 경계를 잡아야 한다**는 일반 원칙으로 확장 가능) →
  `pack_systems`(동적 min-slide 균형 패킹: 같은 배율로 최대 3개 시스템까지 겹침 없이 들어가고
  그것이 실제로 슬라이드 수를 줄이거나 균형을 개선할 때만 3개씩 묶는다 — 순수 "들어가면 무조건
  3개" greedy는 오히려 불균형한 3+1 분할을 만들 수 있어 채택하지 않음; 참조용 고정 2개/장
  `group_systems`는 남아 있음) → 저작권 크롭(`resolve_copyright_crop`, 트리거
  "Administered by"/"Adm. by" 유무와 무관하게 꼬리 마침표를 잉크런 분석으로 제거
  `_trim_trailing_period_px`, 배치 높이는 288032×0.9=259229 EMU로 전곡 고정·폭은 곡별 크롭
  종횡비 유지) → 헤더 run 조립(`build_header_runs`, 출처 표시명만 `SOURCE_DISPLAY`로 치환 —
  "야훼 이레"→"야훼이레", `SOURCES` 키/CLI/조회 함수는 원본 유지; 제목이 헤더 폭을 넘치면
  `_autosize_header`가 자간 spc→0 → 그래도 넘치면 제목 run만 폰트 축소, 궁서 `batang.ttc`
  index 2 Pillow 측정) → `duplicate_slide` 복제 조립(`build_hymn_pptx`). 헤더 run 삽입은
  `missa_psalm_score_image._insert_run_before_end_para_rpr`(endParaRPr 앞 삽입) 재사용.
  고정 템플릿 자산 `assets/청년미사_성가_template.pptx`(`tools/build_청년미사_template.py`로 생성),
  나주노 저작권 크롭 좌표는 `assets/나주노_copyright_bbox_cache.json`(개발자 PC에서 Tesseract OCR
  **1회** 생성 — 재생성 툴 `tools/build_나주노_copyright_cache.py`, 개발 전용; **런타임·exe는 캐시만
  읽고 Tesseract를 호출하지 않는다**; 도구가 만드는 값은 초안일 뿐이며 447 항목처럼 손제작 샘플과
  대조해 사람이 손으로 확정한 값이 있을 수 있다 — 도구를 재실행하면 그 hand-tuned 값이 도구
  기본값으로 덮인다). 이 모듈도 **완전 독립 leaf**로 본 파이프라인 미접촉이며,
  테스트는 `test_missa_progression.py`에만 있고(회귀 테스트 승격은 미실시 — 정식 회귀 아님).
- `test_missa_progression.py` — 아직 안정화되지 않은 신규 동작을 먼저 명세하는 프로그레션
  테스트(TDD red→green). 안정화되면 `test_missa_regression.py`로 승격.
- `test_missa_regression.py` — 주일/평일 통합 + 단위 회귀 테스트 (`pytest`로 실행).
- `config.json` — `onedrive_hymn_folder`(악보 성가 PPT 경로).
- `missa_to_ppt.spec` + `dist/` — PyInstaller 빌드(`missa_to_ppt.exe`, GUI 모드).
- `docs/` — 요구사항·구현 계획(버전 번호 없이 항상 최신 상태 유지, 변경 이력은 문서 맨 끝
  부록 참고), `archive/`(v1.0~v1.3 과거 버전 원문), `ai-readiness-check/`,
  `ooxml-pitfalls-log.md`(아래 OOXML 함정 규칙들의 "발견 경위" 전문 — CLAUDE.md에는 규칙
  본문과 한 줄 요약만 남기고 상세 사고 경위는 여기로 분리했다, 2026-09-13).

**날짜 폴더(`output/YYYYMMDD/`):** 미사별 작업 디렉터리. `OUTPUT_ROOT`(`missa_gui.py`) 상수가
가리키는 `output/` 폴더 밑에 날짜별로 위치한다. JSON, 템플릿·결과 pptx, 화답송 악보 pptx,
`log/`를 포함. 입력/출력 산출물이라 매번 커밋하지 않으며, 회귀 테스트 픽스처로 쓰이는 3개
(`20260624` 평일·`20260705` 성수축복·`20260712` 주일)만 git으로 추적한다(`.gitignore` 참고).
그 외 날짜 폴더는 로컬에만 남기고 커밋·푸시하지 않는다.

## OOXML XML 조작 규칙 (필수 적용)

### 자식 요소 순서 엄수
OOXML은 각 요소의 자식 순서를 XML Schema sequence로 강제한다.
순서가 틀리면 PowerPoint가 "손상된 파일" 오류를 표시한다.

**`<a:pPr>` (CT_TextParagraphProperties) 자식 순서:**
```
lnSpc → spcBef → spcAft → buClr → buSz → buFont → bu* → tabLst → defRPr → extLst
```

### 절대 금지: `pPr.append(element)`
`append()`는 요소를 맨 끝에 추가하므로 `lnSpc`가 `spcAft` 뒤에 오게 된다.
반드시 `insert(0, element)` 또는 스키마 순서에 맞는 위치에 `insert`를 사용한다.

```python
# ❌ 틀린 방법
pPr.append(pptx_parse_xml('<a:lnSpc .../>'))

# ✅ 올바른 방법 (lnSpc는 pPr의 첫 번째 자식)
pPr.insert(0, pptx_parse_xml('<a:lnSpc .../>'))
```

### 기존 요소 교체 시
`remove()` 후 `insert(0, new_element)`를 사용한다. `append()`로 재추가하지 않는다.

**발견 경위 (2026-07-14):** `_set_reading_text`의 `lnSpc.append()`로 재현(상세:
`docs/ooxml-pitfalls-log.md`).

### 이 원칙은 `<a:pPr>`뿐 아니라 순서 스키마를 가진 모든 요소에 적용된다

`append()` 금지 원칙을 "pPr에 요소를 추가할 때"로만 기억하면, 순서가 강제되는 **다른** 요소를
건드릴 때 같은 함정에 다시 걸린다. 예: `<a:p>` (CT_TextParagraph)의 자식 순서도
`pPr? → (run|br)* → endParaRPr?`로 강제된다. 단락에서 기존 run/br만 지우고 `endParaRPr`는
남긴 채 새 run을 `p.append(new_r)`로 추가하면, 실제 순서가 `pPr → endParaRPr → r`가 되어
run이 endParaRPr **뒤에** 오는 스키마 위반이 된다.

```python
# ❌ 틀린 방법 — endParaRPr가 이미 남아있는 <a:p>에 append
p.append(new_r)   # 결과: pPr, endParaRPr, r (스키마 위반)

# ✅ 올바른 방법 — endParaRPr가 있으면 그 앞에 삽입
eprp = p.find(qn('a:endParaRPr'))
if eprp is not None:
    eprp.addprevious(new_r)
else:
    p.append(new_r)  # endParaRPr가 없으면 맨 끝이 곧 올바른 위치
```

새로운 요소 타입을 다룰 때는 "pPr 자식 순서"라는 암기된 규칙을 그대로 대입하지 말고, 해당
요소의 OOXML 스키마 시퀀스를 확인한 뒤 `append()`가 마지막 자식(시퀀스상 옵션인 꼬리 요소,
예: `endParaRPr`, `extLst`) 뒤에 요소를 놓지는 않는지 점검한다.

**발견 경위 (2026-08-29):** `_replace_title_paragraph`의 `p.append(new_r)`로 재현, 독립
코드 리뷰가 실제 렌더링 XML을 열어보고서야 발견(상세: `docs/ooxml-pitfalls-log.md`).

## 슬라이드 복사 시 배경/서식 재설정 금지

`copy_slide_from_prs()`는 원본 슬라이드의 배경(레이아웃/마스터 상속분까지 해석한 실제 배경)을
이미 정확히 복사한다. 복사 직후 `_set_slide_bg_black()`을 또 호출하면 방금 복사한 원본 서식을
하드코딩된 검정으로 덮어써 버린다. 다른 프레젠테이션에서 슬라이드를 복사해 온 경우
(`copy_slide_from_prs`)에는 배경을 다시 강제하지 않는다. 같은 프레젠테이션 내부 템플릿을
복제한 경우(`insert_slide_copy`/`duplicate_slide`)는 원본이 이미 이 문서의 스타일을 따르므로
`_set_slide_bg_black()`을 걸어도 안전하다.

**발견 경위 (2026-07-16):** `update_화답송()`이 `copy_slide_from_prs()` 직후
`_set_slide_bg_black()`을 호출해 원본 서식이 사라짐(상세: `docs/ooxml-pitfalls-log.md`).

## 슬라이드 복제 시 `<p:sld>`의 `showMasterSp` 속성도 원본과 맞춘다

슬라이드를 복제할 때(`duplicate_slide`/`insert_slide_copy`)는 `p:cSld` 안의 `p:bg`(배경)뿐
아니라 최상위 `<p:sld>` 요소 자체의 `showMasterSp` 속성도 원본과 동일하게 맞춰야 한다.
`showMasterSp="0"`은 슬라이드 마스터에 배치된 요소(배경 장식·로고 등)를 이 슬라이드에서 숨기라는
뜻으로, `p:bg`와는 **별개의** 속성이다 — 두 슬라이드가 같은 layout/master를 쓰고 `p:bg`가
동일해도, 한쪽에만 `showMasterSp="0"`이 있으면 마스터 상속 요소의 노출 여부가 달라져 색이 다르게
보인다. `p:bg` 복사만으로는 이 차이가 커버되지 않는다.

`python-pptx`의 `prs.slides.add_slide()`가 만드는 새 슬라이드의 `<p:sld>`에는 이 속성이 아예
없다(속성 부재 = 기본값 "마스터 요소 노출"). 따라서 복제 시 원본의 값을 명시적으로 옮겨야 한다.
`showMasterSp`는 `p:sld`의 네임스페이스 프리픽스 없는 unqualified attribute이므로 `qn()` 없이
`src_slide.element.get('showMasterSp')` / `new_slide.element.set('showMasterSp', ...)`로 처리한다
(자식 요소 조작이 아니라 요소 자체의 속성이라 순서 스키마 함정과는 무관). 원본에 속성이 없으면
새 슬라이드에도 넣지 않는다(불필요한 속성 주입 금지 — `add_slide` 기본 상태를 그대로 둔다).

**향후 빈 슬라이드를 삽입할 때도 이 원칙이 그대로 적용된다** — 새 blank 슬라이드를 어떤 원본에서
복제하든, 원본의 `showMasterSp` 상태를 따라가야 마스터 상속 요소의 노출이 원본과 일치한다.

**발견 경위 (2026-09-13):** `duplicate_slide()`가 `p:bg`는 복사하면서 `showMasterSp`는
빠뜨려, 실사용자 육안 검수에서 구분 슬라이드 색이 미세하게 다르게 보임(상세:
`docs/ooxml-pitfalls-log.md`).

## 도형 분류 시 "빈 텍스트"가 falsy임을 주의

빈 문자열(`""`)은 파이썬에서 falsy다. 도형을 텍스트 키워드로 분류하는 코드(`'ending' if kw in txt
else 'content'` 같은 패턴)에서 `if key == 'content' and txt:` 식으로 "본문이 있는 content만
스킵" 조건을 걸면, 텍스트가 비어 있는 배경 도형(`BlackBg` 등)도 `content`로 분류된 채 스킵되지
않고 그대로 처리 대상에 들어간다. 배경 도형처럼 텍스트 분류 로직에 절대 걸리면 안 되는 도형은
`shape.name == 'BlackBg'`처럼 이름으로 명시적으로 먼저 걸러낸다.

**발견 경위 (2026-07-16):** `_align_ending_slides_to_제2독서()`에서 `BlackBg`가 `content`로
오분류돼 잘못 리사이즈됨(상세: `docs/ooxml-pitfalls-log.md`).

## 텍스트 키워드로 도형을 식별할 때 부분 문자열 충돌 주의

`kw in shape.text_frame.text` 같은 부분 문자열 포함검사로 특정 도형(제목·라벨 등)을 찾을 때,
그 키워드가 **다른** 도형의 텍스트에도 우연히 포함되어 있으면 도형 순회 순서에 의존하는
암묵적 안전성만 남는다. 첫 매칭에서 `return`하는 코드는 의도한 도형이 항상 먼저 순회되는 동안만
안전하고, 도형 순서가 바뀌거나(슬라이드 복사·재배치) 텍스트가 수정되면 조용히 엉뚱한 도형이
매칭되어 그 내용을 덮어쓴다. 키워드는 "의도한 도형에만 나타나는" 형태(공백·구두점 포함 등
더 구체적인 패턴)로 좁혀서, 순회 순서와 무관하게 안전하도록 만든다.

**발견 경위 (2026-08-29):** `_update_title`의 "화답송"(공백 없음) 키워드가 저작권 표기 도형의
문구에도 우연히 매칭됨, 독립 리뷰가 지적(상세: `docs/ooxml-pitfalls-log.md`).

## post-write 재조정 이후 값을 참조할 때는 최신 상태를 다시 측정

여러 슬라이드 조정 단계가 순차 실행되는 파이프라인(예: `replace_reading_slides()` →
`_rebalance_reading_slides_post_write()` → 종료 슬라이드 병합 판단)에서, 뒤 단계가 앞 단계의
**계획값**(`units_pages` 등, 실제 반영 전 값)을 참조하면 안 된다. 앞 단계가 슬라이드를
추가/재배치할 수 있으므로, 실제 물리 슬라이드에서 다시 측정(`_count_slide_lines()` 등)해야 한다.

**발견 경위 (2026-07-16):** 종료 텍스트 병합 판단이 재조정 이전 계획값(`units_pages[-1]`)을
써서 실제 마지막 슬라이드와 어긋남(상세: `docs/ooxml-pitfalls-log.md`).

## 도형 위치 계산: 1줄 높이를 "박스 height ÷ 고정 줄 수"로 역산하지 않는다

텍스트박스의 1줄 높이가 필요할 때 `box.height // LINES_PER_SLIDE`(고정 줄 수로 나눔)로
역산하면 안 된다. 본문 박스 height는 슬라이드마다 다르고 "항상 그 줄 수만큼 꽉 차 있다"는
보장이 없다 — 줄 수가 적은(짧은) 슬라이드의 박스는 여전히 클 수 있어, 나눗셈이 실제보다 작은
1줄 높이를 만들어낸다. 1줄 높이는 실제 폰트 메트릭으로 직접 계산한다: `_get_slide_render_params()`
가 준 Pillow 폰트의 `getmetrics()`(ascent+descent, 96dpi px)를 EMU로 환산(1px=9525EMU)하고
lnSpc(spcPct) 배율을 곱한다(`_content_line_height_emu()`). Pillow/폰트 미존재 시에만 기존
나눗셈으로 폴백한다.

**주의:** Pillow의 ascent+descent는 PowerPoint 실제 줄 높이보다 약 14~18% 작다. 겹침 방지
여백(GAP=2줄)이 이 과소추정을 흡수하며, 이 조합은 본문 12줄까지 안전하다(병합 대상은 5줄
이하라 마진 방대). GAP을 줄이거나 line_count>12로 벗어나는 변경 시 재검토 필요 — 보정치 유도
전체 수식은 `docs/ooxml-pitfalls-log.md` 참고.

**발견 경위 (2026-09-02):** `_reposition_merged_ending_shapes()`의
`line_height = content_shape.height // 9` 나눗셈으로 재현, 실사용자 육안 검수로 발견(상세:
`docs/ooxml-pitfalls-log.md`).

## 한글 줄 수 계산: TTC 폰트 인덱스와 Pillow 커닝 한계

`_get_slide_render_params()`가 Pillow로 폰트 폭을 측정해 실제 PowerPoint 렌더링 줄 수를
추정한다. 여기서 두 가지를 주의한다.

1. **TTC(트루타입 컬렉션) 인덱스는 실제 파일에서 직접 확인한다.** `batang.ttc`/`gulim.ttc`처럼
   한 파일에 여러 서체가 들어 있는 경우, 인덱스를 추측하지 말고
   `ImageFont.truetype(path, size, index=N).getname()`으로 실제 순서를 확인한다.
2. **이 환경의 Pillow는 `libraqm`(커닝) 미지원이라 텍스트 폭을 실제보다 넓게 계산한다.**
   `_RENDER_WIDTH_CALIBRATION`(현재 1.03)으로 보정한다. 폰트나 크기 조합이 바뀌면
   `test_missa_regression.py`의 `test_no_reading_slide_line_overflow`로 재검증하고,
   필요하면 실측 기반으로 재보정한다.

**발견 경위:** TTC 인덱스를 추측(반대로 가정)했다가 줄 수 계산이 틀림, Pillow 폭 과대평가를
실측으로 확인해 보정치 도입(상세: `docs/ooxml-pitfalls-log.md`).

## 단락 분리·병합 함수는 run 개수를 2개로 가정하지 않는다

절 하나 = run 2개(오렌지 절 번호 + 흰색 본문)라고 가정하는 코드는 여러 절이 하나의 논리
단락으로 병합된 경우(`layout_units_on_slides()`의 `is_continuation` 병합) 깨진다. 이때 단락은
[오렌지][흰색][오렌지][흰색]... 처럼 run이 3개 이상이 된다. 단락을 두 조각으로 자르는 함수는
반드시 "분리 지점이 몇 번째 run의 몇 번째 글자인지"를 실제로 계산해서, 그 지점 앞뒤 run들의
서식(rPr)을 각각 그대로 유지해야 한다. `runs[0]`/`runs[1]`처럼 인덱스를 고정하거나 "마지막
run만 남기고 나머지 삭제" 같은 방식은 텍스트는 보존해도 중간 run의 색상(서식)을 조용히
잃어버린다 — 예외 없이 실행되므로 육안 검수 전까지 발견되지 않는다.

**발견 경위 (2026-07-26):** `_split_para_at_lines()`의 인덱스 고정 삭제 방식이 51/52절 병합
단락에서 52절 오렌지 run의 색을 지움(상세: `docs/ooxml-pitfalls-log.md`).
`_missing_orange_verse_numbers()`(§검증) 추가로 이런 사례를 자동으로 잡는다.

## run별 텍스트를 재작성할 때 옛 run 길이로 통짜 재배치하지 말고 의미 단위 경계로 나눈다

여러 run에 걸친 앞부분 텍스트를 "run 서식은 보존하고 텍스트만 교체"하려고, 새 문자열을 옛 run의
글자 수만큼씩 순서대로 잘라 넣는 방식은 **길이가 바뀌면** 조용히 색을 오염시킨다. 새 텍스트가
옛 텍스트보다 길거나 짧으면 문자 위치가 통째로 밀리고, 원래 의미가 다른 구간(라벨 vs 구분자 vs
숫자)의 run 경계가 새 텍스트에서는 전혀 다른 위치로 이동한다. 그 결과 한 구간의 마지막 글자가
옆 구간의 run(우연히 다른 색일 수 있음)으로 밀려 그 색을 물려받는다. 이 코드는 예외를 던지지
않으므로 육안 검수 전까지 발견되지 않는다.

**규칙:** 여러 의미 구간(라벨·구분자·숫자처럼 색/역할이 다른 구간)에 걸친 텍스트를 재작성할
때는, 호출부가 이미 알고 있는 의미 경계(정규식 그룹 등)를 **재배치의 하드 경계**로 삼아 각
구간을 독립적으로 재배치한다. 한 구간의 길이 변화(증가/감소)는 그 구간이 점유한 run 안에서만
흡수되어야 하고, 다른 구간이 점유하던 run에 절대 닿으면 안 된다. `_update_prefix_in_runs()`는
`segments=[(old_text, new_text), ...]`를 받아 세그먼트별로 독립 재배치한다.

**발견 경위 (2026-09-02):** `_update_성가_header()`가 라벨 `2차봉헌`→`2차 봉헌`(1자 증가)
정규화 시 통짜 재배치를 해서, 라벨 마지막 글자가 옆 구분자 run의 회색을 물려받음(상세:
`docs/ooxml-pitfalls-log.md`). "단락 분리·병합 함수는 run 개수를 2개로 가정하지 않는다"와
같은 계열의 함정이다 — "run 서식을 보존한다"는 목적은 같아도 "어느 run이 어느 새 글자를
받는가"의 배정 규칙이 부정확하면 색이 조용히 샌다.

## 검증: 텍스트 존재가 아니라 절 번호별 오렌지색 렌더링을 확인

`validate()`는 "오렌지색 run이 프레젠테이션 어딘가에 하나라도 있는지"가 아니라, JSON
`content`를 `parse_into_verse_units()`로 파싱해 나온 절 번호 전체가 해당 독서/복음 슬라이드
범위 안에서 실제로 오렌지색 run으로 렌더링됐는지 절 번호 단위로 대조한다
(`_missing_orange_verse_numbers()`). 텍스트만 남고 색이 사라지는 위 버그처럼, "텍스트 존재
여부"만 보는 검증은 이런 회귀를 통과시킨다.

## 이미지에서 여러 그룹을 크롭할 때는 전역 경계가 아니라 그룹별 경계를 쓴다

한 이미지(페이지 스캔본 등) 안에 여러 그룹(예: 악보 시스템 묶음)을 순서대로 잘라낼 때,
`detect_content_bounds()`류의 **전역** 좌우 경계를 모든 그룹에 공통으로 쓰면 안 된다. 그 이미지
안에 그룹과 무관한 다른 줄(제목·저작권 표기 등)이 그룹보다 더 넓게 뻗어 있으면, 전역 경계가 그
넓은 줄에 맞춰지고 모든 그룹의 크롭이 실제 내용보다 불필요하게 넓어져(그 그룹 우측에 흰 여백이
남아) 크롭이 "타이트하지 않게" 된다. 그룹별로 잘라야 하는 코드는 반드시 **그 그룹 자신의 행/열
범위 안에서만** 잉크(내용) 경계를 다시 계산해야 한다(예: `_group_content_bbox(gray, group)` —
그룹의 top..bottom 행 범위 안의 잉크 열 min/max로 left/right 산출).

**발견 경위 (2026-09-11):** `_content_crop_box`가 전역 `detect_content_bounds(gray)`로 447번
악보를 크롭하면서, 페이지 맨 아래 저작권 줄이 전역 우측 경계를 오염시켜 모든 시스템 묶음
크롭이 ~10% 더 넓게 잘림(상세: `docs/ooxml-pitfalls-log.md`). `_group_content_bbox`로 해결,
`pack_systems`의 fit 판정도 이 헬퍼를 공유해야 했다(전역 경계를 쓰면 종횡비 과소평가로 오판
가능).

## 회귀 테스트

`test_missa_regression.py`에 주일(20260712)·평일(20260624) 통합 테스트와 핵심 함수 단위
테스트가 있다. 독서·복음 줄 수 계산, 화답송/성가 처리 분기, 배경색, 절 번호 오렌지색처럼 이
문서에 기록된 버그와 직결된 항목을 검증하므로, 관련 로직을 수정하면
`pytest test_missa_regression.py -v`를 실행해 회귀 여부를 확인한다.

## 하네스: missa_ppt 개발

**목표:** 성인미사 전용으로 짜인 `missa_to_ppt.py`를 청년미사·어린이미사 등 새 미사 유형과
새 성당/본당으로 안전하게 확장한다. 확장 과정에서 이 문서에 기록된 OOXML 함정이 재발하지
않도록 설계·TDD 구현·독립 리뷰·검증을 분리된 전문 에이전트가 담당한다. 새 기능은 구현 전에
실패하는 프로그레션 테스트(`test_missa_progression.py`)로 먼저 명세하고, 안정화되면
회귀 테스트(`test_missa_regression.py`)로 승격한다.

**트리거:** 새 미사 유형/성당 지원 추가, PPT 템플릿 구조 확장, `missa_to_ppt.py`/
`missa_to_json.py`의 섹션 처리·OOXML 로직 구현 또는 버그 수정 요청 시 `mass-ppt-dev` 스킬을
사용하라. 단순 질문(코드 설명, 문서 조회)은 직접 응답 가능.

**구성:** `.claude/agents/`(mass-template-architect·ppt-ooxml-specialist·ooxml-code-reviewer·
regression-qa), `.claude/skills/`(mass-ppt-dev 오케스트레이터 + mass-template-analysis·
tdd-progression-testing·ooxml-pitfalls·boundary-verification·docs-sync).

**변경 이력:**
| 날짜 | 변경 내용 | 대상 | 사유 |
|------|----------|------|------|
| 2026-08-25 | 초기 구성 (에이전트 3명 + 오케스트레이터 1개 + 스킬 4개) | 전체 | 청년미사·어린이미사·타 성당 확장을 앞두고 설계/구현/검증 역할 분리 필요 |
| 2026-08-25 | TDD 도입(`tdd-progression-testing` 스킬 신설, `test_missa_progression.py`↔`test_missa_regression.py` 분리, regression-qa에 승격 절차 추가) + 독립 코드 리뷰 단계 추가(`ooxml-code-reviewer` 에이전트 신설, specialist→reviewer→qa 파이프라인으로 확장) | 에이전트 4명, 스킬 5개 | 프로그레션(신규 동작 명세)과 회귀(기존 동작 보호) 테스트의 목적이 다름을 명확히 분리하고, 구현자 본인이 못 보는 사각지대를 잡을 독립 리뷰 단계 필요 |
| 2026-08-30 | 하네스 첫 실제 기능(화답송 이미지 지원) 완료 후 harness Phase 6(with-skill vs without-skill 비교, `_workspace/04_phase6_validation_report.md`) 실시 → 결과 반영: (1) `ooxml-code-reviewer`에 "검증 우선순위 0"(스펙 자체가 원본 자료와 일치하는지 독립 재검증, 스펙에서 복사된 기대값 신뢰 금지) 추가, (2) `mass-template-architect`에 "실측 검증"의 정의를 통계적 타당성이 아니라 개별 판단 지점의 직접 확인으로 명시, (3) `ooxml-pitfalls`에 "완성 XML 통째 추가"가 append/순서 함정을 구조적으로 피하는 대안이라는 항목과 `validate_pptx_structure()`가 요소 순서를 검사하지 않는다는 주의사항 추가 | ooxml-code-reviewer.md, mass-template-architect.md, ooxml-pitfalls/SKILL.md | with-skill(독립 리뷰 4라운드)조차 1라운드에서 설계 스펙 자체의 오류(바라인 오검출)를 못 잡고 실사용자 육안 검수로 뒤늦게 발견됨 — "스펙과 일치하는가"만 보는 리뷰로는 부족하고 원본 자료 재검증이 필요함을 실측으로 확인. without-skill 베이스라인은 독립 검증 부재로 같은 종류의 결함(3세트 중 2세트가 실제로는 마디 경계 아님)을 "정확함"으로 자체 오판·보고 |
