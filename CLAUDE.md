# CLAUDE.md — missa_ppt 프로젝트

## 프로젝트 구조

매일미사 웹데이터로 미사 PPT를 자동 생성하는 도구.

**파이프라인:** `missa_to_json.py`(missa.cbck.or.kr 크롤링 → `missa_YYYYMMDD.json`)
→ `missa_to_ppt.py`(JSON + 성가번호로 참조/템플릿 PPT를 수정해 결과 PPT 생성).

**핵심 파일:**
- `missa_to_json.py` — 독서·복음·화답송·성가 등 섹션을 크롤링해 JSON 출력.
- `missa_to_ppt.py` — 메인 생성기. 슬라이드 복사·독서 줄 나눔·절 번호 색상·`validate()` 등.
- `test_missa_regression.py` — 주일/평일 통합 + 단위 회귀 테스트 (`pytest`로 실행).
- `config.json` — `onedrive_hymn_folder`(악보 성가 PPT 경로).
- `missa_to_ppt.spec` + `dist/` — PyInstaller 빌드(`missa_to_ppt.exe`, GUI 모드).
- `docs/` — 요구사항·구현 계획(버전 번호 없이 항상 최신 상태 유지, 변경 이력은 문서 맨 끝
  부록 참고), `archive/`(v1.0~v1.3 과거 버전 원문), `ai-readiness-check/`.

**날짜 폴더(`YYYYMMDD/`):** 미사별 작업 디렉터리. JSON, 템플릿·결과 pptx,
화답송 악보 pptx, `log/`를 포함. 입력/출력 산출물이라 매번 커밋하지 않으며, 회귀 테스트
픽스처로 쓰이는 3개(`20260624` 평일·`20260705` 성수축복·`20260712` 주일)만 git으로 추적한다
(`.gitignore` 참고). 그 외 날짜 폴더는 로컬에만 남기고 커밋·푸시하지 않는다.

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

**이 규칙의 발견 경위 (2026-07-14):**
`_set_reading_text`에서 `lnSpc`를 `append()`로 추가해 `spcAft` 뒤에 놓이게 됐다.
특정 참조 PPT(`spcAft`를 포함한 단락 서식)에서만 오류가 재현되어 원인 파악에 오랜 시간이 걸렸다.

## 슬라이드 복사 시 배경/서식 재설정 금지

`copy_slide_from_prs()`는 원본 슬라이드의 배경(레이아웃/마스터 상속분까지 해석한 실제 배경)을
이미 정확히 복사한다. 복사 직후 `_set_slide_bg_black()`을 또 호출하면 방금 복사한 원본 서식을
하드코딩된 검정으로 덮어써 버린다. 다른 프레젠테이션에서 슬라이드를 복사해 온 경우
(`copy_slide_from_prs`)에는 배경을 다시 강제하지 않는다. 같은 프레젠테이션 내부 템플릿을
복제한 경우(`insert_slide_copy`/`duplicate_slide`)는 원본이 이미 이 문서의 스타일을 따르므로
`_set_slide_bg_black()`을 걸어도 안전하다.

**발견 경위 (2026-07-16):** `update_화답송()`의 악보 슬라이드 분기에서 `copy_slide_from_prs()`
직후 `_set_slide_bg_black()`을 호출해 원본 화답송 악보 PPT의 서식이 사라졌다.

## 도형 분류 시 "빈 텍스트"가 falsy임을 주의

빈 문자열(`""`)은 파이썬에서 falsy다. 도형을 텍스트 키워드로 분류하는 코드(`'ending' if kw in txt
else 'content'` 같은 패턴)에서 `if key == 'content' and txt:` 식으로 "본문이 있는 content만
스킵" 조건을 걸면, 텍스트가 비어 있는 배경 도형(`BlackBg` 등)도 `content`로 분류된 채 스킵되지
않고 그대로 처리 대상에 들어간다. 배경 도형처럼 텍스트 분류 로직에 절대 걸리면 안 되는 도형은
`shape.name == 'BlackBg'`처럼 이름으로 명시적으로 먼저 걸러낸다.

**발견 경위 (2026-07-16):** `_align_ending_slides_to_제2독서()`에서 `BlackBg`가 `content`로
오분류되어 종료전용 템플릿의 작은 콘텐츠 자리표시자 크기로 잘못 리사이즈됐다.

## post-write 재조정 이후 값을 참조할 때는 최신 상태를 다시 측정

여러 슬라이드 조정 단계가 순차 실행되는 파이프라인(예: `replace_reading_slides()` →
`_rebalance_reading_slides_post_write()` → 종료 슬라이드 병합 판단)에서, 뒤 단계가 앞 단계의
**계획값**(`units_pages` 등, 실제 반영 전 값)을 참조하면 안 된다. 앞 단계가 슬라이드를
추가/재배치할 수 있으므로, 실제 물리 슬라이드에서 다시 측정(`_count_slide_lines()` 등)해야 한다.

**발견 경위 (2026-07-16):** 종료 텍스트 병합 여부 판단이 `_page_visual_lines(units_pages[-1])`
(재조정 이전 계획)을 썼다가, 재조정으로 슬라이드가 추가되면서 실제 마지막 슬라이드 내용과
어긋나 병합이 되어야 할 때 안 되는 문제가 발생했다.

## 한글 줄 수 계산: TTC 폰트 인덱스와 Pillow 커닝 한계

`_get_slide_render_params()`가 Pillow로 폰트 폭을 측정해 실제 PowerPoint 렌더링 줄 수를
추정한다. 여기서 두 가지를 주의한다.

1. **TTC(트루타입 컬렉션) 인덱스는 실제 파일에서 직접 확인한다.** `batang.ttc`/`gulim.ttc`처럼
   한 파일에 여러 서체가 들어 있는 경우, "Batang이 index 0일 것"이라고 가정하지 말고
   `ImageFont.truetype(path, size, index=N).getname()`으로 실제 순서를 확인한다. 이 환경에서는
   0=Batang, 1=BatangChe / 0=Gulim, 1=GulimChe였다(반대로 가정했다가 폭을 과대평가해 줄 수
   계산이 틀렸음).
2. **이 환경의 Pillow는 `libraqm`(커닝) 미지원이라 텍스트 폭을 실제보다 넓게 계산한다.**
   `_RENDER_WIDTH_CALIBRATION`(현재 1.03)으로 보정한다. 폰트나 크기 조합이 바뀌면
   `test_missa_regression.py`의 `test_no_reading_slide_line_overflow`로 재검증하고,
   필요하면 실측 기반으로 재보정한다. `libraqm` 설치는 이 환경(Windows Store Python + pip,
   conda 없음)에서 비공식 바이너리가 필요해 권장하지 않는다.

## 단락 분리·병합 함수는 run 개수를 2개로 가정하지 않는다

절 하나 = run 2개(오렌지 절 번호 + 흰색 본문)라고 가정하는 코드는 여러 절이 하나의 논리
단락으로 병합된 경우(`layout_units_on_slides()`의 `is_continuation` 병합) 깨진다. 이때 단락은
[오렌지][흰색][오렌지][흰색]... 처럼 run이 3개 이상이 된다. 단락을 두 조각으로 자르는 함수는
반드시 "분리 지점이 몇 번째 run의 몇 번째 글자인지"를 실제로 계산해서, 그 지점 앞뒤 run들의
서식(rPr)을 각각 그대로 유지해야 한다. `runs[0]`/`runs[1]`처럼 인덱스를 고정하거나 "마지막
run만 남기고 나머지 삭제" 같은 방식은 텍스트는 보존해도 중간 run의 색상(서식)을 조용히
잃어버린다 — 예외 없이 실행되므로 육안 검수 전까지 발견되지 않는다.

**발견 경위 (2026-07-26):** `_split_para_at_lines()`가 앞부분은 `runs[0]`(절 번호)+`runs[1]`
(본문)만 남기고 `runs[2:]`를 삭제, 뒷부분은 마지막 run만 남기고 나머지를 삭제하는 방식이었다.
51절과 52절이 continuation으로 병합된 단락(런: [51-오렌지][본문][52-오렌지][본문])이 슬라이드
경계에서 분리되면서, 중간의 52절 오렌지 run이 흰색 본문 run에 흡수되어 텍스트("52")는 남았지만
색이 사라졌다. `_missing_orange_verse_numbers()`(§검증) 추가로 이런 사례를 자동으로 잡는다.

## 검증: 텍스트 존재가 아니라 절 번호별 오렌지색 렌더링을 확인

`validate()`는 "오렌지색 run이 프레젠테이션 어딘가에 하나라도 있는지"가 아니라, JSON
`content`를 `parse_into_verse_units()`로 파싱해 나온 절 번호 전체가 해당 독서/복음 슬라이드
범위 안에서 실제로 오렌지색 run으로 렌더링됐는지 절 번호 단위로 대조한다
(`_missing_orange_verse_numbers()`). 텍스트만 남고 색이 사라지는 위 버그처럼, "텍스트 존재
여부"만 보는 검증은 이런 회귀를 통과시킨다.

## 회귀 테스트

`test_missa_regression.py`에 주일(20260712)·평일(20260624) 통합 테스트와 핵심 함수 단위
테스트가 있다. 독서·복음 줄 수 계산, 화답송/성가 처리 분기, 배경색, 절 번호 오렌지색처럼 이
문서에 기록된 버그와 직결된 항목을 검증하므로, 관련 로직을 수정하면
`pytest test_missa_regression.py -v`를 실행해 회귀 여부를 확인한다.
