# CLAUDE.md — missa_ppt 프로젝트

## 프로젝트 구조

매일미사 웹데이터로 미사 PPT를 자동 생성하는 도구.

**파이프라인:** `missa_to_json.py`(missa.cbck.or.kr 크롤링 → `missa_YYYYMMDD.json`)
→ `missa_to_ppt.py`(JSON + 성가번호로 참조/템플릿 PPT를 수정해 결과 PPT 생성).

**핵심 파일:**
- `missa_to_json.py` — 독서·복음·화답송·성가 등 섹션을 크롤링해 JSON 출력.
- `missa_to_ppt.py` — 메인 생성기. 슬라이드 복사·독서 줄 나눔·절 번호 색상·`validate()` 등.
- `missa_psalm_score_image.py` — 화답송 악보 원본 이미지(PNG/JPG) → 슬라이드 변환(1차
  마일스톤, 순수 함수 묶음, `missa_to_ppt.py`를 import하지 않는 단방향 의존). 진입점
  `render_화답송_score_slide()`. 고정 템플릿 자산은 `assets/화답송_악보_template.pptx`
  (`tools/build_화답송_template.py`로 생성). 아직 `missa_to_ppt.py`와 배선되지 않은 독립
  모듈 — 통합(입력창/`find_files()`/`update_화답송()` 연결)은 2차 마일스톤.
- `test_missa_progression.py` — 아직 안정화되지 않은 신규 동작을 먼저 명세하는 프로그레션
  테스트(TDD red→green). 안정화되면 `test_missa_regression.py`로 승격.
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

**발견 경위 (2026-08-29):** `missa_psalm_score_image.py`의 `_replace_title_paragraph`가
제목 단락의 run/br만 지우고 `endParaRPr`는 남긴 뒤 새 run을 `append()`해, `pPr → endParaRPr →
r` 순서로 렌더링됐다. `validate_pptx_structure()`(XML 제어문자·끊어진 rId만 검사)와 42개
프로그레션 테스트 전부가 이 스키마 위반을 잡지 못했고, 독립 코드 리뷰(`ooxml-code-reviewer`)가
실제 렌더링 XML을 직접 열어보고서야 발견했다. 구현자 본인은 "pPr을 직접 조작하지 않았으니
안전하다"고 자평했으나, append 위험을 "pPr 자식 순서"라는 기억된 형태로만 점검해 같은 원리가
`<a:p>` 자식 레벨(run vs endParaRPr)에도 적용됨을 일반화하지 못했다.

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

## 텍스트 키워드로 도형을 식별할 때 부분 문자열 충돌 주의

`kw in shape.text_frame.text` 같은 부분 문자열 포함검사로 특정 도형(제목·라벨 등)을 찾을 때,
그 키워드가 **다른** 도형의 텍스트에도 우연히 포함되어 있으면 도형 순회 순서에 의존하는
암묵적 안전성만 남는다. 첫 매칭에서 `return`하는 코드는 의도한 도형이 항상 먼저 순회되는 동안만
안전하고, 도형 순서가 바뀌거나(슬라이드 복사·재배치) 텍스트가 수정되면 조용히 엉뚱한 도형이
매칭되어 그 내용을 덮어쓴다. 키워드는 "의도한 도형에만 나타나는" 형태(공백·구두점 포함 등
더 구체적인 패턴)로 좁혀서, 순회 순서와 무관하게 안전하도록 만든다.

**발견 경위 (2026-08-29):** `missa_psalm_score_image.py`의 `_update_title`이 제목 도형을
`"화 답 송" in t or "화답송" in t`(공백 없는 형태 포함)로 찾았는데, 화답송 악보 템플릿의
저작권 표기 도형에 "…박원주 <**화답송**과 시편의 노래>…"라는 문구가 있어 공백 없는 "화답송"
키워드가 이 도형에도 매칭됐다. 제목 도형이 저작권 도형보다 먼저 순회되는 현재 도형 순서
덕분에만 우연히 안전했고, 순서가 바뀌면 저작권 문구가 제목으로 조용히 덮어써질 수 있었다.
독립 코드 리뷰(`ooxml-code-reviewer`)가 지적해 키워드를 공백이 있는 `"화 답 송"` 형태(제목
도형에만 나타남)로 좁혀 해결했다.

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
