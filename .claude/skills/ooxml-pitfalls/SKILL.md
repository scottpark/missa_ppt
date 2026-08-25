---
name: ooxml-pitfalls
description: "missa_to_ppt.py에서 python-pptx/lxml로 OOXML(a:pPr, spTree, rPr 등)을 직접 조작하기 전에 반드시 확인하는 체크리스트. pPr에 요소 추가, 슬라이드 복사, 도형 텍스트 분류, 단락/run 분리, PowerPoint COM 자동화 코드를 건드릴 때 반드시 사용. 이 체크리스트를 건너뛰면 PowerPoint에서 '손상된 파일'로 뜨거나, 예외 없이 조용히 서식이 사라지는 버그가 재발한다."
---

# OOXML Pitfalls — missa_to_ppt.py 구현 전 체크리스트

이 프로젝트의 OOXML 버그는 전부 **예외를 던지지 않고 조용히 실패**했다(파일은 저장되고 열리지만
서식이나 텍스트가 틀림). 그래서 육안 검수 전까지 발견되지 않았다. 아래 항목은 `CLAUDE.md`에
기록된 실제 발견 경위를 근거로 한다 — 원문은 `CLAUDE.md`를 참고하고, 여기서는 "코드를 쓰기
직전에 무엇을 확인해야 하는지"만 행동 단위로 정리한다.

## 1. `<a:pPr>`에 요소를 추가/교체할 때

**증상이 나면**: PowerPoint가 파일을 열 때 "복구할 수 있는 콘텐츠가 있습니다" 경고.

- [ ] `pPr.append(element)`를 쓰고 있지 않은가? → **금지**. `pPr`의 자식은 스키마 순서
  (`lnSpc → spcBef → spcAft → buClr → buSz → buFont → bu* → tabLst → defRPr → extLst`)를
  지켜야 한다. `lnSpc`를 append하면 `spcAft` 뒤로 밀려 순서 위반이 된다
- [ ] 새 요소를 넣을 때 `insert(0, element)`(맨 앞이 맞는 경우) 또는 스키마 순서에 맞는
  위치에 `insert`를 쓰고 있는가?
- [ ] 기존 요소를 교체할 때 `remove()` 후 `insert(0, new_element)`를 쓰고 있는가?
  (`append()`로 재추가하면 안 됨)

## 2. 다른 프레젠테이션에서 슬라이드를 복사할 때 (`copy_slide_from_prs()` 계열)

**증상이 나면**: 복사해온 슬라이드의 배경이 하드코딩된 검정으로 덮이거나, 도형/텍스트가
확대되어 잘려 보임.

- [ ] 복사 직후 `_set_slide_bg_black()`을 또 호출하고 있지 않은가? → `copy_slide_from_prs()`가
  이미 원본 배경(레이아웃/마스터 상속까지 해석)을 정확히 복사한다. 다시 호출하면 그 결과를
  덮어쓴다
  - 예외: **같은 프레젠테이션 내부** 템플릿을 복제한 경우(`insert_slide_copy`/`duplicate_slide`)는
    원본이 이미 이 문서 스타일을 따르므로 `_set_slide_bg_black()`을 걸어도 안전
- [ ] 원본과 타겟의 슬라이드 크기(가로세로)가 다를 가능성이 있는가? → `scale_x`/`scale_y`
  보정이 도형 위치·크기뿐 아니라 폰트 크기(`sz`)와 단락 여백(`marL`/`marR`/`indent`/
  `defTabSz`)까지 함께 축소하는지 확인(텍스트 잘림 방지, `text_scale = min(scale_x, scale_y)`)

## 3. 도형을 텍스트 키워드로 분류할 때

**증상이 나면**: 배경 사각형(`BlackBg` 등)이 콘텐츠 텍스트박스로 오분류되어 엉뚱한 크기로
리사이즈됨.

- [ ] 빈 문자열(`""`)은 파이썬에서 falsy라는 점을 감안했는가? `if key == 'content' and txt:`
  같은 조건은 "텍스트가 있는 content만 스킵"하려는 의도였어도, 텍스트가 비어 있는 배경 도형이
  `content`로 분류된 채 그대로 통과해버린다
- [ ] 배경 도형처럼 절대 걸리면 안 되는 도형은 `shape.name == 'BlackBg'`처럼 **이름으로 먼저**
  명시적으로 걸러내고 있는가?

## 4. 여러 단계가 순차 실행되는 파이프라인에서 앞 단계 결과를 참조할 때

**증상이 나면**: 슬라이드가 추가/재배치된 뒤에도 이전 "계획값"을 써서 실제 상태와 어긋남.

- [ ] 참조하려는 값이 "계획 단계의 값"(예: `units_pages` 같은 사전 배분 결과)인가, 아니면
  "실제로 슬라이드에 반영된 뒤 다시 측정한 값"인가? 앞 단계가 슬라이드를 추가/삭제/재배치할
  수 있다면 반드시 재측정 함수(`_count_slide_lines()` 등)로 최신 상태를 다시 읽는다

## 5. 단락/run을 두 조각으로 분리할 때

**증상이 나면**: 텍스트는 정상인데 절 번호나 특정 구간의 색상만 조용히 사라짐.

- [ ] "절 하나 = run 2개(절 번호+본문)"라고 가정하고 있지 않은가? 여러 절이 continuation으로
  하나의 논리 단락에 병합되면 run이 3개 이상(`[오렌지][본문][오렌지][본문]...`)이 될 수 있다
- [ ] `runs[0]`/`runs[1]` 같은 고정 인덱스나 "마지막 run만 남기고 나머지 삭제" 방식을 쓰고
  있지 않은가? → 분리 지점이 몇 번째 run의 몇 번째 글자인지 실제로 계산해, 그 앞뒤 run들의
  서식(rPr, 색상 포함)을 각각 그대로 유지해야 한다

## 6. PowerPoint COM 자동화 (`ppt_com_verify.py`)를 건드릴 때

- [ ] `win32com.client.gencache.EnsureDispatch`를 쓰고 있지 않은가? → **절대 금지**. 최초
  실행한 PowerPoint 버전에 고정된 typelib 캐시가 생겨, 빌드된 exe가 다른 버전(2007~365)의
  최종 사용자 PC에서 깨진다. late-binding(`win32com.client.Dispatch`)만 사용
- [ ] 슬라이드별로 검증 결과를 캡(cap)해서 건너뛰는 로직을 추가하려는 것은 아닌가? → 과거
  "슬라이드별 캡이 거짓 성공을 보고하는 버그"가 실제로 있었다. 조건(경계값+enabled)을 만족하면
  매번 실제로 COM에 물어야 하며, 재시도 방지는 `max_adjust`/`_MAX_SWEEPS` 같은 전역 안전장치가
  담당한다
- [ ] 빈 프레젠테이션을 계속 열어두는 `_keepalive_presentation`을 건드리고 있다면: 열린
  프레젠테이션이 0개가 되면 PowerPoint 프레임 창 자체가 사라져 `Presentations.Open`이 실패한다

## 7. 한글 줄 수 계산(Pillow)을 건드릴 때

- [ ] TTC(트루타입 컬렉션) 폰트의 인덱스를 "이럴 것"이라고 가정하지 않고
  `ImageFont.truetype(path, size, index=N).getname()`으로 실제 확인했는가?
- [ ] 이 환경의 Pillow는 `libraqm`(커닝) 미지원이라 폭을 과대평가한다 — `_RENDER_WIDTH_CALIBRATION`
  보정 계수를 건드렸다면 `test_no_reading_slide_line_overflow`로 재검증했는가?

## 저장 후 필수 확인

코드를 다 쓴 뒤에는 항상:

```
python -c "from missa_to_ppt import validate_pptx_structure; print(validate_pptx_structure('<출력경로>'))"
```

빈 리스트가 아니면 구조 손상이 있다는 뜻이다. `regression-qa` 에이전트가 이 확인을 다시
하겠지만, 구현 직후 스스로도 먼저 확인해 왕복을 줄인다.
