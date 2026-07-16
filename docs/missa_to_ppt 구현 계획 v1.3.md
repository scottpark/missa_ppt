# missa_to_ppt.py 구현 계획 v1.3 — 버그 수정 및 회귀 테스트

- 대상 파일: `missa_to_ppt.py` (~6,600줄), `test_missa_regression.py` (신규)
- 작성일: 2026-07-16
- 이전 버전: `docs/archive/missa_to_ppt 구현 계획 v1.2 (평일미사 추가).md`

---

## 1. 배경

v1.2에서 계획한 평일미사 지원은 커밋(`평일미사 지원 추가`)되어 기능적으로 동작했으나,
2026-07-16 실제 주일·평일 날짜로 생성해 육안 검수하는 과정에서 v1.2 요구사항 자체는
이미 정의돼 있던(9줄 제한, 화답송 서식 보존 등) 항목들이 구현상 정확히 지켜지지 않는
버그 6건이 발견되어 함께 수정했다. 이 문서는 그 수정 내역과, 재발 방지를 위해 추가한
회귀 테스트 스위트를 기록한다.

---

## 2. 버그 수정 이력 (2026-07-16)

### 2.1 화답송 악보 슬라이드 서식 미적용

**증상**: 주일미사 화답송 악보 슬라이드에 원본 화답송 악보 PPT의 서식(배경 포함)이 적용되지 않음.

**원인**: `update_화답송()`에서 `copy_slide_from_prs()`로 악보 슬라이드를 복사한 직후
`_set_slide_bg_black()`을 호출하고 있었다. `copy_slide_from_prs()`는 원본 배경(레이아웃/마스터에서
상속된 배경까지 포함해 해석한 실제 배경)을 이미 정확히 복사하도록 설계된 함수인데, 뒤이은
`_set_slide_bg_black()` 호출이 그 결과를 하드코딩된 검정으로 덮어써버렸다. 이 호출은 평일미사
작업 중 텍스트 슬라이드에도 일괄 적용하며 실수로 악보 슬라이드 분기에도 들어간 것으로 보인다.

**수정**: `update_화답송()`의 악보 슬라이드 분기(`if i % 2 == 0:`)에서 `_set_slide_bg_black()` 호출 제거.
텍스트 슬라이드 분기는 그대로 유지(같은 프레젠테이션 내부 템플릿 복제이므로 배경 강제가 안전함).

### 2.2 Post-write 재조정 로직이 슬라이드당 9줄을 못 채우는 문제 (연쇄 초과)

**증상**: 독서/복음 본문 슬라이드가 10~13줄까지 넘치는데도 재조정되지 않고 그대로 저장됨.

**원인**: `_rebalance_reading_slides_post_write()`의 초과 처리 분기가 "마지막 단락을 다음
슬라이드로 이동 → 이동 시 다음 슬라이드도 넘치면 롤백 후 단락 분리"까지만 시도했다. 앞 슬라이드의
초과분이 계속 뒤로 밀리며 누적되는 연쇄 상황에서는, 마지막 단락 하나를 분리해도 전체 초과분을
해소하지 못하는 경우(`keep = last_p_lines - excess`가 0 이하)가 생기는데, 이때 "분리 불가" 경고만
남기고 그대로 방치했다.

**수정**: 인접 슬라이드로 이동·분리 모두 불가능한 경우, 현재 위치 뒤에 새 슬라이드를 삽입해
초과분을 옮기는 폴백을 추가했다(마지막 슬라이드 초과 시 이미 쓰던 것과 동일한 패턴). 삽입 후
`n_content`를 갱신하고 처음부터 재스캔한다.

### 2.3 화답송 종료 슬라이드 병합 판단이 재조정 이전 값을 사용

**증상**: 마지막 본문 슬라이드가 실제로는 짧은데도(예: 1~2줄) "종료 텍스트 병합" 조건(5줄 이하)을
만족하지 못한 것으로 판단되어, 종료 텍스트("주님의 말씀입니다...")가 별도의 빈 슬라이드에 원본
위치 그대로(엉뚱한 곳에) 남는 문제.

**원인**: `replace_reading_slides()`의 병합 여부 판단이 `_page_visual_lines(units_pages[-1])`
(post-write 재조정 **이전**에 계획된 마지막 페이지 줄 수)를 사용했다. 그런데 이 판단 직전에 실행되는
`_rebalance_reading_slides_post_write()`가 슬라이드를 추가로 삽입/재배치할 수 있어, 실제 마지막
슬라이드의 내용이 계획과 달라질 수 있다.

**수정**: `_count_slide_lines(prs.slides[last_content_idx])`로 실제 마지막 물리 슬라이드의
렌더링된 줄 수를 확인하도록 변경. `last_content_idx = content_start + needed + n_rebalance_inserted - 1`.

### 2.4 병합된 종료 슬라이드의 배경 도형이 잘못된 크기로 축소됨

**증상**: 독서·복음 종료 텍스트가 통합된 마지막 슬라이드의 검정 배경(`BlackBg`)이 슬라이드
전체가 아니라 훨씬 작은 영역만 덮음.

**원인**: `_align_ending_slides_to_제2독서()`가 슬라이드의 텍스트 도형을 종료 텍스트("주님의
말씀입니다" 등 키워드 포함) 여부로 `ending`/`content`로 분류하는데, `BlackBg` 도형은 빈 텍스트
프레임을 가지고 있어 `content`로 오분류되었다. 이후 "본문+ending 통합 슬라이드의 content shape는
위치 조정 제외" 스킵 조건이 `key == 'content' and txt`였는데, `BlackBg`의 `txt`가 빈 문자열이라
falsy가 되어 스킵되지 않고, 제2독서 종료전용 템플릿의 (전체 화면보다 작은) 콘텐츠 자리표시자
크기로 덮어써졌다.

**수정**: `BlackBg` 도형을 이름으로 명시적으로 제외(참조 위치 수집 시, 적용 시 모두).

### 2.5 폰트 인덱스 반전으로 인한 줄 수 계산 오차

**증상**: 위 2.2 수정 후에도 특정 슬라이드가 실제로는 8줄인데 Pillow 계산은 9줄로 판단해,
post-write 재조정이 다음 슬라이드에서 내용을 더 끌어오지 않고 멈춤.

**원인**: `_get_slide_render_params()`의 `_FONT_MAP`에서 TTC(트루타입 컬렉션) 파일 내부 인덱스가
반대로 매핑돼 있었다. 실제 `C:\Windows\Fonts\batang.ttc`는 index 0=Batang, index 1=BatangChe인데
코드는 반대로 가정하고 있었다(gulim.ttc도 동일). 참조 PPT가 "바탕"(Batang, 가변폭)을 쓰는데
실제로는 "바탕체"(BatangChe, 고정폭이라 더 넓음)로 폭을 측정해 실제보다 넓게(=줄이 더 필요한
것으로) 계산했다.

**수정**: `_FONT_MAP`의 Batang/BatangChe, Gulim/GulimChe 인덱스를 실제 파일 순서에 맞게 교정.

### 2.6 Pillow의 커닝 미지원으로 인한 잔여 폭 오차

**증상**: 폰트 인덱스 수정 후에도 일부 슬라이드에서 여전히 실제보다 1줄 많게 계산됨(2~4% 폭
과대평가).

**원인**: 이 환경의 Pillow는 `libraqm`(커닝·복합 텍스트 셰이핑) 미지원 빌드라, 글자 advance width를
단순 합산해 실제 PowerPoint(DirectWrite 엔진) 렌더링보다 폭을 살짝 넓게 계산한다.
`libraqm` 설치는 이 환경(Windows Store Python + pip, conda 없음)에서 비공식 바이너리 설치가
필요해 위험도가 높고, 설치해도 DirectWrite와 100% 일치를 보장하지 않는다.

**수정**: 실측 확인된 7개 슬라이드(20260712 복음)를 기준으로 보정 계수를 0~6% 범위에서 스윕
테스트해, 1.03~1.04에서만 전체 일치함을 확인. `_RENDER_WIDTH_CALIBRATION = 1.03`을
`_get_slide_render_params()`의 `box_px` 계산에 곱하는 방식으로 적용.

> 이 값은 실측 기반 경험적 보정치이며, 폰트·크기 조합이 크게 달라지면 재보정이 필요할 수 있다.
> `test_missa_regression.py`의 `test_no_reading_slide_line_overflow`가 이 값의 유효성을
> 지속적으로 검증한다.

### 2.7 (관련 선행 변경) missa_to_json.py — title/chapter_verse 분리

독서·복음 JSON의 `title`에 장·절 번호까지 섞여 있던 것을 `title`(성서명)과
`chapter_verse`(장·절)로 분리하고, `content`의 절 사이 개행을 제거해 하나의 문자열로 합쳤다.
위 2.1~2.6 수정이 절 단위 텍스트 처리·병합 로직을 다루므로, 이 사전 변경과 호환되게 확인했다.

---

## 3. 회귀 테스트 스위트 (`test_missa_regression.py`)

### 3.1 목적

주일미사·평일미사 처리 로직의 회귀 테스트. 향후 토요일미사 지원이 추가되거나 다른 기능이
변경됐을 때, 기존 주일/평일 로직이 깨지지 않았는지 확인하는 기준선.

### 3.2 실행

```
pip install pytest
pytest test_missa_regression.py -v
```

날짜 폴더에 `~$*.pptx` 잠금 파일이 있으면(PowerPoint에서 파일이 열려 있음) 해당 케이스는
자동으로 skip 처리된다.

### 3.3 테스트 데이터

- 주일: `20260712` (2026-07-12, 연중 제15주일)
- 평일: `20260624` (2026-06-24, 성 요한 세례자 탄생 대축일, 수요일)

두 폴더 모두 참조 PPT·JSON·성가 파일이 미리 준비되어 있어야 하며, `MASS_CASES` 리스트에
새 항목(예: 토요일미사 케이스)을 추가하면 통합 테스트 전체가 파라미터화되어 자동으로 함께 실행된다.

### 3.4 구성

**단위 테스트** — 순수 함수, PPT 생성 없이 즉시 실행

| 클래스 | 검증 대상 |
|---|---|
| `TestIsSundayMass` | `is_sunday_mass()` 요일 판단. 토요일이 현재 "평일미사"로 처리됨을 명시적으로 고정 — 토요일 로직 추가 시 이 테스트가 회귀 신호가 됨 |
| `TestWrapLineCount` | `_wrap_line_count()` 줄바꿈 계산 |
| `TestParseIntoVerseUnits` | `parse_into_verse_units()` 절 번호 파싱 |

**통합 테스트** — `generated_case` fixture(모듈 스코프, `MASS_CASES` 파라미터화)가 각 케이스에
대해 실제로 `missa_to_ppt.py`를 CLI로 실행해 PPT를 생성한 뒤, 아래 항목을 검증한다.

| 테스트 | 검증 내용 | 관련 버그 |
|---|---|---|
| `test_output_opens_without_corruption` | 출력 파일이 손상 없이 열리는지 | — |
| `test_liturgy_type_detected_correctly` | 로그에 미사 유형이 올바르게 찍히는지 | — |
| `test_required_sections_present` | 필수 섹션(입당송·독서·화답송·복음·영성체송)이 모두 있는지 | — |
| `test_no_reading_slide_line_overflow` | 독서·복음 본문 슬라이드가 9줄을 넘지 않는지 | §2.2, §2.5, §2.6 |
| `test_blackbg_covers_full_slide` | `BlackBg` 도형이 항상 슬라이드 전체를 덮는지 | §2.4 |
| `test_화답송_pattern_matches_mass_type` | 화답송이 주일(악보/텍스트 교대)·평일(텍스트만) 패턴에 맞는지 | §2.1 |
| `test_hymn_score_copy_matches_mass_type` | 성가가 주일(악보 복사)·평일(번호만 갱신) 패턴에 맞는지 | — |

### 3.5 확장 방법 (토요일미사 등 신규 케이스 추가 시)

1. `MASS_CASES`에 새 케이스(날짜·`is_sunday` 플래그·성가 번호) 추가
2. 토요일 전용 로직이 생기면(예: `is_saturday_vigil_mass()`) `TestIsSundayMass`에 대응하는
   단위 테스트 케이스 추가, 기존 `test_saturday_currently_treated_as_weekday`는 새 기대값으로 갱신
3. 토요일미사만의 새로운 처리 분기(예: 미사 후 기도가 평일과 다르게 동작 등)가 생기면
   통합 테스트에 해당 검증 함수 추가
4. `pytest test_missa_regression.py -v`로 전체 재실행 → 기존 주일/평일 케이스가 여전히
   통과하는지 확인 후 병합
</content>
