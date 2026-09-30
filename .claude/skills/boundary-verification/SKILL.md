---
name: boundary-verification
description: "missa_ppt 프로젝트에서 두 컴포넌트가 만나는 경계면(JSON 스키마↔소비 코드, find_sections() 키↔실제 템플릿 슬라이드, config.json 스키마↔실제 읽는 코드)이 실제로 맞물리는지 교차 비교하는 검증 스킬. 각각 개별로는 정상이어도 연결 지점에서 어긋나는 버그를 잡는다. 새 기능 구현 직후, 회귀 테스트 실행 전, '검증해줘' 요청 시 사용."
---

# Boundary Verification — 경계면 정합성 검증

이 프로젝트에서 실제로 발생한 버그들의 공통점: 각 함수는 개별적으로 봤을 때 "맞게" 동작했지만,
두 함수가 만나는 지점의 가정이 어긋나 있었다. 예를 들어 `_split_para_at_lines()`는 텍스트를
정확히 보존했지만 "절 하나 = run 2개"라는 암묵적 가정이 실제 데이터(continuation 병합 단락)와
맞지 않아 서식이 사라졌다 — 함수 자체를 리뷰해서는 못 잡고, 그 함수가 실제로 받는 입력의 모양을
같이 봐야 잡을 수 있는 버그였다.

## 원칙: 한쪽만 읽지 않는다

아래 각 경계면은 **양쪽을 동시에 열어** 비교한다. 한쪽만 보고 "존재하니까 됐다"고 판단하지 않는다.

## 경계면 1: JSON 스키마 ↔ 소비 코드

**왼쪽(생산자)**: `missa_to_json.py`의 `parse_missa()`/`parse_section()`이 실제로 만드는 dict의
키와 값 형태(`title`/`chapter_verse` 분리 여부, `content`의 개행 유지 여부 등)
**오른쪽(소비자)**: `missa_to_ppt.py`의 `get_json_data()` 호출 이후 `json_data['...']`로
접근하는 모든 지점

```
grep -n "json_data\[" missa_to_ppt.py
grep -n "^def parse_\|data\[" missa_to_json.py
```

두 목록을 나란히 놓고, 소비 코드가 참조하는 키가 생산 코드에서 실제로 채워지는지, 값의 형태
(문자열 vs dict, 개행 유무)가 소비 코드의 처리 방식과 맞는지 확인한다. 새 미사 유형이 JSON에
새 필드를 요구한다면 이 경계면부터 깨진다.

## 경계면 2: `find_sections()` 키 ↔ 실제 템플릿 슬라이드

**왼쪽**: `find_sections()`가 반환하는 dict 키(어떤 텍스트 패턴으로 탐지하는지)
**오른쪽**: 실제 참조 PPT의 슬라이드 텍스트

```
python .claude/skills/boundary-verification/scripts/dump_find_sections.py "<참조PPT경로>"
```

출력된 `키 => slide N   텍스트조각들`을 보고, 그 슬라이드가 정말 그 섹션이 맞는지 사람이
확인한다. 새 템플릿(청년미사 등)에서 텍스트 패턴이 조금만 달라도(예: "복 음" 옆에 "GOSPEL"이
병기) 탐지가 엉뚱한 슬라이드를 가리킬 수 있다 — 이 스크립트 없이는 코드 리뷰만으로 못 잡는다.

## 경계면 3: `config.json` 스키마 ↔ 실제 읽는 코드

**왼쪽**: `config.json`의 실제 키
**오른쪽**: `_load_config()`, `get_onedrive_hymn_folder()`, `_com_verification_enabled()` 등이
`config.get(...)`으로 읽는 키

```
cat config.json
grep -n "config.get\|config\[" missa_to_ppt.py
```

새 미사 유형/성당을 위해 config 필드를 추가했다면, 실제로 그 필드를 읽는 코드가 존재하는지
(설계만 하고 구현을 빠뜨리지 않았는지), 그리고 필드가 없을 때의 기본값 처리가 문서
(`docs/missa_to_ppt 구현 계획.md` §3.2)와 일치하는지 확인한다.

## 경계면 4: 설계서 함수 명세 ↔ 실제 구현

**왼쪽**: `_workspace/01_architect_design.md`의 "완전 신규"/"파라미터 분기 필요" 표에 적힌
함수 시그니처
**오른쪽**: 실제 구현된 함수 시그니처(`grep -n "^def "`)

설계서에 명시된 파라미터가 실제 구현에 그대로 반영됐는지, 이름이 바뀌었다면
`_workspace/02_specialist_impl_notes.md`에 그 사유가 기록돼 있는지 확인한다.

## 회귀 테스트 실행

경계면 검증 후 반드시 기존 회귀 테스트를 실행해 기존 케이스(20260712 주일, 20260624 평일)를
깨뜨리지 않았는지 확인한다:

```
pytest tests/test_missa_regression.py tests/test_ppt_com_verify.py -v
```

새 미사 유형/성당이 추가됐다면 `MASS_CASES`(`tests/test_missa_regression.py`)에 새 케이스를 추가하고
같은 명령으로 재실행한다. `~$*.pptx` 잠금 파일이 있는 폴더는 자동 skip되므로, PowerPoint에서
해당 파일을 열어둔 채 테스트를 실행하지 않는다.

## 리포트 형식

`_workspace/03_qa_report.md`에 아래 형식으로 기록한다:

```markdown
## 경계면 검증 결과
| 경계면 | 상태 | 비고 |
|--------|------|------|
| JSON↔소비코드 | 통과/실패/미검증 | |
| find_sections↔템플릿 | 통과/실패/미검증 | |
| config↔읽는코드 | 통과/실패/미검증 | |
| 설계서↔구현 | 통과/실패/미검증 | |

## 회귀 테스트
(pytest 출력 요약, 실패 시 재현 방법)
```
