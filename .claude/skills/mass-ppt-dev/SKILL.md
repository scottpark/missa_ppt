---
name: mass-ppt-dev
description: "missa_ppt 프로젝트(미사 PPT 자동 생성 도구)의 개발·확장·버그 수정 오케스트레이터. TDD(설계→프로그레션 테스트 선작성→구현→독립 코드 리뷰→회귀 검증) 파이프라인으로 새 미사 유형(청년미사·어린이미사 등) 지원을 코드로 구현, 새 성당/본당 지원 추가, PPT 템플릿 구조 확장, missa_to_ppt.py/missa_to_json.py의 섹션 처리·OOXML 슬라이드 조작 로직 구현 또는 버그 수정 시 반드시 사용. '청년미사 지원 만들어줘/구현해줘', '어린이미사 지원 추가해줘', '이 본당용 템플릿 지원해줘', 'PPT 생성 로직 고쳐줘', '독서 슬라이드 줄바꿈 버그 고쳐줘' 같은 요청에 트리거. 특정 날짜의 미사 PPT를 실제로 생성하는 단순 실행 요청(예: '이번 주 청년미사 PPT 만들어줘', '20260830 미사 PPT 뽑아줘')은 이미 구현된 기능이면 missa_to_ppt.py를 직접 실행하면 되므로 이 스킬의 대상이 아니다 — 아직 지원하지 않는 유형/성당이라 코드 작업이 필요할 때만 트리거. 후속 작업(설계 다시 검토, 일부 미사유형만 재구현, 이전 설계 기반 보완, 프로그레션/회귀 테스트 추가·재실행, 리뷰만 다시, 문서 동기화만 다시)에도 반드시 이 스킬을 사용."
---

# Mass PPT Dev Orchestrator

missa_ppt 프로젝트의 확장(새 미사 유형·새 성당) 및 유지보수 작업을 **설계 → TDD 구현(프로그레션
테스트 선작성) → 독립 코드 리뷰 → 회귀 검증/승격 → 문서동기화** 파이프라인으로 조율하는
오케스트레이터.

## 실행 모드: 에이전트 팀 (기본) — 소규모 작업은 경량 경로로 하이브리드 축소

이 프로젝트는 `missa_to_ppt.py` 단일 파일(~7,200줄)에 로직이 집중돼 있어 여러 에이전트가
동시에 같은 파일을 편집하면 충돌한다. 그래서 **파이프라인 패턴**을 팀 모드로 운영한다 — 순차
실행이 기본이지만, 구현 중 설계 재검토·리뷰 왕복·검증 실패로 되돌아가는 경우가 실제로 잦으므로
`SendMessage`로 실시간 왕복이 가능한 팀 모드가 서브 에이전트보다 유리하다.

## 에이전트 구성

| 팀원 | 에이전트 타입 | 역할 | 스킬 | 출력 |
|------|-------------|------|------|------|
| mass-template-architect | mass-template-architect | 새 템플릿 조사·설계 (코드 수정 안 함) | mass-template-analysis | `_workspace/01_architect_design.md` |
| ppt-ooxml-specialist | ppt-ooxml-specialist | TDD 구현(프로그레션 테스트 선작성) + 문서 동기화 | tdd-progression-testing, ooxml-pitfalls, docs-sync | 코드 변경 + `test_missa_progression.py` + `_workspace/02_specialist_impl_notes.md` |
| ooxml-code-reviewer | ooxml-code-reviewer | specialist와 독립된 시각의 코드 리뷰 | ooxml-pitfalls | `_workspace/02b_review_report.md` |
| regression-qa | regression-qa | 회귀 검증 + 경계면 교차 검증 + 프로그레션 테스트 승격 | boundary-verification | `_workspace/03_qa_report.md` |

**왜 리뷰어가 구현자와 분리돼 있는가**: 같은 에이전트가 테스트와 구현을 모두 쓰면 자신의
암묵적 가정(예: "절 하나 = run 2개")을 테스트에도 똑같이 반영해버려, 테스트가 통과해도 그
가정 자체가 틀렸을 때는 못 잡는다. `ooxml-code-reviewer`는 팀에 별도 인스턴스로 스폰되므로
기본적으로 specialist의 대화 맥락을 공유하지 않는다 — 게다가 에이전트 정의 자체가
"specialist의 자체 설명(`02_specialist_impl_notes.md`)을 독립 판단 이후에만 읽는다"는 원칙을
강제한다.

## 워크플로우

### Phase 0: 컨텍스트 확인 (후속 작업 지원)

1. 프로젝트 루트의 `_workspace/` 존재 여부 확인
2. 실행 모드 결정:
   - **`_workspace/` 미존재** → 초기 실행. Phase 1로 진행
   - **`_workspace/` 존재 + 사용자가 부분 수정 요청**(예: "구현만 다시", "리뷰만 다시",
     "문서만 갱신") → 부분 재실행. 해당 에이전트만 재호출하고 나머지 산출물은 그대로 둔다
   - **`_workspace/` 존재 + 새 미사유형/성당 요청** → 새 실행. 기존 `_workspace/`를
     `_workspace_{YYYYMMDD_HHMMSS}/`로 이동 후 Phase 1 진행
3. 부분 재실행 시: 이전 산출물 경로(`_workspace/01_architect_design.md` 등)를 재호출하는
   에이전트의 프롬프트에 포함해, 기존 결과를 읽고 개선하도록 지시

### Phase 1: 요청 분석 및 팀 필요 여부 판단 (Triage)

1. 사용자 요청이 다음 중 무엇인지 판단한다:
   - **구조적 확장** (새 미사 유형, 새 성당, 새 섹션, 템플릿 구조 변경) → Phase 2로 진행(전체 팀)
   - **국소적 버그 수정** (기존 함수 1~2개, 새 템플릿/구조 변경 없음) → **경량 경로**: 전체
     팀을 구성하지 않고 아래 순서로 서브 에이전트(`Agent` 도구, `model: "opus"`)를 순차 호출:
     `ppt-ooxml-specialist`(TDD로 수정, `test_missa_progression.py`에 재현 테스트 선작성) →
     `ooxml-code-reviewer`(리뷰) → 리더가 `boundary-verification` 절차 직접 수행 또는
     `regression-qa` 1회 호출 → Phase 6(문서 동기화)로 진행. **경량 경로에서도 TDD와 독립
     리뷰는 생략하지 않는다** — 국소적이라고 해서 품질 게이트를 건너뛰지 않는다
2. 구조적 확장으로 판단되면: 참조 PPT 파일 경로를 확보한다. 사용자가 아직 제공하지 않았다면
   `AskUserQuestion`으로 요청한다(추측으로 진행하지 않는다 — 이 프로젝트는 "참조 PPT의 실제
   구조를 실측 확인"하는 것이 핵심 원칙이다)
3. `_workspace/` 생성(초기 실행 시), 참조 PPT 경로/미사유형명을 `_workspace/00_request.md`에
   기록

### Phase 2: 팀 구성 (구조적 확장인 경우만)

```
TeamCreate(
  team_name: "mass-ppt-dev-team",
  members: [
    { name: "mass-template-architect", agent_type: "mass-template-architect", model: "opus",
      prompt: "<대상 미사유형/성당명>, 참조 PPT 경로: <경로>. mass-template-analysis 스킬로 조사 후 _workspace/01_architect_design.md 작성" },
    { name: "ppt-ooxml-specialist", agent_type: "ppt-ooxml-specialist", model: "opus",
      prompt: "architect의 설계서(_workspace/01_architect_design.md)가 완성되면 tdd-progression-testing 스킬로 항목별 red→green→refactor 진행. ooxml-pitfalls 스킬 필독. 완료 후 ooxml-code-reviewer에게 리뷰 요청" },
    { name: "ooxml-code-reviewer", agent_type: "ooxml-code-reviewer", model: "opus",
      prompt: "specialist의 리뷰 요청을 받으면 _workspace/02_specialist_impl_notes.md보다 diff와 _workspace/01_architect_design.md를 먼저 읽고 독립적으로 판단할 것. ooxml-pitfalls 체크리스트를 감사 렌즈로 적용" },
    { name: "regression-qa", agent_type: "regression-qa", model: "opus",
      prompt: "리뷰 통과 통지를 받으면 boundary-verification 스킬로 검증하고 pytest 전체(회귀+진행중 프로그레션) 실행 후 안정된 프로그레션 테스트를 MASS_CASES로 승격, _workspace/03_qa_report.md 작성" }
  ]
)
```

```
TaskCreate(tasks: [
  { title: "템플릿 조사·설계", description: "...", assignee: "mass-template-architect" },
  { title: "TDD 구현", description: "...", assignee: "ppt-ooxml-specialist", depends_on: ["템플릿 조사·설계"] },
  { title: "독립 코드 리뷰", description: "...", assignee: "ooxml-code-reviewer", depends_on: ["TDD 구현"] },
  { title: "회귀 검증·승격", description: "...", assignee: "regression-qa", depends_on: ["독립 코드 리뷰"] }
])
```

> 팀원 4명, 작업 4개 — "소규모(5~10개 작업): 2~3명" 기준보다 약간 크지만, 설계/구현/리뷰/검증이
> 요구하는 전문성이 서로 명확히 달라 통합보다 분리가 낫다고 판단했다(에이전트 분리 기준: 전문성).

### Phase 3: 설계 (architect)

**실행 방식**: architect 단독 작업. 완료 시 리더와 specialist에게 SendMessage 통지.

### Phase 4: TDD 구현 (specialist)

**실행 방식**: specialist가 `_workspace/01_architect_design.md`의 "6. 테스트 대상 행동 목록"
항목마다 `tdd-progression-testing` 스킬의 red→green→refactor 사이클을 반복한다. 설계 결함을
발견하면 architect에게 SendMessage로 재설계 요청(Phase 3로 되돌아간 것으로 간주, 리더가
TaskUpdate로 반영). 모든 항목이 green이 되면 ooxml-code-reviewer에게 리뷰 요청.

### Phase 5: 독립 코드 리뷰 (ooxml-code-reviewer)

**실행 방식**: reviewer가 diff + 설계 스펙을 specialist의 자체 설명보다 먼저 읽고 독립적으로
판단(에이전트 정의의 "컨텍스트 격리 원칙" 참고). 발견 사항을 specialist에게 SendMessage로
전달(파일:라인 + 실패 시나리오). specialist 수정 → 재리뷰. **최대 2회 왕복**, 그래도 이견이
남으면 reviewer가 리더에게 에스컬레이션(심각도 판단을 리더가 대신하거나 사용자에게 확인).
리뷰 통과 시 regression-qa에게 SendMessage 통지.

### Phase 6: 회귀 검증 및 프로그레션 테스트 승격 (regression-qa)

**실행 방식**: qa가 `boundary-verification` 스킬로 경계면 교차 검증 + 전체 pytest 실행. 안정된
`test_missa_progression.py` 항목을 `test_missa_regression.py`의 `MASS_CASES`/테스트 클래스로
옮기고 원본에서 제거(졸업). 실패 발견 시 specialist에게 SendMessage로 재작업 요청(**최대 2회
재시도**, 초과 시 리더에게 에스컬레이션).

### Phase 7: 문서 동기화

1. `_workspace/03_qa_report.md`가 "통과"면 ppt-ooxml-specialist(경량 경로면 리더 본인)가
   `docs-sync` 스킬로 `docs/missa_to_ppt 요구사항.md`/`docs/missa_to_ppt 구현 계획.md`/
   (필요 시) `CLAUDE.md`를 갱신
2. 두 문서의 "부록: 변경 이력"에 오늘 날짜로 항목 추가됐는지 확인

### Phase 8: 정리 및 보고

1. 팀 모드였다면: 팀원들에게 종료 요청(SendMessage) → `TeamDelete`
2. `_workspace/` 보존(중간 산출물 삭제하지 않음 — 사후 검증·감사 추적용)
3. 사용자에게 요약 보고: 무엇을 구현했는지, 리뷰에서 발견/수정된 사항, 회귀 테스트 결과,
   승격된 프로그레션 테스트 목록, 갱신된 문서 목록
4. **커밋/푸시는 사용자가 명시적으로 요청할 때만** 수행한다(git 안전 정책, 이 하네스의 어떤
   에이전트도 자동으로 커밋하지 않는다)

## 데이터 흐름

```
[리더] → Phase1 triage → (경량) specialist→reviewer 순차 서브 호출
                        → (구조적) TeamCreate
                                     │
                architect ──SendMessage/설계완료──> specialist
                                                        │
                                    (TDD 사이클: red→green→refactor 반복)
                                                        │
                                              SendMessage/리뷰요청
                                                        ↓
                                                    reviewer
                                                        │
                                    이견 시 SendMessage 왕복(최대 2회) ──┐
                                                        │              │
                                              SendMessage/리뷰통과      └→ specialist 재작업
                                                        ↓
                                                       qa
                                                        │
                                    실패 시 SendMessage 왕복(최대 2회) → specialist 재작업
                                                        ↓
                                          _workspace/03_qa_report.md (+ 승격 결과)
                                                        ↓
                                              [리더] docs-sync 확인 → 보고
```

## 에러 핸들링

| 상황 | 전략 |
|------|------|
| architect가 참조 PPT를 찾지 못함 | 추측 금지, 리더가 AskUserQuestion으로 사용자에게 재확인 |
| specialist가 테스트보다 구현을 먼저 씀(TDD 위반 발견) | 리더 또는 reviewer가 지적, 테스트를 먼저 쓰고 실패를 확인한 뒤 재작업 |
| specialist 구현이 기존 회귀(20260712/20260624/20260705)를 깨뜨림 | 롤백 후 원인 분석, 새 기능이 기존 동작을 희생해선 안 됨 |
| reviewer와 specialist가 심각도에 이견(2회 왕복 후에도) | 리더에게 에스컬레이션, 필요 시 사용자에게 판단 요청 |
| qa 2회 재시도 후에도 실패 | 리더에게 에스컬레이션, 사용자에게 미해결 항목 명시하고 진행 여부 확인 |
| PowerPoint COM 사용 불가 환경 | COM 전용 검증만 "미검증(환경 제약)"으로 리포트, 실패로 취급하지 않음 |
| 팀원 유휴/중단 | 리더가 SendMessage로 상태 확인 후 재시작 또는 작업 재할당 |

## 테스트 시나리오

### 정상 흐름 (구조적 확장)
1. 사용자: "청년미사 지원 추가해줘, 참조 PPT는 청년미사/180726.pptx야"
2. Phase 1: 구조적 확장으로 판단, `_workspace/` 생성
3. Phase 2: 4명 팀 구성, 4개 작업 등록(의존관계 포함)
4. Phase 3: architect가 템플릿 조사 후 설계서 작성(테스트 대상 행동 목록 포함)
5. Phase 4: specialist가 행동 목록 항목마다 red→green→refactor 반복, 전부 green
6. Phase 5: reviewer가 diff를 독립적으로 리뷰, 사소한 지적 1건 → specialist 수정 → 재리뷰 통과
7. Phase 6: qa가 경계면 검증 + 전체 pytest 통과 확인, 안정된 프로그레션 테스트 3건을 MASS_CASES로 승격
8. Phase 7: 문서 동기화 완료
9. Phase 8: 팀 정리, 사용자에게 요약 보고(커밋은 하지 않음)

### 정상 흐름 (경량 경로)
1. 사용자: "제1독서 줄 수 계산에서 오프바이원 버그 고쳐줘"
2. Phase 1: 국소적 버그 수정으로 판단, 팀 미구성
3. ppt-ooxml-specialist를 서브 에이전트로 호출 — 버그를 재현하는 실패 테스트를 먼저 작성 후 수정
4. ooxml-code-reviewer를 서브 에이전트로 호출해 독립 리뷰
5. 리더가 boundary-verification 절차 직접 수행(또는 qa 1회 서브 호출)
6. Phase 7: 문서 동기화(영향 있는 경우만)
7. 사용자에게 보고

### 에러 흐름
1. Phase 5에서 reviewer가 "specialist의 새 함수가 run 개수를 2개로 가정해 CLAUDE.md 함정을
   재발시킴" 발견
2. specialist에게 SendMessage로 파일:라인 + 재현 시나리오 전달
3. specialist 수정 → 재리뷰 통과 → qa에게 통지
4. Phase 6에서 qa가 pytest 실행 중 기존 회귀 케이스(20260624) 실패 발견
5. specialist에게 SendMessage로 재작업 요청 → 수정 → 재검증 → 여전히 실패
6. 2회 재시도 소진 → qa가 리더에게 에스컬레이션
7. 리더가 사용자에게 "회귀 실패가 해결 안 됨, architect 재설계 필요할 수 있음"으로 보고하고
   계속 여부 확인
