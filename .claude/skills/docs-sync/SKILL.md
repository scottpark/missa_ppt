---
name: docs-sync
description: "missa_ppt에서 코드 구현 완료 후 docs/SPECS/ · docs/ARCHITECTURE.md · docs/IMPLEMENTATION_PLAN.md · CLAUDE.md를 실제 코드와 일치시키는 스킬(새 4단계 문서 체계: PRD·SPECS·ARCHITECTURE·IMPLEMENTATION_PLAN). 기능 구현 완료 직후, 커밋 전, '문서 갱신해줘'·'문서 확인해줘' 요청 시 사용. 요구사항 변경으로 문서를 먼저 고치는 절차(승인 게이트)는 mass-ppt-dev 스킬의 Phase A가 담당하고, 이 스킬은 구현 후 정합성 확인만 한다."
---

# Docs Sync — 구현 후 문서-코드 정합성 확인

문서 체계(2026-10-06~): `docs/PRD.md`(한 장) · `docs/SPECS/SPEC_*.md`(기능별) · `docs/ARCHITECTURE.md`
(구조·기술 사양) · `docs/IMPLEMENTATION_PLAN.md`(진행 중 작업, 임시). 버전 번호 없이 항상 최신 상태를 유지하며,
변경 이력은 git 커밋 이력이 대신한다(옛 "부록: 변경 이력"과 `archive/legacy-2026-10-06/`의 변경이력 파일은 과거 기록).

**원칙: 요구사항 변경은 코드보다 문서가 먼저**(PRD/SPEC → ARCHITECTURE → PLAN, 각 사용자 승인; `mass-ppt-dev` Phase A).
이 스킬은 승인된 문서대로 구현된 뒤 **구현 중 생긴 차이**를 정합시키는 단계다. 승인된 SPEC과 코드가 어긋나면
조용히 SPEC을 고치지 말고 사용자에게 알리고 문서 변경 절차를 밟는다.

## 절차
1. **변경 정리**: 추가/삭제/시그니처 변경 함수, 새 config.json 키·CLI 인수, 달라진 기존 동작.
2. **함수 목록**: `bash .claude/skills/docs-sync/scripts/list_functions.sh` 출력과 `docs/ARCHITECTURE.md`의
   함수 목록 절을 비교해 빠진 함수를 적절한 카테고리에 추가.
3. **서술 갱신** (문서와 코드가 다르면 코드를 `grep`/`Read`로 확인한 뒤 고친다):
   - `docs/ARCHITECTURE.md` — 구현 관점(파이프라인·함수·데이터 모델)
   - `docs/SPECS/SPEC_*.md` — 구현 중 확정된 세부 동작·인수 기준(승인된 범위를 넘는 변경은 사용자 확인)
   - `docs/IMPLEMENTATION_PLAN.md` — 완료 항목 체크, 작업 종료 시 비우거나 `archive/`로 이동
   - `CLAUDE.md` — 새 OOXML 함정은 규칙 한 줄+날짜만, 전문은 `docs/ooxml-pitfalls-log.md`에 같은 작업 단위로 기록
4. 하네스 구조(에이전트·스킬·모델·하네스 규칙)가 바뀐 경우 CLAUDE.md "하네스" 절 갱신과 함께 `docs/missa_to_ppt-module-notes.md` 마지막 절 "하네스(mass-ppt-dev) 변경 이력" 표에 행을 **반드시** 추가한다.

## 하지 않는 것
- 버전 번호 파일 생성, 문서 끝 변경 이력 부록 추가(git 이력으로 대체).
- 승인 없이 PRD 변경.
