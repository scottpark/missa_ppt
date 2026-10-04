# missa_to_ppt 구현 계획 변경 이력 (상세)

`docs/missa_to_ppt 구현 계획.md`(2026-10-01 성인+청년 통합 이후 항상 최신 상태로 유지되는
본문)의 날짜순 변경 이력을 여기로 분리했다 — 본문은 "지금 어떻게 동작하는가"만 담고, "언제·
왜·어떤 버그를 거쳐 지금 형태가 됐는지"의 상세 경위는 이 파일에서 찾는다. 2026-10-01 통합 전
원문(성인 구현 계획, 청년미사 2단계 구현 계획)은 `docs/archive/`에 보관되어 있다.

## 성인미사 구현 계획 — 부록 A (v1.0~2026-09-27)

| 시점 | 주요 변경 | 원문 |
|---|---|---|
| v1.0 | 최초 구현. 주일미사 전용 CLI. 독서/복음 파싱·배분, 슬라이드 복사 유틸, 서식 보존 함수, PPT2007 호환성 처리의 기본 골격 확립 | `docs/archive/missa_to_ppt 구현 계획 v1.0.md` |
| v1.1 | 대화형 팝업 모드, `config.json` 기반 OneDrive 성가 폴더 관리, 종료 슬라이드 통합(`merge_threshold`)과 동적 위치 재조정 추가 | `docs/archive/missa_to_ppt 구현 계획 v1.1.md` |
| v1.2 | 평일미사 지원 추가 — `is_sunday_mass()` 날짜 기반 자동 판단, UI 재배열, 화답송/성가/미사 후 기도 파라미터화 | `docs/archive/missa_to_ppt 구현 계획 v1.2 (평일미사 추가).md` |
| v1.3 | v1.2까지 정의됐던 동작이 실제로는 지켜지지 않던 버그 6건 수정, 회귀 테스트 스위트 신설 | `docs/archive/missa_to_ppt 구현 계획 v1.3.md` |
| v1.4 (2026-08-04) | 절 번호 서식 유실 버그 수정 + `_missing_orange_verse_numbers()` 검증 추가. 성가 악보 파일 필수 검증. 슬라이드 크기 차이 스케일 보정. PowerPoint COM 실측 검증 하이브리드 신설(`ppt_com_verify.py`, `_split_and_adjust_via_com` 등). "슬라이드별 캡이 거짓 성공을 보고하는 버그"와 "give-up 롤백 시 단락 순서가 뒤바뀌는 버그" 동시 수정 | (병합됨) |
| 2026-08-25 | 문서 버전 번호 매기기 중단. 성가번호 필수/형식 검증. 평일미사 OneDrive 조회 생략. 스케일 보정에 폰트/여백 포함. 성가 헤더 라벨 자동 교정 | (이 문서) |
| 2026-08-28 | 화답송 악보 이미지 입력 1차 마일스톤 — `missa_psalm_score_image.py` 신규(워터마크 크롭, 오선 밴드, 마디선 탐지, 무손실 분할, 고정 템플릿 배치) | (이 문서) |
| 2026-08-29 | 화답송 악보 이미지 입력 2차 마일스톤(파이프라인 배선) 완료·승인. 마디선 오검출 방지 필터, 좌/우 정렬+공통 배율 채택 | (이 문서) |
| 2026-09-02 | 종료 텍스트박스 겹침 버그 수정(`_content_line_height_emu()` 폰트 실측). `missa_to_ppt.py`(7,342줄) 모듈 분리 리팩토링(5개 모듈로 분리, 순수 이동·동작 무변경), 죽은 코드 삭제, 골든 마스터 diff 도구 도입 | (이 문서) |
| 2026-09-02 | 성가 헤더 라벨 재작성 시 run 색상 오염 버그 수정 — `_update_prefix_in_runs()` 시그니처를 세그먼트 기반으로 전환 | (이 문서) |
| 2026-09-09 | 공지사항 슬라이드 삽입 기능 추가(`insert_공지사항()`) | (이 문서) |
| 2026-09-10 | 청년미사 지원 1단계 착수 — `missa_youth_gospel.py`/`missa_youth_hymn_pdf.py` 신규 leaf 모듈. 회귀 승격은 사용자 지시로 생략 | 이 문서 + 청년 1단계 요구사항 |
| 2026-09-11 | 청년미사 1단계 후속 v2(6개 항목) — 미사 전 섹션 확장, 헤더 자동 축소, 저작권 크롭 일반화, 동적 패킹, 우측 여백 버그 수정 | 이 문서 + 청년 1단계 요구사항 |
| 2026-09-13 | 공지사항 뒤 구분 슬라이드 색상 불일치 버그 수정(`showMasterSp` 속성도 복제 시 이전) | (이 문서) |
| 2026-09-17 | 청년미사 2단계 Track A(§18) — 공유 함수에 `mass_type` 분기 추가. 회귀 25개 승격(62→87) | 이 문서 §18 + 청년 2단계 구현 계획 |
| 2026-09-17 | 청년미사 2단계 main() 배선 완료 — 날짜 이중 처리, 청년 전용 저장 파일명, 방어 게이트. 회귀 9개 승격(87→96) | 이 문서 §18.7~18.8 + 청년 2단계 구현 계획 |
| 2026-09-18 | 독서/복음 post-write 재조정이 비-마지막 슬라이드를 대량 미달인 채 방치하는 버그 수정 — best-so-far 추적 도입, `_try_absorb_underfull()` 헬퍼로 통합 | (이 문서) + `docs/ooxml-pitfalls-log.md` |
| 2026-09-24 | 청년미사 확장 B/A/C/D 4개 그룹. 성인 경로 공유 함수 2곳도 수정(`duplicate_slide()` 레이아웃, `_adjust_fit_if_needed()` 안전 마진). 독립 리뷰 2라운드(1건 재작업). 프로그레션 40개 승격, 전체 스위트 322개 중 319 passed·3 deselected | (이 문서) + 청년 2단계 구현 계획 |
| 2026-09-26 (E그룹) | 실사용자 육안 검수 후속 5건(E1~E5) 중 4건 반영, 전부 청년 전용 | 청년 2단계 구현 계획 |
| 2026-09-26 (F그룹) | 청년미사 출력 폴더 정합성 버그 3건 + OneDrive remoteItem 주소 해석 버그 수정 | (이 문서) + 청년 2단계 구현 계획 |
| 2026-09-26 (G그룹) | 영문 복음 `_PIL_WRAP_SAFETY` 반대 방향 버그(G1), OneDrive 앱 인스턴스 캐시화(G2) | (이 문서) + `docs/ooxml-pitfalls-log.md` |
| 2026-09-26 (J그룹) | CLI/GUI/배포 개선 7배치(J1~J10) — 미사 유형 선택 팝업 삭제·`--미사유형` CLI 플래그로 대체(§1 전면 재작성, 리뷰 라운드11이 '어린이' 가드 우회 버그 발견·수정), OneDrive 브라우저 단순화, PowerPoint 백그라운드 안내, `VERSION`/업데이트 확인 링크 | (이 문서) + CLAUDE.md |
| 2026-09-27 (K그룹) | OneDrive 브라우저 재귀 프리페치(K1) + 로컬 캐시 재사용(K2). 리뷰 라운드17이 캐시 계층 혼용 버그 발견·수정 | (이 문서 §16) + CLAUDE.md |
| 2026-09-27 (L그룹) | 프리페치 조기 트리거(L1) + 다운로드 즉시반환·백그라운드화(L2). 리뷰 라운드19가 "마지막 선택만 유효" 디스크 수준 위반 발견·수정 | (이 문서 §16) + CLAUDE.md |

## 청년미사 2단계 구현 계획 — 부록 (2026-09-16~27)

| 날짜 | 변경 내용 |
|---|---|
| 2026-09-16 | 최초 작성. 설계 산출물 병합, `is_sunday` 4축 분리·업데이트 메커니즘·배포 형태 확정 |
| 2026-09-17 | Track A(§5~§10) TDD 구현·리뷰 2라운드·QA 완료. §8.3 호출부 치환표 부호 반전 오류를 구현 단계에서 발견·수정 |
| 2026-09-17 | main() 통합 배선 완료 — 팝업 취소 시 무경고 저장 버그 발견·수정(두 팝업 모두 `RuntimeError` + 이중 게이트) |
| 2026-09-18 | `missa_onedrive.py` 모듈 작성 완료(리더 직접 구현). 로컬 우선→Graph 미러링→수동 선택 3단 폴백으로 전환. `ppt_com_verify.py` COM 창 플리커 수정 |
| 2026-09-18 | Graph 미러링 캐시 역방향 업로드 구현. 구현 중 독립 판정 로직이 병행 테스트와 어긋나 MSAL 로그인 대기로 멈추는 사고 발견·즉시 수정 |
| 2026-09-24 | B/A/C/D 4개 그룹. 리뷰 1라운드에서 D2(페이지 정규화 범위 과다)·C1(나주노 151 bbox) 재작업 지적 → D2 opt-in화, C1 재반박(원문 대조) → 2라운드 통과 |
| 2026-09-26 (E그룹) | E1(OneDrive 자동 업로드 신규)·E2(제목 날짜 버그)·E4(복음환호송 run 오조립)·E5(영문 정렬 복구) 반영. E3(화살표 삭제)는 원인 조사 후 철회 |
| 2026-09-26 (F그룹) | F1(한글 JSON 저장 위치)·F2(영문 JSON 미저장) 버그 수정 + remoteItem 주소 해석 버그 근본 원인 규명·수정, 실제 Graph API 왕복으로 업로드 성공 확인 |
| 2026-09-26 (G그룹) | G1(`_PIL_WRAP_SAFETY` 반대 방향 버그, 실사용자 스크린샷으로 재검증 요청받아 발견) + G2(MSAL 앱 인스턴스 매번 생성 — 3~6초 지연 원인) |
| 2026-09-26 (H그룹) | 폴더 확장 화살표 단일 클릭 무반응 버그 — `<Double-1>`만 바인딩된 것을 `<<TreeviewOpen>>` 가상 이벤트로 통일해 수정 |
| 2026-09-26 (I그룹) | OneDrive 선택 후 경로 표시가 46자 폭에 안 맞아 파일명이 잘리던 문제 — 표시 문자열과 실제 처리값을 분리(CLAUDE.md 신규 원칙 도출) |
| 2026-09-27 (K그룹) | K1(재귀 프리페치, 사전 실측 기반 병렬화 안 함 결정) + K2(로컬 캐시 재사용). 리뷰 라운드17이 K1/K2 캐시 시너지가 핵심 불변조건("OneDrive가 더 최신이면 무조건 재다운로드")을 깨는 경로를 재현해 발견 → 신선도 판단을 항상 실시간 조회로 전환 |
| 2026-09-27 (L그룹) | L1(프리페치 조기 트리거) + L2(파일 선택 즉시 반환 + 백그라운드 다운로드). 리뷰 라운드19가 서브프로세스 Tk 프로브로 "같은 파일명 재선택 시 구세대가 최신 파일을 조용히 덮어씀" 버그를 재현해 발견 → blob 최신화와 목적지 커밋을 분리, 커밋을 락 안 generation 재확인 후로 이동 |

## 2026-10-01 — 성인+청년 구현 계획 문서 통합 (이번 변경)

성인미사 구현 계획(`missa_to_ppt 구현 계획.md`, §1~§17 + 부록 A/B)과 청년미사 2단계 구현
계획(별도 문서, §1~§13 + 부록)을 한 문서로 병합하고, 공통/성인 전용/청년 전용 3부로
재구성했다. 이 통합은 문서 재구성·압축이며 코드 변경은 없다.

**현재 코드와 대조해 발견한 드리프트(수정 완료)**: 위 표들이 서술하는 §2.2/§2.4/§5.2/§6.2
"나주노·야훼이레/가톨릭성가는 `get_onedrive_hymn_folder()`/`get_onedrive_youth_hymn_root()`로
**폴더 전체**를 Graph API 미러링한다"는 설명은 2026-09-27 이후 더 이상 사실이 아니다.
`missa_gui.find_youth_onedrive_hymn_file()`/`_find_onedrive_file()`/`_youth_onedrive_local_
sync_dir()`(신규)와 `missa_content_updaters.resolve_youth_hymn_pptx()`의 실제 코드를 직접
읽어 확인한 결과, 세 출처 모두 이제 **곡 하나만 온디맨드로 찾는 방식**(로컬 캐시 →
`_find_onedrive_file()` 재귀 검색)을 쓴다(사용자 관찰 — 성가 5곡만 필요한데 가톨릭성가 폴더
34곡이 전부 캐시로 내려오는 것이 낭비였음). 이번 통합 문서(§3.5(B))는 이 현재 상태를
반영했다. `get_onedrive_hymn_folder()`/`get_onedrive_youth_hymn_root()`(전체 미러링) 자체는
코드에서 삭제되지 않고 남아 있으며, 성인미사 가톨릭성가 조회와 §3.9의 캐시 역업로드 판단
(`get_onedrive_youth_hymn_remote_path()`)용으로만 계속 쓰인다. OneDrive 개별 파일 브라우저
(§3.5(A), K/L그룹)는 이 전환과 무관한 별개 메커니즘으로 그대로 유지된다.

원문은 `docs/archive/`로 이동: `missa_to_ppt 구현 계획 (2026-10-01 통합 전 원문).md`은 남기지
않음(성인 구현 계획은 새 문서가 그 파일 자체를 대체), `청년미사 2단계 구현 계획 (배포 전환·
OneDrive 인증·전체 파이프라인 통합) (2026-10-01 통합 전 원문).md`.

부록 B(재발 방지 규칙 요약, `CLAUDE.md` 상세)는 그대로 `CLAUDE.md`를 단일 출처로 유지하고
이 문서에서는 중복하지 않는다.

## 2026-10-02 — 청년미사 접속 모드(로컬 폴더 우선/로그인 폴백) 설계·구현 완료

`docs/missa_to_ppt 구현 계획.md` §3.4.1 신설: 모드 판정 단일 지점(`get_youth_ppt_mode()` +
`_YOUTH_PPT_FOLDER`), 변경 대상 코드 전수 조사 표(`_run_gui_mode`, `_ask_combined_input_popup`,
`find_youth_onedrive_hymn_file`, `resolve_youth_hymn_pptx`, `_upload_youth_output_to_onedrive`),
로컬 모드에서 `missa_onedrive`/`msal` 미import 제약, 기존 테스트 모킹 갱신(규칙 20) 계획.
사용자 리뷰 후 같은 날 구현했다(`missa_gui.py`·`missa_to_ppt.py`·`missa_content_updaters.py`, 신규 테스트 14개).

## 2026-10-03 — `_prevent_widow_tails()` 추가 (post-write 재조정 끝)

원인: `_split_and_adjust_via_com()`이 9줄 달성만 성공 조건으로 삼고 꼬리 줄 길이를 보지 않음. 수정: 모든
스윕·병합 후 1회 widow 방지(COM 줄 텍스트 `measure_line_texts()`, 오버플로 롤백, 후행 공백 rstrip 통일).
`_split_para_at_lines()`에서 `_wrap_word_buckets()`/`_split_runs_at()`를 동작 불변으로 추출. 테스트는
`tests/test_missa_progression.py`의 W그룹(`test_w0`~`test_w11`, 14개). CLAUDE.md 규칙 25 추가.

## 2026-10-03 — 슬라이드 단위 채우기로 독서·복음 배분 구조 교체 (위 `_prevent_widow_tails` 항목을 대체)

`replace_reading_slides(prs, s, e, units, template_idx, ...)`가 units를 받아 슬라이드마다 COM 실측
(`ppt_com_verify.measure_line_starts`)으로 정확히 9줄을 확정(`_measure_slide_cut`, 폴백 `_pil_line_starts`).
제거: `layout_units_on_slides(_pil)`, `_verify_and_rebalance_pages`, `_rebalance_reading_slides_post_write`,
`_try_absorb_underfull`, `_split_and_adjust_via_com`, `_split_para_at_lines`, `_prevent_widow_tails`,
`_count_slide_lines(_verified/_rendered)`, `_rendered_wrap_count`, `_PIL_WRAP_SAFETY` 등. 종료 병합 판정을
마지막 슬라이드의 실측 줄 수로 변경. 실측(실제 PowerPoint): 성인 20261004 복음 줄 수 9,9,8,6→9,9,9,5
(9.6초/COM 14회 → 6.1초/4회), 청년 20260913 영문 복음 9,9,8,9,4→9,9,9,9,4(11.3초/19회 → 7.3초/5회).
테스트는 SF그룹(가짜 COM + 실제 COM)으로 이전·재작성(상세 표는
`_workspace/refactor_슬라이드단위_채우기/impl_notes.md`). CLAUDE.md 규칙 25를 이 구조 규칙으로 교체.

## 2026-10-03 — 'PPT 문서' 루트·찾아보기 시작 폴더·결과창 복사 버튼 구현 (§3.4.2)

`missa_gui.py`에 성인 루트(`get_adult_ppt_folder`)·`get_ppt_root`·시작 폴더 해석(`browse_initial_dir` 등)·화답송 폴더
기억·결과 복사(`copy_result_to_onedrive`, `_copy_result_with_ui`) 추가, `_show_result_window`에 버튼 추가,
`_run_with_progress_window`에 5초 안내 라벨 추가, Graph 브라우저 `start_subpath`. `missa_to_ppt.py`에서
`_upload_youth_output_to_onedrive`·`[8.5]` 호출·`_show_powerpoint_background_notice` 삭제, `_last_output_mass` 기록,
성인 루트 선택 호출 추가. 테스트: `test_missa_youth_local_mode.py` 확장, `j7a~c`·`e1` 갱신, e_full_output 픽스처의
스텁 복구 시점을 yield 이전으로 이동(k2g/k2h·g2 순서 오염 해소), conftest에 성인 상태 격리.

## 2026-10-04 — 회귀 픽스처 결과 PPT·log를 git 추적에서 제외, 테스트는 임시 폴더에 생성

`output/{20260624,20260705,20260712}`는 입력 파일만 추적(`.gitignore`에 결과 PPT·`log/` 패턴 추가, 기존 추적분은
`git rm --cached`). 신규 `tests/_gen_helper.py`가 입력만 임시 `output/{date}/`로 복사하고 cwd를 임시 폴더로 해
`missa_to_ppt.py`를 실행한다. 회귀 `_generate`/`_find_output`과 공지사항 삽입 테스트의 `_FIXTURES`가 이를 사용.


## 2026-10-04 — 청년미사 성가 교체 성능 개선 + 결과창 버튼 위치

실측(템플릿 127장·성가 5곡, 같은 조건): `replace_성가_youth` 25.4초 → 약 4초. 원인은 (1) `find_sections(youth)`가 호출당 약 1초
(성가 5종 × 슬라이드 전체를 htype마다 따로 훑으며 python-pptx 도형 객체를 반복 생성)인데 곡마다·htype마다 10번 호출, (2) 이미
만든 나주노/야훼이레 성가 파일이 있어도 제목을 PDF에서 조회(야훼 이레는 574쪽 `search_for` ≈ 2.5초). 조치: `_scan_youth_slides()`로
슬라이드당 1회 스캔을 5종이 공유, 남는 슬롯이 있을 때만 재탐색, `입당송`/`영성체송` 탐색은 이미 추출한 `texts` 재사용, 제목은
입력값 → 기존 파일명 → PDF 순. 결과창 '원드라이브로 복사' 버튼은 '파일 열기' 바로 옆·같은 글꼴로 이동(§3.4.2).

## 2026-10-05 — 성인 소스 zip 배포 + 배포 도구 통합

`tools/build_dist_zip.py`를 유형별(`성인|청년|all`) 소스 zip 빌더로 재작성(exe 빌더였음). 신규 `run_missa_성인.bat`·성인
README, `create_shortcut.py`/`install.bat` 성인·청년 공용화(`run_missa_*.bat` 존재로 유형 판단). 삭제: `missa_to_ppt.spec`,
`tools/build_shortcuts.py`, `tools/build_youth_dist_zip.py`, `run_missa.bat`. README 2개를 `docs/`로 이동(빌더가 zip 최상위로
복사). 성인 zip(약 0.8MB)을 압축 해제해 평일 20260624 생성·검증 통과까지 단독 실행 확인.

## 2026-10-05 — 배치 파일 CRLF 보장 + README 자동 갱신

성인 zip의 `install.bat`이 LF만 있는 줄바꿈이라 한글+`chcp 65001`에서 cmd가 줄 경계를 잘못 읽어 오류('issa_ppt is not
recognized…')가 났다. `tools/build_dist_zip.py`의 `_write_file()`이 `.bat`을 항상 CRLF로 정규화해 zip에 넣고, 저장소에
`.gitattributes`(`*.bat text eol=crlf`)를 추가. `missa_updater.py`에 `_UPDATE_README_GLOB`(설치 폴더에 이미 있는 안내서만
갱신) 추가. 옛 exe 산출물(`dist/missa_to_ppt.exe`·`.zip`) 삭제. 테스트: `tests/test_build_dist_zip.py`(신규), updater README 테스트.
성인·청년 zip을 풀어 `install.bat`을 실제 실행해 오류 없음·유형별 바로가기 생성을 확인.

(같은 날 추가) `install.bat`의 안내 문구가 두 유형을 함께 보여주던 것을, 폴더의 `run_missa_*.bat` 존재로 유형을 판단해 해당 유형("성인미사 PPT"/"청년미사 PPT")만 표시하도록 변경. 실제 cmd 실행 테스트 추가.
