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
  `if __name__ == '__main__':`은 `_dispatch_cli_or_gui()`(2026-09-26) 하나만 부른다 — 미사
  유형 선택 팝업은 삭제됐고, `--미사유형`(기본 '성인', '어린이'는 미지원 안내 후 종료) CLI
  플래그로만 정해진다. date가 있으면 `main()`(CLI), 없으면 `_run_gui_mode()`(GUI 입력 흐름,
  처리 시작 직전 `_show_powerpoint_background_notice()`로 PowerPoint 백그라운드 실행을
  안내)로 간다. `_build_arg_parser()`가 인자 정의를 `parse_args()`와 공유한다(상세는
  `docs/missa_to_ppt 구현 계획.md` §1).
- `missa_ooxml_utils.py` — 슬라이드/도형 복사·rId 매핑·배경 상속·텍스트런 조작 등 OOXML 저수준
  프리미티브. 프로젝트 내부 의존성이 없는 leaf 모듈. `HYMN_TYPES`도 여기 위치(entry↔content_updaters
  순환 임포트 방지).
- `missa_gui.py` — Tkinter 팝업 전부 + `is_sunday_mass`/config.json 로드·저장/
  `get_onedrive_hymn_folder`. 이 함수들은 GUI 팝업과 진입점 양쪽에서 호출되어 entry에 그대로
  두면 순환 임포트가 생기므로 gui 모듈에 흡수했다(entry→gui 단방향). `OUTPUT_ROOT`도 entry·gui
  양쪽에서 쓰여 여기 위치. **2026-09-26**: 미사 유형 선택 팝업(`_ask_mass_type_popup`)과
  `config.json`의 `last_mass_type`은 삭제됐다(미사 유형은 이제 CLI `--미사유형`으로만 온다).
  `_on_check_update(parent)`(그 팝업의 클로저였던 업데이트 확인 로직)는 독립 함수로 남아
  `_ask_combined_input_popup()`의 확인 버튼 아래 새 행(`_get_app_version()`이 읽은 버전
  텍스트 + `tk.Label(cursor='hand2')` 링크, 기존 `tk.Button`에서 전환)에 재배선됐다.
  **2026-09-27(K1)**: `_prefetch_onedrive_children()`(§J5)을 재귀로 확장 — 3개 허용 루트
  폴더의 직계 자식뿐 아니라 모든 하위 폴더/파일까지 한 번에 캐시에 채운다(실측: 실제 트리
  API 호출 9회·하위폴더 6개·파일 31개·최대 깊이 2·순차 총 22초 — 이 규모면 병렬화 없이
  순차 재귀로 충분). `<<TreeviewOpen>>`이 거치는 `_cached_onedrive_children()`이 이미
  "캐시에 있으면 네트워크 스킵" 계약을 갖고 있어 재귀로 캐시를 더 깊이 채우는 것만으로
  하위 폴더 탐색도 캐시 히트가 된다. **2026-09-27(K2)**: OneDrive에서 받는 파일(공지사항/
  시작기도/미사후기도/성가 PPT 등, 파일 종류로 분기 없음)을 날짜 폴더와 별개인 공유 캐시
  폴더(`_ONEDRIVE_CACHE_ROOT = 'onedrive_cache'`)에 원본 1부로 보관해, 매니페스트
  (`onedrive_cache/manifest.json`)에 기록된 lastModifiedDateTime이 OneDrive 쪽과 **문자열
  완전일치**할 때만 재다운로드를 스킵한다(`_resolve_onedrive_download()`). 신선도 판단은
  K1의 프리페치 캐시와 무관하게 매번 `missa_onedrive.get_item_metadata()`로 실시간
  재확인한다 — 리뷰에서 K1 캐시 스냅샷을 실시간값으로 오인해 핵심 불변조건("OneDrive가 더
  최신이면 무조건 재다운로드")이 깨지는 버그가 발견돼(아래 "캐시 계층 간 신선도 혼용 금지"
  참고) K1 캐시는 트리 표시용으로만 분리했다. 상세: `_workspace/02_specialist_impl_notes.md`
  "K그룹" 절. **2026-09-27(L1)**: `_start_onedrive_prefetch_thread()` 신설(K1 재귀
  프리페치를 시작하던 인라인 스레드 코드를 공유 헬퍼로 추출) — `_ask_combined_input_popup()`
  이 창을 띄우는 시점(청년미사만)으로 트리거를 앞당겨, 로그인/네트워크 상태와 무관하게 첫
  입력창이 항상 즉시 뜨도록 보장한다(스레드 시작 자체가 논블로킹). `_ask_onedrive_file_
  browser_popup()`도 같은 헬퍼를 호출해 2차 안전망을 유지. **2026-09-27(L2)**: OneDrive
  파일 선택(`_on_choose()`)이 다운로드를 기다리지 않고 결정적 목적지 경로만 계산해 즉시
  반환하고, 실제 다운로드는 generation 추적 백그라운드 스레드(`_download_worker`)로
  넘긴다 — `_ensure_onedrive_blob_fresh()`(공유 캐시 blob만 최신화, dest 커밋 없음)와
  `_commit_onedrive_blob_to_dest()`(순수 복사)로 분리해, 같은 슬롯을 같은 파일명으로
  재선택해도(이 앱의 정상 관례 — 시작기도/미사후기도/공지사항 등이 매주 다른 폴더에서
  같은 이름으로 반복됨) 구세대(느린) 다운로드가 최신 파일을 덮어쓰지 못하게 한다 — 커밋은
  `_download_worker`가 `_download_lock` 안에서 generation을 재확인한 경우에만 수행한다.
  최초 구현은 이 재확인이 메모리 상태만 보호하고 디스크 커밋은 무조건 수행해, 리뷰
  라운드19에서 "확인 성공 선언 후 최신 파일이 조용히 구식 내용으로 역전"되는 버그가
  서브프로세스 Tk 프로브로 직접 재현·확정됐다 — "generation을 재확인했다"는 사실과 "그
  재확인이 실제로 보호하는 대상이 최종 결과물(디스크 파일)인가, 메모리 상태뿐인가"는 다른
  질문이었다. 커밋을 락 안·generation 일치 후로 옮겨 수정, 라운드20 재리뷰(코드
  되돌리기로 직접 재현 후 재통과 확인) 통과. 상세: `_workspace/02_specialist_impl_notes.md`
  "L그룹"/"L그룹 리뷰 라운드19 재작업" 절.
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
  `"복음"`)은 하위호환 래퍼로 유지. `missa_to_json.py`는 여전히 import하지 않지만,
  **`missa_to_ppt.py`에는 배선 완료**(청년미사 §D1) — `main()`이 `mass_type == 'youth'`일 때
  `missa_youth_gospel.get_youth_mass(content_date_str)`를 호출하고
  `merge_youth_gospel_content()`로 기존 `json_data`와 병합한다.
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
  기본값으로 덮인다). `missa_to_ppt.py`를 직접 import하지는 않지만(단방향 leaf 유지),
  **`missa_content_updaters.py`의 `resolve_youth_hymn_pptx()`가 나주노/야훼 이레 캐시-미스
  경로에서 이 모듈을 호출**해 새 성가 PPT를 만들므로 파이프라인과 완전히 무접촉은 아니다.
  C1(151 오선 밴드 오검출)·C2(`duplicate_slide` 레이아웃)는 이 경로의 실사용 버그였고, 수정
  테스트(`test_c1_*`/`test_c2_*`)는 2026-09-24 `test_missa_regression.py`로 승격됐다. 그 외
  2026-09-11 v2 확장분(M/Hf/S/C/P/Hd 계열)은 아직 `test_missa_progression.py`에만 있다(회귀
  승격 미실시). **후속 조치(미착수):** `resolve_youth_hymn_pptx()`는 OneDrive에 같은 이름
  패턴(`'{출처} 성가 {번호} '`)의 기존 파일이 있으면 그대로 재사용하고 재생성하지 않는다 —
  이미 사용자 OneDrive에 캐시돼 있는 나주노 151/471 PPTX는 C1/C2 크롭·레이아웃 수정이
  자동 반영되지 않으므로, 다음에 그 곡이 쓰이기 전에 해당 파일을 수동 삭제해 재생성을 유도해야
  한다(상세: `_workspace/02_specialist_impl_notes.md`, `docs/청년미사 2단계 구현 계획.md`
  2026-09-24 항목).
- `missa_onedrive.py` — (청년미사 2단계) 성당 공용 계정(brokenbaykccppt@gmail.com) 하나로
  Microsoft Graph API(Device Code Flow)에 접근하는 완전 독립 leaf 모듈(다른 프로젝트 모듈을
  import하지 않는다 — `missa_gui.py`/`missa_content_updaters.py`/`missa_to_ppt.py`/
  `missa_youth_hymn_pdf.py`가 이 모듈을 호출하는 방향만 있다). `list_children`/`ensure_folder`/
  `upload_file`/`download_file`. **주의**: 이 계정의 OneDrive에서 최상위 경로 `PPT 문서`는
  진짜 폴더가 아니라 **다른 드라이브를 가리키는 공유 바로가기(remoteItem)**다 — 경로 기반
  주소(`root:/PPT 문서/...:`)는 이 바로가기를 통과하지 못해 하위 경로 조회가 404/422로
  실패한다. `_resolve_remote_base()`/`_ensure_base_resolved()`가 첫 경로 세그먼트의
  remoteItem 여부를 실제 GET으로 확인해 `_REMOTE_ITEM_CACHE`(프로세스 수명 캐시)에 저장하고,
  remoteItem이면 `_item_url()`이 `/drives/{driveId}/items/{itemId}/...` 주소로 전환한다
  (`_item_url()` 자체는 순수 함수로 유지 — 해석은 별도 함수가 담당). "경로 문자열만 올바르게
  이어붙이면 Graph가 알아서 찾아줄 것"이라고 가정하면 이 함정이 재발한다 — 최상위 경로가
  공유 바로가기인지 아닌지는 실제 API 응답(`remoteItem` 파셋 유무)으로만 알 수 있다.
  **발견 경위 (2026-09-26):** E1(청년미사 결과 PPT OneDrive 업로드)이 실제 라이브 Graph
  API에서 `404 Client Error`로 실패, 코디네이터가 직접 Graph API를 호출해 `PPT 문서`가
  remoteItem임을 실측 확인(상세: `_workspace/02_specialist_impl_notes.md` "F그룹 후속").
  **성능(2026-09-26, G2)**: `get_access_token()`이 호출마다 `msal.PublicClientApplication(...)`
  을 새로 생성하면 그 생성 자체가 약 0.9~1.0초 걸린다(`AUTHORITY`가 네트워크 기반이라
  생성자 안에서 매번 instance discovery를 재수행하는 것으로 추정) — OneDrive 커스텀 파일
  브라우저(`missa_gui._ask_onedrive_file_browser_popup`)처럼 폴더 하나 열 때마다 여러 HTTP
  요청(`_headers()`)이 겹치는 경로에서 이 오버헤드가 누적돼 체감 3~6초 지연으로 나타났다.
  `_APP_CACHE`(모듈 레벨 `[app, cache]`)와 `_get_app()`으로 프로세스 수명 동안 앱 인스턴스를
  1회만 생성해 재사용하도록 수정(MSAL 공식 권장 패턴) — `list_children`/`ensure_folder`/
  `upload_file`/`download_file`/`_resolve_remote_base` 등 `_headers()`를 거치는 모든 호출부가
  자동으로 이득을 본다. 실측: 앱 생성 단독 1.131초 → 캐시 재사용 0.000000초, 반복
  `list_children` 호출 6.4초(콜드) → 2.0~2.2초(재사용, 순수 네트워크 왕복만 남음). 상세:
  `_workspace/02_specialist_impl_notes.md` "G그룹" 절.
  **2026-09-27(K2)**: `get_item_metadata(path)` 신설 — children 없이 항목 메타데이터만
  GET(404→FileNotFoundError, `list_children`/`download_file`과 동일 에러 관례).
  `missa_gui._resolve_onedrive_download()`가 로컬 캐시 재사용 여부를 판단할 때
  `lastModifiedDateTime`/`eTag`를 실시간으로 확인하는 용도.
- `tests/test_missa_progression.py` — 아직 안정화되지 않은 신규 동작을 먼저 명세하는
  프로그레션 테스트(TDD red→green). 안정화되면 `tests/test_missa_regression.py`로 승격.
- `tests/test_missa_regression.py` — 주일/평일 통합 + 단위 회귀 테스트 (`pytest`로 실행).
- `conftest.py`(저장소 루트) — `tests/`에서 루트의 `missa_gui` 등을 import할 수 있도록
  `sys.path`에 루트를 추가(2026-10-01, 저장소 정리).
- `config.json` — `onedrive_hymn_folder`(악보 성가 PPT 경로).
- `missa_to_ppt.spec` + `dist/` — PyInstaller 빌드(`missa_to_ppt.exe`) + 배포 zip
  (`missa_to_ppt.zip`/`missa_ppt_청년미사.zip`) 산출물 보관 폴더(2026-10-01부터, 이전엔
  루트에 직접 흩어져 있었다).
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

## 같은 프레젠테이션 내부 복제는 "빈 레이아웃 이름 검색"이 아니라 원본 레이아웃을 그대로 재사용

`duplicate_slide()`처럼 **같은** 프레젠테이션 안에서 슬라이드를 복제할 때는
`src_slide.slide_layout`이 항상 유효하므로 그것을 그대로 재사용해야 한다. `_blank_layout(prs)`
류(레이아웃 목록을 순회해 이름이 'Blank'/'빈 화면'/'blank'인 것을 **순서대로** 찾아 첫 매칭을
반환하는 헬퍼)로 대체하면 안 된다 — 한 프레젠테이션에 이름이 다른 "빈" 레이아웃이 여러 개
공존할 때(예: 슬라이드 자신은 '빈 화면' 레이아웃을 쓰는데, `prs.slide_layouts` 목록에는
'Blank'가 '빈 화면'보다 먼저 나옴) 원본과 다른 레이아웃이 골라져, 같은 곡/섹션 안에서 복제된
슬라이드끼리 레이아웃이 갈라진다. `_blank_layout()`은 **다른** 프레젠테이션에서 복사해 올 때
(`copy_slide_from_prs()`처럼 소스에 매칭되는 레이아웃 이름이 타겟에 없는 경우)의 최후 폴백으로만
쓴다.

**발견 경위 (2026-09-24):** `duplicate_slide()`가 항상 `_blank_layout(prs)`를 썼다가, 나주노
151번 성가 슬라이드 14(원본, '빈 화면')와 15~17(복제본, 'Blank'로 갈라짐)의 배경/서식이
미세하게 달라 보이는 문제로 재현·발견(상세: `docs/ooxml-pitfalls-log.md`).

## 여러 곳에서 재사용하는 fit-보정 함수의 "추정 vs 실제" 안전 마진은 이미 실측된 보정치를 재사용한다

텍스트가 도형/슬라이드 경계를 넘는지 판정하는 함수(예: 줄 수 추정 → 필요 높이 계산 → 여유
공간과 비교)에서, 추정치와 실제 여유 공간의 차이가 근소(수 % 이내)한 경우는 추정 자체가
가진 체계적 오차(Pillow/char-count 근사가 실제 PowerPoint 렌더링보다 작게 나오는 경향, 아래
"한글 줄 수 계산" 항목 참고) 때문에 실제로는 넘치는데도 "안 넘친다"고 판정해 조정 로직이
트리거되지 않을 수 있다. 이럴 때 새 임의의 안전 마진 값을 발명하지 말고, 이미 이 프로젝트에서
실측으로 확정된 보정치(아래 "Pillow ascent+descent가 실제 PowerPoint 줄 높이보다 14~18%
작다")를 그대로 재사용해 추정 높이에 곱한다 — 서로 다른 함수라도 같은 근본 원인(폰트 메트릭
과소평가)에서 나온 오차이므로 같은 보정치가 적용된다.

**발견 경위 (2026-09-24):** `_adjust_fit_if_needed()`(화답송/복음환호송/영성체송 공용 겹침
자동 조정 함수)가 char-count 추정 높이와 실제 여유 공간의 차이가 2.7%뿐인 청년미사 영성체송
실사용 문구에서 조정을 트리거하지 못해 마지막 줄이 슬라이드 밖으로 나감. `_FIT_HEIGHT_SAFETY_
MARGIN=1.15`(아래 "한글 줄 수 계산" 항목의 14~18% 보정치 재사용)로 수정(상세:
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

## 텍스트 배치를 결정하는 단계(페이지네이션)와 검증하는 단계는 반드시 같은 박스 폭을 써야 한다

"어느 텍스트가 어느 슬라이드에 들어갈지" 정하는 페이지네이션과, 배치 후 "실제로 맞는지"
확인하는 검증(post-write 재조정·COM 실측)이 같은 `_get_slide_render_params()`의 box_px에서
출발해도, 한쪽에만 추가 안전 마진(보정 계수)을 곱하면 두 단계가 서로 다른 폭을 기준으로 계산
하게 된다. 검증 쪽이 더 넓은 폭을 쓰면, 페이지네이션은 실제로 들어가는 단어까지 안전 마진 때문에
다음 슬라이드로 밀어내는 **과잉 보수적** 오배치를 만든다 — 오버플로 검증("9줄 이내")에는 걸리지
않는 반대 방향 오차라 육안 검수로만 드러난다. 이런 보정 계수를 새로 넣거나 유지할 때는 "이
계수가 막던 문제가 다른 메커니즘(예: post-write 재조정 강화)으로 이미 해소되지 않았는지"를
재확인해야 한다 — 이미 해소됐다면 이 계수는 더 이상 아무것도 막지 못하면서 새 반대 방향 버그만
만드는 상태가 된다.

**발견 경위 (2026-09-26):** `layout_units_on_slides_pil()`의 `_PIL_WRAP_SAFETY=0.97`이
페이지네이션에만 곱해지고 검증 경로(`_rendered_wrap_count`/COM)에는 적용되지 않아, 실사용자
20260926 청년미사 슬라이드 54에서 실제로 들어가는 단어("father's")가 다음 슬라이드로 밀려남.
`_PIL_WRAP_SAFETY`가 원래 막던 반대 방향 문제(20260913 영문 복음 4번째 슬라이드, 2026-09-17
도입)는 2026-09-18 `_split_and_adjust_via_com()` best-so-far 강화로 이미 별도 해소돼 있었음을
COM 재검증으로 확인 후 `0.97`→`1.0`으로 무력화(상세: `docs/ooxml-pitfalls-log.md`,
`_workspace/02_specialist_impl_notes.md` "G그룹" 절).

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

## 재시도 루프는 "정확히 맞지 않으면 포기"가 아니라 최선의 non-overflow 결과를 보존한다

여러 번 재시도(keep ±1처럼 파라미터를 조금씩 조정)해서 정확한 목표값에 수렴시키려는 루프에서,
마지막 재시도가 실패(더 조정할 여지가 없어 `None` 반환 등)했다고 해서 **그 전에 이미 확보했던,
불변조건(예: 오버플로 금지)은 지키면서 목표에 가장 가까웠던 결과까지 통째로 버리고 완전
원상복구하면 안 된다.** "포기할 바에는 아무것도 안 하느니만 못하다"는 원칙은 **불변조건이
위협받을 때만**(예: 모든 시도가 오버플로) 적용해야지, "정확히 딱 맞지는 않지만 안전한 범위
안의 최선"까지 포기 대상으로 삼으면 안 된다. 매 시도마다 "지금까지 본 것 중 불변조건을 지키는
최선"을 `best_*` 변수로 기억해 두고, 최종적으로 정확히 맞지 않으면 그 best를 재적용해서
반환한다.

이 원칙은 흡수·이동 같은 "이웃에서 부족분을 채우는" 로직에도 적용된다: 이웃의 유일한 자원을
통째로 가져가면 이웃이 완전히 비워진다는 이유로 아무것도 하지 않고 건너뛰면, 그 자리는 계속
크게 미달인 채로 남는다. 빈 이웃을 만드는 것 자체가 아니라 **빈 이웃이 결과물에 그대로 남는
것**이 진짜 문제이므로, 통째로 가져간 뒤 빈 이웃을 삭제해서 병합하는 것이 "아무것도 안 하는
것"보다 낫다.

**발견 경위 (2026-09-18):** `_split_and_adjust_via_com()`의 재시도 루프가 keep=3에서 이미
"8줄(목표 9줄에 1줄 미달, 오버플로 아님)"을 확보했는데도, keep=4 재시도가 그 단락의 워드랩
버킷 소진으로 실패하자 8줄 결과까지 버리고 완전 포기 → `_rebalance_reading_slides_post_write()`
의 재조정 루프가 이를 "이번 스윕은 변화 없음"으로 오인해 조기 수렴 종료, 제2독서 콘텐츠
슬라이드가 6줄/4줄로 크게 미달한 채 남음(실사용자 육안 검수로 발견, 상세:
`docs/ooxml-pitfalls-log.md`). best-so-far 추적으로 수정.

## Tkinter Treeview 등 GUI 위젯을 "펼치는" 방법이 여러 개일 때, 그중 하나에만 이벤트를 걸지 않는다

`ttk.Treeview` 같은 위젯은 사용자가 노드를 여는 방법이 하나가 아니다 — 행 텍스트를 더블클릭하는
것과 확장 화살표(▶/▼ indicator)를 단일 클릭하는 것은 **서로 다른 이벤트**다. 후자는 위젯이
자체적으로 처리하는 네이티브 Tk 토글이라, 앱 코드가 `<Double-1>`만 바인딩해두면 이 경로로는
전혀 호출되지 않는다. lazy-load(자리표시자를 실제 자식으로 교체)처럼 "노드가 열릴 때" 반드시
실행돼야 하는 로직을 트리거별로 따로따로 잡으려 하면, 놓친 트리거 경로에서 자리표시자만 보이고
실제 데이터가 영원히 로드되지 않는 채로 남는다 — 예외가 나지 않으므로 그 트리거를 직접 써보기
전까지는 발견되지 않는다.

**규칙:** "펼침" 같이 여러 트리거가 있는 상태 변화를 감지해야 할 때는, 트리거 방법(마우스
더블클릭/화살표 클릭/키보드)을 각각 따로 바인딩하는 대신, 위젯이 트리거 방법과 무관하게
발생시키는 **가상 이벤트**(Treeview의 경우 `<<TreeviewOpen>>`)가 있는지 먼저 찾는다. 있으면
그 가상 이벤트 하나에만 로직을 걸고, 개별 물리 이벤트(`<Double-1>` 등)는 그 위젯의 다른 목적
(예: 파일 선택)에만 남겨둔다. 가상 이벤트 핸들러 시점에는 위젯이 이미 네이티브 상태 전환(예:
open 여부)을 마친 뒤이므로, 핸들러 안에서 그 상태를 다시 수동으로 토글하면 안 된다(되돌리는
결과가 될 수 있음).

**테스트 방법:** 마우스 좌표를 흉내 내는 대신 `tree.focus(iid)`로 대상을 지정한 뒤
`tree.event_generate('<<TreeviewOpen>>')`로 가상 이벤트 자체를 직접 발생시킨다 — 이것이 바로
Tk가 어떤 트리거 방법에서든 실제로 내보내는 이벤트이므로, 사용자가 어떤 방식으로 펼쳤든
이 방법으로 충실히 재현된다.

**발견 경위 (2026-09-26):** `missa_gui._ask_onedrive_file_browser_popup()`의 폴더 lazy-load
fetch가 `<Double-1>`에만 바인딩돼 있어, 확장 화살표 단일 클릭으로 폴더를 펼치면 자리표시자
"(불러오는 중...)"만 보이고 실제 `list_children()` fetch가 영구히 일어나지 않음(실사용자
스크린샷으로 발견). `<<TreeviewOpen>>` 바인딩으로 수정(상세: `docs/ooxml-pitfalls-log.md`).

## 표시 문자열과 실제 처리값은 반드시 분리한다

Tkinter `Entry`의 `textvariable`처럼, 사용자에게 보여주는 문자열과 그 위젯의 "값"이 같은
변수를 공유하는 위젯이 많다. 이 값을 사용자 친화적으로 다듬고 싶다는 이유로(예: 너무 긴
경로를 줄여 보여주기) 그 변수 자체에 다듬은 문자열을 넣어버리면, 나중에 그 변수를 읽어
실제 처리(파일 열기, 경로 복사 등)에 쓰는 코드가 존재하지 않는 값을 그대로 쓰게 된다 —
표시용으로 다듬은 문자열은 대개 더 이상 유효한 파일 경로나 식별자가 아니기 때문이다. 이
오류는 예외를 던지지 않고(존재하지 않는 경로를 그냥 실패시키거나, 최악의 경우 조용히
아무 일도 안 함) 조용히 잘못된 동작을 만들어, 화면상 표시가 그럴듯해 보일수록 오히려
늦게 발견된다.

**규칙:** "이 값을 다듬어 보여주고 싶다"는 요구가 생기면, 위젯이 들고 있는 변수 자체를
바꾸지 말고 **표시 전용 값**과 **실제 처리값**을 처음부터 별도 변수/딕셔너리로 분리한다.
실제 처리 로직이 값을 읽는 지점은(여러 곳에 흩어져 있을 수 있음) 전부 "실제 처리값이
있으면 그것을, 없으면 기존 표시값을(둘이 같았던 기존 경로와의 호환)" 우선순위로 읽는
단일 헬퍼를 거치게 해서, 그 우선순위를 깜빡하는 호출부가 생기지 않게 한다.

이 원칙은 "여러 단계가 같은 소스에서 파생된 값을 써야 한다"는 더 넓은 계열(위 "텍스트
배치를 결정하는 단계(페이지네이션)와 검증하는 단계는 반드시 같은 박스 폭을 써야 한다"
항목 참고)의 변형이다 — 그 항목은 "두 계산 경로가 몰래 다른 입력을 쓰면 안 된다"는
것이었고, 이 항목은 "한 값의 '보여주는 버전'과 '쓰는 버전'이 갈라지면 안 된다"는 것이다.

**발견 경위 (2026-09-26):** `missa_gui._ask_combined_input_popup(mass_type='youth')`에서
OneDrive로 고른 파일의 로컬 다운로드 경로(`output\{date}_youth\{파일명}`)가 Entry 폭
밖으로 잘려 파일명이 안 보이는 문제(실사용자 스크린샷)를 고치려고, 표시를 OneDrive
상대경로로 바꾸려 했다. Entry의 `textvariable`이 곧 `on_ok()`가 읽는 `vars_[key]`라서,
그 변수 자체를 상대경로로 바꾸면 `on_ok()`가 존재하지 않는 경로("13. 기도문\위령 성월
기도.pptx"는 실제 파일시스템 경로가 아님)를 그대로 파일 처리에 넘기는 회귀가 생길 뻔했다
— 실제 로컬 경로를 별도 dict(`_od_real_paths`)에 저장하고, on_ok()의 모든 필드 조회를
`_resolved_field_value()`(실제 경로 우선) 단일 지점으로 통일해 해결(상세:
`docs/ooxml-pitfalls-log.md`).

## 여러 배치에서 계속 확장되는 함수에 새 부작용 호출을 추가할 때는, 그 함수를 모킹 없이 실제로 실행하는 기존 테스트를 찾아 함께 갱신한다

`_run_gui_mode()`처럼 여러 작업 라운드에 걸쳐 계속 새 단계(팝업, 안내창, 로그인 등)가
추가되는 함수에 또 하나의 실제 부작용(예: 실제 `messagebox.showinfo()`를 띄우는 안내 호출)을
끼워 넣을 때, 그 함수를 이미 호출하고 있는 기존 테스트가 이 새 호출을 모킹하지 않으면
헤드리스(pytest) 환경에서 진짜 모달 창이 뜬 채 무한 대기하다 타임아웃으로만 발견된다 —
예외가 나지 않고 그냥 멈추므로, 원인이 방금 추가한 새 호출인지조차 로그만으로는 알기 어렵다.

**규칙:** 함수에 새 호출을 추가하기 전에, 그 함수를 인자 전체를 모킹하지 않고 실제로
호출하는 기존 테스트를 grep으로 찾는다(`monkeypatch.setattr(mtp, '_run_gui_mode', ...)`처럼
함수 자체를 스텁 처리한 테스트는 영향받지 않지만, 실제 함수 본문을 실행하는 테스트는 모두
갱신 대상이다). 찾은 테스트에는 새로 추가한 호출도 다른 단계들과 동일하게 monkeypatch로
스텁 처리하고, 호출 순서 검증 리스트에도 끼워 넣는다.

**발견 경위 (2026-09-26):** `missa_to_ppt._run_gui_mode()`에 `_show_powerpoint_background_
notice()`(§J7, PowerPoint 백그라운드 실행 안내) 호출을 추가한 뒤, 이미 존재하던
`test_j3m_run_gui_mode_skips_mass_type_popup_and_uses_given_kr`(§J3, `_run_gui_mode('청년')`을
모킹 없이 실제로 호출)가 새 호출을 몰라 진짜 `messagebox.showinfo()`가 뜬 채 60초 타임아웃까지
멈춤(상세: `docs/ooxml-pitfalls-log.md`).

## Tkinter 합성 클릭 이벤트(`event_generate`)는 좌표와 `update_idletasks()` 없이는 조용히 무시될 수 있다

`widget.event_generate('<Button-1>')`처럼 좌표 없이 합성 마우스 클릭 이벤트를 발생시켜
`bind('<Button-1>', ...)` 콜백을 테스트하려 하면, 이 환경(Windows Tk)에서는 예외 없이
그냥 아무 일도 일어나지 않는다 — "표준 Tkinter API니까 당연히 될 것"이라는 가정이 배신당하는
지점이다. 좌표(`x=`, `y=`)를 줘도, 위젯이 실제 geometry(배치)를 아직 확정하지 못한 시점이면
여전히 무시될 수 있다.

**규칙:** 클릭/키 이벤트를 합성해 테스트할 때는 (1) 반드시 위젯 내부 좌표(`x=`, `y=`)를
지정하고, (2) `event_generate()` 직전에 그 위젯에 `update_idletasks()`를 호출해 geometry를
먼저 확정한다. 이 두 조건 중 하나라도 빠지면 콜백이 트리거되지 않는데 예외도 나지 않으므로,
"위젯이 존재하는지"만 확인하는 테스트는 이 회귀(존재는 하지만 클릭이 안 먹히는 배선 누락)를
잡지 못한다 — 실제로 이벤트를 발생시켜 콜백이 호출됐는지까지 검증해야 한다. 합성 이벤트가
예상대로 작동하지 않을 때는 표준 API 호출 하나로 해결하려 하지 말고, 최소 재현 스크립트
(별도 Tk 창 + 위젯 하나)로 좁혀서 정확히 어떤 인자/순서가 필요한지 실측한다.

**발견 경위 (2026-09-26):** `test_j10e_clicking_update_link_actually_invokes_on_check_update`
(§J10, '🔄 업데이트 확인' Label 클릭 시뮬레이션)가 `event_generate('<Button-1>')`(좌표 없이)
로는 구현이 맞는데도 계속 실패 — 최소 재현으로 좌표(`x=5, y=5`) 필요를 먼저 확인했지만,
실제 팝업의 깊이 중첩된 위젯에서는 그것만으로도 부족해 `update_idletasks()`를 이벤트 발생
직전에 호출해야 한다는 것까지 다시 최소 재현으로 좁혀 찾음(상세: `docs/ooxml-pitfalls-log.md`).

## PowerPoint `Application.Visible=True`는 프레임 창을 실제로 보이게 하지만, 포그라운드(포커스)를 가져가지는 않는다(실측 확인)

`ppt_com_verify.py`의 `Application.Visible=True`(변경 금지 제약, 위 "PowerPoint COM
자동화" 관련 설명 참고)가 사용자가 지금 쓰고 있는 창의 포커스를 빼앗아 갈지는 추측이 아니라
직접 측정해야 한다. `EnumWindows`로 실제 프레임 창(클래스명 `PP12FrameClass`)을 찾아
`IsWindowVisible()`/`IsIconic()`으로 확인한 결과, 이 창은 실제로 visible해지고(작업표시줄에
나타남) 최소화되지도 않지만, `GetForegroundWindow()`는 Dispatch 직후와 `Presentations.
Open(WithWindow=False)` 실측 도중 모두 원래 창을 그대로 유지했다(2회 반복 확인). 즉 "창이
전혀 뜨지 않는다"는 주장은 사실과 다르고(실제로 뜬다), "포커스를 가져간다"는 주장도 사실과
다르다(가져가지 않는다) — 사용자에게 안내할 때는 정확히 "작업표시줄에 나타날 수 있지만
포커스는 가져가지 않는다"고만 말해야 한다. 참고로 PowerPoint late-bound COM Application
객체에는 `HWND` 속성 자체가 없다(Word/Excel과 다름, `com_error: Member not found`) — 프레임
창을 찾으려면 OS 레벨 `EnumWindows` + 클래스명으로 찾아야 한다.

**발견 경위 (2026-09-26):** §J7(PowerPoint 백그라운드 실행 안내) 구현 전, 기존 안내 문구
"화면에 창이 뜨지 않습니다"가 실측과 맞는지 직접 검증하려고 real Dispatch + `EnumWindows`
스크립트를 작성해 확인 — 문구가 사실과 다름을 확인해 `ppt_com_verify.py`/새 GUI 안내 문구
둘 다 정정(상세: `docs/ooxml-pitfalls-log.md`).

## 서로 다른 목적의 캐시 계층이 겹칠 때, 한 계층의 스냅샷을 다른 계층의 최종 판단에 재사용하지 않는다

같은 원격 자원(OneDrive 파일 등)에 대해 캐시 계층이 두 개 이상 존재할 때(예: 트리 UI 표시용
프리페치 캐시 + 로컬 파일 재사용 여부를 결정하는 신선도 캐시), 한 계층이 이미 채워둔 값을
"공짜 API 호출 절약"이라는 이유로 다른 계층의 최종 판단에 그대로 재사용하면 안 된다. 두
계층은 채워지는 시점과 갱신 주기가 다르다 — 트리 표시용 캐시는 "그 폴더를 마지막으로 조회한
시점"에 고정되고 그 세션이 끝날 때까지 갱신되지 않는 경우가 많은 반면, 재사용 여부를
최종 결정하는 판단은 항상 "지금 이 순간의 실제 원격 상태"를 알아야 한다. 두 계층이 우연히
서로 일치해도(둘 다 같은 과거 시점의 스냅샷이라면 당연히 일치한다) 그 일치는 "최신 상태와
일치한다"는 뜻이 아니다 — 이 착각이 "원격이 변경됐으면 반드시 재확인/재다운로드"라는 불변
조건을 조용히 깨뜨린다(예외 없이 잘못된 값을 반환하므로 발견이 늦다).

**규칙:** 여러 캐시 계층이 같은 자원을 공유할 때는, 각 계층이 "어떤 질문에 답하기 위해
존재하는지"를 먼저 구분한다. "지금 화면에 뭘 보여줄까"(UI 표시용, 약간의 지연 허용)와
"지금 이 파일을 재사용해도 안전한가"(최종 판단용, 지연 허용 불가)는 다른 질문이다. 후자의
판단은 전자의 캐시가 이미 있어도 항상 실시간 재확인 경로를 거쳐야 하고, 전자의 캐시는 후자의
판단 함수 시그니처에서 완전히 제외하거나(가장 안전) 최소한 "이 값을 최종 판단에 쓰면 안 된다"
는 것을 코드와 docstring 양쪽에 명시해야 한다.

**발견 경위 (2026-09-27):** K1(OneDrive 브라우저 재귀 프리페치)이 트리 표시용으로 채운
`cache[parent_path]`(프리페치 시점에 고정된 스냅샷)를, K2(로컬 캐시 재사용 판단)의
`_lookup_remote_item_metadata()`가 "이미 있으니 재사용"이라는 이유로 실시간 조회
(`get_item_metadata()`) 대신 그대로 반환하도록 구현했다. 독립 코드 리뷰가 직접
`_resolve_onedrive_download()`를 호출해 재현: 로컬 매니페스트 기록과 K1 캐시가 둘 다
"지난주" 스냅샷으로 서로 일치하는데(신선해 보임), 실제 OneDrive는 그 사이 바뀐 상태로
스텁 — 재다운로드가 스킵되고 stale 콘텐츠가 반환됨을 확인. "OneDrive가 더 최신이면 이름이
같아도 반드시 재다운로드"라는 사용자 명시 핵심 불변조건이 정확히 이 경로에서 깨졌다.
`_lookup_remote_item_metadata()`를 완전히 삭제하고 `_resolve_onedrive_download()`가 K1
캐시와 무관하게 항상 `get_item_metadata()`를 직접 호출하도록 수정(상세:
`docs/ooxml-pitfalls-log.md`, `_workspace/02_specialist_impl_notes.md` "K그룹 리뷰
라운드17 재작업" 절).

## 회귀 테스트

`test_missa_regression.py`에 주일(20260712)·평일(20260624) 통합 테스트와 핵심 함수 단위
테스트가 있다. 독서·복음 줄 수 계산, 화답송/성가 처리 분기, 배경색, 절 번호 오렌지색처럼 이
문서에 기록된 버그와 직결된 항목을 검증한다.

**모든 테스트 파일은 `tests/` 폴더 아래 있다**(2026-10-01, 저장소 루트 정리 — 런타임 .py
13개는 `_SCRIPT_DIR` 상대경로 의존 때문에 루트에 그대로 둔다). 저장소 루트의 `conftest.py`
하나가 `tests/` 안에서 `import missa_gui` 등이 되도록 루트를 `sys.path`에 추가해주므로,
`pytest`를 저장소 루트에서 실행하기만 하면 평소처럼 동작한다.

**선택 실행 우선, 전체 실행은 최종 게이트 1회(2026-09-24 정책 변경)**: `tests/test_missa_regression.py`는
`TestIsSundayMass`/`TestWrapLineCount`/`TestSplitAndAdjustViaCom`/`TestComVerify` 등 기능별
클래스로 이미 나뉘어 있고, OneDrive/GUI-OneDrive/COM검증/업데이터는 별도 파일
(`tests/test_missa_onedrive.py`/`tests/test_missa_gui_onedrive.py`/`tests/test_ppt_com_verify.py`/
`tests/test_missa_updater.py`)로 분리돼 있다. 코드를 수정하는 동안에는 그 변경과 직접·간접
관련된 클래스/파일만 `-k`나 `::클래스명`으로 좁혀 돌린다(예: GUI 팝업만 바꿨으면
`pytest tests/test_missa_gui_onedrive.py -v`, 줄 수 계산 로직을 바꿨으면
`pytest tests/test_missa_regression.py -k "WrapLineCount or SplitAndAdjustViaCom" -v`). 전체
스위트(`pytest tests/ -v`, 총 322개·3~4분)는 **리뷰 통과 후
최종 게이트로 딱 한 번만** 돌려 관련 없어 보이는 곳까지 포함한 전수 확인을 한다 — 반복되는
개발 중간 확인마다 매번 전체를 돌리지 않는다(느려서 피드백 루프가 길어지고, PowerPoint COM을
실제로 띄우는 느린 테스트가 UI 전용 변경에서도 매번 실행되는 낭비가 컸다).

**알려진 환경 함정(2026-09-18 발견, 미해결 — 청년미사 2단계 관련이라 이번 작업 범위 밖):**
`test_y2_H1_naju_resolves_and_header_reflects_htype`·`test_y2_H2_cache_reuse_skips_regeneration`·
`test_y2_C1_real_combo_full_pipeline_structure_valid`는 `resolve_youth_hymn_pptx()`의
나주노/야훼이레 캐시-미스 경로에서 `get_onedrive_youth_hymn_remote_path()`/
`missa_onedrive.upload_file()`을 모킹하지 않는다. 이 머신처럼 실제 Azure AD 자격 증명이
`config.json`에 설정돼 있으면 MSAL 기기 코드 플로우(`obtain_token_by_device_flow`)가 실제로
호출되어 대화형 로그인 없이는 **영구히 멈춘다**(`py-spy dump`로 스택 확인). 전체 스위트를 돌릴
때는 이 3개를 `--deselect`로 제외해야 한다. 수정하려면 해당 테스트들이
`get_onedrive_youth_hymn_remote_path`(또는 `missa_onedrive.upload_file`)를 명시적으로
모킹해야 한다 — 청년미사 2단계 작업 시 함께 처리할 것.

**알려진 이슈(2026-09-24 발견, 미해결 — 이번 B/A/C/D 작업 범위 밖):** `ppt_com_verify.py`가
띄우는 PowerPoint COM `Application` 핸들에 `atexit`/`finally` 기반의 확실한 정리 경로가 없어,
pytest 프로세스가 비정상 종료(watchdog 강제종료, COM 티어다운 중 예외 등)하면 `POWERPNT.EXE`가
컨트롤러 없이 좀비로 남을 수 있다. 이번 세션 중 2회 실제 발생(`taskkill //F //IM POWERPNT.EXE
//T`로 수동 정리) — 한 번은 이전 실행이 600초 watchdog에 걸려 강제종료된 뒤, 다른 한 번은
`shutdown()`→`_discard_app()`의 티어다운 단계에서 `Windows fatal exception: code 0x800401fd`
(RPC_E_DISCONNECTED)로 크래시하면서 발생(단, 이 크래시 자체는 pytest 실행을 중단시키지
않았다 — 전체 스위트는 정상 완주). 좀비 `POWERPNT.EXE`가 남은 채 다음 COM 테스트를 돌리면
새 인스턴스가 기존 좀비와 충돌해 응답 없이 멈출 수 있으므로, COM 관련 테스트(`TestComVerify`/
`TestSplitAndAdjustViaCom`/`TestCountSlideLinesVerified` 등)를 돌리기 전에는 `tasklist`로
남은 `POWERPNT.EXE`가 없는지 먼저 확인하고, 있으면 정리 후 시작하는 습관이 필요하다. **코드
수정은 이번 범위에서 하지 않았다** — 후속 조치 후보는 `ppt_com_verify.shutdown()` 호출부를
`atexit.register()`로 등록하거나, pytest 세션 fixture의 teardown을 `try/finally`로 감싸
프로세스 자체가 비정상 종료돼도 `Application.Quit()`이 시도되도록 만드는 것.

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
