# 모듈별 구현 상세 노트

`CLAUDE.md`의 "프로젝트 구조"는 각 모듈을 한 줄로만 소개한다. 이 문서는 그중 설계 이유·
과거 변경 경위·비직관적인 내부 동작까지 알아야 안전하게 손댈 수 있는 모듈들의 상세를
모아둔다. 새 모듈 노트를 추가할 때는 이 파일에 새 절을 만들고, `CLAUDE.md`의 해당 모듈
한 줄 설명 끝에 "(상세: 이 문서)" 포인터만 남긴다.

---

## `missa_to_ppt.py` — 진입점

CLI 인자 파싱(`parse_args`)·파일 탐색(`find_files`)·JSON 로드(`get_json_data`)·`main()`/
`__main__` 흐름만 담당한다. 실제 처리 로직은 `missa_ooxml_utils`/`missa_gui`/
`missa_reading_layout`/`missa_sections`/`missa_content_updaters` 5개 모듈에 있고, 이
파일은 이들을 조합해 호출한다(2026-09 모듈 분리 리팩토링, 순수 이동·동작 무변경).

`if __name__ == '__main__':`은 `_dispatch_cli_or_gui()`(2026-09-26) 하나만 부른다 — 미사
유형 선택 팝업은 삭제됐고, `--미사유형`(기본 '성인', '어린이'는 미지원 안내 후 종료) CLI
플래그로만 정해진다. date가 있으면 `main()`(CLI), 없으면 `_run_gui_mode()`(GUI 입력 흐름,
처리 시작 직전 `_show_powerpoint_background_notice()`로 PowerPoint 백그라운드 실행을
안내)로 간다. `_build_arg_parser()`가 인자 정의를 `parse_args()`와 공유한다(상세는
`docs/missa_to_ppt 구현 계획.md` §1).

## `missa_ooxml_utils.py` — OOXML 저수준 프리미티브

슬라이드/도형 복사·rId 매핑·배경 상속·텍스트런 조작. 프로젝트 내부 의존성이 없는 leaf
모듈. `HYMN_TYPES`도 여기 위치(entry↔content_updaters 순환 임포트 방지).

## `missa_gui.py` — Tkinter 팝업 + 설정/OneDrive 오케스트레이션

Tkinter 팝업 전부 + `is_sunday_mass`/config.json 로드·저장/`get_onedrive_hymn_folder`.
이 함수들은 GUI 팝업과 진입점 양쪽에서 호출되어 entry에 두면 순환 임포트가 생기므로 gui
모듈에 흡수했다(entry→gui 단방향). `OUTPUT_ROOT`도 entry·gui 양쪽에서 쓰여 여기 위치.

**2026-09-26**: 미사 유형 선택 팝업(`_ask_mass_type_popup`)과 `config.json`의
`last_mass_type`은 삭제됐다(미사 유형은 이제 CLI `--미사유형`으로만 온다).
`_on_check_update(parent)`(그 팝업의 클로저였던 업데이트 확인 로직)는 독립 함수로 남아
`_ask_combined_input_popup()`의 확인 버튼 아래 새 행(`_get_app_version()`이 읽은 버전
텍스트 + `tk.Label(cursor='hand2')` 링크)에 재배선됐다.

**2026-09-27(K1)**: `_prefetch_onedrive_children()`이 재귀로 확장돼 3개 허용 루트 폴더의
모든 하위 폴더/파일을 한 번에 캐시에 채운다(실측: API 호출 9회·하위폴더 6개·파일 31개·
최대 깊이 2·순차 총 22초 — 이 규모면 병렬화 불필요). `<<TreeviewOpen>>`이 거치는
`_cached_onedrive_children()`의 "캐시에 있으면 네트워크 스킵" 계약 덕분에 재귀로 캐시를
더 깊이 채우는 것만으로 하위 폴더 탐색도 캐시 히트가 된다.

**2026-09-27(K2)**: OneDrive에서 받는 파일(공지사항/시작기도/미사후기도/성가 PPT 등)을
날짜 폴더와 별개인 공유 캐시 폴더(`_ONEDRIVE_CACHE_ROOT = 'onedrive_cache'`)에 원본
1부로 보관해, 매니페스트(`onedrive_cache/manifest.json`)에 기록된 lastModifiedDateTime이
OneDrive 쪽과 **문자열 완전일치**할 때만 재다운로드를 스킵한다(`_resolve_onedrive_
download()`). 신선도 판단은 K1의 프리페치 캐시와 무관하게 매번
`missa_onedrive.get_item_metadata()`로 실시간 재확인한다 — K1 캐시 스냅샷을 실시간값으로
오인해 "OneDrive가 더 최신이면 무조건 재다운로드" 불변조건이 깨지는 버그가 리뷰에서
발견됐다(상세 사고 경위: `docs/ooxml-pitfalls-log.md` "서로 다른 목적의 캐시 계층이
겹칠 때" 절).

**2026-09-27(L1)**: `_start_onedrive_prefetch_thread()` 신설 — K1 재귀 프리페치를
`_ask_combined_input_popup()`이 창을 띄우는 시점(청년미사만)으로 앞당겨, 로그인/네트워크
상태와 무관하게 첫 입력창이 항상 즉시 뜨도록 보장한다(스레드 시작 자체가 논블로킹).
`_ask_onedrive_file_browser_popup()`도 같은 헬퍼를 호출해 2차 안전망을 유지.

**2026-09-27(L2)**: OneDrive 파일 선택(`_on_choose()`)이 다운로드를 기다리지 않고 결정적
목적지 경로만 계산해 즉시 반환하고, 실제 다운로드는 generation 추적 백그라운드 스레드
(`_download_worker`)로 넘긴다 — `_ensure_onedrive_blob_fresh()`(공유 캐시 blob만 최신화)
와 `_commit_onedrive_blob_to_dest()`(순수 복사)로 분리해, 같은 슬롯을 같은 파일명으로
재선택해도(시작기도/미사후기도/공지사항 등이 매주 다른 폴더에서 같은 이름으로 반복되는
이 앱의 정상 관례) 구세대(느린) 다운로드가 최신 파일을 덮어쓰지 못하게 한다. 커밋은
`_download_worker`가 `_download_lock` 안에서 generation을 재확인한 경우에만 수행한다 —
최초 구현은 재확인이 메모리 상태만 보호하고 디스크 커밋은 무조건 수행해, "확인 성공 선언
후 최신 파일이 조용히 구식 내용으로 역전"되는 버그가 있었다(서브프로세스 Tk 프로브로
직접 재현·수정 확인).

**2026-09-27(2차)**: 가톨릭성가/나주노/야훼이레 성가 조회가 폴더 전체 미러링에서 곡 단위
온디맨드 조회로 전환됐다. `find_youth_onedrive_hymn_file()`이 (1) 로컬 캐시
(`_SCRIPT_DIR/cache/...`) → (2) `_find_onedrive_file()`로 OneDrive 실시간 검색, 2단계만
수행한다(로컬 동기화 폴더 우선검색 단계는 사용자 요청으로 제거됨). `get_onedrive_hymn_
folder()`(성인미사 전용, 여전히 폴더 전체 미러링)는 이 변경과 무관 — `missa_to_ppt.py`의
`mass_type == 'adult'` 경로에서만 쓰인다. `get_onedrive_youth_hymn_root()`/`get_onedrive_
youth_hymn_remote_path()`는 이 전환 이후 프로덕션 코드에서 더 이상 호출되지 않는다(테스트
전용으로만 남음 — 삭제 여부는 아직 결정 안 됨).

## `missa_reading_layout.py` — 독서·복음 레이아웃 엔진

절 파싱(`parse_into_verse_units`)부터 슬라이드 분배(`layout_units_on_slides`), Pillow
추정 + PowerPoint COM 실측 줄 수 계산, 기록 후 재조정까지 독서/복음 처리 파이프라인
전체.

## `missa_sections.py` — 섹션 탐색·검증

`find_sections`(섹션 탐색), `validate`/`validate_pptx_structure`/
`strip_ppt2007_incompatible`(검증).

## `missa_content_updaters.py` — 섹션별 콘텐츠 갱신

화답송·성가·입당송·복음환호송·영성체송·시작기도문·미사후기도 등 섹션별 콘텐츠 갱신
(`update_*`/`replace_*` 함수들). 청년미사 성가 5종×4출처 통합은 `replace_성가_youth()`/
`resolve_youth_hymn_pptx()`가 담당(청년미사 §참고).

## `missa_psalm_score_image.py` — 화답송 악보 이미지 변환

화답송 악보 원본 이미지(PNG/JPG) → 슬라이드 변환. 순수 함수 묶음,
`missa_to_ppt.py`를 import하지 않는 단방향 의존. 진입점 `render_화답송_score_slide()`.
고정 템플릿 자산은 `assets/화답송_악보_template.pptx`(`tools/build_화답송_template.py`로
생성). 수작업 PPT가 없을 때 자동 폴백된다.

## `missa_youth_gospel.py` — 청년미사 영문 복음/독서 조회

universalis.com에서 날짜별 **미사 전 섹션**(First reading·Responsorial Psalm·Second
reading·Gospel Acclamation·Gospel)을 조회해 JSON으로 반환. `fetch_gospel_html`(리다이렉트
차단 — 3xx면 `GospelFetchError`, 과거 날짜가 조용히 오늘자로 리다이렉트되는 함정 차단),
`parse_mass`(stdlib `html.parser`, 섹션 마커=`align="left"` th/참조=그 뒤 첫
`align="right"` th, 영어 키만 사용·한글 키 없음, Psalm은 reference만, Gospel
Acclamation의 "Or:" 대체 환호송은 `"or"` 중첩), `get_youth_mass`(기본 출력 파일명
`missa_en_YYYYMMDD.json`, 한글 미사 `missa_YYYYMMDD.json`과 구분). `parse_gospel`/
`get_youth_gospel`(Gospel 섹션만, 한글 키 `"복음"`)은 하위호환 래퍼로 유지.
`missa_to_ppt.py`의 `main()`이 `mass_type == 'youth'`일 때 `get_youth_mass(content_
date_str)`를 호출하고 `merge_youth_gospel_content()`로 기존 `json_data`와 병합한다.

## `missa_youth_hymn_pdf.py` — PDF 성가집 → 성가 슬라이드 PPT

PDF 성가집(나주노=스캔 이미지 / 야훼 이레=텍스트 레이어) → 성가 슬라이드 PPT 생성.

파이프라인: 번호 조회 → 이미지 오선 밴드 탐지(`detect_system_bands`, 콘텐츠폭 기준
row_fill>0.40) → `_group_content_bbox`(그룹 **행 범위 안의 잉크 열**만으로 좌우 경계를
잡는 per-group 크롭 — 전역 `detect_content_bounds`를 쓰면 그 그룹과 무관한 다른 줄이
폭을 오염시킨다, 상세: `docs/ooxml-pitfalls-log.md` "이미지에서 여러 그룹을 크롭할 때"
절) → `pack_systems`(동적 min-slide 균형 패킹: 같은 배율로 최대 3개 시스템까지 겹침 없이
들어가고 그것이 실제로 슬라이드 수를 줄이거나 균형을 개선할 때만 3개씩 묶는다) → 저작권
크롭(`resolve_copyright_crop`, 꼬리 마침표를 잉크런 분석으로 제거하는
`_trim_trailing_period_px`, 배치 높이는 288032×0.9=259229 EMU로 전곡 고정·폭은 곡별
크롭 종횡비 유지) → 헤더 run 조립(`build_header_runs`, 출처 표시명만 `SOURCE_DISPLAY`로
치환 — "야훼 이레"→"야훼이레"; 제목이 헤더 폭을 넘치면 `_autosize_header`가 자간 spc→0 →
그래도 넘치면 제목 run만 폰트 축소, 궁서 `batang.ttc` index 2 Pillow 측정) →
`duplicate_slide` 복제 조립(`build_hymn_pptx`).

고정 템플릿 자산 `assets/청년미사_성가_template.pptx`(`tools/build_청년미사_template.py`로
생성), 나주노 저작권 크롭 좌표는 `assets/나주노_copyright_bbox_cache.json`(개발자 PC에서
Tesseract OCR 1회 생성 — 재생성 툴 `tools/build_나주노_copyright_cache.py`, 개발 전용;
런타임·exe는 캐시만 읽고 Tesseract를 호출하지 않는다; 도구가 만드는 값은 초안일 뿐이며
447 항목처럼 손제작 샘플과 대조해 사람이 손으로 확정한 값이 있을 수 있으므로, 도구를
재실행하면 그 hand-tuned 값이 도구 기본값으로 덮인다).

`missa_content_updaters.py`의 `resolve_youth_hymn_pptx()`가 나주노/야훼 이레 캐시-미스
경로에서 이 모듈을 호출해 새 성가 PPT를 만든다(2026-09-27부터: OneDrive/로컬 캐시 어디서도
곡을 못 찾았을 때만).

## `missa_onedrive.py` — Microsoft Graph API 접근

성당 공용 계정(brokenbaykccppt@gmail.com) 하나로 Microsoft Graph API(Device Code Flow)에
접근하는 완전 독립 leaf 모듈(다른 프로젝트 모듈을 import하지 않는다). `list_children`/
`ensure_folder`/`upload_file`/`download_file`/`get_item_metadata`.

**주의**: 이 계정의 OneDrive에서 최상위 경로 `PPT 문서`는 진짜 폴더가 아니라 **다른
드라이브를 가리키는 공유 바로가기(remoteItem)**다 — 경로 기반 주소(`root:/PPT 문서/...:`)
는 이 바로가기를 통과하지 못해 하위 경로 조회가 404/422로 실패한다. `_resolve_remote_
base()`/`_ensure_base_resolved()`가 첫 경로 세그먼트의 remoteItem 여부를 실제 GET으로
확인해 `_REMOTE_ITEM_CACHE`(프로세스 수명 캐시)에 저장하고, remoteItem이면 `_item_url()`이
`/drives/{driveId}/items/{itemId}/...` 주소로 전환한다. "경로 문자열만 올바르게 이어붙이면
Graph가 알아서 찾아줄 것"이라고 가정하면 이 함정이 재발한다 — 최상위 경로가 공유
바로가기인지 아닌지는 실제 API 응답(`remoteItem` 파셋 유무)으로만 알 수 있다.

**성능(2026-09-26)**: `get_access_token()`이 호출마다 `msal.PublicClientApplication(...)`을
새로 생성하면 그 생성 자체가 약 0.9~1.0초 걸린다(`AUTHORITY`가 네트워크 기반이라 생성자
안에서 매번 instance discovery를 재수행하는 것으로 추정). `_APP_CACHE`(모듈 레벨
`[app, cache]`)와 `_get_app()`으로 프로세스 수명 동안 앱 인스턴스를 1회만 생성해 재사용
하도록 수정(MSAL 공식 권장 패턴). 실측: 앱 생성 단독 1.131초 → 캐시 재사용 0.000000초.

**2026-09-27**: `get_item_metadata(path)` 신설 — children 없이 항목 메타데이터만 GET
(404→FileNotFoundError, `list_children`/`download_file`과 동일 에러 관례).
`missa_gui._resolve_onedrive_download()`가 로컬 캐시 재사용 여부를 판단할 때
`lastModifiedDateTime`/`eTag`를 실시간으로 확인하는 용도.

## `missa_updater.py` — GitHub 기반 자동 업데이트

배포 형태를 PyInstaller exe에서 "Python 인터프리터 + run_missa.bat 런처"로 전환하면서
(청년미사), `.py` 파일 교체만으로 즉시 최신화가 가능해졌다. git CLI 대신 GitHub REST API
+ 브랜치 zip 아카이브 다운로드만으로 구현한다(저장소가 public이라 인증 불필요).

`apply_update()`는 화이트리스트(`_UPDATE_FILE_GLOBS = ["*.py", "*.spec", "VERSION"]`,
`_UPDATE_DIRS = ["assets"]`)에 해당하는 파일/폴더만 설치 위치에 덮어쓴다 — `config.json`/
`output/`/`reference/`/`cache/`/`.git/` 등 로컬 전용 상태는 자동으로 보존된다. `VERSION`은
2026-10-01부터 화이트리스트에 포함돼(이전엔 `.py`/`.spec`만) 업데이트 후 화면에 표시되는
버전 번호도 함께 동기화된다.

## `ppt_com_verify.py` — PowerPoint COM 실측 검증

Pillow 추정 줄 수를 실제 PowerPoint COM 자동화로 검증. `Application.Visible=True`는
변경 금지 제약(버전 호환성) — 프레임 창이 visible해지지만 포커스는 가져가지 않는다(실측
확인, 상세: `docs/ooxml-pitfalls-log.md` "PowerPoint Application.Visible=True" 절).

## 저장소 구조 (2026-10-01 정리)

- `tests/` — 테스트 파일 전부(런타임 .py 13개는 `_SCRIPT_DIR` 상대경로 의존 때문에 루트에
  그대로 둔다).
- `conftest.py`(저장소 루트) — `tests/`에서 루트의 `missa_gui` 등을 import할 수 있도록
  `sys.path`에 루트를 추가.
- `dist/` — `missa_to_ppt.spec`(PyInstaller)과 배포 zip 빌드 스크립트의 산출물
  (`missa_to_ppt.exe`/`missa_to_ppt.zip`/`missa_ppt_청년미사.zip`) 보관 폴더. 전부
  재생성 가능한 빌드 결과물이라 git에서 제외(`.gitignore`).
- `output/YYYYMMDD[_youth]/` — 미사별 작업 디렉터리(`OUTPUT_ROOT`, `missa_gui.py`).
  입력/출력 산출물이라 매번 커밋하지 않으며, 회귀 테스트 픽스처 3개(`20260624` 평일·
  `20260705` 성수축복·`20260712` 주일)만 git으로 추적한다.
- `docs/` — 요구사항·구현 계획(버전 번호 없이 항상 최신 상태 유지, 변경 이력은 별도
  변경이력 문서 참고), `archive/`(과거 버전 원문), `ooxml-pitfalls-log.md`(OOXML 함정
  규칙의 상세 본문+발견 경위), 이 문서(모듈별 구현 상세).

## 하네스(mass-ppt-dev) 변경 이력

`CLAUDE.md` "하네스: missa_ppt 개발" 절이 참조하는 에이전트/스킬 구성 자체의 변경 이력.
순수 기능 구현(새 미사 유형 지원 등)은 이 표의 대상이 아니다 — 에이전트/스킬 구성이
바뀐 경우만 기록한다.

| 날짜 | 변경 내용 | 대상 | 사유 |
|------|----------|------|------|
| 2026-08-25 | 초기 구성 (에이전트 3명 + 오케스트레이터 1개 + 스킬 4개) | 전체 | 청년미사·어린이미사·타 성당 확장을 앞두고 설계/구현/검증 역할 분리 필요 |
| 2026-08-25 | TDD 도입(`tdd-progression-testing` 스킬 신설, `test_missa_progression.py`↔`test_missa_regression.py` 분리, regression-qa에 승격 절차 추가) + 독립 코드 리뷰 단계 추가(`ooxml-code-reviewer` 에이전트 신설, specialist→reviewer→qa 파이프라인으로 확장) | 에이전트 4명, 스킬 5개 | 프로그레션(신규 동작 명세)과 회귀(기존 동작 보호) 테스트의 목적이 다름을 명확히 분리하고, 구현자 본인이 못 보는 사각지대를 잡을 독립 리뷰 단계 필요 |
| 2026-08-30 | 하네스 첫 실제 기능(화답송 이미지 지원) 완료 후 harness Phase 6(with-skill vs without-skill 비교, `_workspace/04_phase6_validation_report.md`) 실시 → 결과 반영: (1) `ooxml-code-reviewer`에 "검증 우선순위 0"(스펙 자체가 원본 자료와 일치하는지 독립 재검증, 스펙에서 복사된 기대값 신뢰 금지) 추가, (2) `mass-template-architect`에 "실측 검증"의 정의를 통계적 타당성이 아니라 개별 판단 지점의 직접 확인으로 명시, (3) `ooxml-pitfalls`에 "완성 XML 통째 추가"가 append/순서 함정을 구조적으로 피하는 대안이라는 항목과 `validate_pptx_structure()`가 요소 순서를 검사하지 않는다는 주의사항 추가 | ooxml-code-reviewer.md, mass-template-architect.md, ooxml-pitfalls/SKILL.md | with-skill(독립 리뷰 4라운드)조차 1라운드에서 설계 스펙 자체의 오류(바라인 오검출)를 못 잡고 실사용자 육안 검수로 뒤늦게 발견됨 — "스펙과 일치하는가"만 보는 리뷰로는 부족하고 원본 자료 재검증이 필요함을 실측으로 확인. without-skill 베이스라인은 독립 검증 부재로 같은 종류의 결함(3세트 중 2세트가 실제로는 마디 경계 아님)을 "정확함"으로 자체 오판·보고 |
