# missa_to_ppt.py 구현 계획 v1.0

## 1. 개요

방식: 참조 PPT를 python-pptx로 열어 JSON + 성가/시작기도 PPT로 내용을 교체 후 새 PPT 저장.

```
python missa_to_ppt.py 20260628 [--입당 55 --봉헌 216 --성체 163 --2차봉헌 205 --파견 19]
```

성가번호가 생략되면 tkinter 팝업으로 입력. `--test` 플래그로 팝업 없이 기본값 테스트 실행 가능.

출력: `YYYYMMDD_<liturgy>.pptx`

## 2. 처리 흐름

1. Step 1: 성가번호 확인 (CLI 인수 없으면 tkinter 팝업)
2. Step 2: `YYYYMMDD/` 폴더에서 참조 PPT, 시작기도 PPT, 화답송 악보 PPT, 성가 PPT 탐색
3. Step 3: `missa_YYYYMMDD.json` 로드 (없으면 `missa_to_json.py` 자동 실행)
4. Step 4: 참조 PPT 열기
5. Step 5: 전례 텍스트 교체 (제목·입당송·제1독서·화답송·제2독서·복음환호송·복음)
6. Step 6: 종료 슬라이드 텍스트박스 위치 정렬 (제1독서·복음 → 제2독서 기준)
7. Step 7: 영성체송 교체
8. Step 8: 시작기도문 슬라이드 교체
9. Step 9: 성가 슬라이드 교체 (5종)
10. Step 10: 저장 및 검증

## 3. 섹션 탐색 (find_sections)

PPT 전체 슬라이드 텍스트를 스캔하여 각 섹션의 슬라이드 인덱스를 dict로 반환.

- 입당송: "입당송" 키워드
- 시작기도: 입당송 직전 비어있지 않은 슬라이드 범위 — 시작기도 앞 빈 슬라이드는 포함하지 않음
- 제1독서·제2독서: "제1독서"/"제2독서" 키워드 → title 인덱스 + content 범위 (start, end)
- 화답송: "화답송" 키워드
- 복음환호송: "복음 환호송"/"복음환호송" 키워드
- 복음: shape에 "복 음" 단독 텍스트 존재 → title + content 범위
- 영성체송·성가 divider도 동일 방식으로 탐색

## 4. 독서/복음 슬라이드 교체

### 4-1. 절 파싱 (parse_into_verse_units)

- content에 `\n`이 없으면 절 번호(숫자 + 한글/따옴표) 경계에서 regex split
- 각 줄을 `{ text, verse_num }` 단위로 변환
- "주님의 말씀입니다"·"◎ 하느님"·"◎ 그리스도님" 등 종료 텍스트는 파싱 단계에서 제외

### 4-2. 슬라이드 배분 (layout_units_on_slides)

- 각 verse unit을 `CHARS_PER_LINE`(26)자 기준 display line으로 분해
- 공백 기준 단어 경계 우선 분리 (불가능하면 강제 분리)
- `LINES_PER_SLIDE`(8)줄씩 슬라이드로 묶음 — 절 경계와 무관하게 꽉 채움
- 절 번호 오렌지색은 해당 숫자가 display line 첫머리에 올 때만 적용

### 4-3. 슬라이드 수 조정 및 ending 슬라이드 처리 (replace_reading_slides)

- `n_ending`: content 범위 뒤에서 `_has_ending_text()`로 연속 ending 슬라이드 수 감지
- `n_usable = 기존 슬라이드 수 - n_ending`
- `needed > n_usable`: `insert_slide_copy()`로 부족분 추가, `_set_slide_bg_black()` 호출
- `needed < n_usable`: 초과분 슬라이드 삭제
- 본문 슬라이드: `_set_reading_text()`로 텍스트 설정
- ending 슬라이드: `_clear_text_frame()`로 content box만 비움 (ending 텍스트 보존)

### 4-4. ending 슬라이드 위치 정렬 (_align_ending_slides_to_제2독서)

- 제2독서 마지막 ending 슬라이드의 텍스트박스 위치·크기를 기준으로 수집
- 수집 기준: ending 키워드 포함 여부로 "ending" / "content" 두 종류로 구분
- 제1독서·복음 ending 슬라이드의 대응 텍스트박스 위치·크기를 동일하게 설정

## 5. 복음환호송 교체 (update_복음환호송)

- 참조 PPT의 기존 단락 3개를 템플릿으로 추출
  - `para[0]` = 첫 ◎ 알렐루야 (lnSpc=110%, tabLst 없음)
  - `para[1]` = ○ 구절 (lnSpc=120%, tabLst 있음)
  - `para[2]` = 마지막 ◎ 알렐루야 (lnSpc=120%, tabLst 있음)
- ◎ 줄: `allel_count`로 첫/마지막 구분 → 해당 템플릿 deepcopy, run 텍스트만 교체
- ○ 줄: `para[1]` deepcopy, `run[0]`("○ \t") 유지, `run[1]`에 새 내용 설정, `run[2+]` 제거
- 결과: 대한체(◎)·바탕체(○ 내용) run 구조가 참조 PPT와 동일하게 유지됨

## 6. 시작기도문 교체 (replace_시작기도문)

- `sections["시작기도_start"]` ~ `sections["시작기도_end"]` 범위 슬라이드만 삭제
- `시작기도_start` 앞의 빈 슬라이드는 범위에 포함되지 않으므로 자동 보존
- 시작기도문.pptx 슬라이드를 `copy_slide_from_prs()`로 해당 위치에 삽입
- 삽입된 슬라이드마다 `_set_slide_bg_black()` 호출

## 7. 성가 교체 (replace_성가)

- 5종 순서대로 처리, 매번 `find_sections()` 재호출로 인덱스 변화 반영
- divider 슬라이드 번호 업데이트 (`_update_성가_divider_number`)
- 기존 content 슬라이드 삭제 후 성가 PPT 슬라이드를 `copy_slide_from_prs()`로 삽입
- 헤더 타입/번호 수정 (`_update_성가_header`, run 서식 보존)

## 8. 슬라이드 복사 유틸

- `duplicate_slide(prs, src_idx)`: 슬라이드 복제 — shape 트리 + 이미지 rel + 배경 복사, 맨 끝에 추가
- `insert_slide_copy(prs, position, src_idx)`: duplicate 후 원하는 position으로 이동
- `copy_slide_from_prs(target, pos, source, src_idx)`: 다른 Presentation에서 슬라이드 복사 (레이아웃명 매칭)
- `_copy_image_rels(src, dst)`: image/hdphoto rel 복사, 파트명 충돌 방지를 위해 고유 파트명 생성
- `_effective_bg(slide)`: 슬라이드→레이아웃→마스터 순으로 effective 배경(p:bg) 탐색

## 9. 배경 검정 설정 (_set_slide_bg_black)

- `p:cSld` 아래에 `p:bg`(solidFill 검정) 삽입
- `spTree` 최하단에 전체 슬라이드 크기의 검정 사각형(BlackBg) 삽입 — 레이아웃 shape 가림

## 10. 서식 보존 핵심 함수

- `_set_single_para_text(tf, text)`: 단일 단락 텍스트 교체 — pPr·rPr 보존
- `_replace_para_text_clone(para, text)`: 단락 run 재구성 — 첫 run deepcopy 후 텍스트 교체
- `_set_reading_text(tf, units)`: 절 번호 오렌지 run + 본문 run 생성 — pPr 보존
- `_update_prefix_in_runs(para, old, new)`: run 경계 유지하며 앞부분 텍스트만 교체 — rPr 보존
- `_clear_text_frame(tf)`: 텍스트박스 내용 삭제, 빈 단락 하나 유지
- `_has_ending_text(slide)`: "주님의 말씀입니다" 등 ending 키워드 슬라이드 판별
- `_align_ending_slides_to_제2독서(prs, sec)`: 제1독서·복음 ending 슬라이드 텍스트박스 위치를 제2독서 기준으로 정렬

## 11. PPT 2007 호환성 처리 (strip_ppt2007_incompatible)

- changesInfo 파일 및 관계 제거 (PowerPoint 2016 공동작성 추적 파일)
- `a14:imgProps` 블록 및 빈 `extLst` 제거
- 고아 image/hdphoto Relationship 제거
- 외부 비디오 링크 제거
- app.xml 슬라이드 카운트 수정

## 12. 의존성

- python-pptx
- lxml
- requests
- beautifulsoup4
- tkinter (표준 라이브러리)
</content>
