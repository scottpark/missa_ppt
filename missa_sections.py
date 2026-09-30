"""
missa_to_ppt.py 모듈 분리 리팩토링 Phase 4 — 섹션 탐색 + 검증.

find_sections()가 참조 PPT 안에서 입당송·독서·화답송·복음·성가 등 각 섹션의
슬라이드 위치를 찾고, validate()가 생성된 PPT가 JSON 내용과 실제로 일치하는지
(절 번호 오렌지색 렌더링 포함) 확인한다.
"""
from __future__ import annotations

import io
import re
from pathlib import Path

from pptx.enum.shapes import MSO_SHAPE_TYPE

from missa_ooxml_utils import (
    HYMN_TYPES, _slide_text, all_slide_texts, find_slide_with_text, find_shape_exact_text,
)
from missa_reading_layout import ORANGE, parse_into_verse_units


# 성가 5종의 라벨 키워드(슬라이드 텍스트 매칭용). missa_content_updaters.py의 title/헤더
# 재구성 함수들도 같은 매핑을 써야 하므로 모듈 상수로 노출한다(사본을 따로 두면 한쪽만
# 갱신되는 드리프트 위험).
HYMN_LABEL_KW = {
    '입당': '입 당', '봉헌': '봉 헌', '성체': '성 체',
    '2차봉헌': '2차 봉헌', '파견': '파 견',
}


# ─────────────────────────────────────────────────────────────────────────────

# 슬라이드 XML 유틸

# ─────────────────────────────────────────────────────────────────────────────



def find_content_range(prs, title_idx: int) -> tuple:

    """title_idx 다음 콘텐츠 슬라이드 범위 (start, end exclusive)."""

    NEXT_SECTION = [

        '화 답 송', '화답송', '제 2 독서', '제2독서',

        '복음 환호송', '복음환호송', '영성체송',

        '봉 헌', '성 체', '2차 봉헌', '파 견',

    ]

    start = title_idx + 1

    end = start

    n = len(prs.slides)

    for i in range(start, n):

        t = _slide_text(prs.slides[i]).strip()

        if not t:

            break

        if any(kw in t for kw in NEXT_SECTION):

            break

        end = i + 1

    return start, end





def find_복음_content_range(prs, title_idx: int) -> tuple:

    """복음 제목 슬라이드 다음 콘텐츠 슬라이드 범위."""

    STOP = ['영성체송', '봉 헌', '봉헌']

    start = title_idx + 1

    end = start

    n = len(prs.slides)

    for i in range(start, n):

        t = _slide_text(prs.slides[i]).strip()

        if not t:

            break

        if any(kw in t for kw in STOP):

            break

        end = i + 1

    return start, end





def _find_youth_title_shape(slide, type_kw: str):
    """청년미사용 성가 title(구 코드의 'divider') 도형을 반환(없으면 None).

    기존 `_is_hymn_divider()`는 "라벨+순수숫자" 2단락을 기대하지만, 청년 템플릿은
    "라벨+제목+출처문장" 3단락(번호가 다른 텍스트와 섞임)이라 매칭되지 않는다(실측).
    번호 조건을 빼고 라벨 텍스트만 보면, 114번("파 견" 인사말, PLACEHOLDER 도형)이 116번
    (진짜 파견 title, AUTO_SHAPE 도형)보다 먼저 매칭되는 충돌이 생긴다 — 실측 결과 진짜
    title 6장(12/65/103/106/109/116)은 전부 AUTO_SHAPE 도형 안에 라벨이 있어, shape 타입으로
    충돌 없이 구분할 수 있다(CLAUDE.md "텍스트 키워드로 도형을 식별할 때 부분 문자열 충돌
    주의"와 같은 계열의 함정 — 여기서는 shape 타입을 부가 신호로 좁혀 회피)."""
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        if shape.shape_type != MSO_SHAPE_TYPE.AUTO_SHAPE:
            continue
        for para in shape.text_frame.paragraphs:
            t = para.text.strip()
            if t == type_kw or t == type_kw.replace(' ', ''):
                return shape
    return None


def _is_hymn_title_youth(slide, type_kw: str) -> bool:
    return _find_youth_title_shape(slide, type_kw) is not None


def _slide_has_picture(slide) -> bool:
    return any(shape.shape_type == MSO_SHAPE_TYPE.PICTURE for shape in slide.shapes)


def _find_hymn_songs_youth(prs, kw: str) -> list:
    """청년미사 성가 title~content 범위를 곡 단위 리스트로 반환.

    콘텐츠 범위는 "빈 슬라이드까지" 대신 "PICTURE 도형이 있는 슬라이드까지"로 판정한다
    (실측: 입당 content 13-16은 전부 악보 Picture가 있고, 뒤이은 구분 슬라이드 17은
    "▶" 텍스트만 있는 AUTO_SHAPE라 비어있지 않지만 Picture는 없다 — 텍스트 유무 기준으로는
    17번이 콘텐츠로 잘못 흡수된다).

    성체처럼 두 번째 곡이 구분 슬라이드 없이 바로 이어지는 경우(103→104,105→106)에만
    다음 곡을 계속 찾는다 — 두 번째 title이 첫 곡의 content_end 바로 그 자리에 있을 때만
    "같은 페이지에 이어지는 곡"으로 보고, 그렇지 않으면(다른 htype 라벨이 우연히 뒷부분에
    다시 나타나는 경우 등) 확장하지 않는다.

    **알려진 제약 (독립 리뷰에서 발견, 02b_review_report.md §의심 1)**: "기타" 출처 콘텐츠
    슬라이드(예: 107번)는 가사 placeholder뿐이라 원래 악보 Picture가 없어야 정상이지만,
    참조 PPT의 107번은 우연히 작은 저작권 워터마크 Picture(104/105와 동일 위치·크기)가
    남아 있어 이 PICTURE 기준을 통과한다. 향후 누군가 그 잔여 이미지를 정리하거나, 다른
    본당 템플릿의 "기타" 콘텐츠 슬라이드가 정말로 Picture 없이 텍스트만 있으면
    `content_end == content_start`(빈 범위)가 되어 `_replace_one_youth_song`의 기타 헤더
    갱신 게이트(`if cs < ce:`)가 조용히 스킵된다(예외 없음). 새 템플릿을 붙일 때는 이
    가정이 여전히 성립하는지 먼저 실측 확인할 것."""
    n = len(prs.slides)
    songs = []
    search_from = 0
    while True:
        title_idx = -1
        for i in range(search_from, n):
            if _is_hymn_title_youth(prs.slides[i], kw):
                title_idx = i
                break
        if title_idx == -1:
            break
        cs = title_idx + 1
        ce = cs
        j = cs
        while j < n and _slide_has_picture(prs.slides[j]):
            ce = j + 1
            j += 1
        songs.append({'title_idx': title_idx, 'content_start': cs, 'content_end': ce})
        search_from = ce
        if not (search_from < n and _is_hymn_title_youth(prs.slides[search_from], kw)):
            break
    return songs


def _find_복음환호송_middle(texts) -> int:
    """'복음 환호송' 라벨이 붙은 연속 슬라이드 블록 중, 그날 구절(○로 시작)이 있는
    가운데 슬라이드만 반환. 앞/뒤 고정 슬라이드는 라벨은 같아도 '○'가 없다(실측: 48/50은
    고정 알렐루야뿐, 49만 '○ 너는 베드로이다...'로 시작). 성인 템플릿(복음환호송이 1장뿐,
    라벨+○가 같은 슬라이드에 공존)에서도 참이라 mass_type 분기 없이 공용으로 쓴다."""
    for i, t in enumerate(texts):
        if '복음 환호송' not in t and '복음환호송' not in t:
            continue
        if '○' in t:
            return i
    return -1


def find_sections(prs, mass_type: str = 'adult') -> dict:

    """PPT 내 모든 섹션 위치를 찾아 dict로 반환.

    mass_type='adult'(기본값)는 기존 동작을 100% 보존한다. mass_type='youth'는 성가 5종의
    title/content 탐지를 `_find_hymn_songs_youth()`로 교체하고 `{htype}_songs` 리스트를
    추가로 채운다(기존 flat 키 `{htype}_divider`/`_content_start`/`_content_end`는 첫 곡을
    가리키는 alias로 그대로 유지 — `insert_공지사항()` 등 기존 호출부 하위 호환)."""

    sections = {'title': 0}

    n = len(prs.slides)

    texts = all_slide_texts(prs)



    # 입당송

    i = find_slide_with_text(prs, '입당송')

    if i >= 0:

        sections['입당송'] = i



    # 시작기도: 입당송 직전 비어있지 않은 슬라이드들

    if '입당송' in sections:

        입당송_idx = sections['입당송']

        ptr = 입당송_idx - 1

        while ptr >= 0 and not texts[ptr].strip():

            ptr -= 1

        end_excl = ptr + 1

        start = ptr

        while start > 0 and texts[start - 1].strip():

            start -= 1

        if start < end_excl:

            sections['시작기도_start'] = start

            sections['시작기도_end'] = end_excl



    # 제1독서

    for i in range(n):

        t = texts[i]

        if '제 1 독서' in t or '제1독서' in t:

            sections['제1독서_title'] = i

            s, e = find_content_range(prs, i)

            sections['제1독서_start'] = s

            sections['제1독서_end'] = e

            break



    # 화답송

    for i in range(n):

        t = texts[i]

        if '화 답 송' in t or '화답송' in t:

            sections['화답송_start'] = i

            end = i + 1

            for j in range(i + 1, n):

                tj = texts[j]

                if not tj.strip():

                    break

                if '화 답 송' in tj or '화답송' in tj:

                    end = j + 1

                else:

                    break

            sections['화답송_end'] = end

            break



    # 제2독서

    for i in range(n):

        t = texts[i]

        if '제 2 독서' in t or '제2독서' in t:

            sections['제2독서_title'] = i

            s, e = find_content_range(prs, i)

            sections['제2독서_start'] = s

            sections['제2독서_end'] = e

            break



    # 복음환호송 (가운데 슬라이드: '○'로 시작하는 그날 구절이 있는 슬라이드)

    idx = _find_복음환호송_middle(texts)

    if idx >= 0:

        sections['복음환호송'] = idx



    # 복음 (제목 슬라이드: '복 음' 단독 shape)

    for i in range(n):

        slide = prs.slides[i]

        if find_shape_exact_text(slide, '복 음'):

            sections['복음_title'] = i

            s, e = find_복음_content_range(prs, i)

            sections['복음_start'] = s

            sections['복음_end'] = e

            break



    # 영성체송

    i = find_slide_with_text(prs, '영성체송')

    if i >= 0:

        sections['영성체송'] = i



    # 성가 섹션 (divider + content)

    HYMN_KEYWORDS = HYMN_LABEL_KW

    if mass_type == 'youth':

        for htype, kw in HYMN_KEYWORDS.items():

            songs = _find_hymn_songs_youth(prs, kw)

            if not songs:

                continue

            sections[f'{htype}_songs'] = songs

            # 하위 호환 alias: 기존 flat 키는 첫 곡을 가리킨다(insert_공지사항() 등 재사용)
            first = songs[0]

            sections[f'{htype}_divider'] = first['title_idx']

            sections[f'{htype}_content_start'] = first['content_start']

            sections[f'{htype}_content_end'] = first['content_end']

    else:

        for htype, kw in HYMN_KEYWORDS.items():

            for i in range(n):

                slide = prs.slides[i]

                if _is_hymn_divider(slide, kw):

                    sections[f'{htype}_divider'] = i

                    cs = i + 1

                    ce = cs

                    for j in range(cs, n):

                        if not texts[j].strip():

                            break

                        ce = j + 1

                    sections[f'{htype}_content_start'] = cs

                    sections[f'{htype}_content_end'] = ce

                    break



    return sections





def _is_hymn_divider(slide, type_kw: str) -> bool:

    """성가 divider 슬라이드 판별: 타입 텍스트 단락 AND 순수 숫자 단락이 존재."""

    has_type = False

    has_number = False

    for shape in slide.shapes:

        if not shape.has_text_frame:

            continue

        for para in shape.text_frame.paragraphs:

            t = para.text.strip()

            if t == type_kw or t == type_kw.replace(' ', ''):

                has_type = True

            elif re.match(r'^\d+$', t):

                has_number = True

    return has_type and has_number



# ─────────────────────────────────────────────────────────────────────────────

# 검증

# ─────────────────────────────────────────────────────────────────────────────



def validate_pptx_structure(pptx_path: str) -> list:
    """PPTX 내부 XML 구조 검증 — 깨진 이미지 참조와 잘못된 XML 문자를 탐지한다.

    PowerPoint가 '슬라이드에 일부 텍스트, 이미지, 개체 등이 손상되어 표시할 수 없습니다'
    오류를 띄우는 두 가지 원인을 미리 포착한다:
      1. XML 1.0에서 허용되지 않는 제어 문자 (0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F, 0x7F)
      2. 슬라이드 XML이 참조하는 rId가 .rels 파일에 정의되지 않은 경우 (깨진 이미지 링크)
    """
    import zipfile as _zip

    _INVALID_XML = re.compile(rb'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]')

    issues = []

    with _zip.ZipFile(str(pptx_path), 'r') as z:
        names = set(z.namelist())

        for name in sorted(names):
            if not re.match(r'ppt/slides/slide\d+\.xml$', name):
                continue

            content = z.read(name)
            slide_num = re.search(r'slide(\d+)\.xml$', name).group(1)

            # 1. 잘못된 XML 문자 검사
            m = _INVALID_XML.search(content)
            if m:
                bad_byte = m.group(0)[0]
                issues.append(f'슬라이드 {slide_num}: 잘못된 XML 문자 (0x{bad_byte:02X})')

            # 2. rId 참조 vs .rels 정의 대조
            used_rids = set(re.findall(rb'r:[a-zA-Z]+=[\'"](rId\d+)[\'"]', content))
            if not used_rids:
                continue

            rels_name = f'ppt/slides/_rels/slide{slide_num}.xml.rels'
            if rels_name not in names:
                issues.append(f'슬라이드 {slide_num}: .rels 파일 없음 ({len(used_rids)}개 참조 미검증)')
                continue

            rels_content = z.read(rels_name)
            defined_rids = set(re.findall(rb'Id="(rId\d+)"', rels_content))
            broken = used_rids - defined_rids
            if broken:
                broken_str = ', '.join(sorted(b.decode() for b in broken))
                issues.append(f'슬라이드 {slide_num}: 끊어진 이미지 참조 ({broken_str})')

    return issues



def _orange_verse_numbers_in_range(prs, start: int, end: int) -> set:
    """[start, end) 슬라이드 범위에서 오렌지색으로 표시된 절 번호 집합을 수집."""
    found = set()
    for idx in range(start, end):
        slide = prs.slides[idx]
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    try:
                        if run.font.color.rgb != ORANGE:
                            continue
                    except Exception:
                        continue
                    m = re.match(r'^(\d+(?:,\d+)?)', run.text.strip())
                    if m:
                        found.add(m.group(1))
    return found


def _missing_orange_verse_numbers(prs, start: int, end: int, content: str) -> list:
    """content에서 파싱한 절 번호 중, [start, end) 범위에 오렌지색으로 나타나지 않은 것들.

    절 번호 run이 병합/분리 과정에서 다른 run에 흡수되면 텍스트는 남아도 색상만
    사라질 수 있어(2026-07-26 발견), 전체 텍스트 존재 여부가 아니라 실제 오렌지색
    run으로 렌더링됐는지를 절 번호 단위로 확인한다."""
    expected = {
        u['verse_num'] for u in parse_into_verse_units(content)
        if u.get('verse_num')
    }
    if not expected:
        return []
    found = _orange_verse_numbers_in_range(prs, start, end)
    return sorted(expected - found, key=lambda v: [int(x) for x in v.split(',')])


def validate(prs, json_data: dict) -> bool:
    # (2026-09-17) 화답송 라벨 체크가 항상 실행되도록 단순화되면서(§4.3/§8.6) 이 함수
    # 본문 어디에서도 "주일/평일" 축을 더 이상 참조하지 않는다. 예전엔 is_sunday 매개변수를
    # 받았는데, 남겨두면 호출부가 이름과 반대 의미의 값을 넘겨도(실제로 그랬었다 — 독립
    # 리뷰에서 발견) 아무 문제가 드러나지 않아 조용한 함정이 된다. 파라미터 자체를 제거해
    # 그 함정을 구조적으로 없앴다(CLAUDE.md 계열 원칙: 안 쓰는 파라미터를 "혹시 몰라서"
    # 남기지 않는다).

    texts = all_slide_texts(prs)

    errors = []

    warnings = []



    liturgy = json_data.get('liturgy', '')

    if not any(liturgy in t for t in texts):

        errors.append(f'liturgy 없음: {liturgy}')



    # 화답송 라벨은 평일 포맷에서도 존재한다(실측 확인) — is_sunday로 게이팅할 근거가 없어
    # 항상 체크하도록 단순화했다(청년미사 2단계 설계 §4.3/§8.6, 애매한 축을 표시형식 플래그에
    # 얹는 다섯 번째 암묵적 결합을 만들지 않기 위함). 평일 템플릿 라벨은 "화 답 송"(공백 포함)
    # 형식이라(실측: 20260624), 주일 템플릿의 공백 없는 "화답송"과 두 변형 모두 확인해야
    # 오탐 경고가 나지 않는다(find_sections()의 화답송 탐지와 동일한 OR 패턴).
    check_kws = ['입당송', '영성체송']

    for kw in check_kws:

        if not any(kw in t for t in texts):

            warnings.append(f'{kw} 슬라이드 없음')

    if not any('화답송' in t or '화 답 송' in t for t in texts):

        warnings.append('화답송 슬라이드 없음')



    # 독서·복음 절 번호 오렌지색 확인 (절 번호 단위로 정확히 대조)

    sections = find_sections(prs)

    for label, start_key, end_key in [
        ('제1독서', '제1독서_start', '제1독서_end'),
        ('제2독서', '제2독서_start', '제2독서_end'),
        ('복음', '복음_start', '복음_end'),
    ]:

        reading = json_data.get(label)

        if not reading or not reading.get('content'):

            continue

        if start_key not in sections or end_key not in sections:

            warnings.append(f'{label} 슬라이드 위치를 찾을 수 없어 절 번호 확인 생략')

            continue

        missing = _missing_orange_verse_numbers(
            prs, sections[start_key], sections[end_key], reading['content']
        )

        if missing:

            warnings.append(f'{label} 절 번호 오렌지색 누락: {", ".join(missing)}')



    print('\n=== 검증 결과 ===')

    if errors:

        for e in errors:

            print(f'  [오류] {e}')

    if warnings:

        for w in warnings:

            print(f'  [경고] {w}')

    if not errors and not warnings:

        print('  모든 검증 통과!')



    return len(errors) == 0





# ─────────────────────────────────────────────────────────────────────────────

# 메인

# ─────────────────────────────────────────────────────────────────────────────



def _strip_slide_xml(content: bytes) -> bytes:

    """슬라이드 XML에서 호환성 문제 요소 제거 (strip_ppt2007_incompatible 내부용)."""

    # <a:ext ...><a14:imgProps ...>...</a14:imgProps></a:ext> 블록 제거

    # (python-pptx가 효과 이미지를 복사 못 해 r:embed가 단절됨)

    content = re.sub(

        rb'<a:ext uri="[^"]+"><a14:imgProps\b.*?</a14:imgProps>\s*</a:ext>',

        b'',

        content,

        flags=re.DOTALL,

    )

    # imgProps 제거 후 남는 빈 <a:extLst></a:extLst> 제거

    content = re.sub(rb'<a:extLst>\s*</a:extLst>', b'', content)

    # <a:videoFile r:link="..."/> 제거

    content = re.sub(rb'<a:videoFile[^/]*/>\s*', b'', content)

    # <a:hlinkClick ... action="ppaction://media" .../> 제거

    content = re.sub(

        rb'<a:hlinkClick[^>]+action="ppaction://media"[^>]*/>\s*',

        b'',

        content,

    )

    # 비디오 제거 후 남는 고아 미디어 타이밍 제거
    # videoFile / hlinkClick 제거 후 <p:timing> 안의 <p:video> 블록과
    # presetClass="mediacall" 애니메이션이 남으면 PowerPoint가 깨진 개체로 인식함

    if b'<p:video>' in content or b'presetClass="mediacall"' in content:

        content = re.sub(rb'<p:timing>.*?</p:timing>\s*', b'', content, flags=re.DOTALL)

    return content





def strip_ppt2007_incompatible(pptx_path: str) -> None:

    """PPT 2007에서 에러를 유발하는 요소 제거.



    제거 대상:

    - ppt/changesInfos/ (PowerPoint 2016 공동 작성 추적 파일)

    - [Content_Types].xml 의 changesInfo Override

    - ppt/_rels/presentation.xml.rels 의 changesInfo Relationship

    - 슬라이드 내 외부 비디오 링크 (MP4 등 — PPT 2007 미지원)

    - <a14:imgProps> 블록 및 제거 후 남는 빈 <a:extLst>

    - 슬라이드 rels에서 고아 image/hdphoto Relationship 제거

    - app.xml 슬라이드 카운트 수정

    """

    import zipfile as _zip



    path = Path(pptx_path)

    data = path.read_bytes()



    # 1패스: 슬라이드 XML 전처리 후 사용 중인 rId 수집 (고아 rel 제거용)

    slide_used_rids: dict = {}

    with _zip.ZipFile(io.BytesIO(data), 'r') as _z:

        actual_slide_count = sum(

            1 for n in _z.namelist()

            if re.match(r'^ppt/slides/slide\d+\.xml$', n)

        )

        for name in _z.namelist():

            if re.match(r'ppt/slides/slide\d+\.xml$', name):

                stripped = _strip_slide_xml(_z.read(name))

                used = set(re.findall(rb'r:[a-zA-Z]+=[\'"](rId\d+)[\'"]', stripped))

                slide_used_rids[name] = used



    buf = io.BytesIO()

    with _zip.ZipFile(io.BytesIO(data), 'r') as zin:

        with _zip.ZipFile(buf, 'w', _zip.ZIP_DEFLATED) as zout:

            for name in zin.namelist():

                if 'changesInfo' in name:

                    continue

                content = zin.read(name)



                if name == '[Content_Types].xml':

                    content = re.sub(

                        rb'<Override[^>]*changesInfo[^>]*/>\s*',

                        b'',

                        content,

                    )

                elif name == 'ppt/_rels/presentation.xml.rels':

                    content = re.sub(

                        rb'<Relationship[^>]*changesInfo[^>]*/>\s*',

                        b'',

                        content,

                    )

                elif name == 'docProps/app.xml':

                    # Slides 카운트를 실제 슬라이드 수로 수정

                    content = re.sub(

                        rb'<Slides>\d+</Slides>',

                        f'<Slides>{actual_slide_count}</Slides>'.encode(),

                        content,

                    )

                    # 비디오 제거 후 MMClips도 0으로

                    content = re.sub(rb'<MMClips>\d+</MMClips>', b'<MMClips>0</MMClips>', content)

                elif re.match(r'ppt/slides/_rels/slide\d+\.xml\.rels', name):

                    # 슬라이드 rels에서 외부 비디오 Relationship 제거

                    content = re.sub(

                        rb'<Relationship[^>]+Type="[^"]*relationships/video[^"]*"[^>]*/>\s*',

                        b'',

                        content,

                    )

                    # 고아 image/hdphoto Relationship 제거

                    # (shape는 삭제됐으나 rel 항목이 남은 경우)

                    slide_name = name.replace('_rels/', '').replace('.rels', '')

                    used_rids = slide_used_rids.get(slide_name, set())



                    def _drop_orphan(m, _used=used_rids):

                        attrs = m.group(0)

                        rid_m = re.search(rb'Id="(rId\d+)"', attrs)

                        type_m = re.search(rb'Type="([^"]+)"', attrs)

                        if not rid_m or not type_m:

                            return attrs

                        rid = rid_m.group(1)

                        rtype = type_m.group(1)

                        if (b'image' in rtype or b'hdphoto' in rtype) and rid not in _used:

                            return b''

                        return attrs



                    content = re.sub(rb'<Relationship\s[^>]*/>', _drop_orphan, content)

                elif re.match(r'ppt/slides/slide\d+\.xml', name):

                    content = _strip_slide_xml(content)



                zout.writestr(name, content)



    path.write_bytes(buf.getvalue())





