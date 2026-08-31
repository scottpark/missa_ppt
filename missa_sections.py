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

from missa_ooxml_utils import (
    _slide_text, all_slide_texts, find_slide_with_text, find_shape_exact_text,
)
from missa_reading_layout import ORANGE, parse_into_verse_units


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





def find_sections(prs) -> dict:

    """PPT 내 모든 섹션 위치를 찾아 dict로 반환."""

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



    # 복음환호송

    for i in range(n):

        t = texts[i]

        if '복음 환호송' in t or '복음환호송' in t:

            sections['복음환호송'] = i

            break



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

    HYMN_KEYWORDS = {

        '입당': '입 당', '봉헌': '봉 헌', '성체': '성 체',

        '2차봉헌': '2차 봉헌', '파견': '파 견',

    }

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


def validate(prs, json_data: dict, is_sunday: bool = True) -> bool:

    texts = all_slide_texts(prs)

    errors = []

    warnings = []



    liturgy = json_data.get('liturgy', '')

    if not any(liturgy in t for t in texts):

        errors.append(f'liturgy 없음: {liturgy}')



    check_kws = ['입당송', '영성체송']
    if is_sunday:
        check_kws.append('화답송')

    for kw in check_kws:

        if not any(kw in t for t in texts):

            warnings.append(f'{kw} 슬라이드 없음')



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





