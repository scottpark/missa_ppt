"""
missa_to_ppt.py 모듈 분리 리팩토링 Phase 3 — 독서/복음 레이아웃 엔진.

절 파싱(parse_into_verse_units)부터 슬라이드 분배(layout_units_on_slides), 실제
렌더링 줄 수 계산(Pillow 추정 + PowerPoint COM 실측), 슬라이드 기록 후 재조정
(_rebalance_reading_slides_post_write)까지 독서/복음 처리 파이프라인 전체.
"""
from __future__ import annotations

import atexit
import copy
import re

from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from xml.sax.saxutils import escape as _xml_escape

from missa_ooxml_utils import (
    _find_content_shape, _has_ending_text, _clear_text_frame, _para_append_run,
    _set_slide_bg_black, _build_com_probe_pptx, insert_slide_copy, delete_slide,
)
from missa_gui import _com_verification_enabled


CHARS_PER_LINE = 27   # 32pt 바탕체 24.8cm 텍스트박스 실측 기준 약 27자/줄

LINES_PER_SLIDE = 9

# Pillow(libraqm 미지원)는 커닝을 반영하지 못해 실제 PowerPoint 렌더링보다
# 텍스트 폭을 넓게 계산한다. 실측 슬라이드(2026-07-12 연중 제15주일 복음) 기준
# 1.03~1.04에서만 전체 일치 — 1.02 이하는 과소보정, 1.05 이상은 과보정됨.
_RENDER_WIDTH_CALIBRATION = 1.03

_PILLOW_FONT_CACHE: dict = {}  # (font_name, font_size_pt) → PIL ImageFont or None

ORANGE = RGBColor(255, 192, 0)


_SENTENCE_ENDERS = frozenset(

    '.!?'

    '\u3002\uff01\uff1f'  # 。！？

    '\u201c\u201d'         # unicode double quotes

    '\u2018\u2019'         # unicode single quotes

    '\u300d\u300f'         # 」』

    '"\'' 

)





def _ends_sentence(text: str) -> bool:

    """독서 본문 텍스트가 문장 종결로 끝나는지."""

    t = text.rstrip()

    if not t:

        return True

    return t[-1] in _SENTENCE_ENDERS





def parse_into_verse_units(content: str) -> list:

    """

    콘텐츠를 절 단위로 파싱.

    \\n이 있으면 줄 단위로, 없으면 내장 절 번호(숫자+ 한글)로 분리.

    '주님의 말씀입니다.' 등 마무리 텍스트는 제거.

    반환: [{'text': str, 'verse_num': str, 'is_continuation': bool}]

      is_continuation=True: 이전 unit이 문장 미완성으로 끝나 이 unit이 같은 줄에서 이어져야 함

    """

    STRIP_START = ('주님의 말씀입니다', '◎ 하느님', '◎ 그리스도님')



    if '\n' in content:

        lines = [l.strip() for l in content.split('\n')]

    else:

        # \n 없는 연속 텍스트: 절 번호 경계에서 분리

        # \d{1,3}(?:,\d+)? 로 장,절 형식(예: 10,1)도 지원

        # 문자 클래스에 유니코드 따옴표(U+2018/2019/201C/201D) 포함

        lines = re.split(

            r'(?<!\d)\s+(?=\d{1,3}(?:,\d+)?\s+[^\d\s])',

            content

        )

        lines = [l.strip() for l in lines]



    units = []

    prev_sentence_end = True

    for line in lines:

        if not line:

            continue

        if any(line.startswith(p) for p in STRIP_START):

            continue

        m = re.match(r'^(\d+(?:,\d+)?)\s+', line)

        verse_num = m.group(1) if m else ''

        is_continuation = not prev_sentence_end

        units.append({'text': line, 'verse_num': verse_num, 'is_continuation': is_continuation})

        prev_sentence_end = _ends_sentence(line)



    return units





def _visual_lines(text: str) -> int:
    """문자 수 기반 줄 수 추정 (영성체송 높이 조정 등 레이아웃 외 용도 전용)."""
    return max(1, (len(text) + CHARS_PER_LINE - 1) // CHARS_PER_LINE)


def _wrap_line_count(text: str) -> int:
    """단어 경계 word-wrap 시뮬레이션으로 줄 수 계산.

    layout_units_on_slides의 청킹 로직과 동일하므로 1 dl = 1줄이 보장된다.
    독서·복음 슬라이드 줄 수 계산에 사용한다.
    """
    if not text.strip():
        return 0
    count = 0
    pos = 0
    n = len(text)
    while pos < n:
        end = pos + CHARS_PER_LINE
        if end >= n:
            pos = n
        else:
            space = text.rfind(' ', pos, end + 1)
            if space > pos:
                pos = space + 1
            else:
                pos = end
        count += 1
    return count


def _page_visual_lines(page: list) -> int:
    """display_lines 리스트의 시각적 줄 수 — dl 1개 = 1줄."""
    return len(page)



def layout_units_on_slides(units: list) -> list:

    """절 단위를 슬라이드로 묶음.

    is_continuation=True인 unit은 이전 unit에 합쳐 하나의 논리 단락으로 처리.
    논리 단락은 CHARS_PER_LINE 기준 단어 경계에서 display_lines로 분해하여
    LINES_PER_SLIDE개씩 슬라이드로 묶음 (절 중간 분리 허용).

    반환: [[dl, ...], ...]
      dl: {'text', 'verse_num', 'extra_verses', 'new_para'}
        new_para=True  → 새 PPT paragraph 시작
        new_para=False → 이전 paragraph에 이어 붙임 (문장 연속, paragraph break 없음)
      extra_verses: [(pos_in_dl_text, verse_num), ...] — 오렌지색 절 번호 위치

    """

    # is_continuation 처리: 연속 절을 하나의 논리 단락으로 합침
    merged = []

    for unit in units:

        text = unit['text']

        verse_num = unit.get('verse_num', '')

        is_continuation = unit.get('is_continuation', False)

        if is_continuation and merged:

            prev = merged[-1]

            sep = '' if prev['text'].endswith(' ') else ' '

            join_pos = len(prev['text']) + len(sep)

            extra = list(prev.get('extra_verses', []))

            if verse_num:

                extra.append((join_pos, verse_num))

            merged[-1] = {

                'text': prev['text'] + sep + text,

                'verse_num': prev['verse_num'],

                'extra_verses': extra,

            }

        else:

            merged.append({

                'text': text,

                'verse_num': verse_num,

                'extra_verses': [],

            })

    # 각 논리 단락을 display_lines로 분해 (CHARS_PER_LINE 기준 단어 경계 분리)
    display_lines = []

    for mu in merged:

        mu_text = mu['text']

        mu_extra = mu.get('extra_verses', [])

        mu_verse = mu['verse_num']

        pos = 0

        first_dl = True

        while pos < len(mu_text):

            end = pos + CHARS_PER_LINE

            if end >= len(mu_text):

                chunk = mu_text[pos:]

                next_pos = len(mu_text)

            else:

                # 단어 경계: end 이하에서 마지막 공백
                space = mu_text.rfind(' ', pos, end + 1)

                if space > pos:

                    chunk = mu_text[pos:space]

                    next_pos = space + 1

                else:

                    chunk = mu_text[pos:end]

                    next_pos = end

            if not chunk:

                pos = next_pos

                continue

            # 이 display_line에 해당하는 extra_verses (머지 텍스트 절대 위치 → 상대 위치)
            chunk_start = pos

            chunk_end = pos + len(chunk)

            chunk_extra = [

                (p - chunk_start, v)

                for p, v in mu_extra

                if chunk_start <= p < chunk_end

            ]

            display_lines.append({

                'text': chunk,

                'verse_num': mu_verse if first_dl else '',

                'extra_verses': chunk_extra,

                'new_para': first_dl,

            })

            pos = next_pos

            first_dl = False

    # dl 1개 = 시각적 1줄 (단어 경계 분리 기준)
    # _visual_lines 기반 추정 대신 dl 개수로 직접 분배
    slides = []
    current_slide = []

    for dl in display_lines:
        if len(current_slide) >= LINES_PER_SLIDE:
            slides.append(current_slide)
            current_slide = []
        current_slide.append(dl)

    if current_slide:
        slides.append(current_slide)

    return slides if slides else [[]]



def _verify_and_rebalance_pages(pages: list, label: str) -> list:
    """layout_units_on_slides() 결과를 검증하고 줄 수 이상 슬라이드를 재조정.

    - 비마지막 슬라이드가 LINES_PER_SLIDE 초과 → 마지막 dl들을 다음 슬라이드로 이동
    - 비마지막 슬라이드가 LINES_PER_SLIDE 미만 → 다음 슬라이드 앞 dl들을 흡수 시도
    - 마지막 슬라이드는 어떤 줄 수여도 건드리지 않음
    """
    i = 0
    while i < len(pages):
        lines = _page_visual_lines(pages[i])
        is_last = (i == len(pages) - 1)

        if lines > LINES_PER_SLIDE:
            if len(pages[i]) > 1:
                moved = []
                while _page_visual_lines(pages[i]) > LINES_PER_SLIDE and len(pages[i]) > 1:
                    moved.insert(0, pages[i].pop())
                if i + 1 < len(pages):
                    pages[i + 1] = moved + pages[i + 1]
                else:
                    pages.append(moved)
                new_lines = _page_visual_lines(pages[i])
                print(f'  [{label}] 슬라이드 {i+1}: {lines}줄 → {new_lines}줄 ({len(moved)}개 항목 다음 슬라이드로 이동)')
                continue  # 현재 슬라이드 재검사
            else:
                print(f'  [{label}] 경고 슬라이드 {i+1}: {lines}줄 (항목 1개라 분리 불가)')

        elif not is_last and lines < LINES_PER_SLIDE:
            absorbed = 0
            while i + 1 < len(pages) and pages[i + 1]:
                test = pages[i] + [pages[i + 1][0]]
                if _page_visual_lines(test) <= LINES_PER_SLIDE:
                    pages[i].append(pages[i + 1].pop(0))
                    absorbed += 1
                    if not pages[i + 1]:
                        pages.pop(i + 1)
                        break
                else:
                    break
            if absorbed:
                new_lines = _page_visual_lines(pages[i])
                print(f'  [{label}] 슬라이드 {i+1}: {lines}줄 → {new_lines}줄 ({absorbed}개 항목 흡수)')

        i += 1

    # 최종 검증 보고
    issues = []
    for i, page in enumerate(pages):
        lines = _page_visual_lines(page)
        is_last = (i == len(pages) - 1)
        if lines > LINES_PER_SLIDE:
            issues.append(f'슬라이드 {i+1}: {lines}줄 (분리 불가)')
        elif not is_last and lines < LINES_PER_SLIDE:
            issues.append(f'슬라이드 {i+1}: {lines}줄 (흡수 불가)')
    if issues:
        print(f'  [{label}] 미해결 이슈: {", ".join(issues)}')

    return pages



def _set_reading_text(tf, units: list, line_spacing: float = None):

    """독서/복음 콘텐츠를 TextFrame에 설정 (절 번호 오렌지색, 서식 보존).

    extra_verses [(pos, verse_num)] 지원: 단락 내 여러 절 번호를 오렌지색으로 처리.
    line_spacing: 줄간격 배수 (예: 1.1 = 110%). None이면 템플릿 그대로.

    """

    from pptx.oxml import parse_xml as pptx_parse_xml

    A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'



    txBody = tf._txBody

    existing_paras = txBody.findall(qn('a:p'))



    template_para_el = existing_paras[0] if existing_paras else None

    template_run_el = None

    if template_para_el is not None:

        runs = template_para_el.findall(qn('a:r'))

        if runs:

            template_run_el = runs[0]



    for p in existing_paras:

        txBody.remove(p)



    ORANGE_FILL_XML = f'<a:solidFill xmlns:a="{A_NS}"><a:srgbClr val="FFC000"/></a:solidFill>'

    _FONT_TAGS = {qn('a:latin'), qn('a:ea'), qn('a:cs'), qn('a:sym')}



    def _orange_run(text_val, tmpl_r):

        """오렌지색 run 생성."""

        r = copy.deepcopy(tmpl_r)

        t_el = r.find(qn('a:t'))

        if t_el is not None:

            t_el.text = text_val

        rPr = r.find(qn('a:rPr'))

        if rPr is None:

            rPr = pptx_parse_xml(f'<a:rPr xmlns:a="{A_NS}"/>')

            r.insert(0, rPr)

        for fill_tag in (qn('a:solidFill'), qn('a:gradFill'), qn('a:pattFill')):

            for sf in rPr.findall(fill_tag):

                rPr.remove(sf)

        # OOXML 스키마: fill은 font(latin/ea/cs)보다 앞에 위치해야 함

        _ins = next((i for i, c in enumerate(rPr) if c.tag in _FONT_TAGS), len(rPr))

        rPr.insert(_ins, pptx_parse_xml(ORANGE_FILL_XML))

        return r



    def _white_run(text_val, tmpl_r):

        r = copy.deepcopy(tmpl_r)

        t_el = r.find(qn('a:t'))

        if t_el is not None:

            t_el.text = text_val

        # 템플릿 run의 fill 색상(오렌지 등)이 복사될 수 있으므로 명시적으로 흰색(bg1)으로 재설정
        rPr = r.find(qn('a:rPr'))

        if rPr is not None:

            for fill_tag in (qn('a:solidFill'), qn('a:gradFill'), qn('a:pattFill')):

                for sf in rPr.findall(fill_tag):

                    rPr.remove(sf)

            WHITE_FILL_XML = f'<a:solidFill xmlns:a="{A_NS}"><a:schemeClr val="bg1"/></a:solidFill>'

            _ins = next((i for i, c in enumerate(rPr) if c.tag in _FONT_TAGS), len(rPr))

            rPr.insert(_ins, pptx_parse_xml(WHITE_FILL_XML))

        return r



    def _add_colored_runs(p, text, verse_num, extra_verses, tmpl_r):

        """텍스트를 오렌지/흰색 runs로 분리하여 단락에 추가.

        verse_num: 텍스트 앞부분 절 번호 (단독 혹은 장,절 형식)

        extra_verses: [(pos, verse_num), ...] — 텍스트 중간 추가 절 번호

        """

        # 오렌지색 구간 수집

        orange_ranges = []

        if verse_num:

            m = re.match(r'^(\d+(?:,\d+)?\s*)', text)

            if m:

                orange_ranges.append((0, len(m.group(1))))

        for pos, vnum in (extra_verses or []):

            if pos < len(text):

                m2 = re.match(r'^(\d+(?:,\d+)?\s*)', text[pos:])

                if m2:

                    orange_ranges.append((pos, pos + len(m2.group(1))))

        orange_ranges.sort()



        if not orange_ranges or tmpl_r is None:
            if tmpl_r is not None:
                r = _white_run(text, tmpl_r)
            else:
                r = pptx_parse_xml(f'<a:r xmlns:a="{A_NS}"><a:t>{_xml_escape(text)}</a:t></a:r>')
            _para_append_run(p, r)

            return



        cur = 0

        for start, end in orange_ranges:

            if cur < start:

                _para_append_run(p, _white_run(text[cur:start], tmpl_r))

            _para_append_run(p, _orange_run(text[start:end], tmpl_r))

            cur = end

        if cur < len(text):

            _para_append_run(p, _white_run(text[cur:], tmpl_r))



    current_p = None

    for unit in units:

        text = unit['text']

        verse_num = unit.get('verse_num', '')

        extra_verses = unit.get('extra_verses', [])

        new_para = unit.get('new_para', True)

        if new_para or current_p is None:

            # 새 단락 (기존 단락 XML 복제)

            if template_para_el is not None:

                new_p = copy.deepcopy(template_para_el)

                for r in new_p.findall(qn('a:r')): new_p.remove(r)

                for br in new_p.findall(qn('a:br')): new_p.remove(br)

            else:

                new_p = pptx_parse_xml(f'<a:p xmlns:a="{A_NS}"/>')

            if line_spacing is not None:

                pPr = new_p.find(qn('a:pPr'))

                if pPr is None:

                    pPr = pptx_parse_xml(f'<a:pPr xmlns:a="{A_NS}"/>')

                    new_p.insert(0, pPr)

                for old_ln in pPr.findall(qn('a:lnSpc')):

                    pPr.remove(old_ln)

                val = int(line_spacing * 100000)

                # lnSpc는 OOXML 스키마상 pPr의 첫 번째 자식이어야 함 (spcAft 앞에 위치)
                pPr.insert(0, pptx_parse_xml(f'<a:lnSpc xmlns:a="{A_NS}"><a:spcPct val="{val}"/></a:lnSpc>'))

            txBody.append(new_p)

            current_p = new_p

            _add_colored_runs(current_p, text, verse_num, extra_verses, template_run_el)

        else:

            # 같은 논리 단락 연속 — 이전 paragraph에 공백+텍스트 이어 붙임 (paragraph break 없음)
            # 공백 run 먼저 추가 후, extra_verses 처리 포함하여 colored runs 추가

            if template_run_el is not None:

                _para_append_run(current_p, _white_run(' ', template_run_el))

            else:

                _para_append_run(current_p, pptx_parse_xml(

                    f'<a:r xmlns:a="{A_NS}"><a:t> </a:t></a:r>'

                ))

            _add_colored_runs(current_p, text, '', extra_verses, template_run_el)




def _count_slide_lines(slide) -> int:
    """슬라이드 본문 shape의 시각적 줄 수 계산 (word-wrap 시뮬레이션 기준)."""
    shape = _find_content_shape(slide)
    if shape is None:
        return 0
    return sum(
        _wrap_line_count(para.text)
        for para in shape.text_frame.paragraphs
        if para.text.strip()
    )


def _rendered_wrap_count(text: str, pil_font, box_px: float) -> int:
    """Pillow 폰트 메트릭으로 word-wrap 줄 수 계산."""
    if not text.strip():
        return 0
    words = text.split(' ')
    lines = 1
    cur = ''
    for word in words:
        candidate = (cur + ' ' + word) if cur else word
        if pil_font.getlength(candidate) <= box_px:
            cur = candidate
        else:
            if cur:
                lines += 1
            cur = word
    return lines


def _get_slide_render_params(slide):
    """슬라이드의 콘텐츠 shape에서 Pillow 폰트와 박스 너비(px)를 반환.
    실패 또는 Pillow 미설치 시 (None, 0.0) 반환."""
    try:
        from PIL import ImageFont
    except ImportError:
        return None, 0.0

    shape = _find_content_shape(slide)
    if shape is None:
        return None, 0.0

    tf = shape.text_frame
    margin_l = tf.margin_left or 0
    margin_r = tf.margin_right or 0
    box_emu = shape.width - margin_l - margin_r
    if box_emu <= 0:
        box_emu = shape.width
    box_px = box_emu / 914400 * 96
    # 보정 계수: Pillow는 libraqm(커닝/복합 텍스트 셰이핑) 미지원 환경에서
    # 글자 advance width를 단순 합산해 실제 PowerPoint(DirectWrite) 렌더링보다
    # 폭을 약간(2~4%) 넓게 계산한다. 실측 확인된 슬라이드 기준으로 보정.
    box_px *= _RENDER_WIDTH_CALIBRATION

    font_name = 'BatangChe'
    font_size_pt = 32.0
    for para_el in tf._txBody.findall(qn('a:p')):
        for r_el in para_el.findall(qn('a:r')):
            rPr = r_el.find(qn('a:rPr'))
            if rPr is not None:
                sz = rPr.get('sz')
                latin = rPr.find(qn('a:latin'))
                if sz:
                    font_size_pt = int(sz) / 100.0
                if latin is not None:
                    fn = latin.get('typeface', '')
                    if fn and not fn.startswith('+'):
                        font_name = fn
                if sz or (latin is not None and latin.get('typeface')):
                    break
        else:
            continue
        break

    cache_key = (font_name, round(font_size_pt, 1))
    if cache_key not in _PILLOW_FONT_CACHE:
        import os
        _FONT_MAP = {
            'Batang':    ('batang.ttc', 0), '바탕':   ('batang.ttc', 0),
            'BatangChe': ('batang.ttc', 1), '바탕체': ('batang.ttc', 1),
            'Gulim':     ('gulim.ttc',  0), '굴림':   ('gulim.ttc',  0),
            'GulimChe':  ('gulim.ttc',  1), '굴림체': ('gulim.ttc',  1),
            'Malgun Gothic': ('malgun.ttf', 0), '맑은 고딕': ('malgun.ttf', 0),
        }
        fname, fidx = _FONT_MAP.get(font_name, ('batang.ttc', 0))
        fpath = os.path.join(r'C:\Windows\Fonts', fname)
        pil_font = None
        if os.path.exists(fpath):
            sz_px = max(1, round(font_size_pt * 96 / 72))
            try:
                pil_font = ImageFont.truetype(fpath, sz_px, index=fidx)
            except Exception:
                pass
        _PILLOW_FONT_CACHE[cache_key] = pil_font

    return _PILLOW_FONT_CACHE.get(cache_key), box_px


def _count_slide_lines_rendered(slide) -> int:
    """Pillow 실제 폰트 메트릭으로 줄 수 계산.
    Pillow 또는 폰트 파일 미존재 시 _count_slide_lines()로 폴백."""
    pil_font, box_px = _get_slide_render_params(slide)
    if pil_font is None:
        return _count_slide_lines(slide)

    shape = _find_content_shape(slide)
    if shape is None:
        return 0

    total = 0
    for para in shape.text_frame.paragraphs:
        text = para.text
        if not text.strip():
            continue
        total += _rendered_wrap_count(text, pil_font, box_px)
    return total


_COM_DISABLED = [False]
_COM_MISMATCH_COUNT: dict = {}
_COM_ATEXIT_REGISTERED = [False]


def _count_slide_lines_verified(prs, slide) -> int:
    """_count_slide_lines_rendered()의 Pillow 추정치가 경계값(LINES_PER_SLIDE)일
    때만 PowerPoint COM 실측으로 재확인한다. COM 불가/실패 시 Pillow 값을 그대로
    반환(영구 폴백) — 어떤 실패 경로도 예외를 밖으로 내보내지 않는다.

    2026-08-04 발견 버그: 예전에는 슬라이드별 불일치 횟수가 일정 캡을 넘으면 이
    함수 자체가 그 슬라이드에 한해 COM을 다시 묻지 않고 Pillow 값을 영구히
    반환했다. 문제는 이 캡이 "한 번의 수정 시도"가 아니라 섹션 전체 처리 동안
    누적되는 전역 카운터였다는 것 — 한 슬라이드가 앞선 무관한 작업에서 이미
    3번 불일치를 겪었다면, 그 뒤 실제로는 성공하지 못한 수정 시도조차 COM을
    건너뛰고 Pillow의 (틀린) 값을 "실측값"인 것처럼 반환해 거짓 성공을
    보고했다(제1독서 슬라이드 19가 실제로는 10줄인데 9줄로 "성공" 처리된
    사례로 실측 확인). 이제는 조건을 만족하는 한 매번 실제로 COM에 묻는다 —
    무한 재시도 방지는 개별 수정 시도 쪽(`_split_and_adjust_via_com`의
    max_adjust, `_rebalance_reading_slides_post_write`의 스윕 횟수 상한)에서
    책임진다."""
    pil_lines = _count_slide_lines_rendered(slide)
    if pil_lines != LINES_PER_SLIDE or _COM_DISABLED[0]:
        return pil_lines
    if not _com_verification_enabled():
        return pil_lines

    try:
        import ppt_com_verify as com
    except ImportError:
        _COM_DISABLED[0] = True
        print('  [경고] pywin32(PowerPoint COM) 미설치 — 이후 Pillow 추정치만으로 줄 수를 계산합니다.')
        return pil_lines

    # probe 생성(_build_com_probe_pptx)은 python-pptx만 쓰는 순수 파이썬 코드라
    # ComVerificationUnavailable이 아닌 예외(예: 복사된 슬라이드에 content shape가
    # 없어 ValueError, 디스크 오류로 OSError 등)를 낼 수 있다. 이는 COM 자체의
    # 문제가 아니라 이 슬라이드 하나에 국한된 문제이므로, COM을 전역적으로
    # 비활성화하지 않고 이번 호출만 Pillow로 폴백한다.
    try:
        probe_path, shape_idx = _build_com_probe_pptx(prs, slide)
    except Exception as e:
        print(f'  [경고] COM 검증용 probe 생성 실패({e}) — 이 슬라이드는 Pillow 추정치를 사용합니다.')
        return pil_lines

    com_failed = False
    try:
        real_lines = com.count_slide_lines(str(probe_path), shape_idx)
    except com.ComVerificationUnavailable as e:
        _COM_DISABLED[0] = True
        print(f'  [경고] PowerPoint COM 실측 실패({e}) — 이후 Pillow 추정치만으로 줄 수를 계산합니다.')
        com_failed = True
    # probe pptx는 프로세스당 1개 경로를 재사용(_com_probe_path)하므로 매 호출 후
    # 삭제하지 않는다 — 정리는 atexit(_cleanup_probe_file)이 프로세스 종료 시 처리한다.

    if not _COM_ATEXIT_REGISTERED[0]:
        # PowerPoint Application 인스턴스는 프로세스 종료 시 반드시 Quit()되어야 한다.
        # atexit는 등록 역순(LIFO)으로 실행되므로, 위 _build_com_probe_pptx()가 이미
        # 등록한 임시 파일 정리(_cleanup_probe_file)보다 반드시 "뒤에" 등록해야
        # PowerPoint가 먼저 종료되고(파일 잠금 해제) 그 다음에 파일이 삭제된다.
        # 반대 순서로 등록하면 Quit() 전에 삭제를 시도해 파일이 잠긴 채로 남을 수 있다.
        atexit.register(com.shutdown)
        _COM_ATEXIT_REGISTERED[0] = True

    if com_failed:
        return pil_lines

    if real_lines != pil_lines:
        sid = slide.slide_id
        _COM_MISMATCH_COUNT[sid] = _COM_MISMATCH_COUNT.get(sid, 0) + 1
        print(f'  [경고] 줄 수 불일치 감지: Pillow={pil_lines}줄, COM 실측={real_lines}줄 '
              f'→ COM 값 채택 (해당 슬라이드 누적 {_COM_MISMATCH_COUNT[sid]}회)')
    return real_lines


def _split_para_at_lines(p_elem, keep_lines: int, pil_font, box_px: float):
    """단락 XML 요소를 word-wrap 기준 keep_lines 줄에서 분리.

    p_elem을 keep_lines 줄에 맞게 수정하고, 나머지 텍스트를 담은
    새 a:p 요소를 반환. 분리 불필요하거나 불가하면 None 반환.

    run 개수를 2개(절 번호+본문)로 가정하지 않는다. 여러 절이 하나의 논리 단락으로
    병합된 경우(예: continuation 절) 절 번호 run이 단락 중간에도 나타날 수 있으므로,
    분리 지점이 어느 run에 속하는지 실제로 찾아 그 run의 서식(rPr, 오렌지색 포함)을
    그대로 유지한 채 텍스트만 자른다."""
    from copy import deepcopy
    A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

    runs = p_elem.findall(f'{{{A}}}r')
    if not runs:
        return None
    run_texts = [
        (r.find(f'{{{A}}}t').text or '') if r.find(f'{{{A}}}t') is not None else ''
        for r in runs
    ]
    full_text = ''.join(run_texts)
    if not full_text.strip():
        return None

    # word-wrap 시뮬레이션 → 줄별 단어 목록
    words = full_text.split(' ')
    line_buckets = []
    cur_words: list = []
    cur_w = 0.0
    for word in words:
        ww = pil_font.getlength(word)
        if cur_words:
            sw = pil_font.getlength(' ')
            if cur_w + sw + ww > box_px:
                line_buckets.append(cur_words)
                cur_words = [word]
                cur_w = ww
            else:
                cur_words.append(word)
                cur_w += sw + ww
        else:
            cur_words = [word]
            cur_w = ww
    if cur_words:
        line_buckets.append(cur_words)

    if len(line_buckets) <= keep_lines:
        return None  # 분리 불필요

    first_text = ' '.join(w for bucket in line_buckets[:keep_lines] for w in bucket)
    rest_text  = ' '.join(w for bucket in line_buckets[keep_lines:]  for w in bucket)
    if not rest_text:
        return None

    # first_text/rest_text는 full_text를 공백 기준으로 재분할한 것이므로,
    # 경계는 full_text[len(first_text)] 위치의 공백이다.
    split_idx = len(first_text)
    rest_start = split_idx + 1 if full_text[split_idx:split_idx + 1] == ' ' else split_idx

    def _locate(pos):
        """full_text상의 문자 위치 → (run 인덱스, run 내부 offset)."""
        acc = 0
        for i, t in enumerate(run_texts):
            if pos <= acc + len(t):
                return i, pos - acc
            acc += len(t)
        return len(run_texts) - 1, len(run_texts[-1])

    front_idx, front_off = _locate(split_idx)
    rest_idx, rest_off = _locate(rest_start)

    # 수정 전에 먼저 deepcopy → rest_p는 원본 run 서식을 그대로 유지
    rest_p = deepcopy(p_elem)
    rest_runs = rest_p.findall(f'{{{A}}}r')

    # 앞부분(p_elem): front_idx까지 run 유지, 그 run은 앞쪽 글자만 남김 (서식은 그 run 고유의 것)
    for r in runs[front_idx + 1:]:
        p_elem.remove(r)
    t_front = runs[front_idx].find(f'{{{A}}}t')
    if t_front is not None:
        t_front.text = run_texts[front_idx][:front_off]

    # 뒷부분(rest_p): rest_idx부터 run 유지, 그 run은 뒤쪽 글자만 남김 (서식은 원본 그대로)
    for r in rest_runs[:rest_idx]:
        rest_p.remove(r)
    t_rest = rest_runs[rest_idx].find(f'{{{A}}}t')
    if t_rest is not None:
        t_rest.text = run_texts[rest_idx][rest_off:]

    return rest_p


def _restore_para_from_backup(p_elem, backup):
    """_split_para_at_lines()로 분리(mutate)된 p_elem을, 분리 전 deepcopy해 둔
    backup 상태로 되돌린다. 자식 요소(run 등)를 전부 backup의 복사본으로
    교체한다 — backup은 이미 올바른 스키마 순서였으므로 순서도 그대로 보존된다."""
    for child in list(p_elem):
        p_elem.remove(child)
    for child in backup:
        p_elem.append(copy.deepcopy(child))


def _split_and_adjust_via_com(prs, cur_slide, p_elem, keep: int, pil_font, box_px: float,
                                place_rest, remove_rest, max_adjust: int = 2):
    """p_elem을 keep 줄에서 분리해 place_rest(rest_p)로 배치한 뒤, cur_slide의
    실제 줄 수를 COM으로 재확인한다.

    _split_para_at_lines()는 "어디서 자를지"를 여전히 Pillow 워드랩으로
    계산하므로, COM 실측이 목표(LINES_PER_SLIDE)와 다르면 분리 지점 자체가
    편향된 것이다. keep을 늘리면(p_elem에 더 많이 남기면) cur_slide에 남는
    줄 수가 늘고, 줄이면 준다 — 이 관계는 p_elem이 cur_slide에 남는 경우
    (초과분 분리)와 다음 슬라이드에서 넘어와 cur_slide에 합쳐지는 경우
    (부족분 흡수) 모두 동일하다. 불일치가 남으면 배치를 되돌리고 keep을
    ±1 조정해 재분리하기를 최대 max_adjust회 반복한다.

    max_adjust회를 다 써도 여전히 LINES_PER_SLIDE를 초과하면(즉 오버플로가
    남으면), 이 시도 전체를 되돌리고 (None, 원래 줄 수)를 반환한다 — 이
    함수가 존재하는 목적 자체가 "9줄로 맞추다가 실수로 넘치는 것"을 막는
    것이므로, 못 맞출 바에는 아무것도 안 하느니만 못하다. 반대로 미달(9줄
    미만)로 끝나는 것은 넘침이 아니므로 그대로 받아들인다 — 안 채워진 줄은
    미관상 아쉬울 뿐 내용이 잘리는 문제가 아니다.

    반환: (최종 rest_p 또는 None, cur_slide의 최종 실제 줄 수)
    """
    # 실패(포기) 시 반환하는 줄 수는 호출부 어디에서도 쓰이지 않는다(모든 호출부는
    # rest_p is None이면 자체적으로 롤백/스킵 처리하지 실제 값을 참조하지 않는다) —
    # 그래서 되돌린 뒤 값을 다시 측정하는 불필요한 COM 호출을 하지 않고 마지막으로
    # 알고 있던 값을 그대로 반환한다.
    original_backup = copy.deepcopy(p_elem)

    def _give_up(last_actual):
        if rest_p is not None:
            remove_rest(rest_p)
        _restore_para_from_backup(p_elem, original_backup)
        return None, last_actual

    rest_p = _split_para_at_lines(p_elem, keep, pil_font, box_px)
    if rest_p is None:
        return None, None
    place_rest(rest_p)

    actual = _count_slide_lines_verified(prs, cur_slide)
    for _ in range(max_adjust):
        if actual == LINES_PER_SLIDE:
            break
        keep += 1 if actual < LINES_PER_SLIDE else -1
        if keep <= 0:
            return _give_up(actual)
        remove_rest(rest_p)
        _restore_para_from_backup(p_elem, original_backup)
        new_rest = _split_para_at_lines(p_elem, keep, pil_font, box_px)
        if new_rest is None:
            rest_p = None
            return _give_up(actual)
        rest_p = new_rest
        place_rest(rest_p)
        actual = _count_slide_lines_verified(prs, cur_slide)

    if actual > LINES_PER_SLIDE:
        return _give_up(actual)

    return rest_p, actual


def _rebalance_reading_slides_post_write(prs, content_start: int, n_content: int, label: str) -> int:
    """독서/복음 본문 슬라이드 기록 후 실제 줄 수 검증 및 재조정.

    비마지막 슬라이드가 LINES_PER_SLIDE 미만(7·8줄) 또는 초과(10줄 이상)이면
    인접 슬라이드와 단락을 이동하여 LINES_PER_SLIDE에 맞춤.
    마지막 슬라이드는 건드리지 않음.
    반환: overflow로 새로 삽입한 슬라이드 수
    """
    if n_content < 2:
        return 0
    n_inserted = 0
    _COM_MISMATCH_COUNT.clear()  # 섹션(제1독서/제2독서/복음)마다 통계용 카운트를 새로 시작

    def _content_paras(slide):
        shape = _find_content_shape(slide)
        if shape is None:
            return []
        return [p for p in shape.text_frame.paragraphs if p.text.strip()]

    def _get_txBody(slide):
        shape = _find_content_shape(slide)
        return shape.text_frame._txBody if shape else None

    # 렌더링 파라미터 (단락 분리 시 사용) — 같은 섹션 내 슬라이드는 동일 shape
    _pil_font, _box_px = _get_slide_render_params(prs.slides[content_start])

    changed = True
    _sweep_guard = 0
    _MAX_SWEEPS = 50  # 무한 루프 방지용 안전장치(정상 케이스는 훨씬 적은 스윕으로 수렴)
    while changed:
        _sweep_guard += 1
        if _sweep_guard > _MAX_SWEEPS:
            print(f'  [경고] [{label}] 재조정이 {_MAX_SWEEPS}회 스윕 내에 수렴하지 않아 중단합니다.')
            break
        changed = False
        for i in range(n_content - 1):  # 마지막 슬라이드 제외
            cur_slide = prs.slides[content_start + i]
            nxt_slide = prs.slides[content_start + i + 1]
            lines = _count_slide_lines_verified(prs, cur_slide)

            if lines == LINES_PER_SLIDE:
                continue

            if lines > LINES_PER_SLIDE:
                excess = lines - LINES_PER_SLIDE
                cur_paras = _content_paras(cur_slide)
                cur_txBody = _get_txBody(cur_slide)
                nxt_txBody = _get_txBody(nxt_slide)

                if len(cur_paras) <= 1:
                    # 단락 1개 → 단락 분리로만 해결 가능
                    if _pil_font is not None:
                        last_p = cur_paras[0]._p
                        last_p_lines = _rendered_wrap_count(cur_paras[0].text, _pil_font, _box_px)
                        keep = last_p_lines - excess
                        if keep > 0:
                            def _place_rest1(rp):
                                first_nxt_p = nxt_txBody.find(qn('a:p'))
                                if first_nxt_p is not None:
                                    first_nxt_p.addprevious(rp)
                                else:
                                    nxt_txBody.append(rp)

                            def _remove_rest1(rp):
                                nxt_txBody.remove(rp)

                            rest_p, new_lines = _split_and_adjust_via_com(
                                prs, cur_slide, last_p, keep, _pil_font, _box_px,
                                _place_rest1, _remove_rest1,
                            )
                            if rest_p is not None:
                                print(f'  [{label}] 슬라이드 {i+1}: {lines}줄 → {new_lines}줄 (단락 분리 후 COM 조정)')
                                changed = True
                                continue
                    print(f'  [{label}] 경고 슬라이드 {i+1}: {lines}줄 (단락 1개, 분리 불가)')
                    continue

                # 마지막 단락을 다음 슬라이드 앞으로 이동 시도
                p_elem = cur_paras[-1]._p
                cur_txBody.remove(p_elem)
                first_nxt_p = nxt_txBody.find(qn('a:p'))
                if first_nxt_p is not None:
                    first_nxt_p.addprevious(p_elem)
                else:
                    nxt_txBody.append(p_elem)
                new_lines = _count_slide_lines_verified(prs, cur_slide)
                nxt_new_lines = _count_slide_lines_verified(prs, nxt_slide)

                # 다음 슬라이드가 마지막이 아닌데도 overflow → 롤백 후 단락 분리 시도
                if nxt_new_lines > LINES_PER_SLIDE and i < n_content - 2:
                    nxt_txBody.remove(p_elem)
                    cur_all_rb = cur_txBody.findall(qn('a:p'))
                    if cur_all_rb:
                        cur_all_rb[-1].addnext(p_elem)
                    else:
                        cur_txBody.append(p_elem)

                    # 단락 분리: 현재 슬라이드에 LINES_PER_SLIDE줄까지 채우고 나머지를 다음으로
                    if _pil_font is not None:
                        last_p = cur_paras[-1]._p
                        last_p_lines = _rendered_wrap_count(cur_paras[-1].text, _pil_font, _box_px)
                        keep = last_p_lines - excess
                        if keep > 0:
                            def _place_rest2(rp):
                                first_nxt_p2 = nxt_txBody.find(qn('a:p'))
                                if first_nxt_p2 is not None:
                                    first_nxt_p2.addprevious(rp)
                                else:
                                    nxt_txBody.append(rp)

                            def _remove_rest2(rp):
                                nxt_txBody.remove(rp)

                            rest_p, new_lines = _split_and_adjust_via_com(
                                prs, cur_slide, last_p, keep, _pil_font, _box_px,
                                _place_rest2, _remove_rest2,
                            )
                            if rest_p is not None:
                                print(f'  [{label}] 슬라이드 {i+1}: {lines}줄 → {new_lines}줄 (단락 분리 후 COM 조정, 나머지 → 다음)')
                                changed = True
                                continue

                    # 인접 슬라이드로 이동/분리 모두 불가 (누적 초과분이 이웃 슬라이드로
                    # 흡수되지 않는 경우) → 현재 위치 뒤에 새 슬라이드를 삽입해 초과분을 옮긴다.
                    insert_slide_copy(prs, content_start + i + 1, content_start + i)
                    new_slide = prs.slides[content_start + i + 1]
                    new_txBody = _get_txBody(new_slide)
                    if new_txBody is not None:
                        for _p in list(new_txBody.findall(qn('a:p'))):
                            new_txBody.remove(_p)
                        moved = 0
                        while _count_slide_lines_verified(prs, cur_slide) > LINES_PER_SLIDE:
                            cur_paras_ins = _content_paras(cur_slide)
                            if len(cur_paras_ins) <= 1:
                                break
                            p_elem2 = cur_paras_ins[-1]._p
                            cur_txBody.remove(p_elem2)
                            first_p = new_txBody.find(qn('a:p'))
                            if first_p is not None:
                                first_p.addprevious(p_elem2)
                            else:
                                new_txBody.append(p_elem2)
                            moved += 1
                        new_lines = _count_slide_lines_verified(prs, cur_slide)
                        added_lines = _count_slide_lines_verified(prs, new_slide)
                        print(f'  [{label}] 슬라이드 {i+1}: {lines}줄 → {new_lines}줄 (새 슬라이드 삽입: {added_lines}줄, {moved}단락 이동)')
                        n_content += 1
                        n_inserted += 1
                        changed = True
                    # 슬라이드 인덱스가 밀렸으므로 이번 pass는 중단하고 처음부터 재스캔
                    break

                print(f'  [{label}] 슬라이드 {i+1}: {lines}줄 → {new_lines}줄 (마지막 단락 → 다음={nxt_new_lines}줄)')
                changed = True

            else:  # lines < LINES_PER_SLIDE
                needed = LINES_PER_SLIDE - lines
                # 다음 슬라이드의 첫 단락을 이동해본 뒤 렌더링 줄 수로 판단 (초과 시 롤백)
                nxt_paras = _content_paras(nxt_slide)
                if len(nxt_paras) == 0:
                    continue  # 다음 슬라이드에 가져올 내용 자체가 없음
                if len(nxt_paras) == 1:
                    # 다음 슬라이드(종료 텍스트 병합 슬라이드 등)에 단락이 하나뿐이어도,
                    # 현재 슬라이드가 마지막이 아니라면 9줄을 채워야 하므로 단락을 분리해
                    # 앞부분만 가져오고 나머지는 다음 슬라이드에 남긴다.
                    if _pil_font is None:
                        continue
                    p_elem = nxt_paras[0]._p
                    first_p_lines = _rendered_wrap_count(nxt_paras[0].text, _pil_font, _box_px)
                    if first_p_lines <= needed:
                        continue  # 통째로 가져가면 다음 슬라이드가 완전히 비게 됨 → 건드리지 않음
                    nxt_txBody = _get_txBody(nxt_slide)
                    cur_txBody = _get_txBody(cur_slide)
                    nxt_txBody.remove(p_elem)
                    cur_all_p = cur_txBody.findall(qn('a:p'))
                    if cur_all_p:
                        cur_all_p[-1].addnext(p_elem)
                    else:
                        cur_txBody.append(p_elem)

                    def _place_rest3(rp):
                        first_nxt_p = nxt_txBody.find(qn('a:p'))
                        if first_nxt_p is not None:
                            first_nxt_p.addprevious(rp)
                        else:
                            nxt_txBody.append(rp)

                    def _remove_rest3(rp):
                        nxt_txBody.remove(rp)

                    rest_p, new_lines = _split_and_adjust_via_com(
                        prs, cur_slide, p_elem, needed, _pil_font, _box_px,
                        _place_rest3, _remove_rest3,
                    )
                    if rest_p is None:
                        # 분리 불가 → p_elem을 원위치(다음 슬라이드)로 되돌린다
                        cur_txBody.remove(p_elem)
                        nxt_txBody.append(p_elem)
                        continue
                    print(f'  [{label}] 슬라이드 {i+1}: {lines}줄 → {new_lines}줄 (다음 슬라이드 단일 단락 분리 흡수)')
                    changed = True
                    continue
                p_elem = nxt_paras[0]._p
                nxt_txBody = _get_txBody(nxt_slide)
                cur_txBody = _get_txBody(cur_slide)
                nxt_txBody.remove(p_elem)
                cur_all_p = cur_txBody.findall(qn('a:p'))
                if cur_all_p:
                    cur_all_p[-1].addnext(p_elem)
                else:
                    cur_txBody.append(p_elem)
                new_lines = _count_slide_lines_verified(prs, cur_slide)
                if new_lines > LINES_PER_SLIDE:
                    # 흡수 시 초과 → 원위치 후 단락 분리 시도
                    cur_txBody.remove(p_elem)
                    first_nxt_p = nxt_txBody.find(qn('a:p'))
                    if first_nxt_p is not None:
                        first_nxt_p.addprevious(p_elem)
                    else:
                        nxt_txBody.append(p_elem)
                    # 단락 분리: 다음 슬라이드 첫 단락에서 needed줄만 가져옴
                    if _pil_font is not None:
                        first_p_lines = _rendered_wrap_count(nxt_paras[0].text, _pil_font, _box_px)
                        if first_p_lines > needed:
                            nxt_txBody.remove(p_elem)
                            cur_all_p2 = cur_txBody.findall(qn('a:p'))
                            if cur_all_p2:
                                cur_all_p2[-1].addnext(p_elem)
                            else:
                                cur_txBody.append(p_elem)

                            def _place_rest4(rp):
                                first_nxt_p2 = nxt_txBody.find(qn('a:p'))
                                if first_nxt_p2 is not None:
                                    first_nxt_p2.addprevious(rp)
                                else:
                                    nxt_txBody.append(rp)

                            def _remove_rest4(rp):
                                nxt_txBody.remove(rp)

                            rest_p, new_lines = _split_and_adjust_via_com(
                                prs, cur_slide, p_elem, needed, _pil_font, _box_px,
                                _place_rest4, _remove_rest4,
                            )
                            if rest_p is not None:
                                print(f'  [{label}] 슬라이드 {i+1}: {lines}줄 → {new_lines}줄 (다음 단락 분리 흡수 후 COM 조정)')
                                changed = True
                                continue
                            # 분리 불가 → p_elem을 원위치(다음 슬라이드)로 되돌린다.
                            # nxt_txBody에는 이미 nxt_paras[1:]가 남아 있으므로 append하면
                            # 순서가 뒤바뀐다 — 반드시 맨 앞에 다시 삽입해야 한다.
                            cur_txBody.remove(p_elem)
                            first_nxt_p3 = nxt_txBody.find(qn('a:p'))
                            if first_nxt_p3 is not None:
                                first_nxt_p3.addprevious(p_elem)
                            else:
                                nxt_txBody.append(p_elem)
                    continue
                print(f'  [{label}] 슬라이드 {i+1}: {lines}줄 → {new_lines}줄 (다음 슬라이드 첫 단락 흡수)')
                changed = True

    # 마지막 콘텐츠 슬라이드가 LINES_PER_SLIDE 초과인 경우 새 슬라이드 삽입
    last_idx = content_start + n_content - 1
    last_lines = _count_slide_lines_verified(prs, prs.slides[last_idx])
    if last_lines > LINES_PER_SLIDE:
        last_paras_check = _content_paras(prs.slides[last_idx])
        if len(last_paras_check) > 1:
            # 과분리 방지: 마지막 단락 이동 후 남는 줄이 LINES_PER_SLIDE//2 이하이면
            # 이동분이 남은 분보다 많은 불균형 분리 → 새 슬라이드 삽입 건너뜀
            _skip_split = False
            if _pil_font is not None:
                _last_para_lines = _rendered_wrap_count(
                    last_paras_check[-1].text, _pil_font, _box_px
                )
                _remaining_est = last_lines - _last_para_lines
                if _remaining_est <= LINES_PER_SLIDE // 2:
                    print(f'  [{label}] 마지막 슬라이드 {n_content}: {last_lines}줄 초과이나 단락 이동 시 {_remaining_est}줄만 남아 불균형 분리 → 유지')
                    _skip_split = True
            if not _skip_split:
                insert_slide_copy(prs, last_idx + 1, last_idx)
                n_inserted += 1
                new_slide = prs.slides[last_idx + 1]
                # new_txBody를 한 번만 조회하여 재사용 (루프 내 재조회 시 빈 shape 인식 실패 방지)
                new_txBody = _get_txBody(new_slide)
                if new_txBody is not None:
                    for _p in list(new_txBody.findall(qn('a:p'))):
                        new_txBody.remove(_p)
                last_txBody = _get_txBody(prs.slides[last_idx])
                moved = 0
                while _count_slide_lines_verified(prs, prs.slides[last_idx]) > LINES_PER_SLIDE:
                    cur_paras2 = _content_paras(prs.slides[last_idx])
                    if len(cur_paras2) <= 1:
                        break
                    if new_txBody is None:
                        break
                    p_elem = cur_paras2[-1]._p
                    last_txBody.remove(p_elem)
                    first_p = new_txBody.find(qn('a:p'))
                    if first_p is not None:
                        first_p.addprevious(p_elem)
                    else:
                        new_txBody.append(p_elem)
                    moved += 1
                new_last_lines = _count_slide_lines_verified(prs, prs.slides[last_idx])
                added_lines = _count_slide_lines_verified(prs, prs.slides[last_idx + 1]) if new_txBody is not None else 0
                print(f'  [{label}] 마지막 슬라이드 {n_content}: {last_lines}줄 → {new_last_lines}줄 (새 슬라이드 삽입: {added_lines}줄, {moved}단락 이동)')

    # ── 연속 단편 병합: 문장 미완성 단락 + 절 번호 없는 다음 단락 → 하나로 합치기 ──
    # _split_para_at_lines 분리 후 rebalance 이동으로 같은 슬라이드에 놓인
    # 연속 단편("하고" / "불렀다." 등)을 단락 간격 없이 한 줄로 이어 붙인다.
    _A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

    def _run_is_orange(r):
        rPr = r.find(f'{{{_A}}}rPr')
        if rPr is None:
            return False
        sf = rPr.find(f'{{{_A}}}solidFill')
        if sf is None:
            return False
        cl = sf.find(f'{{{_A}}}srgbClr')
        return cl is not None and cl.get('val', '').upper() == 'FFC000'

    def _has_verse_run(p):
        rr = p.findall(f'{{{_A}}}r')
        return bool(rr) and _run_is_orange(rr[0])

    def _p_text(p):
        return ''.join(
            (r.find(f'{{{_A}}}t').text or '')
            for r in p.findall(f'{{{_A}}}r')
            if r.find(f'{{{_A}}}t') is not None
        )

    def _ends_sentence(p):
        t = _p_text(p).rstrip()
        return bool(t) and t[-1] in '.!?'

    for si in range(n_content):
        slide_m = prs.slides[content_start + si]
        txBody_m = _get_txBody(slide_m)
        if txBody_m is None:
            continue
        merged_any = True
        while merged_any:
            merged_any = False
            m_paras = _content_paras(slide_m)
            for j in range(len(m_paras) - 1):
                p_c = m_paras[j]._p
                p_n = m_paras[j + 1]._p
                if not _ends_sentence(p_c) and not _has_verse_run(p_n):
                    # 공백 런 추가 후 p_n의 런을 p_c로 이동, p_n 제거
                    c_runs = p_c.findall(f'{{{_A}}}r')
                    if c_runs:
                        space_r = copy.deepcopy(c_runs[-1])
                        space_t = space_r.find(f'{{{_A}}}t')
                        if space_t is not None:
                            space_t.text = ' '
                        _para_append_run(p_c, space_r)
                    for nr in list(p_n.findall(f'{{{_A}}}r')):
                        p_n.remove(nr)
                        _para_append_run(p_c, nr)
                    txBody_m.remove(p_n)
                    print(f'  [{label}] 슬라이드 {si+1}: 연속 단편 병합 (단락 {j+1}+{j+2})')
                    merged_any = True
                    break  # m_paras 변경됐으므로 재시작

    # ── 병합 후 재검증: 단락 병합으로 텍스트가 재배치(reflow)되어
    #    LINES_PER_SLIDE 아래로 떨어진 슬라이드를 다음 슬라이드에서 다시 흡수해 보충 ──
    changed = True
    _sweep_guard = 0
    _MAX_SWEEPS = 50  # 무한 루프 방지용 안전장치(정상 케이스는 훨씬 적은 스윕으로 수렴)
    while changed:
        _sweep_guard += 1
        if _sweep_guard > _MAX_SWEEPS:
            print(f'  [경고] [{label}] 재조정이 {_MAX_SWEEPS}회 스윕 내에 수렴하지 않아 중단합니다.')
            break
        changed = False
        for i in range(n_content - 1):  # 마지막 슬라이드 제외
            cur_slide = prs.slides[content_start + i]
            nxt_slide = prs.slides[content_start + i + 1]
            lines = _count_slide_lines_verified(prs, cur_slide)
            if lines >= LINES_PER_SLIDE:
                continue

            needed_lines = LINES_PER_SLIDE - lines
            nxt_paras = _content_paras(nxt_slide)
            if len(nxt_paras) <= 1:
                continue

            p_elem = nxt_paras[0]._p
            nxt_txBody = _get_txBody(nxt_slide)
            cur_txBody = _get_txBody(cur_slide)
            nxt_txBody.remove(p_elem)
            cur_all_p = cur_txBody.findall(qn('a:p'))
            if cur_all_p:
                cur_all_p[-1].addnext(p_elem)
            else:
                cur_txBody.append(p_elem)
            new_lines = _count_slide_lines_verified(prs, cur_slide)

            if new_lines > LINES_PER_SLIDE:
                # 흡수 시 초과 → 원위치 후 단락 분리 시도
                cur_txBody.remove(p_elem)
                first_nxt_p = nxt_txBody.find(qn('a:p'))
                if first_nxt_p is not None:
                    first_nxt_p.addprevious(p_elem)
                else:
                    nxt_txBody.append(p_elem)
                if _pil_font is not None:
                    first_p_lines = _rendered_wrap_count(nxt_paras[0].text, _pil_font, _box_px)
                    if first_p_lines > needed_lines:
                        nxt_txBody.remove(p_elem)
                        cur_all_p2 = cur_txBody.findall(qn('a:p'))
                        if cur_all_p2:
                            cur_all_p2[-1].addnext(p_elem)
                        else:
                            cur_txBody.append(p_elem)

                        def _place_rest5(rp):
                            first_nxt_p2 = nxt_txBody.find(qn('a:p'))
                            if first_nxt_p2 is not None:
                                first_nxt_p2.addprevious(rp)
                            else:
                                nxt_txBody.append(rp)

                        def _remove_rest5(rp):
                            nxt_txBody.remove(rp)

                        rest_p, new_lines = _split_and_adjust_via_com(
                            prs, cur_slide, p_elem, needed_lines, _pil_font, _box_px,
                            _place_rest5, _remove_rest5,
                        )
                        if rest_p is not None:
                            print(f'  [{label}] 슬라이드 {i+1}: {lines}줄 → {new_lines}줄 (병합 후 보충: 다음 단락 분리 흡수 후 COM 조정)')
                            changed = True
                        else:
                            # 분리 불가 → p_elem을 원위치(다음 슬라이드)로 되돌린다.
                            # nxt_txBody에는 이미 nxt_paras[1:]가 남아 있으므로 append하면
                            # 순서가 뒤바뀐다 — 반드시 맨 앞에 다시 삽입해야 한다.
                            cur_txBody.remove(p_elem)
                            first_nxt_p3 = nxt_txBody.find(qn('a:p'))
                            if first_nxt_p3 is not None:
                                first_nxt_p3.addprevious(p_elem)
                            else:
                                nxt_txBody.append(p_elem)
                continue

            print(f'  [{label}] 슬라이드 {i+1}: {lines}줄 → {new_lines}줄 (병합 후 보충: 다음 슬라이드 첫 단락 흡수)')
            changed = True

    return n_inserted



def replace_reading_slides(prs, content_start: int, content_end: int,

                            units_pages: list, template_idx: int,

                            line_spacing: float = None, merge_threshold: int = 5,

                            label: str = '') -> int:

    """

    독서/복음 콘텐츠 슬라이드를 교체.

    맨 뒤에 '주님의 말씀입니다.' 전용 슬라이드(TYPE B)가 있으면 보존하고

    본문 슬라이드는 그 앞에 삽입.

    units_pages: layout_units_on_slides() 결과

    template_idx: 새 슬라이드 복제 기준 슬라이드 인덱스

    반환: 슬라이드 수 변화 (양수 = 추가됨)

    """

    needed = len(units_pages)

    existing = content_end - content_start



    # 맨 뒤 연속된 종료 슬라이드 감지 ('주님의 말씀입니다.' 등이 있는 슬라이드)

    # 본문 박스가 없는 슬라이드뿐 아니라, 본문 박스와 종료 텍스트가 공존하는 슬라이드도 포함

    n_ending = 0

    for i in range(content_end - 1, content_start - 1, -1):

        if _has_ending_text(prs.slides[i]):

            n_ending += 1

        else:

            break



    n_usable = existing - n_ending

    insert_pos = content_end - n_ending  # 종료 슬라이드 앞에 삽입



    # 본문 슬라이드 수 조정

    if needed > n_usable:

        for k in range(needed - n_usable):

            insert_slide_copy(prs, insert_pos + k, template_idx)

            _set_slide_bg_black(prs.slides[insert_pos + k], prs)

    elif needed < n_usable:

        for _ in range(n_usable - needed):

            delete_slide(prs, content_start + needed)



    # 각 본문 슬라이드에 콘텐츠 설정 (종료 슬라이드는 건드리지 않음)

    for page_i, page_units in enumerate(units_pages):

        slide = prs.slides[content_start + page_i]

        shape = _find_content_shape(slide)

        if shape:

            _set_reading_text(shape.text_frame, page_units, line_spacing=line_spacing)



    # 기록 후 실제 줄 수 검증 및 재조정 (비마지막 슬라이드 7·8·10줄 → 9줄)
    n_rebalance_inserted = 0
    if needed > 1:
        n_rebalance_inserted = _rebalance_reading_slides_post_write(prs, content_start, needed, label)

    # 종료 슬라이드의 본문 텍스트박스 비우기 (참조 PPT 잔여 내용 제거)
    # n_rebalance_inserted: overflow로 삽입된 슬라이드 수만큼 ending 슬라이드 위치가 밀림

    for i in range(content_start + needed + n_rebalance_inserted,
                   content_start + needed + n_rebalance_inserted + n_ending):

        shape = _find_content_shape(prs.slides[i])

        if shape:

            _clear_text_frame(shape.text_frame)



    total_new = needed + n_rebalance_inserted + n_ending

    # 마지막 본문 슬라이드의 줄 수가 merge_threshold 이하이면 ending shape를 본문 슬라이드로 통합
    # (post-write 재조정으로 슬라이드가 추가/재배치될 수 있으므로 units_pages[-1]이 아닌
    #  실제 마지막 본문 슬라이드의 렌더링된 줄 수를 확인해야 한다)
    if n_ending > 0 and units_pages and needed > 0:
        last_content_idx = content_start + needed + n_rebalance_inserted - 1
        last_page_lines = _count_slide_lines(prs.slides[last_content_idx])
        if last_page_lines <= merge_threshold:
            ENDING_KW = ('주님의 말씀입니다', '◎ 하느님', '◎ 그리스도님')
            ending_slide_idx = content_start + needed + n_rebalance_inserted
            ending_slide = prs.slides[ending_slide_idx]
            last_slide = prs.slides[content_start + needed + n_rebalance_inserted - 1]
            spTree = last_slide.shapes._spTree
            for shape in ending_slide.shapes:
                if shape.has_text_frame and any(kw in shape.text_frame.text for kw in ENDING_KW):
                    spTree.append(copy.deepcopy(shape._element))
            for i in range(content_start + needed + n_rebalance_inserted + n_ending - 1,
                           content_start + needed + n_rebalance_inserted - 1, -1):
                delete_slide(prs, i)
            total_new -= n_ending

    return total_new - existing





# ─────────────────────────────────────────────────────────────────────────────

# 제목 슬라이드

# ─────────────────────────────────────────────────────────────────────────────



def _align_ending_slides_to_제2독서(prs, sections: dict):

    """제1독서와 복음의 ending 슬라이드 텍스트박스 위치/크기를 제2독서 기준으로 맞춤."""

    ENDING_KW = ('주님의 말씀입니다', '◎ 하느님', '◎ 그리스도님')



    if '제2독서_start' not in sections or '제2독서_end' not in sections:

        return



    # 제2독서 ending 슬라이드 탐색

    ref_slide = None

    for i in range(sections['제2독서_end'] - 1, sections['제2독서_start'] - 1, -1):

        if _has_ending_text(prs.slides[i]):

            ref_slide = prs.slides[i]

            break



    if ref_slide is None:

        return



    # 제2독서 ending 슬라이드의 텍스트박스 위치 수집

    ref_pos = {}

    for shape in ref_slide.shapes:

        if shape.name == 'BlackBg':

            continue

        if not shape.has_text_frame:

            continue

        txt = shape.text_frame.text.strip()

        key = 'ending' if any(kw in txt for kw in ENDING_KW) else 'content'

        if key not in ref_pos:

            ref_pos[key] = (shape.left, shape.top, shape.width, shape.height)



    if not ref_pos:

        return



    # 제1독서, 복음 ending 슬라이드에 동일 위치 적용

    for start_key, end_key in [('제1독서_start', '제1독서_end'), ('복음_start', '복음_end')]:

        if start_key not in sections or end_key not in sections:

            continue

        for i in range(sections[end_key] - 1, sections[start_key] - 1, -1):

            if not _has_ending_text(prs.slides[i]):

                break

            slide = prs.slides[i]

            for shape in slide.shapes:

                if shape.name == 'BlackBg':

                    continue

                if not shape.has_text_frame:

                    continue

                txt = shape.text_frame.text.strip()

                key = 'ending' if any(kw in txt for kw in ENDING_KW) else 'content'

                # 본문+ending 통합 슬라이드의 content shape는 위치 조정 제외
                if key == 'content' and txt:

                    continue

                if key in ref_pos:

                    l, top, w, h = ref_pos[key]

                    shape.left = l

                    shape.top = top

                    shape.width = w

                    shape.height = h





def _reposition_merged_ending_shapes(prs, sections: dict):

    """본문+종료 텍스트가 통합된 슬라이드의 '주님의 말씀입니다' 텍스트박스를

    실제 본문 마지막 줄 아래 한 줄 여백 위치로 재조정한다.

    _align_ending_slides_to_제2독서 실행 후에 호출해야 한다."""

    ENDING_KW = ('주님의 말씀입니다', '◎ 하느님', '◎ 그리스도님')



    for start_key, end_key in [

        ('제1독서_start', '제1독서_end'),

        ('제2독서_start', '제2독서_end'),

        ('복음_start', '복음_end'),

    ]:

        if start_key not in sections or end_key not in sections:

            continue

        s, e = sections[start_key], sections[end_key]

        for i in range(e - 1, s - 1, -1):

            slide = prs.slides[i]

            if not _has_ending_text(slide):

                break

            content_shape = _find_content_shape(slide)

            if content_shape is None:

                continue

            # 실제 본문 줄 수 계산 (빈 슬라이드는 건너뜀)

            line_count = 0

            for para in content_shape.text_frame.paragraphs:

                text = para.text.strip()

                if text:

                    line_count += _wrap_line_count(text)

            if line_count == 0:

                continue

            # ending shape 탐색 및 위치 재조정

            for shape in slide.shapes:

                if shape.has_text_frame and any(kw in shape.text_frame.text for kw in ENDING_KW):

                    line_height = content_shape.height // LINES_PER_SLIDE

                    shape.top = content_shape.top + (line_count + 2) * line_height

                    break
