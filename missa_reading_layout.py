"""
missa_to_ppt.py 모듈 분리 리팩토링 Phase 3 — 독서/복음 레이아웃 엔진.

절 파싱(parse_into_verse_units)부터 슬라이드 단위 채우기(replace_reading_slides — 각 슬라이드의
실제 본문 상자에서 줄 시작 위치를 PowerPoint COM(불가 시 Pillow)으로 실측해 정확히 9줄씩 확정),
실제 렌더링 줄 수 계산과 종료 슬라이드 위치 정렬까지 독서/복음 처리 파이프라인 전체.
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





def _wrap_line_count(text: str) -> int:
    """단어 경계 word-wrap 시뮬레이션으로 줄 수 계산.

    CHARS_PER_LINE 근사라 폰트 메트릭이 없을 때의 폴백 추정이다(종료 병합 판정·종료 상자 배치 계산에 사용).
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


def _merge_continuation_units(units: list) -> list:
    """is_continuation=True인 unit을 이전 unit에 합쳐 하나의 논리 단락으로 만든다.

    절 중간에서 문장이 이어지는지 여부는 언어/렌더링 방식과 무관한 순수 텍스트 판단이라 한글·영문이
    이 병합 단계를 공유한다(`_reading_paras_from_units`가 이 결과로 슬라이드 분할용 문단을 만든다)."""
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

    return merged


def _pil_wrap_spans(text: str, pil_font, box_px: float) -> list:
    """단어 경계 word-wrap을 Pillow 실측 폭으로 수행해 (start, end) 오프셋 리스트로 반환.

    (구 _rendered_wrap_count와 동일한) 그리디 word-wrap 알고리즘(누적 폭이 box_px를 넘기 전까지
    단어를 계속 붙임)이되, 재구성한 문자열이 아니라 원본 슬라이스 오프셋을 돌려준다 — 절 번호
    오렌지 위치(extra_verses)를 원본 텍스트 좌표로 그대로 계산해야 하므로, `cur + ' ' + word`로
    재조립한 문자열은 원본 공백 패턴과 어긋나 `.index()` 매칭이 불안정해질 수 있다(CLAUDE.md
    "옛 run 길이로 통짜 재배치" 계열 함정과 같은 이유로 재구성 대신 슬라이싱을 쓴다)."""
    if not text.strip():
        return []
    tokens = [(m.start(), m.end()) for m in re.finditer(r'\S+', text)]
    if not tokens:
        return []
    spans = []
    cur_start, cur_end = tokens[0]
    for ts, te in tokens[1:]:
        candidate = text[cur_start:te]
        if pil_font.getlength(candidate) <= box_px:
            cur_end = te
        else:
            spans.append((cur_start, cur_end))
            cur_start, cur_end = ts, te
    spans.append((cur_start, cur_end))
    return spans


def _set_reading_text(tf, units: list, line_spacing: float = None, align: str = None):

    """독서/복음 콘텐츠를 TextFrame에 설정 (절 번호 오렌지색, 서식 보존).

    extra_verses [(pos, verse_num)] 지원: 단락 내 여러 절 번호를 오렌지색으로 처리.
    line_spacing: 줄간격 배수 (예: 1.1 = 110%). None이면 템플릿 그대로.
    align: pPr의 algn 속성(예: 'ctr'). None이면 템플릿 pPr을 그대로 두고 건드리지
    않는다(D2 — 청년미사 영문 복음 텍스트박스는 템플릿 자체에 algn이 없어(좌측 기본
    정렬) 상속만으로는 중앙 정렬이 안 돼, 호출부가 명시적으로 지정해야 한다. 성인
    독서/복음 등 기존 호출부는 이 인자를 넘기지 않아 기존 정렬을 그대로 보존한다).

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

            if line_spacing is not None or align is not None:

                pPr = new_p.find(qn('a:pPr'))

                if pPr is None:

                    pPr = pptx_parse_xml(f'<a:pPr xmlns:a="{A_NS}"/>')

                    new_p.insert(0, pPr)

                if line_spacing is not None:

                    for old_ln in pPr.findall(qn('a:lnSpc')):

                        pPr.remove(old_ln)

                    val = int(line_spacing * 100000)

                    # lnSpc는 OOXML 스키마상 pPr의 첫 번째 자식이어야 함 (spcAft 앞에 위치)
                    pPr.insert(0, pptx_parse_xml(f'<a:lnSpc xmlns:a="{A_NS}"><a:spcPct val="{val}"/></a:lnSpc>'))

                if align is not None:
                    # algn은 pPr의 속성(attribute)이지 자식 요소가 아니므로 순서 스키마와
                    # 무관하다 — set()으로 바로 지정 가능.
                    pPr.set('algn', align)

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


def _content_lnspc_factor(content_shape) -> float:
    """콘텐츠 shape 첫 단락의 줄간격(lnSpc spcPct) 배율. 없으면 1.0(단일 간격)."""
    body = content_shape.text_frame._txBody
    for p in body.findall(qn('a:p')):
        pPr = p.find(qn('a:pPr'))
        if pPr is None:
            continue
        ln = pPr.find(qn('a:lnSpc'))
        if ln is None:
            continue
        pct = ln.find(qn('a:spcPct'))
        if pct is not None and pct.get('val'):
            return int(pct.get('val')) / 100000.0
    return 1.0


def _content_line_height_emu(slide, content_shape):
    """콘텐츠 텍스트 1줄 높이를 폰트 메트릭 기반 EMU로 직접 계산.

    기존에는 box_height // LINES_PER_SLIDE로 역산했으나, 실제 본문 박스 높이는
    슬라이드마다 다르고 '항상 9줄 분량'이 아니라서 줄 수가 적은(짧은) 박스에서는
    line_height가 과소 산출돼 종료 텍스트박스가 본문과 겹쳤다. Pillow가 실측한 폰트
    글리프 높이(ascent+descent, 96dpi px)를 EMU로 환산하고 lnSpc 배율을 적용한다.

    Pillow/폰트 파일 미존재로 폰트 객체를 못 얻으면 None을 반환한다(호출부에서 기존
    box_height // LINES_PER_SLIDE 방식으로 안전 폴백 — 종료 텍스트박스가 배치 자체를
    건너뛰는 상황은 없어야 한다)."""
    pil_font, _box_px = _get_slide_render_params(slide)
    if pil_font is None:
        return None
    ascent, descent = pil_font.getmetrics()
    # _get_slide_render_params는 폰트를 96dpi px로 로드한다. 1px = 914400/96 = 9525 EMU.
    glyph_emu = (ascent + descent) * 9525
    return int(round(glyph_emu * _content_lnspc_factor(content_shape)))


def _move_ending_shape_to_next_slide(prs, slide_idx, ending_shape, line_height):
    """종료 텍스트박스가 slide_idx 한 장에 온전히 못 들어갈 때, 도형을 통째로 새
    다음 슬라이드로 옮긴다(텍스트 일부만 걸치는 분리 금지 — 종료 텍스트박스는 항상
    단일 도형이므로 도형 통째 이동을 뜻한다).

    같은 프레젠테이션 내부 슬라이드를 insert_slide_copy로 복제하므로 복제본은 이미
    올바른 배경/레이아웃을 갖는다(copy_slide_from_prs가 아니라서 배경 재설정 함정과
    무관하고, 별도 재설정도 불필요). 복제 후 본문 텍스트만 비워 종료 전용 슬라이드로
    만든다."""
    ENDING_KW = ('주님의 말씀입니다', '◎ 하느님', '◎ 그리스도님')

    insert_slide_copy(prs, slide_idx + 1, slide_idx)
    new_slide = prs.slides[slide_idx + 1]

    # 새 슬라이드는 본문+종료 사본 → 본문을 비워 종료 전용으로 만들고 종료 도형을 상단 배치
    new_content = _find_content_shape(new_slide)
    if new_content is not None:
        new_top = new_content.top + line_height  # 빈 본문 위 한 줄 여백
        _clear_text_frame(new_content.text_frame)
    else:
        new_top = line_height
    for shape in new_slide.shapes:
        if shape.has_text_frame and any(kw in shape.text_frame.text for kw in ENDING_KW):
            shape.top = new_top
            break

    # 원래 슬라이드에서 종료 도형 제거 (본문만 남긴다)
    ending_shape._element.getparent().remove(ending_shape._element)


_COM_DISABLED = [False]
_COM_MISMATCH_COUNT: dict = {}
_COM_ATEXIT_REGISTERED = [False]


def _merged_ending_body_metrics(prs, slide, content_shape) -> tuple:
    """본문+종료 통합 슬라이드에서 종료 텍스트박스를 배치하기 전, 실제 본문 줄 수와
    1줄 높이(EMU)를 함께 구한다. (line_count, line_height_emu) 반환.

    (삭제된) `_count_slide_lines_verified()`는 overflow 경계(LINES_PER_SLIDE)에 걸렸을 때만 COM에
    묻는다 — "9줄을 넘었는지"만 중요한 그 용도에선 합리적이지만, 이 함수(종료 텍스트박스
    위치 계산)는 줄 수·줄 높이 추정이 조금만 틀려도 종료 텍스트박스가 실제 마지막 줄과
    거의 겹치는 시각적 결함으로 바로 드러난다.

    2026-10-03 실사용 사례(청년미사 출력, 제2독서 종료 통합 슬라이드)에서 char-count 근사
    (`_wrap_line_count`)가 실제 5줄(PowerPoint COM `TextRange.Lines.Count` 실측 확인)인
    문단을 4줄로 오산했다 — Pillow 폰트 메트릭 word-wrap(구 `_rendered_wrap_count`)도 같은
    문단을 4줄로 오산해(동일 문단의 둘째 줄을 3줄로 과소 계산) 두 추정 경로 모두 실제보다
    적게 세었다. 이는 "Pillow가 PowerPoint보다 몇 % 작게 잰다"는 기존 보정 계수(14~18%,
    `_content_line_height_emu`가 보정하는 대상)로는 못 잡는, word-wrap 자체의 줄 수 오차다.
    줄 수뿐 아니라 1줄 높이 자체도 `_content_line_height_emu`(Pillow ascent+descent 기반)가
    실측(`TextRange.BoundHeight`)보다 14~18% 작게 나온다 — 두 오차가 겹쳐 GAP=2줄로
    설계했는데 실제로는 거의 0(≈0.09인치)에 배치됐다. 경계값 여부와 무관하게 COM이
    가능하면 항상 실측값(줄 수·BoundHeight 둘 다)을 쓴다 — 이 함수는 섹션(제1독서/제2독서/
    복음)당 병합 슬라이드가 있을 때만, 많아야 1회 호출되므로 비용이 낮다."""
    fallback_count = sum(
        _wrap_line_count(para.text)
        for para in content_shape.text_frame.paragraphs
        if para.text.strip()
    )
    fallback_height = _content_line_height_emu(slide, content_shape)
    if fallback_height is None or fallback_height <= 0:
        fallback_height = content_shape.height // LINES_PER_SLIDE

    if _COM_DISABLED[0] or not _com_verification_enabled():
        return fallback_count, fallback_height

    try:
        import ppt_com_verify as com
    except ImportError:
        _COM_DISABLED[0] = True
        return fallback_count, fallback_height

    try:
        probe_path, shape_idx = _build_com_probe_pptx(prs, slide)
    except Exception as e:
        print(f'  [경고] 종료 텍스트박스 위치 계산용 COM probe 생성 실패({e}) — 추정치를 사용합니다.')
        return fallback_count, fallback_height

    try:
        metrics = com.measure_text_metrics(str(probe_path), shape_idx)
    except com.ComVerificationUnavailable as e:
        _COM_DISABLED[0] = True
        print(f'  [경고] PowerPoint COM 실측 실패({e}) — 이후 종료 텍스트박스 위치는 추정치로 계산합니다.')
        return fallback_count, fallback_height

    if not _COM_ATEXIT_REGISTERED[0]:
        atexit.register(com.shutdown)
        _COM_ATEXIT_REGISTERED[0] = True

    real_lines = metrics['lines']
    real_height = (
        int(round(metrics['bound_height_pt'] * 12700 / real_lines))
        if real_lines > 0 else fallback_height
    )
    if real_lines != fallback_count:
        print(f'  [경고] 종료 텍스트박스 위치: 추정 {fallback_count}줄 vs COM 실측 {real_lines}줄 → COM 값 채택')
    return real_lines, real_height


# ─────────────────────────────────────────────────────────────────────────────
# 슬라이드 단위 채우기 (2026-10-03 구조 교체)
#
# 이전 구조는 본문 전체를 추정(27자/줄, Pillow 폭)으로 9줄씩 먼저 나눠 모든 슬라이드에 기록한 뒤,
# PowerPoint 실측과의 오차를 이웃 슬라이드 간 문단 이동/분리로 사후 보정했다. 추정 오차가 이웃으로
# 전파되어 보정이 수렴한다는 보장이 없었다(예: 비마지막 슬라이드가 8줄로 남거나 마지막 줄에 두 글자만
# 남는 꼬리). 지금은 슬라이드마다 "남은 본문"을 그 슬라이드의 실제 본문 상자에 놓고 실제 줄 시작 위치를
# 실측해 정확히 첫 LINES_PER_SLIDE줄까지만 확정한다 — 줄 경계에서 자르므로 앞 슬라이드 마지막 줄은 항상
# 꽉 찬 줄이고, 그리디 줄바꿈은 줄 시작에서 무기억이라 앞 조각만 따로 렌더링해도 첫 9줄이 그대로다.
# ─────────────────────────────────────────────────────────────────────────────

# 측정용 프리픽스 글자 수 하한(문단 단위로 담는다). 10번째 줄이 안 보이면 2배로 늘려 재측정한다 —
# 본문 전체를 매번 기록하면 슬라이드 수만큼 긴 텍스트를 COM에 태우게 된다.
_MEASURE_PREFIX_CHARS = 1200
_MAX_FILL_SLIDES = 500  # 무한 루프 방지용 안전장치(정상 본문은 수 장)
_VERSE_PREFIX_RE = re.compile(r'^(\d+(?:,\d+)?\s*)')


def _reading_paras_from_units(units: list) -> list:
    """parse_into_verse_units() 결과를 슬라이드 분할용 문단 목록으로 바꾼다.

    {'text': 문단 전체 텍스트, 'orange': [(start, end), ...]} — 오렌지 구간(절 번호)은 _set_reading_text가
    쓰는 것과 같은 규칙(선두 절 번호 + continuation으로 병합된 중간 절 번호)으로 문단 내 절대 오프셋에
    저장한다. run 개수를 가정하지 않으므로 어느 위치에서 잘라도 절 번호 서식이 유실되지 않는다."""
    paras = []
    for mu in _merge_continuation_units(units):
        text = mu['text']
        if not text.strip():
            continue
        ranges = []
        if mu.get('verse_num'):
            m = _VERSE_PREFIX_RE.match(text)
            if m:
                ranges.append((0, len(m.group(1))))
        for pos, _vnum in mu.get('extra_verses', []):
            if pos < len(text):
                m2 = _VERSE_PREFIX_RE.match(text[pos:])
                if m2:
                    ranges.append((pos, pos + len(m2.group(1))))
        ranges.sort()
        paras.append({'text': text, 'orange': ranges})
    return paras


def _split_paras_at(paras: list, para_idx: int, offset: int):
    """paras를 (para_idx번 문단, 문단 내 offset)에서 (앞 조각, 뒤 조각)으로 자른다.

    offset==0이면 그 문단은 통째로 뒤 조각에 들어간다. 앞 조각의 끝 공백은 제거하고(줄 끝 공백이라
    줄 수에 영향이 없다) 뒤 조각은 줄 시작 문자부터 그대로 둔다. 오렌지 구간은 양쪽에서 잘라 옮긴다."""
    front = [dict(p, orange=list(p['orange'])) for p in paras[:para_idx]]
    rest = []
    p = paras[para_idx]
    text = p['text']
    if offset <= 0:
        rest.append({'text': text, 'orange': list(p['orange'])})
    else:
        ftxt = text[:offset].rstrip()
        if ftxt:
            front.append({
                'text': ftxt,
                'orange': [(s, min(e, len(ftxt))) for s, e in p['orange'] if s < len(ftxt)],
            })
        rtxt = text[offset:]
        if rtxt.strip():
            rest.append({
                'text': rtxt,
                'orange': [(max(s - offset, 0), e - offset) for s, e in p['orange'] if e > offset],
            })
    rest.extend({'text': q['text'], 'orange': list(q['orange'])} for q in paras[para_idx + 1:])
    return front, rest


def _paras_to_units(paras: list) -> list:
    """문단 목록을 _set_reading_text 입력 형식으로 바꾼다. 오렌지 구간은 모두 extra_verses
    (pos, 절번호)로 표현한다(선두 절 번호도 pos 0) — _set_reading_text가 같은 정규식으로 run을 나눈다."""
    units = []
    for p in paras:
        text = p['text']
        extra = []
        for s, e in p['orange']:
            m = re.match(r'\d+(?:,\d+)?', text[s:e])
            if m and s < len(text):
                extra.append((s, m.group(0)))
        units.append({'text': text, 'verse_num': '', 'extra_verses': extra, 'new_para': True})
    return units


def _write_reading_paras(shape, paras: list, line_spacing, align):
    if paras:
        _set_reading_text(shape.text_frame, _paras_to_units(paras),
                          line_spacing=line_spacing, align=align)
    else:
        _clear_text_frame(shape.text_frame)


def _pil_line_starts(paras: list, pil_font, box_px: float) -> list:
    """Pillow word-wrap으로 본 줄 시작 위치 [(문단 인덱스, 문단 내 오프셋), ...]."""
    out = []
    for i, p in enumerate(paras):
        for j, (s, _e) in enumerate(_pil_wrap_spans(p['text'], pil_font, box_px)):
            out.append((i, 0 if j == 0 else s))
    return out


def _char_line_starts(paras: list) -> list:
    """폰트 메트릭이 전혀 없을 때만 쓰는 CHARS_PER_LINE 단어 경계 줄 시작 위치."""
    out = []
    for i, p in enumerate(paras):
        text = p['text']
        n = len(text)
        pos = 0
        while pos < n:
            out.append((i, pos))
            end = pos + CHARS_PER_LINE
            if end >= n:
                break
            space = text.rfind(' ', pos, end + 1)
            pos = space + 1 if space > pos else end
    return out


def _locate_offset(paras: list, offset: int) -> tuple:
    """'\\r'로 문단을 이은 전체 텍스트의 오프셋 → (문단 인덱스, 문단 내 오프셋)."""
    base = 0
    for i, p in enumerate(paras):
        if offset < base + len(p['text']) + 1:
            return i, offset - base
        base += len(p['text']) + 1
    return len(paras) - 1, len(paras[-1]['text'])


def _com_measure_line_starts(prs, slide, paras: list):
    """slide에 이미 기록된 paras의 (총 줄 수, 앞 LINES_PER_SLIDE+1줄의 줄 시작 위치)를 PowerPoint COM으로
    실측한다. COM을 쓸 수 없거나 응답이 텍스트와 어긋나면 None(호출부가 Pillow로 폴백) — 어떤 실패 경로도
    예외를 밖으로 내보내지 않는다. COM 자체가 안 되면(ImportError/ComVerificationUnavailable) 전역 비활성화하고,
    이 슬라이드 하나만의 문제(probe 생성 실패, 응답 불일치)면 비활성화하지 않는다."""
    if _COM_DISABLED[0] or not _com_verification_enabled():
        return None
    try:
        import ppt_com_verify as com
    except ImportError:
        _COM_DISABLED[0] = True
        print('  [경고] pywin32(PowerPoint COM) 미설치 — 이후 Pillow 추정치만으로 줄 수를 계산합니다.')
        return None
    measure = getattr(com, 'measure_line_starts', None)
    if measure is None:
        return None

    try:
        probe_path, shape_idx = _build_com_probe_pptx(prs, slide)
    except Exception as e:
        print(f'  [경고] COM 검증용 probe 생성 실패({e}) — 이 슬라이드는 Pillow 추정치를 사용합니다.')
        return None
    try:
        info = measure(str(probe_path), shape_idx, LINES_PER_SLIDE + 1)
    except com.ComVerificationUnavailable as e:
        _COM_DISABLED[0] = True
        print(f'  [경고] PowerPoint COM 실측 실패({e}) — 이후 Pillow 추정치만으로 줄 수를 계산합니다.')
        return None

    if not _COM_ATEXIT_REGISTERED[0]:
        # PowerPoint Application은 프로세스 종료 시 Quit()되어야 하고, probe 임시 파일 정리(atexit)보다
        # "뒤에" 등록해야(LIFO) PowerPoint가 먼저 종료돼 파일 잠금이 풀린 뒤 삭제된다.
        atexit.register(com.shutdown)
        _COM_ATEXIT_REGISTERED[0] = True

    full = '\r'.join(p['text'] for p in paras)
    starts, texts, total = info['starts'], info['texts'], info['lines']
    consistent = (
        bool(starts) and starts[0] == 0 and len(starts) == min(total, LINES_PER_SLIDE + 1)
        and all(a < b for a, b in zip(starts, starts[1:])) and starts[-1] < len(full)
        and len(texts) == len(starts)
        and all(full.startswith(t.rstrip('\r\n '), s) for s, t in zip(starts, texts))
    )
    if not consistent:
        print('  [경고] COM 줄 시작 응답이 기록한 텍스트와 어긋나 이 슬라이드는 Pillow 추정치를 사용합니다.')
        return None
    return total, [_locate_offset(paras, s) for s in starts]


def _measure_slide_cut(prs, slide, remaining: list, write, label: str):
    """slide의 실제 본문 상자에서 remaining의 앞 LINES_PER_SLIDE줄 뒤 절단점을 찾는다.

    반환: (절단점 (문단 인덱스, 문단 내 오프셋) 또는 None, 마지막 측정의 총 줄 수).
    None이면 남은 본문 전부가 한 슬라이드(<= LINES_PER_SLIDE줄)에 들어간다 — 이때 슬라이드에는 이미
    remaining 전체가 기록된 상태다. 절단점이 있으면 슬라이드에는 프리픽스(앞부분 문단들)가 기록돼 있으니
    호출부가 앞 조각으로 다시 기록한다."""
    cap = _MEASURE_PREFIX_CHARS
    while True:
        n, chars = 0, 0
        while n < len(remaining) and (n == 0 or chars < cap):
            chars += len(remaining[n]['text'])
            n += 1
        prefix = remaining[:n]
        write(prefix)

        pil_font, box_px = _get_slide_render_params(slide)
        measured = _com_measure_line_starts(prs, slide, prefix)
        if measured is None:
            starts = (_char_line_starts(prefix) if pil_font is None
                      else _pil_line_starts(prefix, pil_font, box_px))
            total = len(starts)
        else:
            total, starts = measured
            if pil_font is not None:
                pil_total = len(_pil_line_starts(prefix, pil_font, box_px))
                if pil_total != total:
                    sid = slide.slide_id
                    _COM_MISMATCH_COUNT[sid] = _COM_MISMATCH_COUNT.get(sid, 0) + 1
                    print(f'  [경고] 줄 수 불일치 감지: Pillow={pil_total}줄, COM 실측={total}줄 '
                          f'→ COM 값 채택 (해당 슬라이드 누적 {_COM_MISMATCH_COUNT[sid]}회)')

        if total > LINES_PER_SLIDE:
            return starts[LINES_PER_SLIDE], total
        if n == len(remaining):
            return None, total
        cap *= 2


def replace_reading_slides(prs, content_start: int, content_end: int,

                            units: list, template_idx: int,

                            line_spacing: float = None, merge_threshold: int = 5,

                            label: str = '', align: str = None,
                            normalize_page_size: bool = False) -> int:

    """
    독서/복음 콘텐츠 슬라이드를 교체.

    맨 뒤에 '주님의 말씀입니다.' 전용 슬라이드(TYPE B)가 있으면 보존하고
    본문 슬라이드는 그 앞에 삽입.

    units: parse_into_verse_units() 결과. 슬라이드마다 남은 본문을 그 슬라이드의 실제 본문 상자에서
    실측해 정확히 LINES_PER_SLIDE줄까지만 확정하고 나머지로 다음 슬라이드를 같은 방식으로 채운다
    (마지막 슬라이드만 그 이하). 슬라이드가 모자라면 template_idx를 복제해 늘리고 남으면 삭제한다.

    template_idx: 새 슬라이드 복제 기준 슬라이드 인덱스

    반환: 슬라이드 수 변화 (양수 = 추가됨)

    """

    paras = _reading_paras_from_units(units)

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



    # normalize_page_size=False(기본값)이면 템플릿 박스 크기 복사를 건너뛴다 — 성인 픽스처
    # (output/20260712/Template_20260628_...pptx)의 제2독서 콘텐츠 슬라이드 62/63은 원래부터 박스
    # 크기가 서로 달라, 이 정규화를 항상 켜면 청년미사 버그(D2)와 무관한 성인 경로의 기존 지오메트리를
    # 조용히 바꾼다(02b_review_report.md, 2026-09-24 실측). D2가 실제로 필요한 곳(청년미사 영문 복음)만
    # 호출부가 opt-in한다. 슬라이드 단위 채우기는 각 슬라이드의 "실제" 상자에서 측정하므로 이 옵션이
    # 꺼져 있어도 박스 크기가 달라서 줄 수가 어긋나지는 않는다.
    template_shape = _find_content_shape(prs.slides[template_idx]) if normalize_page_size else None

    _COM_MISMATCH_COUNT.clear()  # 섹션(제1독서/제2독서/복음)마다 통계용 카운트를 새로 시작

    remaining = paras
    n_used = 0
    k = 0
    last_total = None  # 마지막 슬라이드를 확정할 때 실측(COM/Pillow)한 줄 수 — 종료 병합 판정에 쓴다
    while k < _MAX_FILL_SLIDES:
        if k < n_usable:
            slide = prs.slides[content_start + k]
        else:
            ins = insert_pos + (k - n_usable)
            insert_slide_copy(prs, ins, template_idx)
            slide = prs.slides[ins]
            _set_slide_bg_black(slide, prs)
        k += 1
        n_used = k

        shape = _find_content_shape(slide)
        if shape is None:
            if k > n_usable:
                raise RuntimeError(
                    f'[{label}] 슬라이드 {k}: 복제한 슬라이드에서 본문 상자를 찾지 못해 남은 본문 '
                    f'{len(remaining)}개 단락을 배치할 수 없습니다')
            continue  # 본문 상자가 없는 기존 슬라이드는 건너뛴다(내용을 잃지 않고 다음 슬라이드로)

        if template_shape is not None and (
            shape.height != template_shape.height or shape.width != template_shape.width
        ):
            shape.width = template_shape.width
            shape.height = template_shape.height
            shape.left = template_shape.left
            shape.top = template_shape.top

        def _write(ps, _shape=shape):
            _write_reading_paras(_shape, ps, line_spacing, align)

        if not remaining:
            _write([])
            last_total = 0
            break

        cut, total = _measure_slide_cut(prs, slide, remaining, _write, label)
        if cut is None:
            print(f'  [{label}] 슬라이드 {k}: 마지막 슬라이드 {total}줄')
            last_total = total
            break
        front, remaining = _split_paras_at(remaining, cut[0], cut[1])
        _write(front)
        split_note = ', 마지막 단락은 줄 경계에서 분할' if cut[1] > 0 else ''
        print(f'  [{label}] 슬라이드 {k}: {LINES_PER_SLIDE}줄 확정 (단락 {len(front)}개{split_note})')

    # 필요한 만큼 채우고 남은 기존 본문 슬라이드(종료 슬라이드 앞쪽)는 삭제
    for _ in range(n_usable - n_used):

        delete_slide(prs, content_start + n_used)

    needed = n_used



    # 종료 슬라이드의 본문 텍스트박스 비우기 (참조 PPT 잔여 내용 제거)

    for i in range(content_start + needed, content_start + needed + n_ending):

        shape = _find_content_shape(prs.slides[i])

        if shape:

            _clear_text_frame(shape.text_frame)



    total_new = needed + n_ending

    # 마지막 본문 슬라이드의 줄 수가 merge_threshold 이하이면 ending shape를 본문 슬라이드로 통합
    # 판정 줄 수는 마지막 슬라이드를 확정할 때 얻은 실측값(COM, 불가 시 Pillow, 폰트도 없으면 27자 근사)이다 —
    # 구 27자/줄 추정은 라틴 본문을 약 2배 과대 추정해 실측 4~5줄도 병합하지 못했다(2026-10-03 사용자 결정).

    if n_ending > 0 and needed > 0:
        last_content_idx = content_start + needed - 1
        if last_total is not None and last_total <= merge_threshold:
            ENDING_KW = ('주님의 말씀입니다', '◎ 하느님', '◎ 그리스도님')
            ending_slide_idx = content_start + needed
            ending_slide = prs.slides[ending_slide_idx]
            last_slide = prs.slides[last_content_idx]
            spTree = last_slide.shapes._spTree
            for shape in ending_slide.shapes:
                if shape.has_text_frame and any(kw in shape.text_frame.text for kw in ENDING_KW):
                    spTree.append(copy.deepcopy(shape._element))
            for i in range(content_start + needed + n_ending - 1,
                           content_start + needed - 1, -1):
                delete_slide(prs, i)
            total_new -= n_ending

    print(f'    [{label}] 본문 {needed}개 슬라이드')
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



    GAP = 2  # 본문 마지막 줄과 종료 텍스트박스 사이 여백 줄 수

    # 섹션을 뒤(복음)부터 처리한다: 종료 텍스트박스가 슬라이드를 못 넘어가 다음
    # 슬라이드를 새로 삽입할 때(방어적 경로) 삽입은 더 높은 인덱스만 밀어내므로,
    # 아직 처리하지 않은 낮은 인덱스 섹션의 sections 값이 무효화되지 않는다.
    for start_key, end_key in [

        ('복음_start', '복음_end'),

        ('제2독서_start', '제2독서_end'),

        ('제1독서_start', '제1독서_end'),

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

            # 실제 본문 줄 수·1줄 높이 계산 (빈 슬라이드는 건너뜀). char-count/Pillow
            # 추정이 틀리면 종료 텍스트박스가 실제 마지막 줄과 거의 겹치는 결함으로 바로
            # 드러나므로, COM이 가능하면 항상 실측값을 쓴다(_merged_ending_body_metrics).

            line_count, line_height = _merged_ending_body_metrics(prs, slide, content_shape)

            if line_count == 0:

                continue

            # ending shape 탐색
            ending_shape = None
            for shape in slide.shapes:
                if shape.has_text_frame and any(kw in shape.text_frame.text for kw in ENDING_KW):
                    ending_shape = shape
                    break
            if ending_shape is None:
                continue

            # 종료 텍스트박스 자체의 시각 줄 수 (일반적으로 2줄: 사제 응답 + 회중 응답)
            ending_span = 0
            for para in ending_shape.text_frame.paragraphs:
                t = para.text.strip()
                if t:
                    ending_span += _wrap_line_count(t)
            if ending_span == 0:
                ending_span = 2

            # 본문(line_count줄) + 여백(GAP줄) + 종료(ending_span줄)이 한 슬라이드
            # 세로 공간(LINES_PER_SLIDE줄)을 넘으면 도형 통째로 다음 슬라이드로 옮긴다
            # (종료 텍스트박스는 반드시 한 슬라이드에 온전히 들어가야 한다 — 분리 금지).
            if line_count + GAP + ending_span > LINES_PER_SLIDE:
                _move_ending_shape_to_next_slide(prs, i, ending_shape, line_height)
            else:
                ending_shape.top = content_shape.top + (line_count + GAP) * line_height
