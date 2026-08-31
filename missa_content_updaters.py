"""
missa_to_ppt.py 모듈 분리 리팩토링 Phase 5 — 화답송·성가·입당송 등 섹션별 콘텐츠 갱신.

참조 PPT의 각 섹션(제목, 입당송, 화답송, 복음환호송, 영성체송, 시작기도문,
미사 후 기도, 성가 5종)을 JSON 데이터로 교체하는 update_*/replace_* 함수들.
"""
from __future__ import annotations

import copy
import re
from pathlib import Path
from xml.sax.saxutils import escape as _xml_escape

from pptx import Presentation
from pptx.oxml.ns import qn

from missa_ooxml_utils import (
    HYMN_TYPES, _replace_para_text_clone, _set_single_para_text,
    _update_book_name_after_br, insert_slide_copy, delete_slide, _set_slide_bg_black,
    _set_화답송_content_text, copy_slide_from_prs, _para_append_run, all_slide_texts,
)
from missa_sections import find_sections


# ─────────────────────────────────────────────────────────────────────────────

# 텍스트 서식 유틸 (v3: clone 방식)


def update_title_slide(prs, json_data: dict):

    slide = prs.slides[0]

    date_str = json_data['date']  # "2026-06-28"

    parts = date_str.split('-')

    year, month, day = parts[0], str(int(parts[1])), str(int(parts[2]))

    liturgy = json_data['liturgy']



    for shape in slide.shapes:

        if not shape.has_text_frame:

            continue

        tf = shape.text_frame

        full = tf.text

        if re.search(r'\d{4}년', full) and '월' in full and '일' in full:

            paras = tf.paragraphs

            if len(paras) >= 1:

                _replace_para_text_clone(paras[0], f'{year}년 {month}월 {day}일')

            if len(paras) >= 2:

                _replace_para_text_clone(paras[1], liturgy)

            break





# ─────────────────────────────────────────────────────────────────────────────

# 입당송

# ─────────────────────────────────────────────────────────────────────────────





def update_입당송(prs, json_data: dict, sections: dict):

    if '입당송' not in sections or not json_data.get('입당송'):

        return

    slide = prs.slides[sections['입당송']]

    content = '◎ \t' + json_data['입당송']['content']



    for shape in slide.shapes:

        if not shape.has_text_frame:

            continue

        t = shape.text_frame.text.strip()

        if t and '입당송' not in t and '전례문' not in t:

            _set_single_para_text(shape.text_frame, content)

            _adjust_fit_if_needed(slide, shape, prs)

            break





# ─────────────────────────────────────────────────────────────────────────────

# 독서 제목 슬라이드 업데이트

# ─────────────────────────────────────────────────────────────────────────────



def update_reading_title_slide(prs, title_idx: int, label_kw: str, reading_title: str):

    """독서/복음 제목 슬라이드에서 성서 제목 교체.

    reading_title: JSON의 title 값 (예: "에제키엘 예언서의 말씀입니다.", "루카가 전한 거룩한 복음입니다.")

    """

    slide = prs.slides[title_idx]



    for shape in slide.shapes:

        if not shape.has_text_frame:

            continue

        raw = shape.text_frame.text.replace(' ', '').replace('\x0b', '').replace('\n', '')

        if label_kw.replace(' ', '') in raw:

            for para in shape.text_frame.paragraphs:

                p = para._p

                brs = p.findall(qn('a:br'))

                if brs:

                    if _update_book_name_after_br(para, reading_title):

                        return

                elif para.text.strip() and label_kw.replace(' ', '') not in para.text.replace(' ', ''):

                    _replace_para_text_clone(para, reading_title)

                    return





def update_복음_title_slide(prs, title_idx: int, gospel_title: str):

    """복음 제목 슬라이드에서 복음 제목 교체.

    gospel_title: JSON의 title 값 (예: "루카가 전한 거룩한 복음입니다.")

    """

    slide = prs.slides[title_idx]

    for shape in slide.shapes:

        if not shape.has_text_frame:

            continue

        for para in shape.text_frame.paragraphs:

            t = para.text.strip()

            if '전한 거룩한 복음입니다' in t:

                p = para._p

                runs = p.findall(qn('a:r'))

                if runs:

                    template_r = runs[0]

                    new_r = copy.deepcopy(template_r)

                    t_el = new_r.find(qn('a:t'))

                    if t_el is not None:

                        t_el.text = gospel_title

                    for r in runs:

                        p.remove(r)

                    for br in p.findall(qn('a:br')):

                        p.remove(br)

                    p.append(new_r)

                return





# ─────────────────────────────────────────────────────────────────────────────

# 화답송

# ─────────────────────────────────────────────────────────────────────────────



def _update_화답송_title_in_slide(slide, new_title: str):

    for shape in slide.shapes:

        if not shape.has_text_frame:

            continue

        t = shape.text_frame.text

        if '화 답 송' in t:

            para = shape.text_frame.paragraphs[0]

            _replace_para_text_clone(para, f'화 답 송   {new_title}')

            return





def update_화답송(prs, json_data: dict, sections: dict, 화답송_pptx_path, 화답송_img_path=None, is_sunday: bool = True):

    if '화답송_start' not in sections or not json_data.get('화답송'):

        return



    화 = json_data['화답송']

    title = 화['title']

    content = 화['content']



    start = sections['화답송_start']

    cur_end = sections['화답송_end']

    cur_total = cur_end - start

    text_tmpl = start + 1 if cur_total > 1 else start



    if not is_sunday:

        # 평일미사 포맷: 슬라이드마다 ◎ 후렴 + ○ 절 (2단락)

        segments = [s.strip() for s in content.split('\n') if s.strip()]

        if not segments:

            return

        refrain_raw = segments[0]   # ◎ 후렴

        verses = segments[1:]       # ○ 절들

        # 후렴 텍스트 정규화 (탭 정렬)

        if refrain_raw.startswith('◎ '):

            refrain_text = '◎\t' + refrain_raw[2:]

        else:

            refrain_text = refrain_raw

        needed = len(verses)    # 슬라이드 수 = 절 수

        # 기존 범위 뒤에 텍스트 템플릿 복제본 needed개 삽입

        for k in range(needed):

            insert_slide_copy(prs, cur_end + k, text_tmpl)

        # 기존 슬라이드 삭제

        for i in range(cur_end - 1, start - 1, -1):

            delete_slide(prs, i)

        # 새 슬라이드 [start, start+needed-1] 채우기

        for i in range(needed):

            idx = start + i

            _update_화답송_title_in_slide(prs.slides[idx], title)

            _set_slide_bg_black(prs.slides[idx], prs)

            verse = verses[i]

            verse_clean = verse.strip()

            if verse_clean.startswith('○ '):

                verse_text = '○\t' + verse_clean[2:]

            elif verse_clean.startswith('◎ '):

                verse_text = '◎\t' + verse_clean[2:]

            else:

                verse_text = verse_clean

            for shape in prs.slides[idx].shapes:

                if not shape.has_text_frame:

                    continue

                t = shape.text_frame.text.strip()

                if not t or '화 답 송' in t or '전례문' in t or 'Responsorial' in t:

                    continue

                _set_화답송_content_text(shape.text_frame, refrain_text, verse_text)

                _adjust_fit_if_needed(prs.slides[idx], shape, prs)

                break

        return



    # 주일미사: 악보(n+1) + 텍스트(n) = 2n+1 슬라이드

    segments = [s.strip() for s in content.split('\n') if s.strip()]

    verses = segments[1:] if len(segments) > 1 else segments

    n = len(verses)



    # 필요 슬라이드: 악보(n+1) + 텍스트(n) = 2n+1

    needed = max(1, 2 * n + 1)



    # 슬라이드 수 조정 (텍스트 템플릿 = start+1)

    if needed > cur_total:

        for k in range(needed - cur_total):

            tmpl = start if (cur_total + k) % 2 == 0 else text_tmpl

            insert_slide_copy(prs, cur_end + k, tmpl)

    elif needed < cur_total:

        for i in range(cur_end - 1, start + needed - 1, -1):

            delete_slide(prs, i)



    # 화답송 악보 로드: 수작업 PPT 우선(기존 동작 완전 보존), 없으면 원본 사진에서 자동

    # 생성한다(2차 마일스톤 — missa_psalm_score_image.render_화답송_score_slide 재사용,

    # 설계서 §4.3). title은 위에서 이미 json_data['화답송']['title']로 구해져 있으므로

    # 그대로 넘기면 JSON 제목 자동 연결이 끝난다(추가 가공 불필요).

    악보_prs = None

    if 화답송_pptx_path and Path(화답송_pptx_path).exists():

        try:

            악보_prs = Presentation(str(화답송_pptx_path))

        except Exception as e:

            print(f'  [경고] 화답송 악보 PPT 로드 실패: {e}')

    elif 화답송_img_path and Path(화답송_img_path).exists():

        try:

            from missa_psalm_score_image import render_화답송_score_slide

            악보_prs = render_화답송_score_slide(화답송_img_path, title)

        except Exception as e:

            print(f'  [경고] 화답송 악보 이미지 처리 실패: {e}')



    # 각 슬라이드 업데이트

    for i in range(needed):

        idx = start + i



        if i % 2 == 0:

            # 악보 슬라이드: 화답송 악보 PPT에서 복사
            # copy_slide_from_prs()가 원본 배경/서식을 이미 그대로 복사하므로
            # _set_slide_bg_black()로 덮어쓰지 않는다 (원본 서식 보존)

            if 악보_prs:

                delete_slide(prs, idx)

                copy_slide_from_prs(prs, idx, 악보_prs, 0)

            _update_화답송_title_in_slide(prs.slides[idx], title)

        else:

            # 텍스트 슬라이드: ○ 절만 표시 (주일미사는 단일 단락)

            _update_화답송_title_in_slide(prs.slides[idx], title)

            _set_slide_bg_black(prs.slides[idx], prs)

            verse_idx = i // 2

            if verse_idx < n:

                verse = verses[verse_idx]

                verse_clean = verse.strip()

                if verse_clean.startswith('○ '):

                    verse_text = '○\t' + verse_clean[2:]

                else:

                    verse_text = verse_clean

                for shape in prs.slides[idx].shapes:

                    if not shape.has_text_frame:

                        continue

                    t = shape.text_frame.text.strip()

                    if t and '화 답 송' not in t and '전례문' not in t and 'Responsorial' not in t:

                        _set_single_para_text(shape.text_frame, verse_text)

                        _adjust_fit_if_needed(prs.slides[idx], shape, prs)

                        break





# ─────────────────────────────────────────────────────────────────────────────

# 복음환호송

# ─────────────────────────────────────────────────────────────────────────────



def update_복음환호송(prs, json_data: dict, sections: dict):

    if '복음환호송' not in sections or not json_data.get('복음환호송'):

        return

    slide = prs.slides[sections['복음환호송']]

    content = json_data['복음환호송']['content']

    lines = [l.strip() for l in content.split('\n') if l.strip()]



    for shape in slide.shapes:

        if not shape.has_text_frame:

            continue

        t = shape.text_frame.text.strip()

        if t and '복음 환호송' not in t and '환호송' not in t and '전례문' not in t and 'ALLELUIA' not in t:

            tf = shape.text_frame

            txBody = tf._txBody

            existing_paras = txBody.findall(qn('a:p'))



            # 템플릿 단락:

            #   para[0] = 첫 ◎ (lnSpc=110%, tabLst 없음)

            #   para[1] = ○ 구절 (lnSpc=120%, tabLst 있음)

            #   para[2] = 마지막 ◎ (lnSpc=120%, tabLst 있음)

            tmpl_allel0 = existing_paras[0] if len(existing_paras) > 0 else None

            tmpl_verse  = existing_paras[1] if len(existing_paras) > 1 else tmpl_allel0

            tmpl_allel2 = existing_paras[2] if len(existing_paras) > 2 else tmpl_allel0



            for p in existing_paras:

                txBody.remove(p)



            from pptx.oxml import parse_xml as pptx_parse_xml

            A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'



            allel_count = 0

            total_allel = sum(1 for l in lines if not l.startswith('○'))



            for line in lines:

                is_verse = line.startswith('○')



                if not is_verse:

                    # ◎ 줄: 적절한 템플릿 단락을 그대로 deepcopy (run 구조·폰트 보존)

                    allel_count += 1

                    tmpl = tmpl_allel2 if (allel_count == total_allel and tmpl_allel2 is not None) else tmpl_allel0

                    if tmpl is None:

                        new_p = pptx_parse_xml(f'<a:p xmlns:a="{A_NS}"/>')

                        _para_append_run(new_p, pptx_parse_xml(f'<a:r xmlns:a="{A_NS}"><a:t>{_xml_escape(line)}</a:t></a:r>'))

                        txBody.append(new_p)

                        continue

                    new_p = copy.deepcopy(tmpl)

                    # 템플릿 텍스트와 다를 때만 run 업데이트

                    tmpl_text = ''.join(

                        (r.find(qn('a:t')).text or '')

                        for r in tmpl.findall(qn('a:r'))

                        if r.find(qn('a:t')) is not None

                    )

                    if tmpl_text.strip() != line.strip():

                        # ◎ [body][punct] 패턴으로 파싱해서 run 텍스트만 교체

                        m = re.match(r'^(◎\s+)(.+?)([.!?]?)$', line)

                        if m:

                            body, punct = m.group(2), m.group(3)

                            new_runs = new_p.findall(qn('a:r'))

                            if len(new_runs) >= 4:   # "◎" " " body punct

                                new_runs[2].find(qn('a:t')).text = body

                                new_runs[3].find(qn('a:t')).text = punct

                            elif len(new_runs) >= 3: # "◎ " body punct

                                new_runs[1].find(qn('a:t')).text = body

                                new_runs[2].find(qn('a:t')).text = punct

                    txBody.append(new_p)



                else:

                    # ○ 줄: para[1] deepcopy 후 run[0]("○ \t") 유지,

                    #        run[1].text = 내용(바탕체), run[2+] 제거

                    if tmpl_verse is None:

                        new_p = pptx_parse_xml(f'<a:p xmlns:a="{A_NS}"/>')

                        text = '○ \t' + line[2:] if line.startswith('○ ') else line

                        _para_append_run(new_p, pptx_parse_xml(f'<a:r xmlns:a="{A_NS}"><a:t>{_xml_escape(text)}</a:t></a:r>'))

                        txBody.append(new_p)

                        continue

                    new_p = copy.deepcopy(tmpl_verse)

                    verse_runs = new_p.findall(qn('a:r'))

                    verse_content = line[2:] if line.startswith('○ ') else line

                    if len(verse_runs) >= 2:

                        # run[0] = "○ \t" 유지, run[1].text = 새 내용, run[2+] 제거

                        t_el = verse_runs[1].find(qn('a:t'))

                        if t_el is not None:

                            t_el.text = verse_content

                        for r in verse_runs[2:]:

                            new_p.remove(r)

                    elif len(verse_runs) == 1:

                        # run[0]만 있으면 텍스트 전체 교체

                        t_el = verse_runs[0].find(qn('a:t'))

                        if t_el is not None:

                            t_el.text = '○ \t' + verse_content

                    txBody.append(new_p)

            _adjust_fit_if_needed(slide, shape, prs)

            break





# ─────────────────────────────────────────────────────────────────────────────

# 텍스트 오버플로우 자동 조정

# ─────────────────────────────────────────────────────────────────────────────



def _find_last_row_top(slide, prs) -> int:

    """슬라이드·레이아웃·마스터에서 전례문 텍스트 상단 위치 반환 (없으면 슬라이드 높이)."""

    LAST_ROW_KW = ('전례문', '한국천주교', '©')

    best = prs.slide_height

    sources = [slide.shapes]

    try:

        sources.append(slide.slide_layout.shapes)

    except Exception:

        pass

    try:

        sources.append(slide.slide_layout.slide_master.shapes)

    except Exception:

        pass

    for shapes in sources:

        for shape in shapes:

            if not hasattr(shape, 'text_frame'):

                continue

            try:

                t = shape.text_frame.text

            except Exception:

                continue

            if any(kw in t for kw in LAST_ROW_KW):

                best = min(best, shape.top)

    return best




def _shape_first_run_font_size_emu(shape) -> int:

    """도형 첫 번째 run 폰트 크기 (EMU). 없으면 32pt."""

    for para in shape.text_frame.paragraphs:

        for run in para.runs:

            if run.font.size:

                return run.font.size

    return int(32 * 12700)




def _shape_first_para_line_spacing_pct(shape) -> int:

    """도형 첫 번째 단락 줄간격 (spcPct 단위). 없으면 100000 (= 100%)."""

    for para in shape.text_frame.paragraphs:

        pPr = para._p.find(qn('a:pPr'))

        if pPr is None:

            continue

        lnSpc = pPr.find(qn('a:lnSpc'))

        if lnSpc is None:

            continue

        el = lnSpc.find(qn('a:spcPct'))

        if el is not None:

            return int(el.get('val', '100000'))

    return 100000




def _set_shape_all_para_line_spacing(shape, spc_pct: int):

    """도형 모든 단락 줄간격을 배수(spcPct) 형식으로 설정."""

    from pptx.oxml import parse_xml as pptx_parse_xml

    A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

    for para in shape.text_frame.paragraphs:

        pPr = para._p.find(qn('a:pPr'))

        if pPr is None:

            pPr = pptx_parse_xml(f'<a:pPr xmlns:a="{A_NS}"/>')

            para._p.insert(0, pPr)

        for old in pPr.findall(qn('a:lnSpc')):

            pPr.remove(old)

        pPr.insert(0, pptx_parse_xml(

            f'<a:lnSpc xmlns:a="{A_NS}"><a:spcPct val="{spc_pct}"/></a:lnSpc>'

        ))




def _set_shape_all_run_font_size(shape, font_size_pt: float):

    """도형 모든 run 폰트 크기 설정 (pt)."""

    sz = str(int(font_size_pt * 100))

    for para in shape.text_frame.paragraphs:

        for run in para.runs:

            rPr = run._r.find(qn('a:rPr'))

            if rPr is not None:

                rPr.set('sz', sz)




def _estimate_text_lines(text: str, font_size_emu: int, box_width_emu: int) -> int:

    """단어 경계 줄바꿈 시뮬레이션으로 추정 줄 수 반환."""

    font_pt = font_size_emu / 12700

    # 한국어 문자 평균 너비: em 크기의 약 0.92배 (전각 문자 기준)

    chars_per_line = max(1.0, (box_width_emu / 12700) / (font_pt * 0.92))

    total = 0

    for para_text in text.split('\n'):

        if not para_text.strip():

            total += 1

            continue

        words = para_text.split(' ')

        lines = 1

        cur = 0.0

        for i, word in enumerate(words):

            wl = len(word) + (1 if i > 0 else 0)

            if cur + wl > chars_per_line and cur > 0:

                lines += 1

                cur = float(len(word))

            else:

                cur += wl

        total += lines

    return max(1, total)




def _adjust_fit_if_needed(slide, content_shape, prs):

    """텍스트가 전례문 줄 또는 슬라이드 마스터 텍스트와 겹치면 자동 조정.

    (a) 줄간격을 배수 1.0으로 변경하고 검증.

    (b) 그래도 겹치면 폰트 크기를 단계적으로 축소.

    """

    boundary = _find_last_row_top(slide, prs)

    available = boundary - content_shape.top

    if available <= 0:

        return

    font_emu = _shape_first_run_font_size_emu(content_shape)

    spc_pct = _shape_first_para_line_spacing_pct(content_shape)

    text = content_shape.text_frame.text

    box_w = content_shape.width

    font_pt = font_emu / 12700

    num_lines = _estimate_text_lines(text, font_emu, box_w)

    est_h = int(num_lines * font_pt * (spc_pct / 100000) * 12700)

    if est_h <= available:

        return

    # (a) 줄간격을 배수 1.0으로 변경

    _set_shape_all_para_line_spacing(content_shape, 100000)

    est_h = int(num_lines * font_pt * 12700)

    if est_h <= available:

        print(f'    [높이조정] 줄간격 → 배수 1.0 (줄수={num_lines}, 추정={est_h//12700:.0f}pt)')

        return

    # (b) 폰트 크기 단계적 축소

    while font_pt > 16:

        font_pt -= 1

        num_lines = _estimate_text_lines(text, int(font_pt * 12700), box_w)

        est_h = int(num_lines * font_pt * 12700)

        if est_h <= available:

            break

    _set_shape_all_run_font_size(content_shape, font_pt)

    print(f'    [높이조정] 줄간격 1.0 + 폰트 {font_pt:.0f}pt (줄수={num_lines}, 추정={est_h//12700:.0f}pt)')




# ─────────────────────────────────────────────────────────────────────────────

# 영성체송

# ─────────────────────────────────────────────────────────────────────────────



def update_영성체송(prs, json_data: dict, sections: dict):

    if '영성체송' not in sections or not json_data.get('영성체송'):

        return

    slide = prs.slides[sections['영성체송']]

    content = json_data['영성체송']['content']



    for shape in slide.shapes:

        if not shape.has_text_frame:

            continue

        t = shape.text_frame.text.strip()

        if t and '영성체송' not in t and '전례문' not in t and 'COMMUNION' not in t:

            _set_single_para_text(shape.text_frame, content)

            _adjust_fit_if_needed(slide, shape, prs)

            break





# ─────────────────────────────────────────────────────────────────────────────

# 시작기도문 교체

# ─────────────────────────────────────────────────────────────────────────────



def replace_시작기도문(prs, 시작기도_path: Path, sections: dict):

    if not 시작기도_path or '시작기도_start' not in sections:

        return



    start = sections['시작기도_start']

    end = sections['시작기도_end']

    n_existing = end - start



    src_prs = Presentation(str(시작기도_path))

    n_src = len(src_prs.slides)



    # 기존 시작기도 슬라이드 삭제 (뒤에서부터)

    for i in range(end - 1, start - 1, -1):

        delete_slide(prs, i)



    # 새 슬라이드 삽입 후 배경 검정으로 설정

    for i in range(n_src):

        copy_slide_from_prs(prs, start + i, src_prs, i)

        _set_slide_bg_black(prs.slides[start + i], prs)



    print(f'  시작기도: {n_existing}장 → {n_src}장')

    return n_src - n_existing





def replace_미사후기도(prs, path: Path, sections: dict) -> int:

    """파견 성가 이후 기존 미사 후 기도 슬라이드를 새 내용으로 교체 (평일미사 전용).

    - path가 None이면 아무것도 하지 않음 (기존 미사 후 기도 보존)
    - path가 있으면 기존 미사 후 기도 슬라이드를 삭제하고 새 내용으로 삽입
    - 위치: 파견 성가(divider+content) → blank divide → 미사 후 기도 → blank → 파견
    """

    if not path or not path.exists():

        return 0

    src_prs = Presentation(str(path))

    n_src = len(src_prs.slides)

    if n_src == 0:

        return 0

    texts = all_slide_texts(prs)

    n = len(prs.slides)

    # 파견 성가 이후 blank divide 위치 파악
    # 파견_content_end = 파견 콘텐츠 직후 인덱스 (blank이거나 콘텐츠 없는 경우 파견_content_start)

    파견_end = sections.get('파견_content_end', -1)

    if 파견_end < 0:

        파견_div = sections.get('파견_divider', -1)

        if 파견_div < 0:

            print('  [경고] 파견 성가 위치를 찾을 수 없어 미사 후 기도를 삽입할 수 없습니다.')

            return 0

        파견_end = 파견_div + 1

    # blank divide 슬라이드 위치 (파견_end)
    # 미사 후 기도 시작: blank 바로 다음

    미사후기도_start = 파견_end + 1 if 파견_end < n and not texts[파견_end].strip() else 파견_end

    if 미사후기도_start >= n:

        # 미사 후 기도 공간 없음: 파견 뒤에 직접 삽입

        for i in range(n_src):

            copy_slide_from_prs(prs, n + i, src_prs, i)

        print(f'  미사 후 기도: {n_src}장 삽입 (파견 이후 끝에 추가)')

        return n_src

    # 기존 미사 후 기도 범위 탐색 (blank 또는 dismissal slide까지)

    DISMISSAL_KW = ['미사가 끝났으니', '파견하나이다', 'DISMISSAL', 'Dismissal', '미사가 끝']

    미사후기도_end = 미사후기도_start

    for i in range(미사후기도_start, n):

        t = texts[i]

        if not t.strip():

            break

        if any(kw in t for kw in DISMISSAL_KW):

            break

        미사후기도_end = i + 1

    # 기존 미사 후 기도 슬라이드 삭제 (역순)

    n_deleted = 미사후기도_end - 미사후기도_start

    for i in range(미사후기도_end - 1, 미사후기도_start - 1, -1):

        delete_slide(prs, i)

    # 새 미사 후 기도 삽입

    for i in range(n_src):

        copy_slide_from_prs(prs, 미사후기도_start + i, src_prs, i)

    print(f'  미사 후 기도: 기존 {n_deleted}장 제거 → 새 {n_src}장 삽입 (슬라이드 {미사후기도_start + 1}부터)')

    return n_src - n_deleted




# ─────────────────────────────────────────────────────────────────────────────

# 성가 교체

# ─────────────────────────────────────────────────────────────────────────────



def _update_성가_divider_number(prs, divider_idx: int, new_number: int):

    """성가 divider 슬라이드의 번호 업데이트."""

    slide = prs.slides[divider_idx]

    for shape in slide.shapes:

        if not shape.has_text_frame:

            continue

        for para in shape.text_frame.paragraphs:

            if re.match(r'^\d+$', para.text.strip()):

                _replace_para_text_clone(para, str(new_number))

                return





def _update_prefix_in_runs(para, old_prefix: str, new_prefix: str):

    """para 앞부분 old_prefix → new_prefix 교체. run의 rPr(색상 등) 보존."""

    old_len = len(old_prefix)

    new_remaining = new_prefix

    pos = 0

    last_a_t = None

    last_after = ''



    for run in para.runs:

        a_t = run._r.find(qn('a:t'))

        if a_t is None:

            continue

        run_text = a_t.text or ''

        run_len = len(run_text)

        if pos >= old_len:

            break

        n = min(pos + run_len, old_len) - pos

        new_portion = new_remaining[:n]

        new_remaining = new_remaining[n:]

        after = run_text[n:]

        a_t.text = new_portion + after

        last_a_t = a_t

        last_after = after

        pos += run_len



    # new_prefix가 old_prefix보다 긴 경우 나머지를 마지막 처리 run에 삽입

    if new_remaining and last_a_t is not None:

        prefix_part = last_a_t.text[:-len(last_after)] if last_after else last_a_t.text

        last_a_t.text = prefix_part + new_remaining + last_after





def _update_성가_header(slide, expected_type: str, new_number: int, slide_no: int = None, n_slides: int = None):

    """성가 헤더 타입/번호 업데이트. run 서식(색상 등) 보존.

    원본 악보 파일 첫 줄의 구분 표기가 이번 주 실제 용도와 다를 수 있다
    (예: 성가 62가 '2차 봉헌'으로 인쇄돼 있지만 이번 주는 입당 성가로 쓰임).
    번호뿐 아니라 구분 라벨도 이번 주 용도(expected_type)에 맞게 고친다."""

    CANONICAL_LABEL = {
        '입당': '입당',
        '봉헌': '봉헌',
        '성체': '성체',
        '2차봉헌': '2차 봉헌',
        '파견': '파견',
    }

    MATCH_PAT = r'^(\s*)(2차\s*봉헌|입당|봉헌|성체|파견)(\s+)(\d+)'

    for shape in slide.shapes:

        if not shape.has_text_frame:

            continue

        for para in shape.text_frame.paragraphs:

            t = para.text

            m = re.match(MATCH_PAT, t)

            if not m:

                continue

            old_prefix = m.group(0)

            correct_label = CANONICAL_LABEL.get(expected_type, expected_type)

            new_prefix = m.group(1) + correct_label + m.group(3) + str(new_number)

            if old_prefix != new_prefix:

                _update_prefix_in_runs(para, old_prefix, new_prefix)

            return

    where = f' (슬라이드 {slide_no}/{n_slides})' if slide_no else ''

    print(f'  [경고] 성가 {new_number}: 첫 줄에서 구분 라벨을 찾지 못함{where}')









def replace_성가(prs, 성가_map: dict, hymn_numbers: dict, copy_scores: bool = True):

    """5종 성가 슬라이드 교체. 각 타입마다 find_sections 재호출로 정확한 위치 파악."""



    for htype in HYMN_TYPES:

        num = hymn_numbers.get(htype)

        # copy_scores=True(주일미사)일 때만 실제 악보 PPT가 필요하다. 평일미사(copy_scores=False)는

        # 악보를 복사하지 않고 번호만 갱신하므로 성가_map에 파일이 없어도 건너뛰지 않는다.

        if copy_scores and htype not in 성가_map:

            print(f'  [{htype}] 성가 PPT 없음, 건너뜀')

            continue

        if not copy_scores and num is None:

            print(f'  [{htype}] 성가 번호 없음, 건너뜀')

            continue



        # 매번 섹션 재탐색 (이전 교체로 인한 인덱스 변화 반영)

        sec = find_sections(prs)



        pptx_path = 성가_map.get(htype)



        dk = f'{htype}_divider'

        csk = f'{htype}_content_start'

        cek = f'{htype}_content_end'



        if dk not in sec:

            print(f'  [{htype}] divider 슬라이드 없음, 건너뜀')

            continue



        div_idx = sec[dk]

        cs = sec[csk]

        ce = sec[cek]

        n_existing = ce - cs



        # divider 번호 업데이트

        if num:

            _update_성가_divider_number(prs, div_idx, num)



        # 기존 콘텐츠 슬라이드 삭제

        for i in range(ce - 1, cs - 1, -1):

            delete_slide(prs, i)



        if copy_scores:

            # 새 성가 PPT 슬라이드 로드 및 삽입

            src_prs = Presentation(str(pptx_path))

            n_src = len(src_prs.slides)

            for i in range(n_src):

                copy_slide_from_prs(prs, cs + i, src_prs, i)

            # 헤더 업데이트

            if num:

                for i in range(n_src):

                    _update_성가_header(prs.slides[cs + i], htype, num, slide_no=i + 1, n_slides=n_src)

            print(f'  [{htype}] 성가 {num}: {n_existing}장 → {n_src}장')

        else:

            print(f'  [{htype}] 성가 {num}: 번호 업데이트 (악보 없음, {n_existing}장 삭제)')





