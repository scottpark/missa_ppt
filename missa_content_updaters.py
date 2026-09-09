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


def insert_공지사항(prs, path: Path, sections: dict) -> int:
    """2차봉헌 성가 콘텐츠 뒤에 공지사항 PPT 슬라이드를 순수 삽입한다(교체 아님).

    반환값: 삽입으로 늘어난 슬라이드 수(공지사항 n_src장 + 뒤 구분선 1장). 미삽입 시 0.

    replace_미사후기도와 달리 기존 슬라이드를 삭제하지 않는다(순수 삽입). 앞쪽 구분선은
    2차봉헌_content_end 자리의 기존 blank를 재사용하고, 뒤쪽 구분선 1장만 새로 추가한다.
    """
    # (a) 입력 없음 — 조용히 skip
    if not path or not Path(path).exists():
        return 0

    # (b) 2차봉헌 슬롯 없음 — find_sections는 슬롯이 없으면 키 자체를 넣지 않는다.
    #     sections['...']로 직접 인덱싱하면 KeyError로 파이프라인이 죽으므로 .get()으로 판별.
    front_blank = sections.get('2차봉헌_content_end')
    if front_blank is None:
        print('  [경고] 2차봉헌 성가 슬롯이 없어 공지사항 삽입을 건너뜁니다.')
        return 0

    # (c) 재사용할 앞쪽 구분선이 실제로 blank인지 방어적 확인(콘텐츠가 끝까지 이어져 blank가
    #     없던 예외 방지). blank는 shapes 0개라 텍스트가 ''(falsy)이다.
    texts = all_slide_texts(prs)
    n = len(prs.slides)
    if front_blank >= n or texts[front_blank].strip():
        print('  [경고] 2차봉헌 콘텐츠 뒤 구분(blank) 슬라이드를 찾을 수 없어 '
              '공지사항 삽입을 건너뜁니다.')
        return 0

    # (d) 공지사항 PPT 로드 방어 (replace_미사후기도 패턴 이식)
    try:
        src_prs = Presentation(str(path))
    except Exception as e:
        print(f'  [경고] 공지사항 PPT 로드 실패: {e}')
        return 0
    n_src = len(src_prs.slides)
    if n_src == 0:
        print('  [경고] 공지사항 PPT에 슬라이드가 없어 삽입을 건너뜁니다.')
        return 0

    # (e) 콘텐츠 삽입: 앞 구분선(front_blank) 바로 뒤부터 순서대로.
    #     다른 프레젠테이션에서 복사한 경우이므로 _set_slide_bg_black()을 부르지 않는다
    #     (copy_slide_from_prs가 이미 원본 배경을 정확히 복사 — CLAUDE.md "배경 재설정 금지").
    for i in range(n_src):
        copy_slide_from_prs(prs, front_blank + 1 + i, src_prs, i)

    # (f) 뒤쪽 구분선 1장: 같은 프레젠테이션 내 앞 구분선 blank를 복제해 콘텐츠 마지막 다음에
    #     삽입. front_blank는 (e)에서 뒤에만 삽입했으므로 인덱스가 그대로 유효하다. 원본 blank는
    #     자체 p:bg 없이 Blank 레이아웃 검정을 상속하므로 _set_slide_bg_black() 불필요.
    insert_slide_copy(prs, front_blank + 1 + n_src, front_blank)

    print(f'  공지사항: {n_src}장 삽입 + 뒤 구분 슬라이드 1장 '
          f'(슬라이드 {front_blank + 2}부터)')
    return n_src + 1




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





def _update_prefix_in_runs(para, segments):

    """para 앞부분을 세그먼트 단위로 교체한다. run의 rPr(색상 등)은 보존하고 텍스트만 바꾼다.

    segments: [(old_text, new_text), ...] — 원본 텍스트 맨 앞부터 순서대로 덮는 구간들.
    각 구간은 원본 run들에서 len(old_text)만큼의 글자를 소비하며, 신규 텍스트는 그 구간이
    차지하던 run 범위 안에서만 재배치된다.

    왜 세그먼트로 나누는가: 라벨/구분자/숫자를 하나의 문자 스트림으로 취급해 옛 run 길이
    기준으로 통짜 재배치하면, 한 구간(예: 라벨)의 글자 수가 바뀔 때 뒤 구간의 글자가 앞 run으로
    (또는 그 반대로) 밀려, 우연히 다른 색이던 run(예: 소스 제작자가 회색으로 칠한 구분자 공백)의
    서식을 엉뚱한 글자가 물려받는다. 호출부가 정규식 그룹 경계(라벨 vs 구분자+숫자)를 세그먼트로
    끊어 넘겨주면, 한 구간의 길이 변화가 다른 구간이 점유하던 run 색상을 침범하지 못한다.
    (CLAUDE.md "단락 분리·병합 함수는 run 개수를 2개로 가정하지 않는다"와 같은 계열의 함정.)
    """

    ats = []

    for run in para.runs:

        a_t = run._r.find(qn('a:t'))

        if a_t is not None:

            ats.append(a_t)

    orig = [a_t.text or '' for a_t in ats]

    pieces = [None] * len(ats)   # None=원본 유지, list=새 텍스트로 교체

    consumed = [0] * len(ats)    # 각 run에서 prefix가 소비한 글자 수(= 꼬리 시작 위치)

    ri = 0   # 현재 run 인덱스

    off = 0  # 현재 run 내 오프셋(원본 글자 기준)

    for old_seg, new_seg in segments:

        rem_old = len(old_seg)

        touched = []  # 이 세그먼트가 걸치는 (run_idx, 소비 글자 수)

        while rem_old > 0 and ri < len(orig):

            avail = len(orig[ri]) - off

            if avail <= 0:

                ri += 1

                off = 0

                continue

            take = min(avail, rem_old)

            if pieces[ri] is None:

                pieces[ri] = []

            touched.append((ri, take))

            rem_old -= take

            off += take

            consumed[ri] = off

            if off >= len(orig[ri]):

                ri += 1

                off = 0

        # new_seg를 이 세그먼트가 점유한 run들에만 분배한다.

        # 마지막 run이 남는 글자(길이 증가분)를 흡수 → 증가/감소가 세그먼트 안에 갇힌다.

        if touched:

            rem_new = new_seg

            last = len(touched) - 1

            for j, (rj, take) in enumerate(touched):

                if j == last:

                    pieces[rj].append(rem_new)

                else:

                    pieces[rj].append(rem_new[:take])

                    rem_new = rem_new[take:]

    for idx, a_t in enumerate(ats):

        if pieces[idx] is None:

            continue

        a_t.text = ''.join(pieces[idx]) + orig[idx][consumed[idx]:]





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

            correct_label = CANONICAL_LABEL.get(expected_type, expected_type)

            # 정규식 그룹 경계를 재배치의 하드 경계로 삼는다: 라벨(선행 공백+라벨)과
            # 구분자+숫자를 별개 세그먼트로 넘겨, 라벨 길이 변화가 숫자 쪽 run 색상을(또는
            # 반대 방향으로) 침범하지 못하게 한다. (bugfix_성가_label_run_color)

            label_old = m.group(1) + m.group(2)

            label_new = m.group(1) + correct_label

            number_old = m.group(3) + m.group(4)

            number_new = m.group(3) + str(new_number)

            if label_old != label_new or number_old != number_new:

                _update_prefix_in_runs(para, [(label_old, label_new), (number_old, number_new)])

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





