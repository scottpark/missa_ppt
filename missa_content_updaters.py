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
from pptx.oxml import parse_xml as pptx_parse_xml
from pptx.oxml.ns import qn

from missa_ooxml_utils import (
    HYMN_TYPES, _replace_para_text_clone, _set_single_para_text,
    _update_book_name_after_br, insert_slide_copy, delete_slide, _set_slide_bg_black,
    _set_화답송_content_text, copy_slide_from_prs, _para_append_run, all_slide_texts,
    duplicate_slide,
)
from missa_sections import find_sections, HYMN_LABEL_KW, _find_youth_title_shape


# ─────────────────────────────────────────────────────────────────────────────

# 텍스트 서식 유틸 (v3: clone 방식)


def update_title_slide(prs, json_data: dict, date_str: str = None):

    """제목 슬라이드 날짜/전례명 표시. `date_str`(선택, "YYYY-MM-DD")을 넘기면 그 값을
    표시하고, 생략하면 기존처럼 `json_data['date']`를 그대로 쓴다.

    청년미사는 콘텐츠(독서·복음 등) 조회에 +1일(토요일→일요일)을 쓰기 때문에
    `json_data['date']`가 항상 일요일이다 — 이 값을 그대로 표시하면 사용자가 실제로 입력한
    토요일 날짜와 화면에 다른 날짜가 뜬다(2026-09-26 실사용자 산출물로 재현). 콘텐츠 조회
    자체는 건드리지 않고 "화면 표시"만 호출부(`missa_to_ppt.py`)가 원래 입력 date_str을
    변환해 넘기도록 분리한다."""

    slide = prs.slides[0]

    if date_str is None:

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



def update_복음환호송(prs, json_data: dict, sections: dict, mass_type: str = 'adult'):

    if '복음환호송' not in sections or not json_data.get('복음환호송'):

        return

    slide = prs.slides[sections['복음환호송']]

    content = json_data['복음환호송']['content']

    lines = [l.strip() for l in content.split('\n') if l.strip()]

    if mass_type == 'youth':
        # 청년 템플릿의 '복음환호송' 가운데 슬라이드는 앞뒤에 별도 고정 'ALLELUIA'/
        # '알렐루야' 장식 슬라이드가 이미 있어, 이 슬라이드 자체의 구절 도형은 원래
        # ○ 구절 문단 1개뿐이다(성인 템플릿의 ◎/○/◎ 3문단 구조가 아님). 아래 루프가
        # 성인 3문단 전제로 짜여 있어(문단 1개뿐이면 세 역할이 같은 문단에 별칭돼
        # ◎ 줄이 "○ " 라벨을 물려받아 오염됨, D1), ◎ 줄은 애초에 제외하고 ○ 구절만
        # 남긴다 — 문단 개수가 원본과 같게 유지되어 강제 폰트 축소도 필요 없어진다.
        lines = [l for l in lines if l.startswith('○')]



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

                    # ○ 줄: para[1] deepcopy 후 라벨+탭 run은 유지, 본문 run.text = 내용
                    # (바탕체), 그 뒤 run 제거.
                    #
                    # 라벨/탭/본문의 run 경계는 템플릿마다 다르다 — 성인 템플릿은 "○ \t"가
                    # 라벨+탭이 한 run(run[0])에 합쳐져 있어 본문이 run[1]부터지만, 청년
                    # 템플릿은 라벨(run[0]="○ ")과 탭(run[1]="\t")이 별도 run으로 나뉘어
                    # 본문이 run[2]부터다(2026-09-26 실사용자 산출물 슬라이드 50 재현 —
                    # run 개수만 보고 항상 run[1]=본문으로 가정하면, 청년 템플릿에서는
                    # 탭 문자 자체를 담은 run에 본문 전체를 덮어써 ①탭이 사라져 tabLst
                    # 정렬이 깨지고 ②본문이 탭 run의 폰트(라벨과 같은 계열)로 렌더링돼
                    # 템플릿의 바탕체 본문 서식과 달라진다). run 개수가 아니라 리터럴 탭
                    # 문자('\t')를 담은 run의 실제 위치로 본문 시작 인덱스를 찾는다 — 그런
                    # run이 없으면(성인 템플릿처럼 라벨+탭이 한 run에 결합) 기존 그대로
                    # run[1]을 본문으로 취급해 성인 경로는 완전히 그대로 동작한다.

                    if tmpl_verse is None:

                        new_p = pptx_parse_xml(f'<a:p xmlns:a="{A_NS}"/>')

                        text = '○ \t' + line[2:] if line.startswith('○ ') else line

                        _para_append_run(new_p, pptx_parse_xml(f'<a:r xmlns:a="{A_NS}"><a:t>{_xml_escape(text)}</a:t></a:r>'))

                        txBody.append(new_p)

                        continue

                    new_p = copy.deepcopy(tmpl_verse)

                    verse_runs = new_p.findall(qn('a:r'))

                    verse_content = line[2:] if line.startswith('○ ') else line

                    tab_run_idx = next(
                        (i for i, r in enumerate(verse_runs)
                         if (r.find(qn('a:t')).text or '') == '\t'),
                        None,
                    )
                    body_idx = (tab_run_idx + 1) if tab_run_idx is not None else 1

                    if len(verse_runs) > body_idx:

                        t_el = verse_runs[body_idx].find(qn('a:t'))

                        if t_el is not None:

                            t_el.text = verse_content

                        for r in verse_runs[body_idx + 1:]:

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




# CLAUDE.md "한글 줄 수 계산" 계열 실측 기록: Pillow(및 이 함수의 char-count 근사 둘 다)의
# ascent+descent 기반 줄 높이는 실제 PowerPoint 렌더링보다 약 14~18% 작게 나온다. 청년미사
# 영성체송처럼 안전 마진이 실측 기준 2.7%(전례문까지 여유 4177044 EMU, 추정 필요 높이
# 4064000 EMU)뿐인 근소 초과 케이스는 이 과소추정 때문에 조정이 아예 트리거되지 않아
# 마지막 줄이 슬라이드 밖으로 나간다(D3). 새 임의 값 대신 이미 실측으로 확정된 이 보정치를
# 그대로 재사용한다.
_FIT_HEIGHT_SAFETY_MARGIN = 1.15


def _adjust_fit_if_needed(slide, content_shape, prs):

    """텍스트가 전례문 줄 또는 슬라이드 마스터 텍스트와 겹치면 자동 조정.

    (a) 줄간격을 배수 1.0으로 변경하고 검증.

    (b) 그래도 겹치면 폰트 크기를 단계적으로 축소.

    안전 마진(_FIT_HEIGHT_SAFETY_MARGIN)을 추정 높이에 곱해 비교한다 — 이 도형은
    <a:spAutoFit/>(텍스트에 맞춰 도형이 늘어남)이 걸려 있어, 추정이 실제보다 조금이라도
    작으면 실제 PowerPoint에서는 도형이 경계를 넘어 자란다(D3).

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

    est_h = int(num_lines * font_pt * (spc_pct / 100000) * 12700 * _FIT_HEIGHT_SAFETY_MARGIN)

    if est_h <= available:

        return

    # (a) 줄간격을 배수 1.0으로 변경

    _set_shape_all_para_line_spacing(content_shape, 100000)

    est_h = int(num_lines * font_pt * 12700 * _FIT_HEIGHT_SAFETY_MARGIN)

    if est_h <= available:

        print(f'    [높이조정] 줄간격 → 배수 1.0 (줄수={num_lines}, 추정={est_h//12700:.0f}pt)')

        return

    # (b) 폰트 크기 단계적 축소

    while font_pt > 16:

        font_pt -= 1

        num_lines = _estimate_text_lines(text, int(font_pt * 12700), box_w)

        est_h = int(num_lines * font_pt * 12700 * _FIT_HEIGHT_SAFETY_MARGIN)

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





def _update_성가_header(slide, expected_type: str, new_number: int, slide_no: int = None,
                        n_slides: int = None, label_override: str = None):

    """성가 헤더 타입/번호 업데이트. run 서식(색상 등) 보존.

    원본 악보 파일 첫 줄의 구분 표기가 이번 주 실제 용도와 다를 수 있다
    (예: 성가 62가 '2차 봉헌'으로 인쇄돼 있지만 이번 주는 입당 성가로 쓰임).
    번호뿐 아니라 구분 라벨도 이번 주 용도(expected_type)에 맞게 고친다.

    label_override: 성인미사 관례(CANONICAL_LABEL, '2차봉헌'→'2차 봉헌' 공백 있음)와
    다른 라벨 표기를 쓰는 호출부를 위한 탈출구. 청년미사 콘텐츠 헤더는 공백 없는
    '2차봉헌'을 쓴다(실측: 참조 PPT 110/111번 'Rectangle 11'="2차봉헌 야훼이레 810 ...",
    build_header_runs()가 나주노/야훼이레 경로에 쓰는 라벨도 공백 없음 — 리뷰에서 발견된
    확정 결함: 가톨릭성가 경로만 이 함수를 그대로 썼다가 공백 있는 성인미사 라벨로
    정규화되어 같은 슬라이드 안에서 표기가 갈라짐, 02b_review_report.md §확정 항목)."""

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

            correct_label = label_override if label_override is not None else CANONICAL_LABEL.get(expected_type, expected_type)

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


# ─────────────────────────────────────────────────────────────────────────────
# 청년미사 성가 5종×4출처 통합 (2단계 §6/§2)
# ─────────────────────────────────────────────────────────────────────────────

# title 슬라이드 재구성 시 쓰는 정식 라벨(내부 줄바꿈 없는 한 단락, HYMN_LABEL_KW와 동일 문자열).
_YOUTH_TITLE_LABEL = HYMN_LABEL_KW

# 콘텐츠 헤더('Rectangle 11')의 "{타입} " 접두사 — _update_성가_header()의 CANONICAL_LABEL과
# 동일(2차봉헌만 공백 없는 "2차봉헌" 표기, 실측: 헤더 텍스트 "2차봉헌 야훼이레 810 ...").
_YOUTH_HEADER_LABEL = {
    '입당': '입당', '봉헌': '봉헌', '성체': '성체', '2차봉헌': '2차봉헌', '파견': '파견',
}

_YOUTH_SOURCE_LINE = {
    '나주노': '나는 주님께 노래하리라 성가 {번호}',
    '야훼 이레': '야훼이레 성가 {번호}',
    '가톨릭성가': '가톨릭 성가 {번호}',
}


def resolve_youth_hymn_pptx(entry: dict, htype: str):
    """청년미사 성가 한 곡을 해석해 (Presentation|None, resolved_title)을 반환.

    - 출처='기타': (None, entry['제목']) — PPT 없음. 호출부가 콘텐츠 슬라이드 삭제/삽입을
      완전히 건너뛰어야 한다는 신호(요구사항 §5.3 — 가사 placeholder는 프로그램이 채우지 않음).
    - 출처∈{'나주노','야훼 이레'}: OneDrive 청년미사 성가 폴더에서 '{출처} 성가 {번호} *.pptx'
      패턴으로 기존 파일을 먼저 찾는다(요구사항 §5.2 "찾은 뒤에는 재생성하지 않음"). 없으면
      missa_youth_hymn_pdf.build_hymn_pptx()로 새로 만들어 저장해 다음 주부터 재사용한다.
      두 경로 모두 build_header_runs()를 다시 호출해 이번 주 용도(htype/제목)로 헤더를
      강제 갱신한다 — 저장 시점의 구분과 이번 주 구분이 다를 수 있다(§2.1: 캐시된 PPT를
      그대로 재사용해도 헤더는 항상 최신화되어야 함).
    - 출처='가톨릭성가': OneDrive '성가 {번호}' 정규식으로 탐색. 헤더는 건드리지 않는다 —
      호출부(replace_성가_youth)가 기존 _update_성가_header()로 처리한다(§2.1:
      build_header_runs()와 _update_성가_header()는 run 구조 전제가 달라 서로의 자리에
      쓰면 안 됨).

    세 출처 모두 2026-09-27부터 `missa_gui.find_youth_onedrive_hymn_file()`로 "필요한 곡
    하나만" 찾아 받는다 — 이전에는 `get_onedrive_hymn_folder()`/`get_onedrive_youth_hymn_
    root()`가 폴더 전체를 미러링해서(가톨릭성가 34곡 등) 그날 실제로 쓰는 5곡만 필요한데도
    전부 내려받았다. 성인미사는 이 변경과 무관하다(get_onedrive_hymn_folder()는 여전히
    `mass_type == 'adult'`로 좁혀진 missa_to_ppt.py 호출부에서만 쓰인다).
    """
    출처 = entry.get('출처')
    번호 = entry.get('번호')
    제목 = entry.get('제목')

    if 출처 == '기타':
        if not 제목:
            raise ValueError("'기타' 출처는 제목이 필수입니다.")
        return None, 제목

    if 출처 not in ('나주노', '야훼 이레', '가톨릭성가'):
        raise ValueError(f"알 수 없는 성가 출처: {출처!r}")
    if 번호 is None:
        raise ValueError(f"{출처} 출처는 번호가 필수입니다.")

    from missa_gui import (
        find_youth_onedrive_hymn_file, get_youth_ppt_mode, youth_local_subdir,
        _DEFAULT_ONEDRIVE_HYMN_PATH, _DEFAULT_ONEDRIVE_YOUTH_HYMN_PATH, _load_config,
    )

    if 출처 == '가톨릭성가':
        pattern = rf'성가 {번호}(?!\d)'
        found = find_youth_onedrive_hymn_file(
            remote_path_key='onedrive_hymn_path',
            default_remote_path=_DEFAULT_ONEDRIVE_HYMN_PATH, subfolder=None,
            pattern=pattern, cache_subdir='청년_가톨릭성가', local_kind='catholic_hymn',
        )
        if not found:
            raise FileNotFoundError(f'가톨릭성가 {번호} 성가 PPT를 OneDrive에서 찾을 수 없습니다.')
        # 가톨릭성가는 나주노/야훼이레(PDF 인덱스)와 달리 제목을 조회할 소스 자체가 없고,
        # UI(_ask_youth_hymn_popup, §A6)도 이 출처는 번호만 받아 제목 입력칸을 보여주지
        # 않는다 — `제목 or ''`만 쓰면 항상 빈 문자열이 되어 title 슬라이드에 제목이
        # 빠진다(2026-09-24 실사용자 보고, output/20260919 슬라이드 108 재현). 파일명이
        # 이미 '성가 {번호} {제목}.pptx' 관례를 따르므로(find_files()도 같은 정규식으로
        # 이 파일을 찾음) 매칭된 위치 뒤 나머지를 제목으로 재사용한다.
        match = re.search(pattern, found.name)
        filename_title = found.stem[match.end():].strip()
        resolved_title = 제목 or filename_title
        return Presentation(str(found)), resolved_title

    # 나주노 / 야훼 이레
    import missa_youth_hymn_pdf as yh

    resolved_title = 제목 or yh.find_song_title(출처, 번호)

    subfolder = yh.SOURCES[출처]['onedrive_subfolder']
    prefix = f'{출처} 성가 {번호} '
    existing = find_youth_onedrive_hymn_file(
        remote_path_key='onedrive_youth_hymn_path',
        default_remote_path=_DEFAULT_ONEDRIVE_YOUTH_HYMN_PATH, subfolder=subfolder,
        pattern='^' + re.escape(prefix), cache_subdir=f'청년미사_성가/{subfolder}',
        local_kind='youth_hymn',
    )

    if existing is not None:
        src_prs = Presentation(str(existing))
    else:
        src_prs = yh.build_hymn_pptx(htype, 출처, 번호, resolved_title)
        filename = f'{prefix}{resolved_title}.pptx'

        # 저장 위치도 find_youth_onedrive_hymn_file()과 같은 모드를 따른다(§3.3.1) — 로컬
        # 모드면 'PPT 문서' 아래 성가 폴더에 저장하고(OneDrive 앱이 동기화하므로 직접 업로드
        # 하지 않음), 폴백 모드면 로컬 캐시에 저장한 뒤 Graph로 업로드한다.
        local_mode = get_youth_ppt_mode() == 'local'
        if local_mode:
            dest_dir = youth_local_subdir('youth_hymn', subfolder)
        else:
            from missa_gui import _SCRIPT_DIR
            dest_dir = _SCRIPT_DIR / 'cache' / '청년미사_성가' / subfolder
        local_path = dest_dir / filename
        try:
            dest_dir.mkdir(parents=True, exist_ok=True)
            src_prs.save(str(local_path))
        except Exception as e:
            print(f'  [경고] {출처} {번호} 성가 PPT 캐시 저장 실패: {e}')
        else:
            if not local_mode:
                config = _load_config()
                remote_root = config.get('onedrive_youth_hymn_path', _DEFAULT_ONEDRIVE_YOUTH_HYMN_PATH)
                import missa_onedrive as od
                remote_path = f'{remote_root}/{subfolder}/{filename}'
                try:
                    od.upload_file(remote_path, local_path)
                    print(f'  [OneDrive] {출처} {번호} 성가 PPT를 OneDrive에 업로드했습니다.')
                except Exception as e:
                    print(f'  [경고] {출처} {번호} 성가 PPT를 OneDrive에 업로드하지 못했습니다'
                          f'({e}) — 이 PC에는 저장돼 계속 쓸 수 있지만, 다른 PC와는 공유되지'
                          f' 않습니다. OneDrive 로그인/네트워크 상태를 확인해 주세요.')

    for slide in src_prs.slides:
        yh.build_header_runs(slide, htype, 출처, 번호, resolved_title)

    return src_prs, resolved_title


def _youth_run_templates_by_size(shape):
    """title 도형 안의 기존 run들을 폰트 크기별로 분류해 재구성용 서식 템플릿으로 쓴다.

    실측(§1.2/§2.3): 라벨(5200)/제목(3500,FFC000)/출처문장(2000,FFC000) 세 크기가 곡마다
    일관되게 쓰이지만, 단락 개수·내부 줄바꿈(\\x0b/<a:br/>) 구조는 출처마다 다르다(3단락 vs
    \\x0b 내부 줄바꿈 2단락). 단락 경계가 아니라 run의 sz로 서식을 찾으면 이 구조 차이와
    무관하게 항상 올바른 템플릿을 얻는다."""
    by_size = {}
    for p in shape.text_frame._txBody.findall(qn('a:p')):
        for r in p.findall(qn('a:r')):
            rPr = r.find(qn('a:rPr'))
            sz = rPr.get('sz') if rPr is not None else None
            if sz and sz not in by_size:
                by_size[sz] = r
    return by_size


def _update_youth_title_slide(slide, htype: str, 출처, number, title: str):
    """title 슬라이드를 [라벨, 제목, 출처+번호] 또는(기타) [라벨, 제목] 단락으로 재구성.

    기존 단락 수·내부 줄바꿈 구조에 의존하지 않고 항상 결정론적으로 다시 만든다 —
    run 경계를 가정한 부분 치환은 CLAUDE.md '옛 run 길이로 통짜 재배치' 함정 위험이 크다.
    설계서 §2.3의 "paragraphs[0]을 서식 템플릿으로 재사용" 스케치보다 한 단계 더 정밀하게,
    run의 실제 폰트 크기별로 템플릿을 골라(_youth_run_templates_by_size) 라벨/제목/출처문장의
    서로 다른 크기·색을 그대로 보존한다(그렇지 않으면 제목·출처문장이 라벨 서식(5200pt,
    기본색)을 물려받아 시각적으로 달라진다 — 02_specialist_impl_notes.md 기록)."""
    shape = _find_youth_title_shape(slide, HYMN_LABEL_KW[htype])
    if shape is None:
        print(f'  [경고] 청년미사 {htype} title 슬라이드를 찾지 못함')
        return

    # lxml Element는 len(children)==0일 때 `or`/bool()이 False로 평가돼(FutureWarning 대상)
    # "템플릿 없음"과 "빈 run 요소"를 혼동할 수 있으므로 `is not None`으로 명시 판별한다.
    templates = _youth_run_templates_by_size(shape)
    label_tmpl = templates.get('5200')
    if label_tmpl is None:
        label_tmpl = next(iter(templates.values()), None)
    title_tmpl = templates.get('3500')
    if title_tmpl is None:
        title_tmpl = label_tmpl
    source_tmpl = templates.get('2000')
    if source_tmpl is None:
        # 출처문장(sz=2000) run이 이 슬라이드에 물리적으로 없는 경우(실측: 106번 '성체' 2번째
        # 곡 title — '기타' 전용으로 만들어져 번호/출처문장 자체가 없다). '기타'가 아닌 출처를
        # 이 슬롯에 선택하면(§9.1상 유효) 출처문장 줄이 새로 필요한데, 제목 템플릿(sz=3500)을
        # 그대로 쓰면 출처문장이 제목과 같은 큰 글자로 렌더링된다 — deepcopy 후 sz만 2000으로
        # 강제해 최소한 크기는 스펙과 맞춘다(리뷰 §권장 2 반영).
        source_tmpl = copy.deepcopy(title_tmpl) if title_tmpl is not None else None
        if source_tmpl is not None:
            rPr = source_tmpl.find(qn('a:rPr'))
            if rPr is not None:
                rPr.set('sz', '2000')

    txBody = shape.text_frame._txBody
    paras = txBody.findall(qn('a:p'))
    label_pPr = paras[0].find(qn('a:pPr')) if paras else None
    body_pPr = paras[1].find(qn('a:pPr')) if len(paras) > 1 else label_pPr

    lines = [(HYMN_LABEL_KW[htype], label_tmpl, label_pPr)]
    lines.append((title or '', title_tmpl, body_pPr))
    if 출처 != '기타':
        src_fmt = _YOUTH_SOURCE_LINE.get(출처, '{출처} 성가 {번호}'.replace('{출처}', str(출처)))
        lines.append((src_fmt.format(번호=number), source_tmpl, body_pPr))

    for p in paras:
        txBody.remove(p)

    A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    for text, run_tmpl, pPr_tmpl in lines:
        new_p = pptx_parse_xml(f'<a:p xmlns:a="{A_NS}"/>')
        if pPr_tmpl is not None:
            new_p.insert(0, copy.deepcopy(pPr_tmpl))
        if run_tmpl is not None:
            new_r = copy.deepcopy(run_tmpl)
        else:
            new_r = pptx_parse_xml(f'<a:r xmlns:a="{A_NS}"><a:t/></a:r>')
        t_el = new_r.find(qn('a:t'))
        if t_el is None:
            t_el = pptx_parse_xml(f'<a:t xmlns:a="{A_NS}"/>')
            new_r.append(t_el)
        t_el.text = text
        _para_append_run(new_p, new_r)
        txBody.append(new_p)


def _update_youth_content_header_기타(slide, htype: str, title: str):
    """'기타' 출처 콘텐츠 헤더를 번호 없는 '{타입} {제목}' 포맷으로 갱신.

    _update_성가_header()의 정규식은 라벨 뒤에 숫자가 오는 것을 전제해(예: '입당 447') 번호가
    없는 이 포맷과 맞지 않는다(§6.4 리더 결정: 가사 placeholder는 건드리지 않되 헤더는 매주
    갱신 대상). 가사 placeholder('Text Placeholder 2')는 이 함수가 절대 건드리지 않는다 —
    이름으로 명시적으로 걸러 헤더('Rectangle 11')만 대상으로 한다.

    라벨 자체도 첫 토큰으로 매칭해 갱신 대상에 포함한다(대상 htype 고유 라벨을 리터럴로
    매칭하지 않음) — C4 scratch 재사용(빈 콘텐츠를 성체2 슬롯에서 복제해 다른 htype에 붙여넣는
    패턴)으로 헤더에 원래 다른 htype의 라벨(예: '성체 ...')이 남아있는 채로 들어올 수 있어,
    "이미 자기 라벨로 시작한다"고 가정하면 매칭에 실패해 헤더가 갱신되지 않는다."""
    label = _YOUTH_HEADER_LABEL.get(htype, htype)
    pat = re.compile(r'^(\s*)(\S+)(\s+)(.*)$')
    for shape in slide.shapes:
        if not shape.has_text_frame or shape.name != 'Rectangle 11':
            continue
        for para in shape.text_frame.paragraphs:
            m = pat.match(para.text)
            if not m:
                continue
            label_old = m.group(1) + m.group(2)
            label_new = m.group(1) + label
            rest_old = m.group(3) + m.group(4)
            rest_new = m.group(3) + title
            if label_old != label_new or rest_old != rest_new:
                _update_prefix_in_runs(para, [(label_old, label_new), (rest_old, rest_new)])
            return
    print(f'  [경고] 청년미사 {htype} 콘텐츠 헤더(기타)를 찾지 못함')


def _youth_song_entries_have_기타(성가_선택: dict) -> bool:
    for htype in HYMN_TYPES:
        entries = 성가_선택.get(htype) or []
        if isinstance(entries, dict):
            entries = [entries]
        for entry in entries:
            if entry.get('출처') == '기타':
                return True
    return False


def _capture_youth_blank_content_scratch_slide(prs):
    """출처='기타'를 고른 htype의 콘텐츠를 성체2 슬롯과 동일한 '빈 가사 placeholder 1장'
    모양으로 맞추기 위한 원본을 확보한다(C4). 템플릿에 따라 2차봉헌처럼 기본 콘텐츠가 실제
    악보 그림(야훼이레 810)으로 채워진 htype이 있어, 콘텐츠를 건드리지 않던 예전 방식은 그
    그림이 그대로 남는 버그였다. 성체2 슬롯은 항상 빈 가사 placeholder 1장이므로 이를
    duplicate_slide()로 복제해 맨 끝에 scratch로 보관해 두고, 모든 htype의 '기타' 선택이
    이 scratch를 다시 복제해 쓴다.

    duplicate_slide()는 항상 prs.slides 맨 끝에 추가하고 그 인덱스(len(prs.slides)-1)를
    반환한다. 이후 insert_slide_copy()가 발생시키는 '끝에 추가 후 move'(add_slide로 임시로
    맨 끝에 놓였다가 목표 위치로 이동) 왕복에서도, scratch보다 앞쪽 위치로만 삽입되는 한
    scratch는 항상 최종 슬라이드로 남는다 — 그래서 고정 인덱스를 저장해두지 않고 호출부가
    매번 len(prs.slides)-1로 다시 계산한다(02_specialist_impl_notes.md 참고).

    반환값: scratch 슬라이드 인덱스, 확보 실패(성체2 슬롯이 템플릿에 없거나 이미 1장이
    아님) 시 None.
    """
    sec = find_sections(prs, mass_type='youth')
    songs = sec.get('성체_songs', [])
    if len(songs) < 2:
        return None
    template_slot = songs[1]
    if template_slot['content_end'] - template_slot['content_start'] != 1:
        return None
    duplicate_slide(prs, template_slot['content_start'])
    return len(prs.slides) - 1


def replace_성가_youth(prs, 성가_선택: dict):
    """청년미사 성가 5종(성체는 최대 2곡) 교체. htype마다, 그리고 곡마다 find_sections를
    재탐색해 이전 곡 처리로 인한 슬라이드 인덱스 이동을 항상 반영한다."""
    scratch_present = False
    if _youth_song_entries_have_기타(성가_선택):
        # 어떤 htype이라도 처리되기 전, 성체2 슬롯이 아직 손대지 않은 원본 상태일 때
        # 캡처해야 한다 — HYMN_TYPES 순서상 '성체'가 '2차봉헌'보다 먼저 처리되므로, 루프
        # 안에서 캡처하면 이미 교체된 성체 콘텐츠를 원본으로 오인할 수 있다.
        scratch_present = _capture_youth_blank_content_scratch_slide(prs) is not None

    try:
        for htype in HYMN_TYPES:
            entries = 성가_선택.get(htype) or []
            if isinstance(entries, dict):
                entries = [entries]
            if not entries:
                continue

            for slot_i, entry in enumerate(entries):
                sec = find_sections(prs, mass_type='youth')
                songs = sec.get(f'{htype}_songs', [])
                if slot_i >= len(songs):
                    print(f'  [{htype}] {slot_i + 1}번째 곡 슬롯이 템플릿에 없음, 건너뜀')
                    continue
                _replace_one_youth_song(prs, htype, songs[slot_i], entry, scratch_present)

            # 템플릿엔 있지만 이번 주 입력이 없는 남는 곡(성체 2번째 등)은 title+content를
            # 통째로 삭제한다(요구사항 §5.1: 2번째 곡 입력란을 비우면 1곡만 처리 — 기존
            # 성인미사 "성체 1곡"과 동일 동작). 구분 슬라이드(빈 슬라이드) 자체는 건드리지
            # 않는다 — 그 슬라이드는 이 htype 섹션 전체의 경계이지 이 곡만의 것이 아니다.
            sec = find_sections(prs, mass_type='youth')
            songs = sec.get(f'{htype}_songs', [])
            for extra in reversed(songs[len(entries):]):
                for i in range(extra['content_end'] - 1, extra['title_idx'] - 1, -1):
                    delete_slide(prs, i)
    finally:
        if scratch_present:
            delete_slide(prs, len(prs.slides) - 1)


def _replace_one_youth_song(prs, htype: str, slot: dict, entry: dict, scratch_present: bool = False):
    title_idx = slot['title_idx']
    cs, ce = slot['content_start'], slot['content_end']
    출처 = entry['출처']

    if 출처 == '기타':
        _update_youth_title_slide(prs.slides[title_idx], htype, 출처, None, entry['제목'])
        if scratch_present and cs < ce:
            # 템플릿 기본 콘텐츠(예: 2차봉헌의 실제 악보 그림)를 성체2와 동일한 빈 가사
            # placeholder 1장으로 교체한다 — 헤더 갱신만으로는 본문 그림/텍스트가 남는다(C4).
            for i in range(ce - 1, cs - 1, -1):
                delete_slide(prs, i)
            insert_slide_copy(prs, cs, len(prs.slides) - 1)
            _update_youth_content_header_기타(prs.slides[cs], htype, entry['제목'])
        elif cs < ce:
            # scratch 확보 실패(비정상 템플릿) 시에는 기존 동작으로 폴백 — 헤더만 갱신하고
            # 본문은 건드리지 않는다(그림이 남을 수 있으나, 삭제만 하고 채우지 못하는 것보다
            # 안전하다).
            _update_youth_content_header_기타(prs.slides[cs], htype, entry['제목'])
        return

    src_prs, resolved_title = resolve_youth_hymn_pptx(entry, htype)
    _update_youth_title_slide(prs.slides[title_idx], htype, 출처, entry.get('번호'), resolved_title)

    for i in range(ce - 1, cs - 1, -1):
        delete_slide(prs, i)
    n_src = len(src_prs.slides)
    for i in range(n_src):
        copy_slide_from_prs(prs, cs + i, src_prs, i)

    if 출처 == '가톨릭성가':
        for i in range(n_src):
            # 청년 템플릿 헤더 라벨(_YOUTH_HEADER_LABEL, 공백 없는 '2차봉헌' 등)은 성인미사
            # 관례(CANONICAL_LABEL, 공백 있는 '2차 봉헌')와 다르다 — label_override로 청년
            # 경로 전용 라벨을 강제해 같은 슬라이드 안에서 출처별 표기가 갈라지지 않게 한다
            # (02b_review_report.md 확정 결함 수정).
            _update_성가_header(
                prs.slides[cs + i], htype, entry['번호'], slide_no=i + 1, n_slides=n_src,
                label_override=_YOUTH_HEADER_LABEL.get(htype),
            )
    # 나주노/야훼 이레는 resolve_youth_hymn_pptx() 안에서 build_header_runs()가 이미 헤더를
    # 채웠으므로 여기서 다시 건드리지 않는다(§2.1 — 서로 다른 run 구조 전제를 섞지 않는다).




