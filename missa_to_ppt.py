#!/usr/bin/env python3

"""

missa_to_ppt.py v3



매일미사 JSON 데이터를 기반으로 참조 PPT를 수정하여 새 PPT를 생성합니다.



사용법:

  python missa_to_ppt.py 20260628 [--입당 55 --봉헌 216 --성체 163 --2차봉헌 205 --파견 19]



성가번호 생략 시 tkinter 팝업으로 입력.

"""



import sys

if hasattr(sys.stdout, 'reconfigure'):

    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

    sys.stderr.reconfigure(encoding='utf-8', errors='replace')



import argparse

import atexit

import copy

import io

import json

import os

import re

import subprocess

import sys

import tempfile

import uuid

from pathlib import Path

from xml.sax.saxutils import escape as _xml_escape



from pptx import Presentation

from pptx.enum.shapes import MSO_SHAPE_TYPE

from pptx.oxml.ns import qn

from pptx.util import Emu, Pt
from missa_ooxml_utils import (
    _slide_text, all_slide_texts, delete_slide, move_slide, _blank_layout,
    _copy_spTree, _copy_image_rels, _update_rId_in_spTree, _effective_bg,
    _copy_bg_image_rels, duplicate_slide, insert_slide_copy, copy_slide_from_prs,
    _com_probe_path, _build_com_probe_pptx, find_slide_with_text, find_shape_exact_text,
    _set_slide_bg_black, _find_content_shape, _has_ending_text, _clear_text_frame,
    _para_append_run, _replace_para_text_clone, _set_화답송_content_text,
    _set_single_para_text, _update_book_name_after_br,
)
from missa_gui import (
    OUTPUT_ROOT, _progress_callback, _first_dialog_pos, _last_output_path,
    _last_date_str, _preloaded_inputs, _set_window_icon, _apply_theme, _center_window,
    _load_config, _save_config, _ask_onedrive_path_popup, get_onedrive_hymn_folder,
    _com_verification_enabled, is_sunday_mass, _ask_numbers_popup, _ask_date_popup,
    _ask_input_files_popup, _ask_combined_input_popup, _report_progress,
    _run_with_progress_window, _show_result_window,
)
from missa_reading_layout import (
    LINES_PER_SLIDE, ORANGE, _ends_sentence, parse_into_verse_units, _visual_lines,
    _wrap_line_count, _page_visual_lines, layout_units_on_slides, _verify_and_rebalance_pages,
    _set_reading_text, _count_slide_lines, _rendered_wrap_count, _get_slide_render_params,
    _count_slide_lines_rendered, _count_slide_lines_verified, _split_para_at_lines,
    _restore_para_from_backup, _split_and_adjust_via_com, _rebalance_reading_slides_post_write,
    replace_reading_slides, _align_ending_slides_to_제2독서, _reposition_merged_ending_shapes,
)
from missa_sections import (
    find_content_range, find_복음_content_range, find_sections, _is_hymn_divider,
    validate_pptx_structure, _orange_verse_numbers_in_range, _missing_orange_verse_numbers,
    validate, _strip_slide_xml, strip_ppt2007_incompatible,
)


# ─────────────────────────────────────────────────────────────────────────────

# 상수

# ─────────────────────────────────────────────────────────────────────────────



HYMN_TYPES = ['입당', '봉헌', '성체', '2차봉헌', '파견']

# ─────────────────────────────────────────────────────────────────────────────

# CLI / tkinter 팝업

# ─────────────────────────────────────────────────────────────────────────────



_TEST_HYMN_DEFAULTS = {'입당': 55, '봉헌': 216, '성체': 163, '2차봉헌': 205, '파견': 19}





def _infer_hymn_numbers(numbers: dict) -> dict:

    """누락된 번호를 테스트 기본값으로 채운다."""

    result = dict(numbers)

    for t, v in _TEST_HYMN_DEFAULTS.items():

        if result.get(t) is None:

            result[t] = v

    return result





def parse_args():

    parser = argparse.ArgumentParser(description='매일미사 PPT 생성')

    parser.add_argument('date', help='날짜 YYYYMMDD')

    parser.add_argument('--입당', type=int, default=None)

    parser.add_argument('--봉헌', type=int, default=None)

    parser.add_argument('--성체', type=int, default=None)

    parser.add_argument('--2차봉헌', dest='차봉헌2', type=int, default=None)

    parser.add_argument('--파견', type=int, default=None)

    parser.add_argument('--화답송', dest='화답송_pptx', type=str, default=None,

                        help='화답송 악보 PPT 경로 (기본: 날짜 폴더의 화답송 악보.pptx)')

    parser.add_argument('--미사후기도', dest='미사후기도', type=str, default=None,

                        help='미사 후 기도 PPT 경로 (평일미사 전용)')

    parser.add_argument('--test', action='store_true', help='테스트 모드: 누락 성가 번호를 테스트 기본값으로 채움(검증 생략)')

    args = parser.parse_args()



    numbers = {

        '입당': args.입당,

        '봉헌': args.봉헌,

        '성체': args.성체,

        '2차봉헌': args.차봉헌2,

        '파견': args.파견,

    }



    if args.test:

        # 테스트 모드: 누락 번호를 테스트 기본값으로 채움

        numbers = _infer_hymn_numbers(numbers)

    else:

        # 실사용: 미사 유형별 필수 성가 번호 검증 (누락 시 오류·중단)

        # 주일미사 = 5종 전부, 평일미사 = 2차봉헌 제외 4종 필수 (2차봉헌은 없는 게 정상)

        is_sunday = is_sunday_mass(args.date)

        required = HYMN_TYPES if is_sunday else [t for t in HYMN_TYPES if t != '2차봉헌']

        missing = [t for t in required if numbers.get(t) is None]

        if missing:

            mass_label = '주일미사' if is_sunday else '평일미사'

            print(

                f'오류: {mass_label}에 필요한 성가 번호가 누락되었습니다: {", ".join(missing)}',

                file=sys.stderr,

            )

            sys.exit(1)



    return args.date, numbers, args.화답송_pptx, getattr(args, '미사후기도', None)





def find_files(date_str: str, hymn_numbers: dict, is_sunday: bool = None) -> dict:

    folder = Path(OUTPUT_ROOT) / date_str

    if not folder.is_dir():

        raise FileNotFoundError(f"폴더 없음: {folder}")



    files = {'ref_pptx': None, '시작기도': None, '화답송_pptx': None, '화답송_img': None, '미사후기도': None, '성가': {}}



    # 화답송 악보 PPT

    for f in folder.iterdir():

        if f.suffix.lower() == '.pptx' and '화답송 악보' in f.name and not f.name.startswith('~$'):

            files['화답송_pptx'] = f



    # 화답송 악보 이미지: 수작업 PPT가 있으면 그것을 우선하고(기존 동작 완전 보존) 없을 때만

    # 탐색한다(설계서 §4.1). 촬영 앱이 붙인 타임스탬프 파일명("20260809_043842947.jpg")은

    # 파일명에 '화답송'이 없으므로, 이름 매칭이 안 되면 "폴더 내 유일 이미지"로 폴백한다.

    # 단, 날짜 폴더에 화답송과 무관한 이미지가 여러 개 섞여 있으면 어떤 것인지 모호하므로

    # 폴백하지 않는다 — 잘못된 이미지를 화답송으로 오인해 처리하는 것보다 미검출이 안전하다.

    if not files['화답송_pptx']:

        _HWADAPSONG_IMG_EXTS = {'.png', '.jpg', '.jpeg'}

        named_imgs = [
            f for f in folder.iterdir()
            if f.suffix.lower() in _HWADAPSONG_IMG_EXTS and '화답송' in f.name
            and not f.name.startswith('~$')
        ]

        if named_imgs:

            files['화답송_img'] = named_imgs[0]

        else:

            all_imgs = [
                f for f in folder.iterdir()
                if f.suffix.lower() in _HWADAPSONG_IMG_EXTS and not f.name.startswith('~$')
            ]

            if len(all_imgs) == 1:

                files['화답송_img'] = all_imgs[0]



    if is_sunday is None:

        is_sunday = is_sunday_mass(date_str)



    # 성가 PPT: OneDrive 폴더에서 우선 검색, 없으면 날짜 폴더 fallback (평일미사는 악보 자체가 불필요하므로 조회하지 않는다)

    if is_sunday:

        onedrive_folder = get_onedrive_hymn_folder()

        onedrive_pptxs = [

            f for f in onedrive_folder.rglob('*.pptx')

            if not f.name.startswith('~$')

        ]

        for htype, num in hymn_numbers.items():

            if num is None:

                continue

            found = None

            for f in onedrive_pptxs:

                if re.search(rf'성가 {num}(?!\d)', f.name):

                    found = f

                    break

            if not found:

                for f in folder.iterdir():

                    if f.suffix.lower() == '.pptx' and re.search(rf'성가 {num}(?!\d)', f.name):

                        found = f

                        break

            if found:

                files['성가'][htype] = found



    # 시작기도 PPT

    for f in folder.iterdir():

        if (f.suffix.lower() == '.pptx' and '시작기도' in f.name

                and not f.name.startswith('~$') and not f.name.startswith('_')):

            files['시작기도'] = f



    # 미사 후 기도 PPT
    for f in folder.iterdir():
        if f.suffix.lower() == '.pptx' and '미사후기도' in f.name and not f.name.startswith('~$'):
            files['미사후기도'] = f

    # 참조 PPT: YYYYMMDD_ 로 시작하는 파일 중 현재 날짜가 아닌 것을 우선 선택

    hymn_set = set(files['성가'].values())

    candidates = []

    for f in folder.iterdir():

        if (f.suffix.lower() == '.pptx'

                and f not in hymn_set

                and f != files['시작기도']

                and f != files['화답송_pptx']

                and date_str not in f.name

                and not f.name.startswith('~$')

                and not f.name.startswith('test_')):

            candidates.append(f)

    if candidates:

        # YYYYMMDD_ 패턴 파일 우선

        dated = [f for f in candidates if re.match(r'^\d{8}_', f.name)]

        files['ref_pptx'] = dated[0] if dated else candidates[0]



    return files


def apply_화답송_override(files: dict, override_path_str: str) -> None:
    """--화답송 CLI 오버라이드를 확장자에 따라 화답송_pptx/화답송_img 키로 분기해 반영한다.

    오버라이드는 find_files()가 자동 탐색해 이미 채워둔 반대쪽 키를 명시적으로 비워야 한다 —
    안 그러면 update_화답송()의 "PPT 우선" 규칙 때문에 자동 탐색된 파일이 사용자가 명시적으로
    지정한 오버라이드를 조용히 무시해 버린다(설계서 §4.2).
    """
    override_path = Path(override_path_str)
    if override_path.suffix.lower() in ('.png', '.jpg', '.jpeg'):
        files['화답송_img'] = override_path
        files['화답송_pptx'] = None
    else:
        files['화답송_pptx'] = override_path
        files['화답송_img'] = None




# ─────────────────────────────────────────────────────────────────────────────

# JSON 생성

# ─────────────────────────────────────────────────────────────────────────────



def get_json_data(date_str: str) -> dict:

    json_path = Path(OUTPUT_ROOT) / date_str / f'missa_{date_str}.json'



    if not json_path.exists():

        print(f'[{date_str}] 미사 데이터 가져오는 중...')

        import missa_to_json as _m2j

        _m2j.run(date_str, output_root=OUTPUT_ROOT)



    with open(json_path, 'r', encoding='utf-8') as f:

        return json.load(f)





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





def main():

    if len(sys.argv) == 1:

        if _preloaded_inputs[0] is not None:

            # EXE GUI 흐름: 팝업에서 미리 수집한 값 사용

            _inp = _preloaded_inputs[0]

            _preloaded_inputs[0] = None

            date_str    = _inp['date_str']

            files       = _inp['files']

            hymn_numbers = _inp['hymn_numbers']

            if is_sunday_mass(date_str):

                # 주일미사만 OneDrive 악보 폴더를 조회한다 (평일미사는 악보 자체가 불필요)

                _report_progress(3, 'OneDrive 성가 폴더 검색 중...')

                onedrive_folder = get_onedrive_hymn_folder()

                onedrive_pptxs = [f for f in onedrive_folder.rglob('*.pptx') if not f.name.startswith('~$')]

                for htype, num in hymn_numbers.items():

                    if num is None:

                        continue

                    for f in onedrive_pptxs:

                        if re.search(rf'성가 {num}(?!\d)', f.name):

                            files['성가'][htype] = f

                            break

            (Path(OUTPUT_ROOT) / date_str).mkdir(parents=True, exist_ok=True)

        else:

            # 인수 없이 실행: 대화형 모드 (팝업 순서: 날짜 → 파일 선택 → 성가번호)

            date_str = _ask_date_popup()

            files = _ask_input_files_popup()

            hymn_numbers = _ask_numbers_popup({}, is_sunday_mass(date_str))

            if is_sunday_mass(date_str):

                # 주일미사만 OneDrive 악보 폴더를 조회한다 (평일미사는 악보 자체가 불필요)

                onedrive_folder = get_onedrive_hymn_folder()

                onedrive_pptxs = [f for f in onedrive_folder.rglob('*.pptx') if not f.name.startswith('~$')]

                for htype, num in hymn_numbers.items():

                    if num is None:

                        continue

                    for f in onedrive_pptxs:

                        if re.search(rf'성가 {num}(?!\d)', f.name):

                            files['성가'][htype] = f

                            break

            (Path(OUTPUT_ROOT) / date_str).mkdir(parents=True, exist_ok=True)

    else:

        date_str, hymn_numbers, 화답송_override, 미사후기도_override = parse_args()

        files = find_files(date_str, hymn_numbers, is_sunday_mass(date_str))

        if 화답송_override:

            apply_화답송_override(files, 화답송_override)

        if 미사후기도_override:

            files['미사후기도'] = Path(미사후기도_override)



    is_sunday = is_sunday_mass(date_str)

    print(f'날짜: {date_str}')

    print(f'미사 유형: {"주일미사" if is_sunday else "평일미사"}')

    print(f'성가: {hymn_numbers}')



    _last_date_str[0] = date_str

    # 1. 파일 확인

    _report_progress(5, '파일 확인 중...')

    print('\n[1] 파일 확인...')

    if not files['ref_pptx']:

        print('오류: 참조 PPT 없음', file=sys.stderr)

        sys.exit(1)

    print(f'  참조 PPT: {files["ref_pptx"].name}')

    print(f'  시작기도: {files["시작기도"].name if files["시작기도"] else "없음"}')

    화답송_source = files.get('화답송_pptx') or files.get('화답송_img')

    print(f'  화답송 악보: {화답송_source.name if 화답송_source else "없음"}')

    if is_sunday and not 화답송_source:

        print('  [경고] 주일미사인데 화답송 악보 파일(PPT/이미지)이 없습니다.')

    if files.get('미사후기도'):

        print(f'  미사 후 기도: {files["미사후기도"].name}')

    for k, v in files['성가'].items():

        print(f'  성가({k}): {v.name}')

    # 주일미사 5종(입당/봉헌/성체/2차봉헌/파견) 악보 파일을 (OneDrive/날짜 폴더 어디에서도)

    # 찾지 못한 성가가 있으면 중단한다. 평일미사는 악보를 쓰지 않으므로 검사하지 않는다.

    if is_sunday:

        missing_scores = [
            f'{htype}({hymn_numbers.get(htype)})' for htype in HYMN_TYPES
            if not files['성가'].get(htype)
        ]

        if missing_scores:

            print(f'오류: 주일미사 성가 악보 파일을 찾을 수 없습니다: {", ".join(missing_scores)}', file=sys.stderr)

            sys.exit(1)



    # 2. JSON 생성

    _report_progress(10, 'JSON 로딩 중...')

    print('\n[2] JSON 로딩...')

    json_data = get_json_data(date_str)

    print(f'  전례: {json_data["liturgy"]}')

    print(f'  날짜: {json_data["date"]}')



    # 3. 참조 PPT 열기

    _report_progress(15, '참조 PPT 열기...')

    prs = Presentation(str(files['ref_pptx']))

    print(f'\n[3] 참조 PPT: {len(prs.slides)}개 슬라이드')



    # 4. 전례 텍스트 교체

    _report_progress(20, '전례 텍스트 교체 중...')

    print('\n[4] 전례 텍스트 교체...')



    print('  섹션 위치 파악...')

    sec = find_sections(prs)



    _report_progress(25, '제목 슬라이드...')

    print('  제목 슬라이드...')

    update_title_slide(prs, json_data)



    _report_progress(30, '입당송...')

    print('  입당송...')

    update_입당송(prs, json_data, sec)



    # 제1독서

    _report_progress(35, '제1독서...')

    if '제1독서_title' in sec and json_data.get('제1독서'):

        print('  제1독서...')

        update_reading_title_slide(prs, sec['제1독서_title'], '제1독서', json_data['제1독서']['title'])

        units = parse_into_verse_units(json_data['제1독서']['content'])

        pages = layout_units_on_slides(units)

        pages = _verify_and_rebalance_pages(pages, '제1독서')

        shift = replace_reading_slides(

            prs, sec['제1독서_start'], sec['제1독서_end'],

            pages, sec['제1독서_start'], line_spacing=1.1, label='제1독서'

        )

        print(f'    {len(pages)}개 슬라이드 (변화: {shift:+d})')


    # 섹션 재탐색

    sec = find_sections(prs)



    # 화답송

    _report_progress(45, '화답송...')

    if '화답송_start' in sec and json_data.get('화답송'):

        print('  화답송...')

        update_화답송(prs, json_data, sec, files.get('화답송_pptx'), files.get('화답송_img'), is_sunday=is_sunday)


    # 섹션 재탐색

    sec = find_sections(prs)



    # 제2독서

    _report_progress(55, '제2독서...')

    if '제2독서_title' in sec:

        if json_data.get('제2독서'):

            print('  제2독서...')

            update_reading_title_slide(prs, sec['제2독서_title'], '제2독서', json_data['제2독서']['title'])

            units = parse_into_verse_units(json_data['제2독서']['content'])

            pages = layout_units_on_slides(units)

            pages = _verify_and_rebalance_pages(pages, '제2독서')

            shift = replace_reading_slides(

                prs, sec['제2독서_start'], sec['제2독서_end'],

                pages, sec['제2독서_start'], label='제2독서'

            )

            print(f'    {len(pages)}개 슬라이드 (변화: {shift:+d})')

        else:

            print('  제2독서 없음 → 슬라이드 삭제')

            s, e = sec['제2독서_title'], sec['제2독서_end']

            for i in range(e - 1, s - 1, -1):

                delete_slide(prs, i)

            # 제2독서 앞의 blank divider(s-1)도 삭제 — 연속 blank 방지
            if s > 0 and not _slide_text(prs.slides[s - 1]).strip():

                delete_slide(prs, s - 1)


    # 섹션 재탐색

    sec = find_sections(prs)



    _report_progress(62, '복음환호송 및 복음...')

    print('  복음환호송...')

    update_복음환호송(prs, json_data, sec)




    # 복음

    sec = find_sections(prs)

    if '복음_title' in sec and json_data.get('복음'):

        print('  복음...')

        update_복음_title_slide(prs, sec['복음_title'], json_data['복음']['title'])

        units = parse_into_verse_units(json_data['복음']['content'])

        pages = layout_units_on_slides(units)

        pages = _verify_and_rebalance_pages(pages, '복음')

        shift = replace_reading_slides(

            prs, sec['복음_start'], sec['복음_end'],

            pages, sec['복음_start'], label='복음'

        )

        print(f'    {len(pages)}개 슬라이드 (변화: {shift:+d})')


    # 종료 슬라이드 텍스트박스 위치 정렬 (제1독서·복음 → 제2독서 기준)

    sec = find_sections(prs)

    print('  종료 슬라이드 위치 정렬...')

    _align_ending_slides_to_제2독서(prs, sec)


    # 본문+종료 통합 슬라이드의 ending shape 동적 위치 재조정 (겹침 방지)

    sec = find_sections(prs)

    _reposition_merged_ending_shapes(prs, sec)




    # 영성체송

    _report_progress(75, '영성체송...')

    sec = find_sections(prs)

    print('  영성체송...')

    update_영성체송(prs, json_data, sec)



    # 5. 시작기도문 교체

    _report_progress(80, '시작기도문 교체...')

    sec = find_sections(prs)

    if files.get('시작기도'):

        print('\n[5] 시작기도문 교체...')

        replace_시작기도문(prs, files['시작기도'], sec)



    # 6. 성가 교체

    _report_progress(88, '성가 교체...')

    print('\n[6] 성가 교체...')

    replace_성가(prs, files['성가'], hymn_numbers, copy_scores=is_sunday)



    # 6.5. 미사 후 기도 (평일미사 전용)

    if not is_sunday and files.get('미사후기도'):

        _report_progress(92, '미사 후 기도 교체 중...')

        print('\n[6.5] 미사 후 기도...')

        sec = find_sections(prs)

        replace_미사후기도(prs, files['미사후기도'], sec)



    # 7. 저장

    _report_progress(95, '파일 저장 중...')

    liturgy_safe = re.sub(r'[\\/:*?"<>|]', '_', json_data['liturgy'])

    output_name = f"{date_str}_{liturgy_safe}.pptx"

    output_path = Path(OUTPUT_ROOT) / date_str / output_name

    print(f'\n[7] 저장: {output_path}')


    prs.save(str(output_path))



    # 7.5. PPT 2007 비호환 요소 정리 (깨진 이미지 참조 제거 등)

    _report_progress(96, 'PPT 호환성 정리 중...')
    print('\n[7.5] PPT 호환성 정리...')
    strip_ppt2007_incompatible(str(output_path))



    # 8. 검증

    _report_progress(98, 'PPTX 구조 검증 중...')

    print('\n[8] 검증...')

    xml_issues = validate_pptx_structure(str(output_path))

    if xml_issues:

        print('  [경고] PPTX 구조 문제 발견:')

        for issue in xml_issues:

            print(f'    - {issue}')

    else:

        print('  PPTX 구조 검증 통과!')

    prs2 = Presentation(str(output_path))

    validate(prs2, json_data, is_sunday=is_sunday)



    _last_output_path[0] = output_path

    print(f'\n완료! → {output_path}')





if __name__ == '__main__':

    if len(sys.argv) > 1:

        # CLI 모드: 인수가 있으면 바로 실행 (Python 또는 EXE 모두)

        main()

    else:

        # GUI 모드: 통합 입력 팝업 → 성가번호 팝업 → 진행률 창 → 결과 창

        try:

            _date_str, _files = _ask_combined_input_popup()

        except RuntimeError:

            sys.exit(0)



        try:

            _hymn_numbers = _ask_numbers_popup({}, is_sunday_mass(_date_str))

        except RuntimeError:

            sys.exit(0)



        _preloaded_inputs[0] = {

            'date_str':     _date_str,

            'files':        _files,

            'hymn_numbers': _hymn_numbers,

        }



        _out_text, _err_text = _run_with_progress_window(main)



        if _err_text:

            # 팝업에는 구체적 에러 메시지(_err_text)만 보여준다. 전체 진행 로그(_out_text)는

            # log_text로 넘겨 로그 파일에만 남긴다.

            _show_result_window('오류 발생', _err_text, is_error=True, log_text=_out_text)

        else:

            _show_result_window('완료', _out_text or '')

