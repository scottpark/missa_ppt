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

import json

import re

import sys

from pathlib import Path



from pptx import Presentation
from missa_ooxml_utils import HYMN_TYPES, _slide_text, delete_slide
from missa_gui import (
    OUTPUT_ROOT, _last_output_path, _last_date_str, _preloaded_inputs,
    get_onedrive_hymn_folder, is_sunday_mass, _ask_numbers_popup, _ask_date_popup,
    _ask_input_files_popup, _ask_combined_input_popup, _report_progress,
    _run_with_progress_window, _show_result_window,
)
from missa_reading_layout import (
    parse_into_verse_units, layout_units_on_slides, _verify_and_rebalance_pages,
    replace_reading_slides, _align_ending_slides_to_제2독서, _reposition_merged_ending_shapes,
)
from missa_sections import find_sections, validate_pptx_structure, validate, strip_ppt2007_incompatible
from missa_content_updaters import (
    update_title_slide, update_입당송, update_reading_title_slide, update_복음_title_slide,
    update_화답송, update_복음환호송, update_영성체송, replace_시작기도문, replace_미사후기도,
    replace_성가, insert_공지사항,
)


# ─────────────────────────────────────────────────────────────────────────────

# 상수

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

    parser.add_argument('--공지사항', dest='공지사항', type=str, default=None,

                        help='공지사항 PPT 경로 (2차봉헌 성가 뒤에 삽입)')

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



    # 공지사항 확장자 검증(요구사항 §1.2.1 CLI 분기): .pptx만 허용, 위반 시 오류 종료.
    공지사항_path = getattr(args, '공지사항', None)
    if 공지사항_path and Path(공지사항_path).suffix.lower() != '.pptx':
        print('오류: --공지사항 파일은 PowerPoint(.pptx)만 지정할 수 있습니다.',
              file=sys.stderr)
        sys.exit(1)

    return args.date, numbers, args.화답송_pptx, getattr(args, '미사후기도', None), 공지사항_path





def find_files(date_str: str, hymn_numbers: dict, is_sunday: bool = None) -> dict:

    folder = Path(OUTPUT_ROOT) / date_str

    if not folder.is_dir():

        raise FileNotFoundError(f"폴더 없음: {folder}")



    files = {'ref_pptx': None, '시작기도': None, '화답송_pptx': None, '화답송_img': None, '미사후기도': None, '공지사항': None, '성가': {}}



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

    # 공지사항 PPT (--공지사항 인자가 있으면 main에서 덮어씀)
    for f in folder.iterdir():
        if f.suffix.lower() == '.pptx' and '공지사항' in f.name and not f.name.startswith('~$'):
            files['공지사항'] = f

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

        date_str, hymn_numbers, 화답송_override, 미사후기도_override, 공지사항_override = parse_args()

        files = find_files(date_str, hymn_numbers, is_sunday_mass(date_str))

        if 화답송_override:

            apply_화답송_override(files, 화답송_override)

        if 미사후기도_override:

            files['미사후기도'] = Path(미사후기도_override)

        if 공지사항_override:

            files['공지사항'] = Path(공지사항_override)



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



    # 6.7. 공지사항 슬라이드 삽입 (2차봉헌 성가 뒤). 앞 단계들이 슬라이드를 추가·삭제해
    # 인덱스가 이동했으므로 여기서 find_sections를 반드시 재호출한다(stale 인덱스 방지).

    if files.get('공지사항'):

        _report_progress(93, '공지사항 삽입 중...')

        print('\n[6.7] 공지사항 삽입...')

        sec = find_sections(prs)

        insert_공지사항(prs, files['공지사항'], sec)



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

