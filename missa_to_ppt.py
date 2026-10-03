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

from datetime import datetime, timedelta

from pathlib import Path



from pptx import Presentation
from missa_ooxml_utils import HYMN_TYPES, _slide_text, delete_slide
import missa_youth_gospel
from missa_gui import (
    OUTPUT_ROOT, _last_output_path, _last_date_str, _last_output_key, _last_output_mass, _preloaded_inputs,
    output_folder_key, get_onedrive_hymn_folder, is_sunday_mass, _ask_numbers_popup,
    _ask_date_popup, _ask_input_files_popup, _ask_combined_input_popup,
    _ask_youth_hymn_popup, _report_progress, _run_with_progress_window, _show_result_window,
    ensure_onedrive_login, get_youth_ppt_mode, get_adult_ppt_folder,
)
from missa_reading_layout import (
    parse_into_verse_units,
    replace_reading_slides, _align_ending_slides_to_제2독서, _reposition_merged_ending_shapes,
)
from missa_sections import find_sections, validate_pptx_structure, validate, strip_ppt2007_incompatible
from missa_content_updaters import (
    update_title_slide, update_입당송, update_reading_title_slide, update_복음_title_slide,
    update_화답송, update_복음환호송, update_영성체송, replace_시작기도문, replace_미사후기도,
    replace_성가, replace_성가_youth, insert_공지사항,
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





_YOUTH_HYMN_DISPLAY_TO_SOURCE = {'나주노': '나주노', '야훼이레': '야훼 이레', '가톨릭성가': '가톨릭성가', '기타': '기타'}
# _ask_youth_hymn_popup()의 DISPLAY_TO_SOURCE와 동일한 표기 — 두 곳 모두 작은 고정 매핑이라
# 이 프로젝트 관례상(missa_gui.py도 임포트 대신 인라인 정의) 공유 상수로 추출하지 않고
# 그대로 복제해 CLI 파싱에도 쓴다.


def _parse_youth_hymn_cli_entry(raw: str) -> dict:
    """CLI 문자열 '출처:값'을 성가_선택 entry dict({'출처','번호','제목'})로 변환한다.

    출처 표기는 GUI(_ask_youth_hymn_popup)와 동일하게 나주노/야훼이레/가톨릭성가/기타를
    쓴다(공백 없는 '야훼이레'). 나주노/야훼이레/가톨릭성가는 값이 성가 번호(숫자), 기타는
    값이 제목(텍스트)이다. 형식이 틀리면 ValueError — 호출부(parse_args)가 사용자에게 보일
    오류 메시지로 그대로 감싸 출력한다."""
    if ':' not in raw:
        raise ValueError(f"형식이 올바르지 않습니다(출처:값 형태로 입력): {raw!r}")
    display, value = raw.split(':', 1)
    display, value = display.strip(), value.strip()
    if display not in _YOUTH_HYMN_DISPLAY_TO_SOURCE:
        raise ValueError(
            f"알 수 없는 출처입니다: {display!r} "
            f"({'/'.join(_YOUTH_HYMN_DISPLAY_TO_SOURCE)} 중 선택)"
        )
    출처 = _YOUTH_HYMN_DISPLAY_TO_SOURCE[display]
    entry = {'출처': 출처, '번호': None, '제목': None}
    if 출처 == '기타':
        if not value:
            raise ValueError(f"기타 출처는 제목이 필요합니다: {raw!r}")
        entry['제목'] = value
    else:
        if not value.isdigit():
            raise ValueError(f"{display}는 성가 번호를 숫자로 입력해야 합니다: {raw!r}")
        entry['번호'] = int(value)
    return entry


def _parse_youth_hymn_cli_args(args) -> dict:
    """--y입당/--y봉헌/--y성체(최대 2회)/--y2차봉헌/--y파견 CLI 인자를 성가_선택
    dict(구현 계획 §9.3 형식)로 변환한다. 값이 없는 htype은 결과에서 생략한다 — 필수 여부
    검증은 main()의 기존 missing_htypes 체크(HYMN_TYPES 기준)가 담당하므로 여기서
    중복 검증하지 않는다(단일 검증 지점 유지)."""
    선택 = {}
    for htype, raw in (
        ('입당', args.y_입당), ('봉헌', args.y_봉헌),
        ('2차봉헌', args.y_2차봉헌), ('파견', args.y_파견),
    ):
        if raw:
            선택[htype] = [_parse_youth_hymn_cli_entry(raw)]
    if args.y_성체:
        선택['성체'] = [_parse_youth_hymn_cli_entry(raw) for raw in args.y_성체]
    return 선택


def _cli_replay_token(s: str) -> str:
    return f'"{s}"' if ' ' in s else s


def _build_cli_replay_command(date_str: str, mass_type: str, hymn_numbers: dict, 성가_선택: dict) -> str:
    """방금 이 실행에 쓰인 입력값(날짜·미사유형·성가 번호/선택)을 그대로 재현하는 CLI
    커맨드 한 줄을 만든다.

    GUI로 한 번 실행하면 참조 PPT·시작기도 PPT 등은 output/{날짜}/ 폴더에 남고, CLI의
    find_files()도 같은 폴더를 스캔하므로 그 파일들은 재지정할 필요가 없다 — 재입력이
    필요한 건 사실상 성가 번호/성가 선택뿐이다. 그래서 이 함수는 파일 경로는 다루지 않고
    딱 그 부분만 복사해서 바로 쓸 수 있는 한 줄로 만들어 로그에 남긴다(디버깅 반복 실행
    편의 기능). 공백이 포함된 토큰(기타 성가 제목 등)은 큰따옴표로 감싼다 — cmd.exe/
    PowerShell 기준이고, 제목 안에 큰따옴표 자체가 들어가는 극단적 경우는 다루지 않는다."""
    parts = ['python missa_to_ppt.py', date_str]

    if mass_type == 'youth':
        parts.append('--미사유형 청년')
        _raw_to_display = {v: k for k, v in _YOUTH_HYMN_DISPLAY_TO_SOURCE.items()}
        _flag = {'입당': '--y입당', '봉헌': '--y봉헌', '성체': '--y성체', '2차봉헌': '--y2차봉헌', '파견': '--y파견'}
        for htype in HYMN_TYPES:
            for entry in (성가_선택 or {}).get(htype, []):
                display = _raw_to_display.get(entry['출처'], entry['출처'])
                value = entry['번호'] if entry['번호'] is not None else (entry['제목'] or '')
                parts.append(f'{_flag[htype]} {_cli_replay_token(f"{display}:{value}")}')
    else:
        _flag = {'입당': '--입당', '봉헌': '--봉헌', '성체': '--성체', '2차봉헌': '--2차봉헌', '파견': '--파견'}
        for htype in HYMN_TYPES:
            num = (hymn_numbers or {}).get(htype)
            if num is not None:
                parts.append(f'{_flag[htype]} {num}')

    return ' '.join(parts)


def _build_arg_parser() -> argparse.ArgumentParser:
    """argparse 파서 구성만 담당한다(파싱 실행은 하지 않음) — `parse_args()`(전체 검증
    포함, CLI 실행 경로 전용)와 `__main__`/`main()`의 "date가 실제로 주어졌는가"만 가볍게
    엿보는 지점(§J3, 검증 없이 구조만 파싱) 양쪽이 이 함수를 공유해야, 두 곳의 argparse
    정의가 몰래 갈라지는 사고(예: choices 목록이 서로 다르게 진화)를 원천적으로 막는다."""
    parser = argparse.ArgumentParser(description='매일미사 PPT 생성')

    # §J3 — date를 생략하면(예: `--미사유형 청년`만 주고 날짜 없이 실행) GUI 입력 모드로
    # 진입한다. 기존에는 date가 필수라 생략 시 argparse가 즉시 사용법 오류로 종료했다.
    parser.add_argument('date', nargs='?', default=None,
                        help='날짜 YYYYMMDD (생략하면 GUI 입력 모드로 진입)')

    parser.add_argument('--미사유형', dest='mass_type_kr', choices=['성인', '청년', '어린이'], default='성인',
                        help='미사 유형 (기본: 성인, 어린이는 아직 미지원)')

    parser.add_argument('--입당', type=int, default=None)

    parser.add_argument('--봉헌', type=int, default=None)

    parser.add_argument('--성체', type=int, default=None)

    parser.add_argument('--2차봉헌', dest='차봉헌2', type=int, default=None)

    parser.add_argument('--파견', type=int, default=None)

    parser.add_argument('--y입당', dest='y_입당', type=str, default=None,
                        help='[청년] 입당 성가, 형식: 출처:번호 또는 출처:제목 (예: 나주노:447, 기타:주님의 기도)')

    parser.add_argument('--y봉헌', dest='y_봉헌', type=str, default=None,
                        help='[청년] 봉헌 성가 (--y입당과 동일 형식)')

    parser.add_argument('--y성체', dest='y_성체', type=str, action='append', default=None,
                        help='[청년] 성체 성가 (--y입당과 동일 형식, 최대 2회 지정 가능)')

    parser.add_argument('--y2차봉헌', dest='y_2차봉헌', type=str, default=None,
                        help='[청년] 2차봉헌 성가 (--y입당과 동일 형식)')

    parser.add_argument('--y파견', dest='y_파견', type=str, default=None,
                        help='[청년] 파견 성가 (--y입당과 동일 형식)')

    parser.add_argument('--화답송', dest='화답송_pptx', type=str, default=None,

                        help='화답송 악보 PPT 경로 (기본: 날짜 폴더의 화답송 악보.pptx)')

    parser.add_argument('--미사후기도', dest='미사후기도', type=str, default=None,

                        help='미사 후 기도 PPT 경로 (평일미사 전용)')

    parser.add_argument('--공지사항', dest='공지사항', type=str, default=None,

                        help='공지사항 PPT 경로 (2차봉헌 성가 뒤에 삽입)')

    parser.add_argument('--test', action='store_true', help='테스트 모드: 누락 성가 번호를 테스트 기본값으로 채움(검증 생략)')

    return parser


def parse_args():

    args = _build_arg_parser().parse_args()

    mass_type = _resolve_mass_type(args.mass_type_kr)

    if mass_type == 'youth':

        # 청년미사는 성인미사의 5종 번호 체계(--입당 447 등)를 쓰지 않는다 — 출처별
        # 성가 선택(--y입당 나주노:447 등)을 쓰고, 필수 여부 검증은 main()의 기존
        # missing_htypes 체크에 맡긴다(§9.3, 단일 검증 지점 유지).
        numbers = {}
        try:
            성가_선택 = _parse_youth_hymn_cli_args(args)
        except ValueError as e:
            print(f'오류: {e}', file=sys.stderr)
            sys.exit(1)

    else:

        성가_선택 = None

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

            include_2차봉헌 = resolve_mass_flags('adult', args.date)['include_2차봉헌']

            required = HYMN_TYPES if include_2차봉헌 else [t for t in HYMN_TYPES if t != '2차봉헌']

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

    return (args.date, numbers, args.화답송_pptx, getattr(args, '미사후기도', None), 공지사항_path,
            mass_type, 성가_선택)





def resolve_mass_flags(mass_type: str, date_str: str) -> dict:
    """`is_sunday` 하나가 실제로 겸하던 4개 축을 독립 값으로 분리.

    청년미사 2단계 설계서(§8, 원본 _workspace/청년미사_2단계/01_architect_design.md §4) 실측
    결과 — 옛 `is_sunday = is_sunday_mass(date_str)` 단일 변수는 ①달력상 일요일 여부,
    ②화답송/미사후기도 표시 형식(평일식 여부), ③2차봉헌 포함 여부, ④성가 실제 악보 복사
    여부까지 4개 축을 동시에 제어했다. 청년미사가 원하는 조합(①False/②True/③True/④True)은
    기존 주일(전부 True 계열)도 평일(①②만 False)도 아닌 제3의 조합이라 ②만 이름 붙은
    플래그로 분리해서는 표현할 수 없다 — ③·④는 mass_type에서 직접 도출해, 어느 것도 ①로부터
    암묵적으로 유도되지 않게 한다.

    adult 경로는 4개 값 전부 is_calendar_sunday(또는 그 부정)로만 정의되므로 기존 단일
    `is_sunday` 변수를 썼을 때와 100% 동일하게 동작한다(순수 변수 분리, 로직 변경 없음)."""
    is_calendar_sunday = is_sunday_mass(date_str)
    if mass_type == 'adult':
        use_weekday_display = not is_calendar_sunday
        include_2차봉헌 = is_calendar_sunday
        copy_hymn_scores = is_calendar_sunday
    else:  # 'youth'
        use_weekday_display = True
        include_2차봉헌 = True
        copy_hymn_scores = True
    return {
        'is_calendar_sunday': is_calendar_sunday,
        'use_weekday_display': use_weekday_display,
        'include_2차봉헌': include_2차봉헌,
        'copy_hymn_scores': copy_hymn_scores,
    }


def merge_youth_gospel_content(json_data: dict, en_mass: dict) -> dict:
    """청년미사 복음 JSON 병합: title은 한글 JSON 값 그대로 두고 content만 universalis
    영문 본문으로 교체한다(요구사항 §6.4 — 52번 한글 인트로 슬라이드는 무변경, 본문만 영문).
    원본 json_data를 변경하지 않고 얕은 복사본을 반환한다."""
    merged = dict(json_data)
    merged['복음'] = {**json_data.get('복음', {}), 'content': en_mass['Gospel']['content']}
    return merged


_MASS_TYPE_KR_TO_EN = {'성인': 'adult', '청년': 'youth'}


def _resolve_mass_type(mass_type_kr: str) -> str:
    """`--미사유형`(CLI) 또는 `_run_gui_mode()`가 받는 한글 값('성인'/'청년')을 하위
    파이프라인 전체(`resolve_mass_flags`/`find_sections`/`replace_성가_youth` 등)가
    기대하는 영문 키('adult'/'youth')로 변환한다(2026-09-26 §J3 — 원래는
    `_ask_mass_type_popup()`의 반환값이 유일한 입력원이었으나, 그 팝업 자체가 삭제되고
    CLI 인자가 직접 이 함수를 거친다).

    이 변환을 빠뜨리면 mass_type='청년'이 어떤 `if mass_type == 'youth':` 분기에도 걸리지
    않고 예외 없이 조용히 성인 경로로 빠진다 — Track A QA가 남긴 유일한 미완성 연결점이었다
    (구현 계획 §13-3, CLAUDE.md의 "조용한 오탐" 계열 함정과 동일). 알 수 없는 값은 'adult'로
    폴백한다('어린이'는 호출부가 이 함수에 닿기 전에 이미 걸러내므로 여기 들어오지 않지만,
    방어적으로 안전한 기본값으로 떨어지게 유지한다)."""
    return _MASS_TYPE_KR_TO_EN.get(mass_type_kr, 'adult')


def _youth_content_date_str(date_str: str) -> str:
    """청년미사 날짜 이중 처리(요구사항 §6.1). 사용자는 토요일 날짜를 입력하지만, 미사
    내용(1독서·화답송·영문 복음 등) 조회는 +1일(일요일)로 해야 한다.

    파일명·폴더명(§4)에는 사용자가 입력한 토요일 date_str을 그대로 쓴다 — 이 함수가
    돌려주는 값은 `get_json_data()`/`missa_youth_gospel.get_youth_mass()` 호출에만 쓰는
    "콘텐츠 조회 전용" 날짜다. `resolve_mass_flags('youth', date_str)`에는 이 값이 아니라
    원래 토요일 date_str을 그대로 넘겨야 한다 — 안 그러면 ①달력상 일요일 여부 판정
    (`is_calendar_sunday`)이 오염된다(youth 경로는 이 값을 실제로 쓰지 않아 지금은 무해하지만,
    향후 이 축을 참조하는 로직이 추가되면 조용히 틀린 값이 된다)."""
    return (datetime.strptime(date_str, '%Y%m%d') + timedelta(days=1)).strftime('%Y%m%d')


def find_files(date_str: str, hymn_numbers: dict, copy_hymn_scores: bool = None,
                mass_type: str = 'adult') -> dict:

    # mass_type 기본값 'adult'는 output_folder_key()가 접미사 ''를 반환하게 해, 이 인자를
    # 넘기지 않는 기존 호출부(회귀 테스트 포함) 전부가 이전과 동일한 output/{date_str}/를
    # 그대로 스캔한다(B그룹 — 폴더 명명 규칙 확장, 회귀 픽스처 경로 불변).
    folder = Path(OUTPUT_ROOT) / output_folder_key(date_str, mass_type)

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



    if copy_hymn_scores is None:

        copy_hymn_scores = is_sunday_mass(date_str)



    # 성가 PPT: OneDrive 폴더에서 우선 검색, 없으면 날짜 폴더 fallback (평일미사는 악보 자체가 불필요하므로 조회하지 않는다)

    if copy_hymn_scores:

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





def get_json_data(content_date_str: str, mass_type: str = 'adult', input_date_str: str = None) -> dict:

    """한글 미사 JSON을 로드한다(없으면 크롤링 후 캐시).

    `content_date_str`은 실제 미사 내용이 속한 날짜(청년미사는 +1일) — 크롤링 대상과
    파일명(`missa_{content_date_str}.json`)에 쓰인다(기존 관례 유지). `input_date_str`은
    사용자가 실제로 입력한 날짜 — 저장 **폴더**를 `output_folder_key(input_date_str,
    mass_type)`(청년이면 `{input_date_str}_youth`)로 정하는 데만 쓰인다. 생략하면(`None`)
    `content_date_str`로 폴백하는데, 성인 경로는 두 값이 항상 같으므로(`content_date_str ==
    date_str`) 기존 호출부는 `output/{date_str}/`로 100% 예전과 동일하게 동작한다(회귀
    픽스처 경로 불변 원칙).

    2026-09-26 실사용자 실행(20260926 청년미사) 재현 버그 수정 — 예전에는 폴더도 항상
    `content_date_str` 그대로였다(`Path(OUTPUT_ROOT) / date_str / ...`), 그 결과 청년미사
    JSON이 사용자가 보는 `output/{date_str}_youth/`가 아니라 콘텐츠 조회일(+1일, 일요일)
    날짜 폴더에 저장되어 이번 작업과 무관한 다른 산출물과 섞이고, 청년미사 출력 폴더에는
    아예 나타나지 않았다.

    `missa_to_json.run()`/`save_json()`은 항상 `output_root/date_str/...`로 폴더=파일명
    날짜를 강제 결합하므로(leaf 모듈, mass_type 개념을 모름), 폴더와 파일명 날짜를 분리해야
    하는 이 요구를 그대로 쓸 수 없다 — `fetch_html`/`parse_missa` 순수 함수만 재사용하고
    저장은 이 함수가 직접 한다."""

    folder = Path(OUTPUT_ROOT) / output_folder_key(input_date_str or content_date_str, mass_type)

    json_path = folder / f'missa_{content_date_str}.json'



    if not json_path.exists():

        print(f'[{content_date_str}] 미사 데이터 가져오는 중...')

        import missa_to_json as _m2j

        folder.mkdir(parents=True, exist_ok=True)

        html = _m2j.fetch_html(content_date_str)

        data = _m2j.parse_missa(html, content_date_str)

        with open(json_path, 'w', encoding='utf-8') as f:

            json.dump(data, f, ensure_ascii=False, indent=2)



    with open(json_path, 'r', encoding='utf-8') as f:

        return json.load(f)


def _get_youth_english_mass(content_date_str: str, output_key: str) -> dict:

    """영문 미사 JSON(청년미사 전용)을 로드한다(없으면 조회 후 캐시).

    `missa_youth_gospel.main()`의 CLI 저장 관례(`missa_en_{date_str}.json`)와 파일명을
    맞췄다. 캐시가 있으면 재조회하지 않는다 — `get_json_data()`의 기존 캐시 관례와 같은
    원칙이자, universalis.com이 과거 날짜를 오늘자로 조용히 리다이렉트하는 함정
    (`missa_youth_gospel.GospelFetchError`)을 캐시 히트 시 원천적으로 피하는 효과도 있다.

    2026-09-26 실사용자 실행 재현 버그 수정 — 이전에는 `missa_youth_gospel.get_youth_mass()`
    조회 결과를 메모리에서만 쓰고 버렸다(`missa_youth_gospel.main()`의 CLI 저장 코드는
    파이프라인 호출부가 타지 않는 별도 경로라 실제로 호출된 적이 없다)."""

    json_path = Path(OUTPUT_ROOT) / output_key / f'missa_en_{content_date_str}.json'

    if json_path.exists():

        with open(json_path, 'r', encoding='utf-8') as f:

            return json.load(f)

    data = missa_youth_gospel.get_youth_mass(content_date_str)

    json_path.parent.mkdir(parents=True, exist_ok=True)

    with open(json_path, 'w', encoding='utf-8') as f:

        json.dump(data, f, ensure_ascii=False, indent=2)

    return data





def _main_should_use_preloaded_or_fallback() -> bool:
    """`main()`의 최상위 분기 판단(§J3). `_preloaded_inputs[0]`가 이미 채워져 있으면
    (EXE/GUI 흐름) 무조건 참이고, 아니면 실제 CLI에 date가 주어졌는지를 가볍게 엿봐서
    (`_build_arg_parser()`, 검증 없는 구조 파싱만) 판단한다.

    이 함수를 따로 뽑아 둔 이유: mass_type 선택 팝업이 삭제되면서 `__main__`이 date 없이
    GUI 모드로 들어갈 때도 `--미사유형` 등 CLI 플래그가 붙어 있을 수 있게 됐다(예:
    `python missa_to_ppt.py --미사유형 청년`). 이 경우 `_run_gui_mode()`가
    `_preloaded_inputs[0]`를 채운 뒤 `main()`을 실제 프로세스 sys.argv(플래그가 남아있는
    그 argv) 그대로 호출하므로, 예전처럼 `len(sys.argv) == 1`만으로 "GUI 흐름인가"를
    판단하면 플래그가 있다는 이유만으로 잘못 CLI 분기(`parse_args()`, date=None이라 이후
    로직에서 깨짐)로 빠진다. 참 신호는 오직 `_preloaded_inputs[0]`의 유무이고, 그것도
    아니면 실제로 date가 주어졌는지로만 판단해야 한다(2026-09-26 §J3 실측 버그 수정)."""
    return _preloaded_inputs[0] is not None or _build_arg_parser().parse_args().date is None


def main():

    if _main_should_use_preloaded_or_fallback():

        if _preloaded_inputs[0] is not None:

            # EXE GUI 흐름: __main__ 하단이 미리 수집한 값을 사용한다. mass_type 팝업도
            # 이미 거기서 물었다 — 파일/성가 입력 화면 자체를 mass_type에 따라 분기해야 하므로
            # (성인=_ask_numbers_popup, 청년=_ask_youth_hymn_popup), main()에 들어오기 전에
            # 이미 결정돼 있어야 한다.

            _inp = _preloaded_inputs[0]

            _preloaded_inputs[0] = None

            mass_type    = _inp.get('mass_type', 'adult')

            date_str     = _inp['date_str']

            files        = _inp['files']

            hymn_numbers = _inp.get('hymn_numbers') or {}

            성가_선택     = _inp.get('성가_선택')

            if mass_type == 'adult' and is_sunday_mass(date_str):

                # 주일미사만 OneDrive 악보 폴더를 조회한다 (평일미사는 악보 자체가 불필요).
                # 청년미사는 replace_성가_youth()/resolve_youth_hymn_pptx()가 출처별 탐색·
                # PDF 자동생성까지 전담하므로 files['성가']를 아예 쓰지 않는다 — 여기서
                # 채워봐야 아무도 읽지 않는다.

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

            (Path(OUTPUT_ROOT) / output_folder_key(date_str, mass_type)).mkdir(parents=True, exist_ok=True)

        else:

            # main()을 직접(프로그래밍적으로) 호출하는 경로 — 정상적인 더블클릭 실행은 항상
            # __main__ 하단이 _preloaded_inputs를 채운 뒤 이 함수를 호출하므로, 실제 배포
            # 환경에서는 이 분기를 타지 않는다(그래도 프로그래밍적 호출/테스트를 위해 유지).
            # §J3 — mass_type 선택 팝업 자체가 삭제됐다. 이 분기는 date 없이 main()을 직접
            # 호출한 경우(테스트 등)만 타므로, CLI `--미사유형` 값을 알 방법이 없다 —
            # 팝업 없이 기본값('성인')으로 진행한다.

            mass_type_kr = '성인'

            mass_type = _resolve_mass_type(mass_type_kr)

            date_str = _ask_date_popup()

            files = _ask_input_files_popup()

            if mass_type == 'youth':

                성가_선택 = _ask_youth_hymn_popup()

                hymn_numbers = {}

            else:

                성가_선택 = None

                hymn_numbers = _ask_numbers_popup({}, is_sunday_mass(date_str))

                if is_sunday_mass(date_str):

                    onedrive_folder = get_onedrive_hymn_folder()

                    onedrive_pptxs = [f for f in onedrive_folder.rglob('*.pptx') if not f.name.startswith('~$')]

                    for htype, num in hymn_numbers.items():

                        if num is None:

                            continue

                        for f in onedrive_pptxs:

                            if re.search(rf'성가 {num}(?!\d)', f.name):

                                files['성가'][htype] = f

                                break

            (Path(OUTPUT_ROOT) / output_folder_key(date_str, mass_type)).mkdir(parents=True, exist_ok=True)

    else:

        (date_str, hymn_numbers, 화답송_override, 미사후기도_override, 공지사항_override,
         mass_type, 성가_선택) = parse_args()

        # 청년미사는 files['성가']를 아예 쓰지 않는다(replace_성가_youth가 출처별 탐색·PDF
        # 자동생성을 전담, EXE GUI 흐름과 동일 — main() 앞부분 주석 참고) — copy_hymn_scores를
        # False로 고정해 find_files()가 불필요하게 OneDrive 성가 폴더를 조회하지 않게 한다.
        copy_hymn_scores = (
            resolve_mass_flags('adult', date_str)['copy_hymn_scores'] if mass_type == 'adult' else False
        )
        files = find_files(date_str, hymn_numbers, copy_hymn_scores, mass_type)

        if 화답송_override:

            apply_화답송_override(files, 화답송_override)

        if 미사후기도_override:

            files['미사후기도'] = Path(미사후기도_override)

        if 공지사항_override:

            files['공지사항'] = Path(공지사항_override)



    # is_sunday가 겸하던 4개 축을 독립 값으로 분리(§resolve_mass_flags). adult 경로는 4개 값
    # 전부 is_calendar_sunday(=is_sunday)에서 파생되므로 기존 동작과 100% 동일하다. youth
    # 경로는 date_str(토요일)을 그대로 넘긴다 — resolve_mass_flags 내부의 is_calendar_sunday
    # 판정이 오염되지 않게 하기 위함(_youth_content_date_str 사용은 JSON/영문 복음 조회
    # 전용으로 분리, §6.1).
    mass_flags = resolve_mass_flags(mass_type, date_str)

    is_sunday = mass_flags['is_calendar_sunday']

    content_date_str = _youth_content_date_str(date_str) if mass_type == 'youth' else date_str

    print(f'날짜: {date_str}')

    if mass_type == 'youth':

        print(f'미사 유형: 청년미사 (내용 조회일: {content_date_str})')

        print(f'성가 선택: {성가_선택}')

    else:

        print(f'미사 유형: {"주일미사" if is_sunday else "평일미사"}')

        print(f'성가: {hymn_numbers}')

    print(f'[재실행용 CLI 커맨드] {_build_cli_replay_command(date_str, mass_type, hymn_numbers, 성가_선택)}')

    _last_date_str[0] = date_str
    _last_output_key[0] = output_folder_key(date_str, mass_type)
    _last_output_mass[0] = mass_type

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

    if mass_type == 'adult' and is_sunday and not 화답송_source:

        print('  [경고] 주일미사인데 화답송 악보 파일(PPT/이미지)이 없습니다.')

    if files.get('미사후기도'):

        print(f'  미사 후 기도: {files["미사후기도"].name}')

    print(f'  공지사항: {files["공지사항"].name if files.get("공지사항") else "없음"}')

    for k, v in files['성가'].items():

        print(f'  성가({k}): {v.name}')

    # 주일미사 5종(입당/봉헌/성체/2차봉헌/파견) 악보 파일을 (OneDrive/날짜 폴더 어디에서도)

    # 찾지 못한 성가가 있으면 중단한다. 평일미사는 악보를 쓰지 않으므로 검사하지 않는다.

    # 청년미사는 files['성가']를 애초에 쓰지 않는다(replace_성가_youth가 출처별 탐색·PDF
    # 자동생성을 전담) — mass_type == 'adult'로 명시적으로 좁혀, "토요일=평일"이라는 우연한
    # 달력 일치에 기대지 않는다(CLAUDE.md 암묵적 결합 금지 원칙과 동일선상).

    if mass_type == 'adult' and is_sunday:

        missing_scores = [
            f'{htype}({hymn_numbers.get(htype)})' for htype in HYMN_TYPES
            if not files['성가'].get(htype)
        ]

        if missing_scores:

            print(f'오류: 주일미사 성가 악보 파일을 찾을 수 없습니다: {", ".join(missing_scores)}', file=sys.stderr)

            sys.exit(1)

    elif mass_type == 'youth':

        # 방어적 이중 체크(독립 리뷰 확정 버그 대응, 05_review_report.md). 정상 경로라면
        # `_ask_youth_hymn_popup()`이 취소 시 RuntimeError를 던지므로 성가_선택이 비어 있는
        # 채로 여기까지 오지 않아야 하지만, main()을 프로그래밍적으로 직접 호출하는 경로
        # (테스트 등)에서 빈/부분 성가_선택이 들어오면 replace_성가_youth()가 예외 없이
        # 해당 htype을 조용히 스킵한다(missa_content_updaters.py의 `if not entries: continue`)
        # — 이번 주 성가가 갱신되지 않은 PPT가 경고 없이 "완료"로 저장되는 것을 막는다.

        missing_htypes = [htype for htype in HYMN_TYPES if not (성가_선택 or {}).get(htype)]

        if missing_htypes:

            print(f'오류: 청년미사에 필요한 성가 선택이 누락되었습니다: {", ".join(missing_htypes)}', file=sys.stderr)

            sys.exit(1)



    # 2. JSON 생성

    _report_progress(10, 'JSON 로딩 중...')

    print('\n[2] JSON 로딩...')

    json_data = get_json_data(content_date_str, mass_type=mass_type, input_date_str=date_str)

    if mass_type == 'youth':

        try:

            english_mass = _get_youth_english_mass(
                content_date_str, output_folder_key(date_str, mass_type),
            )

        except Exception as e:

            print(f'오류: 청년미사 영문 복음 조회 실패 ({e})', file=sys.stderr)

            sys.exit(1)

        json_data = merge_youth_gospel_content(json_data, english_mass)

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

    sec = find_sections(prs, mass_type=mass_type)



    _report_progress(25, '제목 슬라이드...')

    print('  제목 슬라이드...')

    # 청년미사는 콘텐츠(독서·복음 등) 조회에 +1일(content_date_str, 일요일)을 쓰므로
    # json_data['date']도 일요일이다 — 화면 표시는 사용자가 실제로 입력한 토요일 date_str
    # 그대로여야 한다(2026-09-26 실사용자 산출물 재현, 콘텐츠 조회 자체는 그대로 유지).
    update_title_slide(
        prs, json_data,
        date_str=f'{date_str[:4]}-{date_str[4:6]}-{date_str[6:]}' if mass_type == 'youth' else None,
    )



    _report_progress(30, '입당송...')

    print('  입당송...')

    update_입당송(prs, json_data, sec)



    # 제1독서

    _report_progress(35, '제1독서...')

    if '제1독서_title' in sec and json_data.get('제1독서'):

        print('  제1독서...')

        update_reading_title_slide(prs, sec['제1독서_title'], '제1독서', json_data['제1독서']['title'])

        units = parse_into_verse_units(json_data['제1독서']['content'])

        shift = replace_reading_slides(

            prs, sec['제1독서_start'], sec['제1독서_end'],

            units, sec['제1독서_start'], line_spacing=1.1, label='제1독서'

        )

        print(f'    슬라이드 수 변화: {shift:+d}')


    # 섹션 재탐색

    sec = find_sections(prs, mass_type=mass_type)



    # 화답송

    _report_progress(45, '화답송...')

    if '화답송_start' in sec and json_data.get('화답송'):

        print('  화답송...')

        # update_화답송()의 is_sunday 파라미터는 "주일 포맷(악보 포함)인가"를 뜻하므로
        # (함수 내부에서 `if not is_sunday: 평일 포맷`), "평일식 표시인가"를 뜻하는
        # mass_flags['use_weekday_display']를 그대로 넘기면 의미가 정반대로 뒤집힌다.
        # 부정(not)해서 넘긴다 — adult에서는 not use_weekday_display == is_calendar_sunday라
        # 기존 동작과 동일하다(구현 노트 참고, 설계서 §8.3 표의 직접 대입은 이 반전을
        # 놓치고 있었다). youth는 use_weekday_display가 항상 True라 is_sunday=False로 호출돼
        # 요구사항 §6.2가 명시한 "악보 없는 텍스트만" 포맷이 그대로 나온다.
        update_화답송(
            prs, json_data, sec, files.get('화답송_pptx'), files.get('화답송_img'),
            is_sunday=not mass_flags['use_weekday_display'],
        )


    # 섹션 재탐색

    sec = find_sections(prs, mass_type=mass_type)



    # 제2독서

    _report_progress(55, '제2독서...')

    if '제2독서_title' in sec:

        if json_data.get('제2독서'):

            print('  제2독서...')

            update_reading_title_slide(prs, sec['제2독서_title'], '제2독서', json_data['제2독서']['title'])

            units = parse_into_verse_units(json_data['제2독서']['content'])

            shift = replace_reading_slides(

                prs, sec['제2독서_start'], sec['제2독서_end'],

                units, sec['제2독서_start'], label='제2독서'

            )

            print(f'    슬라이드 수 변화: {shift:+d}')

        else:

            print('  제2독서 없음 → 슬라이드 삭제')

            s, e = sec['제2독서_title'], sec['제2독서_end']

            for i in range(e - 1, s - 1, -1):

                delete_slide(prs, i)

            # 제2독서 앞의 blank divider(s-1)도 삭제 — 연속 blank 방지
            if s > 0 and not _slide_text(prs.slides[s - 1]).strip():

                delete_slide(prs, s - 1)


    # 섹션 재탐색

    sec = find_sections(prs, mass_type=mass_type)



    _report_progress(62, '복음환호송 및 복음...')

    print('  복음환호송...')

    update_복음환호송(prs, json_data, sec, mass_type=mass_type)




    # 복음

    sec = find_sections(prs, mass_type=mass_type)

    if '복음_title' in sec and json_data.get('복음'):

        print('  복음...')

        update_복음_title_slide(prs, sec['복음_title'], json_data['복음']['title'])

        units = parse_into_verse_units(json_data['복음']['content'])

        # normalize_page_size는 청년 경로에서만 켠다 — 청년 템플릿의 영문 복음 콘텐츠
        # 페이지는 제작자가 샘플 분량에 맞춰 박스 크기를 손으로 다르게 잡아놔 재사용 시
        # 마지막 줄이 잘리는 문제가 있었지만(D2), 성인 한글 독서 경로에는 이 문제가
        # 없다고 확인되지 않았고(오히려 리뷰에서 성인 픽스처의 제2독서 콘텐츠 슬라이드도
        # 서로 다른 박스 크기를 쓰고 있음이 실측됨), 그 경로까지 건드릴 근거가 없다.
        #
        # align은 넘기지 않는다(항상 None) — D2가 처음엔 "텍스트박스를 슬라이드 중앙에
        # 오게 하라"는 요구를 "문단을 가운데 정렬(algn='ctr')하라"로 잘못 구현했다.
        # 박스 위치·크기는 위 normalize_page_size(템플릿 shape 좌표 복사)가 이미 담당하고,
        # 문단 자체는 원래대로 왼쪽 정렬이어야 한다(2026-09-26 실사용자 산출물 슬라이드
        # 54/55 재현 — 영문 복음 텍스트가 줄마다 가운데로 쏠려 보임).
        shift = replace_reading_slides(

            prs, sec['복음_start'], sec['복음_end'],

            units, sec['복음_start'], label='복음',

            align=None,

            normalize_page_size=(mass_type == 'youth'),

        )

        print(f'    슬라이드 수 변화: {shift:+d}')


    # 종료 슬라이드 텍스트박스 위치 정렬 (제1독서·복음 → 제2독서 기준)

    sec = find_sections(prs, mass_type=mass_type)

    print('  종료 슬라이드 위치 정렬...')

    _align_ending_slides_to_제2독서(prs, sec)


    # 본문+종료 통합 슬라이드의 ending shape 동적 위치 재조정 (겹침 방지)

    sec = find_sections(prs, mass_type=mass_type)

    _reposition_merged_ending_shapes(prs, sec)




    # 영성체송

    _report_progress(75, '영성체송...')

    sec = find_sections(prs, mass_type=mass_type)

    print('  영성체송...')

    update_영성체송(prs, json_data, sec)



    # 5. 시작기도문 교체

    _report_progress(80, '시작기도문 교체...')

    sec = find_sections(prs, mass_type=mass_type)

    if files.get('시작기도'):

        print('\n[5] 시작기도문 교체...')

        replace_시작기도문(prs, files['시작기도'], sec)



    # 6. 성가 교체

    _report_progress(88, '성가 교체...')

    if mass_type == 'youth':

        print('\n[6] 성가 교체(청년)...')

        replace_성가_youth(prs, 성가_선택 or {})

    else:

        print('\n[6] 성가 교체...')

        replace_성가(prs, files['성가'], hymn_numbers, copy_scores=mass_flags['copy_hymn_scores'])



    # 6.5. 미사 후 기도 (평일미사/청년미사 전용). §6.5 요구사항 원문 주의사항대로,
    # is_sunday_mass(date_str)의 "우연한" False가 아니라 mass_flags['use_weekday_display']
    # (표시형식 축)로 게이팅해 어느 축에서 파생됐는지 명시적으로 남긴다 — 청년미사는 이 값이
    # 항상 True라 별도 분기 없이 기존 조건 그대로 미사후기도가 삽입된다.
    if mass_flags['use_weekday_display'] and files.get('미사후기도'):

        _report_progress(92, '미사 후 기도 교체 중...')

        print('\n[6.5] 미사 후 기도...')

        sec = find_sections(prs, mass_type=mass_type)

        replace_미사후기도(prs, files['미사후기도'], sec)



    # 6.7. 공지사항 슬라이드 삽입 (2차봉헌 성가 뒤). 앞 단계들이 슬라이드를 추가·삭제해
    # 인덱스가 이동했으므로 여기서 find_sections를 반드시 재호출한다(stale 인덱스 방지).

    if files.get('공지사항'):

        _report_progress(93, '공지사항 삽입 중...')

        print('\n[6.7] 공지사항 삽입...')

        sec = find_sections(prs, mass_type=mass_type)

        insert_공지사항(prs, files['공지사항'], sec)



    # 7. 저장

    _report_progress(95, '파일 저장 중...')

    liturgy_safe = re.sub(r'[\\/:*?"<>|]', '_', json_data['liturgy'])

    if mass_type == 'youth':

        # §4 요구사항: 파일명에는 사용자가 입력한 토요일 date_str을 그대로 쓴다(내용 조회용
        # +1일 변환과는 별개).
        output_name = f"토요일 저녁 청년 주일미사_{date_str}_{liturgy_safe}.pptx"

    else:

        output_name = f"{date_str}_{liturgy_safe}.pptx"

    output_path = Path(OUTPUT_ROOT) / output_folder_key(date_str, mass_type) / output_name

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

    validate(prs2, json_data)



    _last_output_path[0] = output_path

    print(f'\n완료! → {output_path}')





def _show_children_mass_not_supported_message() -> None:
    """§J3 — '어린이' 미사유형은 아직 파이프라인이 없다. 입력 UI를 아예 열지 않고 안내만
    띄운 뒤 정상 종료한다(sys.exit이 아니라 __main__이 그냥 반환하게 함 — 예외적 실패가
    아니라 정상적인 '아직 지원 안 함' 흐름이라 exit code 0으로 조용히 끝나야 한다)."""
    import tkinter as tk
    from tkinter import messagebox
    _root = tk.Tk()
    _root.withdraw()
    messagebox.showinfo('어린이미사', '어린이미사는 아직 지원하지 않습니다.\n다음 업데이트를 기다려 주세요.')
    _root.destroy()


def _run_gui_mode(mass_type_kr: str) -> None:
    """CLI에서 date 없이 실행됐을 때(§J3) GUI 입력 흐름 전체를 담당한다. 원래는 `__main__`
    이 먼저 `_ask_mass_type_popup()`으로 미사 유형을 물은 뒤 이 흐름을 이어갔는데, 그
    팝업 자체가 삭제됐다(§J3) — `mass_type_kr`은 이제 `--미사유형`(기본 '성인') CLI 값을
    그대로 받는다. '어린이'는 호출부(`__main__`)가 이 함수를 부르기 전에 이미 걸러내므로
    여기서는 '성인'/'청년'만 들어온다고 가정한다. 그 아래 흐름(통합 입력 팝업 → 성가번호/
    성가선택 팝업 → 진행률 창 → 결과 창)은 기존과 완전히 동일하다."""
    _mass_type = _resolve_mass_type(mass_type_kr)

    # 청년미사는 통합 입력 팝업(_ask_combined_input_popup) 자체가 OneDrive 커스텀 파일
    # 브라우저(§A5)를 쓰므로, 로그인은 그 팝업을 열기 **전에** 끝나 있어야 한다 — 원래는
    # 성가 교체 워커 스레드 시작 직전(아래)에만 로그인했는데, 그러면 이 팝업이 먼저 열려
    # 로그인 안 된 상태로 '찾아보기'를 누르게 된다(청년미사 성가 자료 조회와 별개로,
    # 파일 브라우저도 같은 access token을 쓰므로 이 시점에 한 번만 끝내면 아래
    # ensure_onedrive_login() 재호출은 캐시 히트로 조용히 통과한다).
    if _mass_type == 'youth' and get_youth_ppt_mode(ask_if_missing=True) == 'fallback':

        try:

            ensure_onedrive_login()

        except Exception as e:

            import tkinter as tk

            from tkinter import messagebox

            _err_root = tk.Tk()

            _err_root.withdraw()

            messagebox.showerror('OneDrive 로그인 실패', f'OneDrive 로그인에 실패했습니다:\n{e}')

            _err_root.destroy()

            sys.exit(0)

    if _mass_type == 'adult':

        # 최초 1회 'PPT 문서' 폴더 선택(취소해도 기존 방식으로 진행). 첫 화면 시작 폴더·결과 복사에 쓰인다.

        get_adult_ppt_folder(ask_if_missing=True)

    try:

        _date_str, _files = _ask_combined_input_popup(mass_type=_mass_type)

    except RuntimeError:

        sys.exit(0)

    try:

        if _mass_type == 'youth':

            _성가_선택 = _ask_youth_hymn_popup()

            _hymn_numbers = {}

        else:

            _hymn_numbers = _ask_numbers_popup({}, is_sunday_mass(_date_str))

            _성가_선택 = None

    except RuntimeError:

        sys.exit(0)

    _preloaded_inputs[0] = {

        'mass_type':    _mass_type,

        'date_str':     _date_str,

        'files':        _files,

        'hymn_numbers': _hymn_numbers,

        '성가_선택':     _성가_선택,

    }

    # 성가 자료(OneDrive) 로그인은 이제 위(_ask_combined_input_popup 호출 전)에서 이미
    # 끝냈다 — 캐시된 토큰이 있으므로 워커 스레드 안에서 다시 필요해져도 조용히
    # 통과한다(2026-09-19 "성가 교체..." 88% 무한 대기 버그 수정 원칙 유지, §A5로 로그인
    # 시점만 더 앞으로 당김).

    _out_text, _err_text = _run_with_progress_window(main)

    if _err_text:

        # 팝업에는 구체적 에러 메시지(_err_text)만 보여준다. 전체 진행 로그(_out_text)는

        # log_text로 넘겨 로그 파일에만 남긴다.

        _show_result_window('오류 발생', _err_text, is_error=True, log_text=_out_text)

    else:

        _show_result_window('완료', _out_text or '')


def _dispatch_cli_or_gui() -> None:
    """`if __name__ == '__main__':`의 실제 진입 판단 로직 전체(§J3) — 함수로 뽑아둔
    이유는 오직 테스트 가능성이다. `if __name__ == '__main__':` 블록 자체는 모듈이
    스크립트로 실행될 때만 도니 단위 테스트로 직접 부를 수 없다(H그룹에서 확립한 "GUI
    트리거 로직은 테스트 가능한 형태로 분리" 원칙과 동일).

    미사 유형 선택 팝업을 완전히 없애고, `--미사유형`(CLI, 기본 '성인')으로 직접 받는다.
    "date가 실제로 주어졌는가"만으로 CLI 전체 모드 vs GUI 입력 모드를 가른다 —
    `--미사유형 청년`만 주고 date를 생략해도(플래그가 있으니 예전 `len(sys.argv) > 1`
    기준으로는 "CLI 모드"로 오판했을 상황) GUI 입력 모드로 들어가야 한다."""
    _args = _build_arg_parser().parse_args()

    # 리뷰 라운드11 지적: '어린이' 가드를 date 분기보다 아래(elif)에 두면 date가 같이
    # 주어진 CLI 호출(`missa_to_ppt.py 20260927 --미사유형 어린이`)이 이 가드를 완전히
    # 우회해 main()으로 직행한다 — main() 내부의 `_resolve_mass_type()`은 매핑에 없는
    # '어린이'를 조용히 'adult'로 폴백하므로, 안내 없이 성인미사 PPT가 그대로 생성된다.
    # date 유무와 무관하게 '어린이'를 최우선으로 걸러야 한다.
    if _args.mass_type_kr == '어린이':

        _show_children_mass_not_supported_message()

    elif _args.date is not None:

        # CLI 모드: date가 주어지면 바로 실행 (Python 또는 EXE 모두)

        main()

    else:

        # GUI 모드: 통합 입력 팝업 → 성가번호/성가선택 팝업 → 진행률 창 → 결과 창.
        # mass_type은 더 이상 팝업으로 묻지 않고 `--미사유형`(기본 '성인') CLI 값을 그대로
        # 쓴다(§J3).
        _run_gui_mode(_args.mass_type_kr)


if __name__ == '__main__':

    _dispatch_cli_or_gui()
