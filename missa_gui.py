"""
missa_to_ppt.py 모듈 분리 리팩토링 Phase 2 — Tkinter 팝업 + 날짜판단/설정.

is_sunday_mass()·config.json 로드/저장·get_onedrive_hymn_folder() 등은 논리적으로는
GUI가 아니지만, GUI 팝업(_ask_combined_input_popup 등)과 진입점(main, parse_args,
find_files) 양쪽에서 모두 호출된다. 진입점에 그대로 두면 entry↔gui 순환 임포트가
생기므로, 이 모듈에 흡수해 entry→gui 단방향 의존만 유지한다.
"""
from __future__ import annotations

import contextlib
import gc
import io
import json
import os
import re
import sys
from pathlib import Path

OUTPUT_ROOT = 'output'  # 날짜 폴더(YYYYMMDD/)를 모아두는 상위 폴더.
# entry(missa_to_ppt.py find_files/get_json_data/main)와 gui(_show_result_window) 양쪽에서
# 쓰여서 여기에 둔다 — entry→gui 단방향 의존만 유지하기 위함(design doc §2.3 참고).

# 미사 유형별 출력 폴더 접미사(2026-09-24 청년미사 확장 §B그룹). 성인=''(무변경 — 회귀
# 픽스처 20260624/20260712/20260705가 가리키는 실제 폴더 경로를 절대 바꾸지 않기 위함),
# 청년='_youth'. 향후 어린이미사를 추가할 때는 이 매핑에 '_children' 한 줄만 추가하면 된다
# (하드코딩 금지 원칙 — output_folder_key()가 이 매핑만 참조한다).
MASS_TYPE_OUTPUT_SUFFIX = {'adult': '', 'youth': '_youth'}


def output_folder_key(date_str: str, mass_type: str = 'adult') -> str:
    """`date_str`(사용자가 입력한 날짜 그대로 — 청년미사는 콘텐츠 조회일(+1일)이 아니라
    이 값을 쓴다, `missa_to_ppt._youth_content_date_str` 참고)에 미사 유형별 접미사를 붙인
    "출력 폴더 키"를 반환한다. 알 수 없는 mass_type은 KeyError 대신 접미사 없음으로
    안전하게 폴백한다(미래에 이 함수 호출부가 늘어나도 방어적으로 동작)."""
    return f'{date_str}{MASS_TYPE_OUTPUT_SUFFIX.get(mass_type, "")}'


def _log_dir_for(date_str: str, output_key: str = None) -> Path:
    """로그 저장 폴더 경로 — `output_key`(main()이 계산한 `output_folder_key()` 결과)가
    있으면 그 폴더를, 없으면(구버전 호출 등) date_str 그대로 폴백한다. tkinter에 의존하지
    않는 순수 함수로 분리해 GUI 없이 단위 테스트할 수 있게 한다."""
    return Path(OUTPUT_ROOT) / (output_key or date_str) / 'log'


# ─── GUI 상태 변수 ───────────────────────────────────────────────────────────
_progress_callback = [None]        # 진행률 콜백 (진행 창 표시 중에만 설정)
_first_dialog_pos  = [None, None]  # 첫 번째 다이얼로그 위치 (x, y)
_last_output_path  = [None]        # main()이 저장한 최종 출력 경로
_last_date_str     = [None]        # main()이 사용한 날짜 문자열
_last_output_key   = [None]        # main()이 사용한 출력 폴더 키(output_folder_key() 결과)
_last_output_mass    = [None]        # main()이 처리한 미사 유형('adult'|'youth') — 결과창 '원드라이브로 복사'용
_preloaded_inputs  = [None]        # EXE 흐름에서 main() 호출 전에 미리 수집한 입력값

# ─── UI 테마 (brokenbaykcc.org Look & Feel) ─────────────────────────────────
_UI = {
    'bg':       '#ffffff',
    'fg':       '#2c2c2c',
    'primary':  '#871b24',
    'pri_dark': '#59001d',
    'fg_light': '#ffffff',
    'font':     'Malgun Gothic',
    'font_sz':  10,
    'font_h':   11,
}

_ICON_PATH = Path(r'C:\Users\Scott\OneDrive\Photos_OneDrive\Misc\성당 로고 아이콘.PNG')
_icon_image = [None]  # PhotoImage를 GC로부터 보호하기 위해 캐시


def _set_window_icon(root):

    """팝업 창 아이콘을 성당 로고 이미지로 설정한다."""

    import tkinter as tk

    if not _ICON_PATH.exists():

        return

    try:

        if _icon_image[0] is None:

            _icon_image[0] = tk.PhotoImage(file=str(_ICON_PATH))

        root.iconphoto(True, _icon_image[0])

    except Exception:

        pass



def _apply_theme(root):
    """창 전체에 brokenbaykcc.org 색상·폰트(Malgun Gothic / #871b24)를 적용한다."""
    root.configure(bg=_UI['bg'])
    root.option_add('*Background',              _UI['bg'])
    root.option_add('*Foreground',              _UI['fg'])
    root.option_add('*Font',                    f'{{{_UI["font"]}}} {_UI["font_sz"]}')
    root.option_add('*Label.background',        _UI['bg'])
    root.option_add('*Label.foreground',        _UI['fg'])
    root.option_add('*Frame.background',        _UI['bg'])
    root.option_add('*Entry.background',        _UI['bg'])
    root.option_add('*Entry.foreground',        _UI['fg'])
    root.option_add('*Entry.relief',            'solid')
    root.option_add('*Entry.borderWidth',       1)
    root.option_add('*Button.background',       _UI['primary'])
    root.option_add('*Button.foreground',       _UI['fg_light'])
    root.option_add('*Button.font',             f'{{{_UI["font"]}}} {_UI["font_sz"]} bold')
    root.option_add('*Button.relief',           'flat')
    root.option_add('*Button.cursor',           'hand2')
    root.option_add('*Button.padX',             10)
    root.option_add('*Button.padY',             4)
    root.option_add('*Button.activeBackground', _UI['pri_dark'])
    root.option_add('*Button.activeForeground', _UI['fg_light'])


def _center_window(root):
    """창을 화면 정중앙에 배치한다."""
    root.update_idletasks()
    w = root.winfo_reqwidth()
    h = root.winfo_reqheight()
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    root.geometry(f'+{max(0, (sw - w) // 2)}+{max(0, (sh - h) // 2)}')



# ─────────────────────────────────────────────────────────────────────────────

# 설정 파일 (config.json)

# ─────────────────────────────────────────────────────────────────────────────



_SCRIPT_DIR = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
# MISSA_CONFIG_FILE: 테스트가 서브프로세스 실행 시 실제 config.json 대신 임시 설정을 쓰게 하는 용도.
CONFIG_FILE = Path(os.environ['MISSA_CONFIG_FILE']) if os.environ.get('MISSA_CONFIG_FILE') else _SCRIPT_DIR / 'config.json'



def _load_config() -> dict:

    if CONFIG_FILE.exists():

        try:

            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:

                return json.load(f)

        except Exception:

            return {}

    return {}



def _save_config(config: dict):

    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:

        json.dump(config, f, ensure_ascii=False, indent=2)



def _ask_onedrive_path_popup(title='성가 라이브러리 경로 설정',
                              message='성가 PPT 파일들이 저장된 OneDrive 폴더를 선택해 주세요.\n(최초 1회만 설정됩니다.)',
                              dialog_title='성가 OneDrive 폴더 선택') -> str:

    """tkinter 폴더 선택 다이얼로그로 OneDrive 폴더 경로 입력."""

    import tkinter as tk

    from tkinter import filedialog, messagebox

    root = tk.Tk()

    root.withdraw()

    messagebox.showinfo(title, message)

    folder = filedialog.askdirectory(title=dialog_title)

    root.destroy()

    return folder


# ─────────────────────────────────────────────────────────────────────────────
# OneDrive 폴더 루트 3종(공용/성인/청년) — 2026-10-18 폴더 구조 개편(SPEC_OneDrive_폴더구조_개편.md)
#
# 'PPT 문서' 단일 루트 대신 `00.공용 PPT`/`10.성인 미사`/`20.청년 미사` 세 로컬 폴더를 각각 config.json에
# 저장한다(운영자 권한에 따라 서로 다른 위치에 있을 수 있어 공통 부모를 가정할 수 없다). 성인 실행은
# 공용+성인, 청년 실행은 공용+청년만 묻는다. OneDrive 로그인(Graph) 폴백은 폐지됐다 — 루트를 못 정하면
# 안내 후 종료한다. 하위 폴더 이름 상수는 이 절에만 둔다(구조가 또 바뀌면 여기 한 곳만 고친다).
# ─────────────────────────────────────────────────────────────────────────────

_ROOT_ROLES = ('common', 'adult', 'youth')
# role -> (config 키, 루트 폴더 이름, 유효 판정용 하위 폴더, 안내문에 쓰는 이름)
_ROOT_SPECS = {
    'common': ('onedrive_common_folder', '00.공용 PPT', '09.가톨릭 성가', '공용'),
    'adult':  ('onedrive_adult_folder',  '10.성인 미사', '10.기타 PPT문서', '성인'),
    'youth':  ('onedrive_youth_folder',  '20.청년 미사', '2.성가', '청년'),
}
_MASS_ROOT_ROLES = {'adult': ('common', 'adult'), 'youth': ('common', 'youth')}
_LEGACY_PPT_FOLDER_KEY = 'onedrive_ppt_folder'   # 옛 'PPT 문서' 단일 루트 — 새 구조를 품고 있을 때만 이관한다.

_CATHOLIC_HYMN_SUBPATH = ('09.가톨릭 성가', '성가-악보버전')
_YOUTH_HYMN_SUBPATH = ('2.성가',)

_ROOTS = {role: None for role in _ROOT_ROLES}   # 프로세스 전역 캐시(검증된 Path)


def _reset_ppt_roots() -> None:
    """루트 캐시를 비운다(테스트·설정 변경 직후용)."""
    for role in _ROOT_ROLES:
        _ROOTS[role] = None


def _find_child_dir(parent: Path, name: str):
    """parent 아래에서 name과 공백 무시 비교로 일치하는 하위 폴더(없으면 None)."""
    exact = parent / name
    if exact.is_dir():
        return exact
    norm = name.replace(' ', '')
    try:
        for child in parent.iterdir():
            if child.is_dir() and child.name.replace(' ', '') == norm:
                return child
    except OSError:
        pass
    return None


def is_valid_ppt_root(role: str, path) -> bool:
    """디렉터리이며 그 안에 역할별 표지 하위 폴더가 있으면 유효한 루트(공백 유무 차이 무시)."""
    try:
        p = Path(path)
        return p.is_dir() and _find_child_dir(p, _ROOT_SPECS[role][2]) is not None
    except Exception:
        return False


def _derive_roots(parent) -> dict:
    """'PPT 문서'처럼 세 루트를 품은 상위 폴더에서 유효한 루트들을 찾아 {role: Path}로 반환한다."""
    found = {}
    try:
        p = Path(parent)
        if not p.is_dir():
            return found
        for role in _ROOT_ROLES:
            child = _find_child_dir(p, _ROOT_SPECS[role][1])
            if child is not None and is_valid_ppt_root(role, child):
                found[role] = child
    except Exception:
        pass
    return found


def get_root(role: str):
    """role('common'|'adult'|'youth')의 로컬 루트 Path 또는 None. 선택창은 절대 띄우지 않는다 —
    선택창은 GUI 진입점이 `ensure_ppt_roots(ask_if_missing=True)`로만 띄운다(다른 호출이 모달 창
    앞에서 멈추지 않게). 캐시 → config 저장값(유효할 때만) 순."""
    if _ROOTS.get(role) is not None:
        return _ROOTS[role]
    saved = _load_config().get(_ROOT_SPECS[role][0], '')
    if saved and is_valid_ppt_root(role, saved):
        _ROOTS[role] = Path(saved)
        return _ROOTS[role]
    return None


def _set_roots(found: dict) -> None:
    config = _load_config()
    for role, path in found.items():
        config[_ROOT_SPECS[role][0]] = str(path)
        _ROOTS[role] = Path(path)
    _save_config(config)


def _migrate_legacy_ppt_folder() -> None:
    """옛 `onedrive_ppt_folder`가 새 구조('PPT 문서' 아래 00./10./20.)를 품고 있으면 비어 있는 루트를
    자동으로 채운다. 옛 구조를 가리키면 아무것도 채우지 않는다(= 무효로 보고 선택창으로 간다).
    옛 `onedrive_hymn_folder`/`onedrive_youth_hymn_folder`는 더 이상 읽지 않는다."""
    config = _load_config()
    legacy = config.get(_LEGACY_PPT_FOLDER_KEY, '')
    if not legacy:
        return
    derived = {r: p for r, p in _derive_roots(legacy).items() if get_root(r) is None}
    if derived:
        _set_roots(derived)
        config = _load_config()
        config.pop(_LEGACY_PPT_FOLDER_KEY, None)
        _save_config(config)


def _select_root_interactively(role: str):
    """role 루트 폴더 선택창. 반환: {role: Path, ...}(상위 'PPT 문서'를 고르면 찾은 모든 루트) 또는
    취소 시 None. 무효한 폴더면 안내 후 다시 묻는다."""
    _key, dir_name, marker, label = _ROOT_SPECS[role]
    message = (f"'{dir_name}' 폴더({label}, 이 컴퓨터에서 OneDrive로 동기화 중인 폴더)를 선택해 주세요.\n"
               f"('PPT 문서' 상위 폴더를 선택해도 됩니다. 최초 1회만 설정됩니다.)")
    while True:
        folder = _ask_onedrive_path_popup(
            title=f"{label} 폴더 설정", message=message, dialog_title=f"'{dir_name}' 폴더 선택",
        )
        if not folder:
            return None
        if is_valid_ppt_root(role, folder):
            found = {role: Path(folder)}
            return found
        derived = _derive_roots(folder)
        if role in derived:
            return derived
        message = (f"선택한 폴더 안에 '{marker}' 폴더가 없습니다.\n"
                   f"'{dir_name}' 폴더(또는 그것들을 품은 'PPT 문서' 폴더)를 다시 선택해 주세요.")


def ensure_ppt_roots(mass_type: str, ask_if_missing: bool = False) -> bool:
    """이번 실행에 필요한 루트(성인=공용+성인, 청년=공용+청년)가 모두 유효한지 확인한다. 없거나 무효한
    루트는 `ask_if_missing`일 때만 선택창으로 받아 저장한다(불필요한 루트는 묻지 않는다). 모두 갖춰지면
    True, 선택이 취소되면 False(호출부가 안내 후 종료)."""
    needed = _MASS_ROOT_ROLES[mass_type]
    _migrate_legacy_ppt_folder()
    for role in needed:
        if get_root(role) is not None:
            continue
        if not ask_if_missing:
            return False
        found = _select_root_interactively(role)
        if not found:
            return False
        _set_roots(found)
        print(f"  [설정] {_ROOT_SPECS[role][3]} 폴더 저장됨: {found[role]}")
    return True


def _join_existing(root, segments):
    """root에서 segments를 따라 내려가되, 이미 있는 폴더는 공백 차이를 무시해 실제 이름을 쓰고, 없는
    폴더부터는 주어진 이름 그대로 이어 붙인다(저장 대상 폴더를 새로 만들 때도 쓰인다)."""
    cur = Path(root)
    exists = True
    for seg in segments:
        if not seg:
            continue
        nxt = _find_child_dir(cur, seg) if exists else None
        if nxt is None:
            exists = False
            cur = cur / seg
        else:
            cur = nxt
    return cur


def resolve_browse_start_dir(root, subpath):
    """root에서 subpath를 따라 내려간다. 중간에 없는 폴더가 있으면 거기까지(최악엔 root)를 반환."""
    if root is None:
        return None
    cur = Path(root)
    for seg in subpath or ():
        nxt = _find_child_dir(cur, seg)
        if nxt is None:
            break
        cur = nxt
    return cur


def _need_root(role: str) -> Path:
    root = get_root(role)
    if root is None:
        raise RuntimeError(
            f"{_ROOT_SPECS[role][3]} 폴더('{_ROOT_SPECS[role][1]}')가 설정되지 않았습니다. "
            "프로그램을 GUI로 실행해 폴더를 먼저 설정해 주세요.")
    return root


def ppt_subdir(kind: str, subfolder: str | None = None, year=None) -> Path:
    """루트로부터 계산한 로컬 하위 경로(하위 폴더 이름 하드코딩은 이 절에만 둔다). kind:
    'catholic_hymn'(공용) | 'youth_hymn'(청년, subfolder=나주노 성가/야훼이레 성가) |
    'adult_output'/'youth_output'(year). 필요한 루트가 없으면 RuntimeError."""
    if kind == 'catholic_hymn':
        return _join_existing(_need_root('common'), _CATHOLIC_HYMN_SUBPATH)
    if kind == 'youth_hymn':
        return _join_existing(_need_root('youth'), _YOUTH_HYMN_SUBPATH + (subfolder or '',))
    if kind == 'adult_output':
        return _join_existing(_need_root('adult'), (str(year),))
    if kind == 'youth_output':
        return _join_existing(_need_root('youth'), (str(year),))
    raise ValueError(f'알 수 없는 kind: {kind!r}')


def get_onedrive_hymn_folder() -> Path:
    """가톨릭 성가(성가-악보버전) 로컬 폴더: `{공용}/09.가톨릭 성가/성가-악보버전`. 공용 루트가 없으면
    RuntimeError(GUI 진입점이 시작 시 `ensure_ppt_roots`로 먼저 받아 둔다)."""
    return ppt_subdir('catholic_hymn')


def find_youth_hymn_file(kind: str, subfolder: str | None, pattern: str) -> Path | None:
    """성가 폴더(`ppt_subdir(kind, subfolder)`)를 재귀 검색해 파일명이 `pattern`(정규식)에 매칭되는 첫
    .pptx를 반환한다. 폴더가 없거나 못 찾으면 None. 잠금 파일(~$)은 제외."""
    local_dir = ppt_subdir(kind, subfolder)
    if not local_dir.is_dir():
        return None
    for f in sorted(local_dir.rglob('*.pptx')):
        if not f.name.startswith('~$') and re.search(pattern, f.name):
            return f
    return None


# ─────────────────────────────────────────────────────────────────────────────
# 찾아보기 시작 폴더와 행별 기억(SPEC §2.2/§2.3). 우선순위: 저장된 폴더(존재할 때만) → 기본 하위 폴더 →
# 해당 루트 → None(OS 기본 위치). 저장은 행별이며 참조 PPT만 성인/청년을 구분한다.
# ─────────────────────────────────────────────────────────────────────────────

_BROWSE_DIRS_KEY = 'browse_dirs'
_LEGACY_PSALM_DIR_KEY = 'last_psalm_dir'   # 행별 기억(`browse_dirs`)이 흡수한 옛 키
# 행 id -> (루트 role, 루트 아래 기본 하위 경로)
_BROWSE_DEFAULTS = {
    '시작기도':      ('common', ('13.기도문',)),
    '미사후기도':    ('common', ('13.기도문',)),
    '공지사항':      ('common', ('11.공지사항 PPT문서',)),
    'adult.ref_pptx': ('adult', ('10.기타 PPT문서', 'PPT 자동화', '1.Template')),
    'youth.ref_pptx': ('youth', ('1.Template',)),
    '화답송_pptx':   ('adult', ('14.화답송 악보',)),
}


def _browse_row_id(mass_type: str, key: str) -> str:
    return f'{mass_type}.{key}' if key == 'ref_pptx' else key


def browse_initial_dir(mass_type: str, key: str):
    """로컬 파일 선택창의 initialdir(Path) 또는 None(저장값·루트 모두 없음 — OS 기본 위치)."""
    row = _browse_row_id(mass_type, key)
    config = _load_config()
    saved = (config.get(_BROWSE_DIRS_KEY) or {}).get(row, '')
    if not saved and row == '화답송_pptx':
        saved = config.get(_LEGACY_PSALM_DIR_KEY, '')
    if saved and Path(saved).is_dir():
        return Path(saved)
    role, subpath = _BROWSE_DEFAULTS.get(row, (mass_type, ()))
    return resolve_browse_start_dir(get_root(role), subpath)


def remember_browse_dir(mass_type: str, key: str, path) -> None:
    """찾아보기로 고른 파일의 폴더를 그 행의 다음 시작 폴더로 기억한다(옛 `last_psalm_dir`은 흡수·삭제)."""
    try:
        config = _load_config()
        dirs = dict(config.get(_BROWSE_DIRS_KEY) or {})
        dirs[_browse_row_id(mass_type, key)] = str(Path(path).parent)
        config[_BROWSE_DIRS_KEY] = dirs
        if _browse_row_id(mass_type, key) == '화답송_pptx':
            config.pop(_LEGACY_PSALM_DIR_KEY, None)
        _save_config(config)
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────────────────
# 결과 PPT '원드라이브로 복사'(결과창 버튼): 성인 `{성인}/{YYYY}`, 청년 `{청년}/{YYYY}`.
# ─────────────────────────────────────────────────────────────────────────────

def result_copy_dir(mass_type: str, year):
    """결과 PPT 복사 대상의 로컬 폴더. 루트를 모르면 None."""
    if get_root(mass_type) is None:
        return None
    return ppt_subdir(f'{mass_type}_output', year=year)


def result_copy_exists(mass_type: str, output_path, date_str: str) -> bool:
    """복사 대상에 같은 이름의 파일이 이미 있는지(덮어쓰기 확인용). 확인 불가면 False."""
    dest_dir = result_copy_dir(mass_type, date_str[:4])
    return dest_dir is not None and (dest_dir / Path(output_path).name).exists()


def copy_result_to_onedrive(mass_type: str, output_path, date_str: str) -> str:
    """결과 PPT를 연도 폴더로 복사하고 저장된 위치(표시용 문자열)를 반환한다. 연도 폴더가 없으면 만든다
    (OneDrive 앱이 동기화). 루트를 모르면 RuntimeError."""
    import shutil
    output_path = Path(output_path)
    dest_dir = result_copy_dir(mass_type, date_str[:4])
    if dest_dir is None:
        raise RuntimeError(f"'{_ROOT_SPECS[mass_type][1]}' 폴더가 설정되지 않았습니다.")
    dest_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(output_path, dest_dir / output_path.name)
    return str(dest_dir)


def _copy_result_with_ui(parent, mass_type: str, output_path, date_str: str) -> None:
    """결과창 '원드라이브로 복사' 버튼 핸들러: 루트 미설정이면 선택창 → 덮어쓰기 확인 → 복사 →
    저장 폴더 안내(실패하면 오류 안내)."""
    from tkinter import messagebox
    if get_root(mass_type) is None:
        if not ensure_ppt_roots(mass_type, ask_if_missing=True):
            messagebox.showwarning('복사 취소', f"'{_ROOT_SPECS[mass_type][1]}' 폴더가 선택되지 않아 복사하지 않았습니다.",
                                   parent=parent)
            return
    if result_copy_exists(mass_type, output_path, date_str):
        if not messagebox.askyesno('덮어쓰기 확인',
                                   f'같은 이름의 파일이 이미 있습니다:\n{Path(output_path).name}\n\n덮어쓸까요?',
                                   parent=parent):
            return
    try:
        dest = copy_result_to_onedrive(mass_type, output_path, date_str)
    except Exception as e:
        messagebox.showerror('복사 실패', f'원드라이브로 복사하지 못했습니다:\n{e}', parent=parent)
        return
    messagebox.showinfo('복사 완료', f'다음 폴더에 저장했습니다:\n{dest}', parent=parent)


_COM_VERIFY_ENABLED_CACHE = [None]


def _com_verification_enabled() -> bool:
    """config.json의 com_verification_enabled 값(기본 True)을 memoize해서 반환한다.
    PowerPoint COM 실측이 실패하는 모든 경로는 이미 자동으로 Pillow 전용 폴백하므로
    기본값은 켜둬도 안전하다 — 이 플래그는 COM을 아예 쓰고 싶지 않은 머신을 위한
    명시적 opt-out이다."""
    if _COM_VERIFY_ENABLED_CACHE[0] is None:
        enabled = bool(_load_config().get('com_verification_enabled', True))
        _COM_VERIFY_ENABLED_CACHE[0] = enabled
        if not enabled:
            print('  [설정] config.json: com_verification_enabled=false → PowerPoint COM 검증 비활성화(Pillow 추정치만 사용)')
    return _COM_VERIFY_ENABLED_CACHE[0]


# ─────────────────────────────────────────────────────────────────────────────

# CLI / tkinter 팝업

# ─────────────────────────────────────────────────────────────────────────────


def is_sunday_mass(date_str: str) -> bool:
    """날짜 문자열(YYYYMMDD)이 일요일이면 True."""
    from datetime import datetime
    return datetime.strptime(date_str, '%Y%m%d').weekday() == 6




# 미사 유형별 성가 출처 목록(청년미사 2단계 §3). 향후 어린이미사 등을 추가할 때는 이
# 딕셔너리에 항목만 추가하면 된다 — 완전한 플러그인 아키텍처는 이번 범위에서 과설계라
# 채택하지 않았다(구현 계획 §3).
MASS_TYPES = {
    '성인': {'label': '성인미사', 'hymn_sources': ['가톨릭성가']},
    '청년': {'label': '청년미사(토요일 저녁)', 'hymn_sources': ['나주노', '야훼 이레', '가톨릭성가', '기타']},
}


def _get_app_version() -> str:
    """VERSION 파일(§J10, repo 루트 = `_SCRIPT_DIR`, 형식 YYYYMMDD-NN)을 읽어 반환한다.
    자동 증가 메커니즘은 두지 않는다 — 배포 시 사람이 직접 VERSION 파일을 갱신한다는
    전제로, 이 함수는 오직 "읽기 + 없을 때 폴백"만 담당한다(코디네이터 명시 지시: 자동
    증가는 이 규모 프로젝트에 과설계)."""
    from datetime import datetime
    version_path = _SCRIPT_DIR / 'VERSION'
    try:
        text = version_path.read_text(encoding='utf-8').strip()
        if text:
            return text
    except OSError:
        pass
    return datetime.now().strftime('%Y%m%d') + '-01'


def format_update_progress(stage: str, done: int, total) -> tuple:
    """업데이트 단계 → ('단계 문구\\n상세 문구', 비율 또는 None). None이면 진행률 미상(애니메이션 막대)."""
    import missa_updater as mu
    mb = 1024 * 1024
    if stage == mu.STAGE_CHECK:
        return '최신 버전 확인 중…\n', None
    if stage == mu.STAGE_DOWNLOAD:
        head = '새 버전 내려받는 중…'
        if not total:
            return f'{head}\n{done / mb:.1f} MB', None
        frac = max(0.0, min(1.0, done / total))
        return f'{head}\n{done / mb:.1f} MB / {total / mb:.1f} MB ({int(frac * 100)}%)', frac
    if stage in (mu.STAGE_EXTRACT, mu.STAGE_INSTALL):
        head = '압축 푸는 중…' if stage == mu.STAGE_EXTRACT else '파일 교체 중…'
        if not total:
            return f'{head}\n', None
        frac = max(0.0, min(1.0, done / total))
        return f'{head}\n{done} / {total}', frac
    if stage == mu.STAGE_DONE:
        return '완료\n', 1.0
    return '\n', None


def _update_worker(func, q) -> None:
    """작업 스레드 본체. 화면은 건드리지 않고 큐에만 넣는다(Tkinter는 메인 스레드 전용)."""
    try:
        result = func(lambda stage, done, total: q.put(('progress', stage, done, total)))
    except Exception as e:  # noqa: BLE001 — 메인 스레드에서 종류별로 처리
        q.put(('error', e))
    else:
        q.put(('result', result))


def _drain_update_queue(q, on_progress, on_result, on_error) -> bool:
    """큐를 비우며 마지막 진행값만 화면에 반영(다운로드 청크마다 갱신하면 UI가 밀린다).
    결과/오류를 처리했으면 True."""
    import queue as _queue
    last = None
    finished = False
    while True:
        try:
            msg = q.get_nowait()
        except _queue.Empty:
            break
        if msg[0] == 'progress':
            last = msg[1:]
        elif msg[0] == 'result':
            if last is not None:
                on_progress(*last)
                last = None
            on_result(msg[1])
            finished = True
            break
        elif msg[0] == 'error':
            on_error(msg[1])
            finished = True
            last = None
            break
    if last is not None:
        on_progress(*last)
    return finished


@contextlib.contextmanager
def _tk_gc_guard():
    """작업 스레드가 도는 동안 Tk 쓰레기가 그 스레드에서 수거되지 않게 한다.

    앞서 닫은 창(입력 창 등)의 Tk 객체는 순환 참조라 곧바로 해제되지 않고 남아 있다가, 작업 스레드가 객체를
    많이 만들 때 도는 GC에서 그 스레드가 수거해 버린다. Tk는 만든 스레드(메인)에서만 해제해야 하므로 그러면
    `Variable.__del__`이 "main thread is not in main loop"를 내고 Tcl 인터프리터가 누수된다(2026-10-25 운영자
    PC 로그). 시작 전에 메인 스레드에서 미리 수거하고, 스레드가 도는 동안은 남은 객체를 GC 대상에서 빼 둔다."""
    gc.collect()
    gc.freeze()
    try:
        yield
    finally:
        gc.unfreeze()
        gc.collect()


def _run_with_progress(parent, title: str, func) -> tuple:
    """모달 진행 창을 띄우고 func(progress)를 작업 스레드에서 실행. 반환 (성공여부, 결과 또는 예외).
    취소 불가이며 작업 중 닫기(X)는 무시한다(파일 교체 도중 중단되면 설치가 반쯤 섞인다)."""
    import queue as _queue
    import threading
    import tkinter as tk
    from tkinter import ttk

    win = tk.Toplevel(parent)
    win.title(title)
    win.resizable(False, False)
    win.transient(parent)
    win.protocol('WM_DELETE_WINDOW', lambda: None)
    _set_window_icon(win)
    frame = tk.Frame(win, padx=24, pady=18)
    frame.pack(fill='both', expand=True)
    stage_var, detail_var = tk.StringVar(), tk.StringVar()
    tk.Label(frame, textvariable=stage_var, anchor='w', font=(_UI['font'], _UI['font_sz'], 'bold')).pack(fill='x')
    bar = ttk.Progressbar(frame, length=360, mode='indeterminate')
    bar.pack(fill='x', pady=8)
    tk.Label(frame, textvariable=detail_var, anchor='w', font=(_UI['font'], _UI['font_sz'])).pack(fill='x')
    bar.start(15)
    stage_var.set('최신 버전 확인 중…' if title == '업데이트 확인' else '준비 중…')
    win.update_idletasks()
    try:
        x = parent.winfo_rootx() + (parent.winfo_width() - win.winfo_reqwidth()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - win.winfo_reqheight()) // 2
        win.geometry(f'+{max(x, 0)}+{max(y, 0)}')
    except tk.TclError:
        pass

    q = _queue.Queue()
    outcome = {}
    mode = {'determinate': False}

    def on_progress(stage, done, total):
        text, frac = format_update_progress(stage, done, total)
        head, _, detail = text.partition('\n')
        stage_var.set(head)
        detail_var.set(detail)
        if frac is None:
            if mode['determinate']:
                bar.configure(mode='indeterminate')
                bar.start(15)
                mode['determinate'] = False
        else:
            if not mode['determinate']:
                bar.stop()
                bar.configure(mode='determinate', maximum=100)
                mode['determinate'] = True
            bar['value'] = frac * 100

    def poll():
        if _drain_update_queue(q, on_progress, lambda v: outcome.update(ok=True, value=v),
                               lambda e: outcome.update(ok=False, value=e)):
            try:
                win.grab_release()
            except tk.TclError:
                pass
            win.destroy()
        else:
            win.after(50, poll)

    with _tk_gc_guard():
        threading.Thread(target=_update_worker, args=(func, q), daemon=True).start()
        win.after(50, poll)
        win.grab_set()
        win.wait_window()
    return outcome['ok'], outcome['value']


def _on_check_update(parent) -> None:
    """'업데이트 확인' 클릭 시 최신 커밋과 비교해 있으면 적용한다.

    §J3 — 원래 `_ask_mass_type_popup()`(미사 유형 선택 팝업, 통째로 삭제됨) 안의 클로저
    `on_check_update()`였다. 그 팝업이 사라지면서 이 로직만 그대로 옮겼다(재구현 아님,
    코디네이터 명시 지시) — 유일한 필수 변경은 클로저가 캡처하던 `root`를 명시적 `parent`
    매개변수로 바꾼 것뿐이다(더 이상 자신을 감싸는 팝업이 없으므로, 호출부가 자신의
    창을 넘겨야 한다). §J10에서 `_ask_combined_input_popup()`의 버전 표시 링크에 이
    함수를 다시 배선한다."""
    import missa_updater as mu
    from tkinter import messagebox
    config = _load_config()

    def _check(progress):
        progress(mu.STAGE_CHECK, 0, None)
        return mu.check_for_update(config)

    ok, value = _run_with_progress(parent, '업데이트 확인', _check)
    if not ok:
        if not isinstance(value, mu.UpdateError):
            raise value
        messagebox.showerror('업데이트 확인 실패', str(value), parent=parent)
        return
    has_update, sha = value
    if not has_update:
        messagebox.showinfo('업데이트 확인', '이미 최신 버전입니다.', parent=parent)
        return
    if not messagebox.askyesno('업데이트 확인', '새 버전이 있습니다. 지금 적용할까요?', parent=parent):
        return
    ok, value = _run_with_progress(
        parent, '업데이트 적용', lambda progress: mu.apply_update(sha, _SCRIPT_DIR, progress))
    if not ok:
        if not isinstance(value, mu.UpdateError):
            raise value
        messagebox.showerror('업데이트 적용 실패', str(value), parent=parent)
        return
    config['last_update_commit'] = sha
    _save_config(config)
    messagebox.showinfo('업데이트 완료', '업데이트가 적용됐습니다. 프로그램을 다시 시작해 주세요.', parent=parent)
    parent.destroy()
    sys.exit(0)


def _ask_youth_hymn_popup(defaults: dict = None) -> dict:
    """청년미사 성가 5종(성체는 최대 2곡)을 출처+번호(또는 제목)로 입력받는다.
    반환: 구현 계획 §9.3 '성가_선택' 형식 — dict[str, list[dict]],
    각 dict는 {'출처': str, '번호': int|None, '제목': str|None}.

    UI 표시명은 '야훼 이레' 대신 '야훼이레'(공백 없음)를 쓰지만, 반환값은
    missa_youth_hymn_pdf.SOURCES의 원본 키('야훼 이레', 공백 포함)로 정규화한다 —
    이 정규화를 UI 계층이 아니라 다른 곳에서 하면 조회가 조용히 실패하므로 여기서 끝낸다
    (구현 계획 §9.3 주의사항)."""
    import tkinter as tk
    from tkinter import messagebox, ttk

    defaults = defaults or {}

    DISPLAY_TO_SOURCE = {'나주노': '나주노', '야훼이레': '야훼 이레', '가톨릭성가': '가톨릭성가', '기타': '기타'}
    SOURCE_DISPLAYS = list(DISPLAY_TO_SOURCE)

    root = tk.Tk()
    root.withdraw()
    root.title('청년미사 성가 입력')
    root.resizable(False, False)
    _set_window_icon(root)
    _apply_theme(root)
    frame = tk.Frame(root, padx=20, pady=16)
    frame.pack(fill='both', expand=True)

    # 컬럼 헤더(요구사항 §A6) — htype 자체(입당/봉헌 등)를 나타내는 0번 열은 헤더가 없다.
    header_font = (_UI['font'], _UI['font_sz'], 'bold')
    tk.Label(frame, text='성가 구분', anchor='w', font=header_font).grid(
        row=0, column=1, sticky='w', padx=(8, 4), pady=(0, 4)
    )
    tk.Label(frame, text='성가 번호', anchor='w', font=header_font).grid(
        row=0, column=2, sticky='w', padx=4, pady=(0, 4)
    )
    tk.Label(frame, text='성가 제목', anchor='w', font=header_font).grid(
        row=0, column=3, sticky='w', padx=4, pady=(0, 4)
    )

    htypes = ['입당', '봉헌', '성체1', '성체2', '2차봉헌', '파견']
    rows = {}

    for i, htype in enumerate(htypes):
        r = i + 1  # 0행은 헤더가 차지
        label = '성체2 (선택)' if htype == '성체2' else htype
        tk.Label(frame, text=label, anchor='w').grid(row=r, column=0, sticky='w', pady=4)

        src_var = tk.StringVar(value=SOURCE_DISPLAYS[0])
        src_combo = ttk.Combobox(frame, textvariable=src_var, values=SOURCE_DISPLAYS,
                                  state='readonly', width=10)
        src_combo.grid(row=r, column=1, sticky='w', padx=(8, 4))

        num_entry = tk.Entry(frame, width=6)
        num_entry.grid(row=r, column=2, sticky='w', padx=4)

        title_entry = tk.Entry(frame, width=22)
        title_entry.grid(row=r, column=3, sticky='w', padx=4)
        title_entry.grid_remove()  # 기본 숨김 — 출처가 '기타'일 때만 보여준다(요구사항 §A6)

        def _on_source_selected(_event=None, num_entry=num_entry, title_entry=title_entry,
                                 src_var=src_var):
            if DISPLAY_TO_SOURCE[src_var.get()] == '기타':
                num_entry.delete(0, tk.END)
                num_entry.configure(state='disabled')
                title_entry.grid()
                title_entry.focus_set()
            else:
                num_entry.configure(state='normal')
                title_entry.grid_remove()

        src_combo.bind('<<ComboboxSelected>>', _on_source_selected)

        d = defaults.get(htype, {})
        if d.get('출처'):
            src_var.set(d['출처'])
        if d.get('번호'):
            num_entry.insert(0, str(d['번호']))
        if d.get('제목'):
            title_entry.insert(0, d['제목'])

        _on_source_selected()  # defaults가 이미 출처='기타'를 채웠을 수 있으므로 초기 동기화

        rows[htype] = (src_var, num_entry, title_entry)

    tk.Label(frame, text='(번호는 나주노/야훼이레/가톨릭성가 필수, 기타는 제목만 입력)',
             fg='#666666', font=(_UI['font'], 8)).grid(
        row=len(htypes) + 1, column=0, columnspan=4, sticky='w', pady=(4, 10))

    result = {}

    def on_ok():
        선택 = {}
        errors = []
        for htype in ['입당', '봉헌', '성체1', '성체2', '2차봉헌', '파견']:
            src_var, num_entry, title_entry = rows[htype]
            display = src_var.get()
            num_raw = num_entry.get().strip()
            title_raw = title_entry.get().strip()

            if htype == '성체2' and not num_raw and not title_raw:
                continue  # 2번째 성체곡은 선택사항 — 완전히 비워두면 스킵

            출처 = DISPLAY_TO_SOURCE[display]
            entry = {'출처': 출처, '번호': None, '제목': title_raw or None}

            if 출처 == '기타':
                if not title_raw:
                    errors.append(f'{htype}: 기타 출처는 제목이 필수입니다')
                    continue
            else:
                if not num_raw or not num_raw.isdigit():
                    errors.append(f'{htype}: 성가 번호를 숫자로 입력해 주세요')
                    continue
                entry['번호'] = int(num_raw)

            key = '성체' if htype in ('성체1', '성체2') else htype
            선택.setdefault(key, []).append(entry)

        if errors:
            messagebox.showerror('입력 오류', '\n'.join(errors), parent=root)
            return
        result['성가_선택'] = 선택
        root.destroy()

    tk.Button(frame, text='확인', command=on_ok, width=10).grid(
        row=len(htypes) + 2, column=0, columnspan=4, pady=10
    )
    root.update_idletasks()
    _w, _h = 460, root.winfo_reqheight()
    _sw, _sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry(f'{_w}x{_h}+{max(0,(_sw-_w)//2)}+{max(0,(_sh-_h)//2)}')
    root.deiconify()
    root.mainloop()
    # 확인 버튼이 아니라 X/Alt+F4로 닫으면 result가 비어 있다. 무경고로 빈 dict({})를
    # 돌려주면 replace_성가_youth(prs, {})가 5종 전부를 조용히 스킵해(entries=[] → continue)
    # 이번 주 성가가 하나도 갱신되지 않은 채 "완료"로 저장된다 — 독립 리뷰가 실제 참조 PPT로
    # 재현한 확정 버그(05_review_report.md). 다른 입력 팝업(_ask_date_popup 등)과 동일하게
    # 명시적으로 취소를 알린다.
    if '성가_선택' not in result:
        raise RuntimeError('청년미사 성가 입력이 취소됐습니다.')
    return result['성가_선택']


def _ask_numbers_popup(defaults: dict, is_sunday: bool = True) -> dict:

    import tkinter as tk

    from tkinter import messagebox



    root = tk.Tk()

    root.withdraw()

    root.title('성가 번호 입력')

    root.resizable(False, False)

    _set_window_icon(root)

    _apply_theme(root)

    frame = tk.Frame(root, padx=24, pady=16)

    frame.pack(fill='both', expand=True)



    labels = ['입당', '봉헌', '성체', '2차봉헌', '파견']

    required_labels = labels if is_sunday else [l for l in labels if l != '2차봉헌']

    entries = {}



    for i, label in enumerate(labels):

        tk.Label(frame, text=label).grid(row=i, column=0, sticky='w', pady=6)

        entry = tk.Entry(frame, width=6)

        entry.grid(row=i, column=1, sticky='w', pady=6, padx=(8, 0))

        if defaults.get(label):

            entry.insert(0, str(defaults[label]))

        entries[label] = entry



    result = {}



    def on_ok():

        raw = {label: entries[label].get().strip() for label in labels}

        missing = [label for label in required_labels if not raw[label]]

        if missing:

            messagebox.showerror(

                '입력 오류',

                f'다음 성가 번호가 입력되지 않았습니다: {", ".join(missing)}',

                parent=root,

            )

            return

        non_numeric = [label for label in labels if raw[label] and not raw[label].isdigit()]

        if non_numeric:

            messagebox.showerror(

                '입력 오류',

                f'성가 번호는 숫자로만 입력해 주세요: {", ".join(non_numeric)}',

                parent=root,

            )

            return

        for label in labels:

            result[label] = int(raw[label]) if raw[label] else None

        root.destroy()



    tk.Button(frame, text='확인', command=on_ok, width=10).grid(

        row=len(labels), column=0, columnspan=2, pady=14

    )

    root.update_idletasks()

    _w, _h = 280, root.winfo_reqheight()

    _sw, _sh = root.winfo_screenwidth(), root.winfo_screenheight()

    root.geometry(f'{_w}x{_h}+{max(0,(_sw-_w)//2)}+{max(0,(_sh-_h)//2)}')

    root.deiconify()

    root.after(50, lambda: entries['입당'].focus_set())

    root.mainloop()

    return result




# ─────────────────────────────────────────────────────────────────────────────

# 파일 탐색

# ─────────────────────────────────────────────────────────────────────────────



def _ask_date_popup() -> str:

    """미사 일자를 YYYYMMDD 형식으로 입력받는 팝업."""

    import tkinter as tk

    from tkinter import messagebox



    root = tk.Tk()

    root.withdraw()

    root.title('미사 일자 입력')

    root.resizable(False, False)

    _set_window_icon(root)

    _apply_theme(root)



    frame = tk.Frame(root, padx=24, pady=20)

    frame.pack(fill='both', expand=True)



    tk.Label(frame, text='미사 일자 (YYYYMMDD):').pack(anchor='w')

    entry = tk.Entry(frame, width=20, font=('', 12))

    entry.pack(pady=8)

    entry.focus_set()



    result = [None]



    def on_ok():

        val = entry.get().strip()

        if not re.match(r'^\d{8}$', val):

            messagebox.showwarning('입력 오류', 'YYYYMMDD 형식으로 입력해 주세요.\n예: 20260628')

            return

        result[0] = val

        root.destroy()



    entry.bind('<Return>', lambda e: on_ok())

    tk.Button(frame, text='확인', command=on_ok, width=10).pack()

    _center_window(root)

    root.deiconify()

    root.mainloop()



    if result[0] is None:

        raise RuntimeError('미사 일자 입력이 취소됐습니다.')

    return result[0]




def _ask_input_files_popup() -> dict:

    """참조 미사 PPT(필수), 시작기도 PPT(선택), 화답송 악보 PPT(필수)를

    하나의 창에서 선택한다."""

    import tkinter as tk

    from tkinter import filedialog, messagebox



    root = tk.Tk()

    root.withdraw()

    root.title('입력 파일 선택')

    root.resizable(False, False)

    _set_window_icon(root)

    _apply_theme(root)



    PPTX_TYPES = [('PowerPoint 파일', '*.pptx'), ('모든 파일', '*.*')]

    # 화답송 악보 행 전용 filetypes (공유 상수 PPTX_TYPES는 다른 행 회귀 방지를 위해 건드리지

    # 않는다 — 설계서 §4.2). 원본 사진(PNG/JPG) 입력을 지원한다.

    HWADAPSONG_TYPES = [

        ('악보 파일', '*.pptx *.png *.jpg *.jpeg'),

        ('PowerPoint', '*.pptx'),

        ('이미지', '*.png *.jpg *.jpeg'),

        ('모든 파일', '*.*'),

    ]

    _IMG_EXTS = ('.png', '.jpg', '.jpeg')



    frame = tk.Frame(root, padx=20, pady=16)

    frame.pack(fill='both', expand=True)



    rows = [

        ('ref_pptx',    '참조 미사 PPT  *',  True),

        ('화답송_pptx', '화답송 악보 PPT/사진', False),

        ('시작기도',    '시작기도 PPT',      False),

        ('미사후기도',  '미사 후 기도 PPT',  False),

    ]



    vars_ = {}

    for i, (key, label, _required) in enumerate(rows):

        tk.Label(frame, text=label, anchor='w', width=18).grid(

            row=i, column=0, sticky='w', pady=6

        )

        var = tk.StringVar()

        vars_[key] = var

        tk.Entry(frame, textvariable=var, width=46, state='readonly').grid(

            row=i, column=1, padx=(8, 6), pady=6

        )

        row_types = HWADAPSONG_TYPES if key == '화답송_pptx' else PPTX_TYPES

        def _browse(v=var, types=row_types):

            path = filedialog.askopenfilename(

                parent=root, filetypes=types

            )

            if path:

                v.set(path)

        tk.Button(frame, text='찾아보기', command=_browse, pady=0).grid(

            row=i, column=2, pady=6

        )



    result = [None]



    def on_ok():

        ref = vars_['ref_pptx'].get()

        if not ref:

            messagebox.showwarning('입력 오류', '참조 미사 PPT를 선택해 주세요.', parent=root)

            return

        화답송 = vars_['화답송_pptx'].get()

        # 확장자로 PPT/이미지 분기 (설계서 §4.2)

        화답송_pptx_val = None

        화답송_img_val = None

        if 화답송:

            _p = Path(화답송)

            if _p.suffix.lower() in _IMG_EXTS:

                화답송_img_val = _p

            else:

                화답송_pptx_val = _p

        시작기도 = vars_['시작기도'].get()

        미사후기도 = vars_['미사후기도'].get()

        result[0] = {

            'ref_pptx':    Path(ref),

            '화답송_pptx': 화답송_pptx_val,

            '화답송_img':  화답송_img_val,

            '시작기도':    Path(시작기도) if 시작기도 else None,

            '미사후기도':  Path(미사후기도) if 미사후기도 else None,

            '성가':        {},

        }

        root.destroy()



    tk.Button(frame, text='확인', command=on_ok, width=10).grid(

        row=len(rows), column=0, columnspan=3, pady=12

    )

    _center_window(root)

    root.deiconify()

    root.mainloop()



    if result[0] is None:

        raise RuntimeError('파일 선택이 취소됐습니다.')

    return result[0]




def _ask_combined_input_popup(mass_type: str = 'adult') -> tuple:

    """미사 일자 + 참조 파일을 하나의 창에서 입력받는다.

    반환: (date_str, files_dict)

    mass_type='youth'(청년미사)는 화답송 악보 PPT/사진 행 자체가 없고(요구사항 §A4 — 청년미사 악보는 성가
    선택 팝업에서 별도 출처로 다룬다) 파일 행 순서가 다르다(§J4). '찾아보기'는 두 유형 모두 로컬 파일
    다이얼로그이며 시작 폴더는 행별 기억 → 기본 하위 폴더 → 루트 순(`browse_initial_dir`).
    mass_type 인자를 생략하면 기존 호출부(성인 전용 경로) 동작과 100% 동일하다."""

    import tkinter as tk

    from tkinter import filedialog, messagebox



    # §J3 — 미사 유형 선택 팝업이 없어져 창 제목만 보고는 어느 유형 입력창인지 구분할
    # 수단이 사라졌다. 제목에 접미사를 붙여 보완한다('어린이'는 이 팝업까지 도달하지
    # 않지만(__main__이 먼저 걸러냄) 향후를 위해 매핑은 남겨둔다).
    _DIALOG_TITLE_SUFFIX = {'adult': ' (성인미사)', 'youth': ' (청년미사)', 'children': ' (어린이 미사)'}
    DIALOG_TITLE = '미사 PPT 파일 자동화 - 브로큰베이교구 한인 천주교회' + _DIALOG_TITLE_SUFFIX.get(mass_type, '')

    PPTX_TYPES = [('PowerPoint 파일', '*.pptx'), ('모든 파일', '*.*')]

    # 화답송 악보 행 전용 filetypes (공유 상수 PPTX_TYPES는 다른 행 회귀 방지를 위해 건드리지

    # 않는다 — 설계서 §4.2). 원본 사진(PNG/JPG) 입력을 지원한다.

    HWADAPSONG_TYPES = [

        ('악보 파일', '*.pptx *.png *.jpg *.jpeg'),

        ('PowerPoint', '*.pptx'),

        ('이미지', '*.png *.jpg *.jpeg'),

        ('모든 파일', '*.*'),

    ]

    _IMG_EXTS = ('.png', '.jpg', '.jpeg')



    root = tk.Tk()

    root.withdraw()

    root.title(DIALOG_TITLE)

    root.resizable(False, False)

    _set_window_icon(root)

    _apply_theme(root)



    frame = tk.Frame(root, padx=20, pady=16)

    frame.pack(fill='both', expand=True)



    # ── 날짜 입력 (레이블과 입력창 같은 행, PPT 입력란과 열 정렬) ──
    tk.Label(frame, text='미사 일자 (YYYYMMDD)', anchor='w').grid(

        row=0, column=0, sticky='w', pady=(6, 20)

    )

    date_var = tk.StringVar()

    date_entry = tk.Entry(frame, textvariable=date_var, width=20)

    date_entry.grid(row=0, column=1, sticky='w', pady=(6, 20), padx=(8, 6), columnspan=2)



    # ── 파일 선택 ──────────────────────────────────────────────
    file_rows = [

        ('ref_pptx',    '참조 미사 PPT  *',  True),

        ('시작기도',    '시작기도 PPT',      False),

        ('화답송_pptx', '화답송 악보 PPT/사진', False),

        ('공지사항',    '공지사항 PPT',      False),

        ('미사후기도',  '미사 후 기도 PPT',  False),

    ]

    if mass_type == 'youth':
        # 청년미사는 화답송 악보 PPT/사진 입력 자체가 없다(요구사항 §A4) — 성가 선택
        # 팝업에서 별도 출처로 다룬다.
        file_rows = [row for row in file_rows if row[0] != '화답송_pptx']
        # §J4 — 청년미사만 [ref_pptx, 시작기도, 미사후기도, 공지사항] 순서로 바꾼다.
        # 성인미사 순서(위 file_rows 원본, 화답송 위치 포함)는 절대 건드리지 않는다 —
        # dict 재배열로 원본 튜플(라벨·필수 여부)을 그대로 보존해 값 자체는 안 바뀐다.
        _rows_by_key = {row[0]: row for row in file_rows}
        file_rows = [_rows_by_key[k] for k in ('ref_pptx', '시작기도', '미사후기도', '공지사항')]

    vars_ = {}

    entries = {}

    for i, (key, label, _req) in enumerate(file_rows):

        r = i + 1

        tk.Label(frame, text=label, anchor='w', width=18).grid(

            row=r, column=0, sticky='w', pady=6

        )

        var = tk.StringVar()

        vars_[key] = var

        entry = tk.Entry(frame, textvariable=var, width=46, state='readonly')

        entry.grid(row=r, column=1, padx=(8, 6), pady=6)

        entries[key] = entry

        row_types = HWADAPSONG_TYPES if key == '화답송_pptx' else PPTX_TYPES

        def _browse(v=var, types=row_types, key=key):

            root.focus_set()

            _kw = {}
            _start = browse_initial_dir(mass_type, key)
            if _start is not None:
                _kw['initialdir'] = str(_start)
            path = filedialog.askopenfilename(parent=root, filetypes=types, **_kw)

            if path:

                v.set(path)
                remember_browse_dir(mass_type, key, path)

        tk.Button(frame, text='찾아보기', command=_browse, pady=0).grid(row=r, column=2, pady=6)



    result = [None]

    def _field_value(key: str) -> str:
        # 청년미사는 화답송_pptx 행 자체가 없어 vars_에 키가 없다 — 빈 문자열로 통일.
        return vars_[key].get() if key in vars_ else ''



    def on_ok():

        date_val = date_var.get().strip()

        if not re.match(r'^\d{8}$', date_val):

            messagebox.showwarning(

                '입력 오류',

                'YYYYMMDD 형식으로 날짜를 입력해 주세요.\n예: 20260628',

                parent=root,

            )

            return

        ref = _field_value('ref_pptx')

        if not ref:

            messagebox.showwarning('입력 오류', '참조 미사 PPT를 선택해 주세요.', parent=root)

            return

        # 청년미사는 화답송_pptx 행 자체가 없어 vars_에 키가 없다(위 file_rows 필터링) —
        # `_resolved_field_value`가 키 부재 시 빈 문자열을 반환하므로 기존 삼항 분기와
        # 동일하게 동작한다.
        화답송 = _field_value('화답송_pptx')

        if not 화답송 and mass_type == 'adult' and is_sunday_mass(date_val):

            if not messagebox.askyesno(

                '확인',

                '주일미사인데 화답송 악보 PPT가 선택되지 않았습니다.\n계속하시겠습니까?',

                parent=root,

            ):

                return

        # 확장자로 PPT/이미지 분기 (설계서 §4.2)

        화답송_pptx_val = None

        화답송_img_val = None

        if 화답송:

            _p = Path(화답송)

            if _p.suffix.lower() in _IMG_EXTS:

                화답송_img_val = _p

            else:

                화답송_pptx_val = _p

        # 공지사항: 파일 선택 필터에 '모든 파일'이 있어 .docx 등을 고를 수 있으므로
        # 확인 시점에 확장자를 검증한다(요구사항 §1.2.1). 빈 값이면 선택 사항이라 통과.
        공지사항 = _field_value('공지사항')

        if 공지사항 and Path(공지사항).suffix.lower() != '.pptx':

            messagebox.showwarning(

                '입력 오류',

                '공지사항 파일은 PowerPoint(.pptx) 파일만 선택할 수 있습니다.\n'

                '다른 파일을 선택해 주세요.',

                parent=root,

            )

            return  # 팝업 유지, 재선택 유도

        시작기도 = _field_value('시작기도')

        미사후기도 = _field_value('미사후기도')

        result[0] = (

            date_val,

            {

                'ref_pptx':    Path(ref),

                '화답송_pptx': 화답송_pptx_val,

                '화답송_img':  화답송_img_val,

                '공지사항':    Path(공지사항) if 공지사항 else None,

                '시작기도':    Path(시작기도) if 시작기도 else None,

                '미사후기도':  Path(미사후기도) if 미사후기도 else None,

                '성가':        {},

            },

        )

        _first_dialog_pos[0] = root.winfo_x()

        _first_dialog_pos[1] = root.winfo_y()

        root.destroy()



    def on_cancel():

        _first_dialog_pos[0] = root.winfo_x()

        _first_dialog_pos[1] = root.winfo_y()

        root.destroy()



    root.protocol('WM_DELETE_WINDOW', on_cancel)

    date_entry.bind('<Return>', lambda e: on_ok())

    tk.Button(frame, text='확인', command=on_ok, width=10).grid(

        row=len(file_rows) + 1, column=0, columnspan=3, pady=14

    )

    # §J10 — 버전 표시 + '업데이트 확인' 링크. tk.Button이 아니라 tk.Label(cursor='hand2'
    # + <Button-1> bind)로 만든다 — 코디네이터 지시로 버튼→링크 전환. 별도 Frame에
    # pack()으로 넣어 grid()를 쓰지 않는다 — grid()를 썼다면 §J4의 파일 행 순서 검증
    # (`test_j4_*`)이 "grid_column==0인 gridded Label"만 세는 필터에 이 라벨들까지
    # 걸려 카운트(성인 5개/청년 4개)가 깨질 뻔했다. bg를 명시적으로 지정하지 않는다 —
    # `_apply_theme()`이 이미 `*Label.background`를 `_UI['bg']`(흰색, 프레임과 동일)로
    # option database에 걸어뒀으므로, 굳이 다른 회색(예: '#f0f0f0')을 직접 주면 흰
    # 프레임 위에서 오히려 도드라져(코디네이터 표현으로 "블랙하게") 보인다.
    _version_row = tk.Frame(frame)
    _version_row.grid(row=len(file_rows) + 2, column=0, columnspan=3, sticky='e', pady=(4, 0))

    tk.Label(_version_row, text=f'v{_get_app_version()}', fg='#999999').pack(
        side='left', padx=(0, 10),
    )

    _update_link = tk.Label(
        _version_row, text='🔄 업데이트 확인', fg=_UI['primary'], cursor='hand2',
    )
    _update_link.pack(side='left')
    _update_link.bind('<Button-1>', lambda e: _on_check_update(root))

    _center_window(root)

    root.deiconify()

    root.after(50, date_entry.focus_set)

    root.mainloop()



    if result[0] is None:

        raise RuntimeError('입력이 취소됐습니다.')

    return result[0]




def _report_progress(pct: int, label: str = ''):

    """main() 내부에서 진행률을 진행 창으로 전달한다."""

    cb = _progress_callback[0]

    if cb:

        cb(pct, label)




# 진행 창 첫머리 안내(§J7 문구 유지 — 실측: PowerPoint 프레임 창은 보이지만 포커스를 가져가지 않는다).
_POWERPOINT_NOTICE_TEXT = (
    '처리 중 슬라이드 줄 수를 정확히 확인하려고 PowerPoint가 백그라운드에서 실행됩니다.\n'
    '작업표시줄에 PowerPoint 아이콘이 잠깐 나타날 수 있지만, 지금 쓰고 있는 창의 포커스를 '
    '가져가지는 않습니다.'
)
_POWERPOINT_NOTICE_SECONDS = 5


def _run_with_progress_window(main_func):

    """main_func 을 워커 스레드로 실행하고 진행률 팝업을 메인 스레드에서 표시한다.

    반환: (output_text, error_text_or_None)

    """

    import tkinter as tk

    from tkinter import ttk

    import threading

    import queue as _queue



    pq = _queue.Queue()

    result = [None, None]



    root = tk.Tk()

    root.withdraw()

    root.title('처리 중...')

    root.resizable(False, False)

    _set_window_icon(root)

    _apply_theme(root)



    frame = tk.Frame(root, padx=20, pady=16)

    frame.pack(fill='both', expand=True)



    # 2026-10-03: 별도 팝업 대신 진행 창 첫머리에 안내를 5초간 보여주고 지운다(그 사이 진행
    # 메시지도 같이 보인다).
    notice = tk.Label(frame, text=_POWERPOINT_NOTICE_TEXT, anchor='w', justify='left',
                      wraplength=380, fg='#666666')
    notice.pack(fill='x', pady=(0, 10))

    def _hide_notice():
        try:
            notice.destroy()
        except Exception:
            pass

    root.after(_POWERPOINT_NOTICE_SECONDS * 1000, _hide_notice)

    label_var = tk.StringVar(value='시작 중...')

    tk.Label(frame, textvariable=label_var, anchor='w').pack(

        fill='x', pady=(0, 6)

    )

    _pb_style = ttk.Style()

    _pb_style.theme_use('default')

    _pb_style.configure(

        'KCC.Horizontal.TProgressbar',

        background=_UI['primary'],

        troughcolor='#e8e8e8',

        bordercolor=_UI['bg'],

        lightcolor=_UI['primary'],

        darkcolor=_UI['primary'],

    )

    pbar = ttk.Progressbar(frame, style='KCC.Horizontal.TProgressbar', length=380, mode='determinate', maximum=100)

    pbar.pack()

    pct_var = tk.StringVar(value='0%')

    tk.Label(frame, textvariable=pct_var).pack(pady=(4, 0))



    log_out = io.StringIO()

    log_err = io.StringIO()

    orig_out, orig_err = sys.stdout, sys.stderr

    sys.stdout, sys.stderr = log_out, log_err



    def worker():

        _progress_callback[0] = lambda p, lbl='': pq.put((p, lbl))

        try:

            main_func()

            full_txt = log_out.getvalue() + log_err.getvalue()

            sys.stdout, sys.stderr = orig_out, orig_err

            pq.put(('__done__', full_txt, None))

        except SystemExit as _e:

            full_txt = log_out.getvalue() + log_err.getvalue()

            err_txt = log_err.getvalue().strip()

            sys.stdout, sys.stderr = orig_out, orig_err

            if _e.code and _e.code != 0:

                # stderr에 찍힌 구체적 에러 메시지(예: "오류: ...")만 팝업에 보여준다.

                # 진행 로그 전체는 full_txt로 별도 보관해 로그 파일에는 남긴다.

                pq.put(('__done__', full_txt, err_txt or f'종료 코드: {_e.code}'))

            else:

                pq.put(('__done__', full_txt, None))

        except Exception as _e2:

            import traceback as _tb

            tb_str = _tb.format_exc()

            full_txt = log_out.getvalue() + log_err.getvalue() + '\n' + tb_str

            sys.stdout, sys.stderr = orig_out, orig_err

            # 팝업에는 예외 메시지만 간결하게, 트레이스백은 로그 파일(full_txt)에만 남긴다.

            pq.put(('__done__', full_txt, f'{type(_e2).__name__}: {_e2}'))



    gc.collect()

    gc.freeze()  # 작업 스레드의 GC가 앞서 닫은 창의 Tk 객체를 수거하지 못하게(_tk_gc_guard 설명 참고)

    t = threading.Thread(target=worker, daemon=True)

    t.start()



    polling = [True]

    _after_id = [None]

    def poll():

        if not polling[0]:

            return

        while True:

            try:

                item = pq.get_nowait()

            except _queue.Empty:

                break

            if isinstance(item, tuple) and item[0] == '__done__':

                _, text, err = item

                polling[0] = False

                if _after_id[0]:

                    try:

                        root.after_cancel(_after_id[0])

                    except Exception:

                        pass

                pbar['value'] = 100

                pct_var.set('100%')

                label_var.set('완료!')

                result[0] = text

                result[1] = err

                def _close():

                    try:

                        root.destroy()

                    except Exception:

                        pass

                root.after(600, _close)

                return

            else:

                pct, lbl = item

                pbar['value'] = pct

                pct_var.set(f'{pct}%')

                if lbl:

                    label_var.set(lbl)

        if polling[0]:

            try:

                _after_id[0] = root.after(80, poll)

            except Exception:

                pass



    _center_window(root)

    root.deiconify()

    _after_id[0] = root.after(80, poll)

    try:

        root.mainloop()

    finally:

        gc.unfreeze()

        gc.collect()

    _progress_callback[0] = None

    return result[0], result[1]


# ─────────────────────────────────────────────────────────────────────────────
# 결과 창
# ─────────────────────────────────────────────────────────────────────────────


# 결과 창 버튼 4개 색(2026-10-04 확정 — 목업 C안): 짙은 버건디(파일 열기) → 버건디(복사, 앱 대표색) →
# 슬레이트 블루(폴더 열기) → 테두리만 있는 닫기. (배경, 글자, 눌림 배경, 테두리) — 외곽선은 닫기만 눈에 보인다.
_RESULT_BTN_COLORS = {
    'open_file':   ('#59001d', '#ffffff', '#3d0013', None),
    'copy':        ('#871b24', '#ffffff', '#6b141c', None),
    'open_folder': ('#3f5a73', '#ffffff', '#2f4458', None),
    'close':       ('#ffffff', '#59001d', '#f3e6ea', '#59001d'),
}


def _result_button(parent, name: str, text: str, command, width: int):
    """결과 창 버튼 한 개. 1px 테두리 색 프레임으로 감싸 반환한다(호출부가 프레임을 pack) — Windows의 tk.Button은
    highlightbackground 외곽선을 그리지 않아서, 테두리 있는 '닫기'를 이렇게 만든다. 테두리 없는 버튼은 프레임을
    배경색과 같게 해 4개의 높이가 같다."""
    import tkinter as tk
    bg, fg, active, edge = _RESULT_BTN_COLORS[name]
    frame = tk.Frame(parent, bg=edge or bg, padx=1, pady=1)
    tk.Button(
        frame, text=text, command=command, bg=bg, fg=fg, activebackground=active, activeforeground=fg,
        font=(_UI['font'], _UI['font_sz'], 'bold'), width=width, relief='flat', bd=0,
        highlightthickness=0, cursor='hand2',
    ).pack()
    return frame


def _show_result_window(title: str, text: str, is_error: bool = False, log_text: str = None) -> None:

    import tkinter as tk

    from tkinter import scrolledtext as _st_mod

    from tkinter import messagebox

    import subprocess as _sp

    from datetime import datetime as _dt



    # 로그 저장 (성공/실패 여부와 무관하게 항상 저장 — 팝업에는 간결한 text만,

    # 파일에는 log_text로 넘어온 전체 로그를 남긴다. log_text가 없으면 text를 그대로 저장)

    date_str = _last_date_str[0]

    output_path = _last_output_path[0]

    if date_str:

        log_dir = _log_dir_for(date_str, _last_output_key[0])

        try:

            log_dir.mkdir(parents=True, exist_ok=True)

            ts = _dt.now().strftime('%H%M%S')

            log_file = log_dir / f'{date_str}_{ts}.txt'

            saved_text = log_text if log_text is not None else text

            log_file.write_text(saved_text.strip(), encoding='utf-8')

        except Exception:

            pass



    if is_error:

        # 전체 로그는 위에서 이미 파일로 저장했으므로, 화면에는 Windows 표준 오류 팝업만 띄운다.

        root = tk.Tk()

        root.withdraw()

        _set_window_icon(root)

        messagebox.showerror(title, text.strip() if text.strip() else '알 수 없는 오류가 발생했습니다.')

        root.destroy()

        return



    root = tk.Tk()

    root.withdraw()

    root.title(title)

    root.minsize(750, 480)

    root.resizable(True, True)

    _set_window_icon(root)

    _apply_theme(root)



    frame = tk.Frame(root)

    frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(12, 4))



    text_widget = _st_mod.ScrolledText(

        frame, wrap=tk.WORD, font=('Consolas', 11), bg='#f9f9f9', fg=_UI['fg']

    )

    text_widget.pack(fill=tk.BOTH, expand=True)

    text_widget.insert(tk.END, text.strip() if text.strip() else '(출력 없음)')

    text_widget.configure(state='disabled')



    btn_frame = tk.Frame(root)

    btn_frame.pack(pady=8)



    if output_path:

        def open_file():

            try:

                _sp.Popen(['start', '', str(output_path)], shell=True)

            except Exception:

                pass

        def open_folder():

            try:

                _sp.Popen(['explorer', '/select,', str(output_path)])

            except Exception:

                pass

        _result_button(btn_frame, 'open_file', '파일 열기', open_file, 12).pack(side=tk.LEFT, padx=6)

        # 2026-10-04: '파일 열기' 바로 옆, 같은 글꼴 크기(성인·청년 공통).
        _mass = _last_output_mass[0]
        if date_str and _mass in ('adult', 'youth'):
            _result_button(
                btn_frame, 'copy', '원드라이브로 복사',
                lambda: _copy_result_with_ui(root, _mass, output_path, date_str), 16,
            ).pack(side=tk.LEFT, padx=6)

        _result_button(btn_frame, 'open_folder', '폴더 열기', open_folder, 12).pack(side=tk.LEFT, padx=6)



    _result_button(btn_frame, 'close', '닫기', root.destroy, 14).pack(side=tk.LEFT, padx=6)

    _center_window(root)

    root.deiconify()

    root.mainloop()
