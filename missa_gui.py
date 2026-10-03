"""
missa_to_ppt.py 모듈 분리 리팩토링 Phase 2 — Tkinter 팝업 + 날짜판단/설정.

is_sunday_mass()·config.json 로드/저장·get_onedrive_hymn_folder() 등은 논리적으로는
GUI가 아니지만, GUI 팝업(_ask_combined_input_popup 등)과 진입점(main, parse_args,
find_files) 양쪽에서 모두 호출된다. 진입점에 그대로 두면 entry↔gui 순환 임포트가
생기므로, 이 모듈에 흡수해 entry→gui 단방향 의존만 유지한다.
"""
from __future__ import annotations

import hashlib
import io
import json
import re
import shutil
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
CONFIG_FILE = _SCRIPT_DIR / 'config.json'



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


# 청년미사 2단계 §2.4 — OneDrive 공용 계정(brokenbaykccppt@gmail.com)의 상대경로 기본값.
# 로컬 동기화 폴더(onedrive_hymn_folder 등)가 설정되지 않은 PC(배포 PC 등)에서
# missa_onedrive(Graph API)로 미러링할 때 쓴다. config.json의 onedrive_hymn_path/
# onedrive_youth_hymn_path 키로 재정의할 수 있다.
_DEFAULT_ONEDRIVE_HYMN_PATH = 'PPT 문서/09.가톨릭 성가/성가-악보버전'
_DEFAULT_ONEDRIVE_YOUTH_HYMN_PATH = 'PPT 문서/20.청년 미사/2.성가'


# ─────────────────────────────────────────────────────────────────────────────
# 청년미사 OneDrive 접속 모드(2026-10-02, 요구사항 §3.3.1): 로컬 'PPT 문서' 폴더 우선,
# 설정이 없거나 무효하면 기존 기기 코드 로그인(Graph API)으로 폴백한다.
# 모드는 실행 시작 시 한 번 결정해 `_YOUTH_PPT_STATE`에 보관하고, 이후 모든 단계(파일 선택·
# 성가 조회·결과 저장)가 같은 값을 읽는다 — 단계마다 따로 판정하면 한쪽은 로컬, 한쪽은
# Graph를 타서 로그인 창이 뒤늦게 뜨는 불일치가 생긴다.
# ─────────────────────────────────────────────────────────────────────────────

_YOUTH_PPT_FOLDER_KEY = 'onedrive_ppt_folder'
_YOUTH_ROOT_FOLDER = '20.청년 미사'
_YOUTH_CATHOLIC_HYMN_SUBPATH = ('09.가톨릭 성가', '성가-악보버전')
_YOUTH_HYMN_SUBPATH = (_YOUTH_ROOT_FOLDER, '2.성가')

_YOUTH_PPT_STATE = {'mode': None, 'folder': None}


def _reset_youth_ppt_mode() -> None:
    """모드 결정 캐시를 비운다(테스트·설정 변경 직후용)."""
    _YOUTH_PPT_STATE['mode'] = None
    _YOUTH_PPT_STATE['folder'] = None


def _is_valid_youth_ppt_folder(path) -> bool:
    """디렉터리이며 그 안에 `20.청년 미사` 폴더가 있으면 유효한 'PPT 문서' 폴더."""
    try:
        p = Path(path)
        return p.is_dir() and (p / _YOUTH_ROOT_FOLDER).is_dir()
    except Exception:
        return False


def _select_youth_ppt_folder_interactively():
    """'PPT 문서' 폴더 선택창을 띄운다. 유효하지 않은 폴더를 고르면 안내 후 다시 선택하게
    하고, 취소하면 None을 반환한다(호출부가 로그인 폴백으로 처리)."""
    message = ("성당 'PPT 문서' 폴더(이 컴퓨터에서 OneDrive로 동기화 중인 폴더)를 선택해 주세요.\n"
               "(최초 1회만 설정됩니다. 선택하지 않고 취소하면 이번에는 OneDrive 로그인 방식으로 진행합니다.)")
    while True:
        folder = _ask_onedrive_path_popup(
            title="'PPT 문서' 폴더 설정", message=message,
            dialog_title="'PPT 문서' 폴더 선택",
        )
        if not folder:
            return None
        if _is_valid_youth_ppt_folder(folder):
            return str(Path(folder))
        message = (f"선택한 폴더 안에 '{_YOUTH_ROOT_FOLDER}' 폴더가 없습니다.\n"
                   "'PPT 문서' 폴더를 다시 선택해 주세요.\n"
                   "(취소하면 OneDrive 로그인 방식으로 진행합니다.)")


def get_youth_ppt_mode(ask_if_missing: bool = False) -> str:
    """청년미사 OneDrive 접속 모드를 반환한다: 'local' | 'fallback'.

    `config.json`의 `onedrive_ppt_folder`가 유효하면 'local'. 아니면 `ask_if_missing`일 때만
    폴더 선택창을 띄우고(선택 성공 시 저장), 취소·무효면 'fallback'. 선택창은 GUI 진입점
    (`_run_gui_mode`)만 `ask_if_missing=True`로 부른다 — 그 외 호출(CLI·하위 함수·테스트)이
    모달 창 앞에서 멈추지 않게 하기 위함. 'local'과 선택창을 거친 결과만 캐시한다."""
    st = _YOUTH_PPT_STATE
    if st['mode'] is not None:
        return st['mode']

    config = _load_config()
    saved = config.get(_YOUTH_PPT_FOLDER_KEY, '')
    if saved and _is_valid_youth_ppt_folder(saved):
        st['mode'], st['folder'] = 'local', Path(saved)
        return 'local'

    if ask_if_missing:
        folder = _select_youth_ppt_folder_interactively()
        if folder:
            config[_YOUTH_PPT_FOLDER_KEY] = folder
            _save_config(config)
            print(f"  [설정] 'PPT 문서' 폴더 저장됨: {folder}")
            st['mode'], st['folder'] = 'local', Path(folder)
            return 'local'
        st['mode'], st['folder'] = 'fallback', None
        return 'fallback'
    return 'fallback'


def youth_local_subdir(kind: str, subfolder: str | None = None, year=None) -> Path:
    """로컬 모드의 하위 경로 조합(하드코딩 문자열은 여기에만 둔다). kind:
    'catholic_hymn' | 'youth_hymn'(subfolder=나주노 성가/야훼이레 성가) | 'output'(year) |
    'start_dir'('PPT 문서' 폴더 자체)."""
    base = _YOUTH_PPT_STATE['folder']
    if base is None:
        raise RuntimeError("청년미사 로컬 모드가 아닙니다('PPT 문서' 폴더 미설정).")
    if kind == 'start_dir':
        return base
    if kind == 'catholic_hymn':
        return base.joinpath(*_YOUTH_CATHOLIC_HYMN_SUBPATH)
    if kind == 'youth_hymn':
        return base.joinpath(*_YOUTH_HYMN_SUBPATH, subfolder or '')
    if kind == 'output':
        return base / _YOUTH_ROOT_FOLDER / str(year)
    raise ValueError(f'알 수 없는 kind: {kind!r}')


# ─────────────────────────────────────────────────────────────────────────────
# 'PPT 문서' 루트·찾아보기 시작 폴더·결과 PPT 복사(2026-10-03, 요구사항 §3.3.2/§3.3.3)
#
# 성인미사도 청년미사와 같은 설정 키(`onedrive_ppt_folder`)로 로컬 'PPT 문서' 폴더를 한 번 고른다
# (성인은 그 안에 `09.가톨릭 성가`가 있어야 유효). 이 루트로부터 (1) 첫 화면 '찾아보기'의 시작
# 폴더, (2) 결과창 '원드라이브로 복사'의 저장 폴더를 계산한다. 하드코딩된 하위 폴더명은 이 절에만 둔다.
# ─────────────────────────────────────────────────────────────────────────────

_ADULT_HYMN_ROOT_FOLDER = '09.가톨릭 성가'
_ADULT_PPT_STATE = {'folder': None}
_LAST_PSALM_DIR_KEY = 'last_psalm_dir'

# 찾아보기 시작 폴더('PPT 문서' 기준 하위 경로). 공통 → 미사 유형별 순으로 조회한다.
_BROWSE_START_SUBPATHS = {
    'common': {
        '시작기도':   ('13.기도문',),
        '미사후기도': ('13.기도문',),
        '공지사항':   ('11.공지사항 PPT문서',),
    },
    'adult': {'ref_pptx': ('10.기타 PPT문서', 'PPT 자동화', '1.Template')},
    'youth': {'ref_pptx': (_YOUTH_ROOT_FOLDER, '1.Template')},
}


def _reset_adult_ppt_folder() -> None:
    """성인미사 'PPT 문서' 폴더 캐시를 비운다(테스트·설정 변경 직후용)."""
    _ADULT_PPT_STATE['folder'] = None


def _is_valid_adult_ppt_folder(path) -> bool:
    """디렉터리이며 그 안에 `09.가톨릭 성가` 폴더가 있으면 유효한 'PPT 문서' 폴더."""
    try:
        p = Path(path)
        return p.is_dir() and (p / _ADULT_HYMN_ROOT_FOLDER).is_dir()
    except Exception:
        return False


def _select_adult_ppt_folder_interactively():
    """성인미사용 'PPT 문서' 폴더 선택창. 무효한 폴더면 다시 묻고, 취소하면 None."""
    message = ("성당 'PPT 문서' 폴더(이 컴퓨터에서 OneDrive로 동기화 중인 폴더)를 선택해 주세요.\n"
               "(최초 1회만 설정됩니다. 선택하지 않고 취소하면 성가·결과 복사 기능은 기존 방식으로 동작합니다.)")
    while True:
        folder = _ask_onedrive_path_popup(
            title="'PPT 문서' 폴더 설정", message=message,
            dialog_title="'PPT 문서' 폴더 선택",
        )
        if not folder:
            return None
        if _is_valid_adult_ppt_folder(folder):
            return str(Path(folder))
        message = (f"선택한 폴더 안에 '{_ADULT_HYMN_ROOT_FOLDER}' 폴더가 없습니다.\n"
                   "'PPT 문서' 폴더를 다시 선택해 주세요.\n(취소하면 기존 방식으로 진행합니다.)")


def get_adult_ppt_folder(ask_if_missing: bool = False):
    """성인미사 'PPT 문서' 폴더(Path) 또는 None. 우선순위: 캐시 → config `onedrive_ppt_folder` →
    기존 `onedrive_hymn_folder`(…/PPT 문서/09.가톨릭 성가/성가-악보버전)에서 역산 → (ask_if_missing일
    때만) 선택창. 선택창으로 고르면 `onedrive_ppt_folder`와 함께 `onedrive_hymn_folder`를
    `{PPT 문서}/09.가톨릭 성가/성가-악보버전`으로 저장한다. 선택창은 GUI 진입점만 `ask_if_missing=True`로
    부른다(다른 호출이 모달 창 앞에서 멈추지 않게)."""
    st = _ADULT_PPT_STATE
    if st['folder'] is not None:
        return st['folder']

    config = _load_config()
    saved = config.get(_YOUTH_PPT_FOLDER_KEY, '')
    if saved and _is_valid_adult_ppt_folder(saved):
        st['folder'] = Path(saved)
        return st['folder']

    hymn = config.get('onedrive_hymn_folder', '')
    if hymn:
        derived = Path(hymn).parent.parent
        if _is_valid_adult_ppt_folder(derived):
            st['folder'] = derived
            return derived

    if ask_if_missing:
        folder = _select_adult_ppt_folder_interactively()
        if folder:
            config[_YOUTH_PPT_FOLDER_KEY] = folder
            config['onedrive_hymn_folder'] = str(
                Path(folder).joinpath(_ADULT_HYMN_ROOT_FOLDER, '성가-악보버전'))
            _save_config(config)
            print(f"  [설정] 'PPT 문서' 폴더 저장됨: {folder}")
            st['folder'] = Path(folder)
            return st['folder']
    return None


def get_ppt_root(mass_type: str):
    """이번 실행의 로컬 'PPT 문서' 폴더(Path) 또는 None(청년 로그인 방식·성인 미설정)."""
    if mass_type == 'youth':
        return _YOUTH_PPT_STATE['folder'] if _YOUTH_PPT_STATE['mode'] == 'local' else None
    return _ADULT_PPT_STATE['folder']


def browse_start_subpath(mass_type: str, key: str):
    """첫 화면 파일 행(key)의 찾아보기 시작 하위 경로(튜플) 또는 None(= 'PPT 문서' 자체)."""
    return (_BROWSE_START_SUBPATHS.get(mass_type, {}).get(key)
            or _BROWSE_START_SUBPATHS['common'].get(key))


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


def psalm_browse_start_dir(root):
    """화답송 찾아보기 시작 폴더: 직전에 고른 화답송 폴더가 있으면 그곳, 없으면 'PPT 문서'."""
    last = _load_config().get(_LAST_PSALM_DIR_KEY, '')
    if last and Path(last).is_dir():
        return Path(last)
    return Path(root) if root is not None else None


def remember_psalm_dir(path) -> None:
    """화답송으로 고른 파일의 폴더를 기억한다(다음 실행의 시작 폴더)."""
    try:
        config = _load_config()
        config[_LAST_PSALM_DIR_KEY] = str(Path(path).parent)
        _save_config(config)
    except Exception:
        pass


def browse_initial_dir(mass_type: str, key: str):
    """로컬 파일 선택창의 initialdir(Path) 또는 None(루트 미설정 — 기본 위치 사용)."""
    root = get_ppt_root(mass_type)
    if mass_type == 'adult' and key == '화답송_pptx':
        return psalm_browse_start_dir(root)
    return resolve_browse_start_dir(root, browse_start_subpath(mass_type, key))


def _result_remote_folder(year) -> str:
    return f'{_ONEDRIVE_BROWSER_ROOT}/{_YOUTH_ROOT_FOLDER}/{year}'


def result_copy_dir(mass_type: str, year):
    """결과 PPT 복사 대상의 로컬 폴더: 청년 `{PPT 문서}/20.청년 미사/{YYYY}`, 성인 `{PPT 문서}/{YYYY}`.
    루트를 모르면 None."""
    root = get_ppt_root(mass_type)
    if root is None:
        return None
    if mass_type == 'youth':
        return Path(root) / _YOUTH_ROOT_FOLDER / str(year)
    return Path(root) / str(year)


def result_copy_exists(mass_type: str, output_path, date_str: str) -> bool:
    """복사 대상에 같은 이름의 파일이 이미 있는지(덮어쓰기 확인용). 확인 불가면 False."""
    name = Path(output_path).name
    dest_dir = result_copy_dir(mass_type, date_str[:4])
    if dest_dir is not None:
        return (dest_dir / name).exists()
    if mass_type == 'youth':
        try:
            import missa_onedrive as od
            od.get_item_metadata(f'{_result_remote_folder(date_str[:4])}/{name}')
            return True
        except Exception:
            return False
    return False


def copy_result_to_onedrive(mass_type: str, output_path, date_str: str) -> str:
    """결과 PPT를 연도 폴더로 복사하고 저장된 위치(표시용 문자열)를 반환한다. 연도 폴더가 없으면 만든다.
    로컬 'PPT 문서'를 알면 로컬 복사(OneDrive 앱이 동기화), 청년 로그인 방식이면 Graph로 같은 위치에
    업로드한다. 성인이 루트를 모르면 RuntimeError."""
    import shutil
    output_path = Path(output_path)
    year = date_str[:4]
    dest_dir = result_copy_dir(mass_type, year)
    if dest_dir is not None:
        dest_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(output_path, dest_dir / output_path.name)
        return str(dest_dir)
    if mass_type == 'youth':
        import missa_onedrive as od
        folder = _result_remote_folder(year)
        od.ensure_folder(folder)
        od.upload_file(f'{folder}/{output_path.name}', output_path)
        return 'OneDrive\\' + folder.replace('/', '\\')
    raise RuntimeError("'PPT 문서' 폴더가 설정되지 않았습니다.")


def _copy_result_with_ui(parent, mass_type: str, output_path, date_str: str) -> None:
    """결과창 '원드라이브로 복사' 버튼 핸들러: 성인이 루트 미설정이면 선택창 → 덮어쓰기 확인 → 복사 →
    저장 폴더 안내(실패하면 오류 안내)."""
    from tkinter import messagebox
    if mass_type == 'adult' and get_ppt_root('adult') is None:
        if get_adult_ppt_folder(ask_if_missing=True) is None:
            messagebox.showwarning('복사 취소', "'PPT 문서' 폴더가 선택되지 않아 복사하지 않았습니다.",
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


def _mirror_onedrive_folder(remote_path: str, local_dir: Path) -> None:
    """remote_path(OneDrive 상대경로) 아래의 모든 파일을 local_dir로 재귀적으로 동기화한다.

    증분 캐시다 — 이미 local_dir에 같은 이름의 파일이 있으면 다시 받지 않는다(성가 PPT는
    한 번 만들어지면 바뀌지 않으므로 매번 재확인할 필요가 없다). 반면 폴더 목록 조회
    (`list_children`)는 매번 새로 해서, 이후 다른 PC가 OneDrive에 새로 추가한 성가도
    다음 실행에서 자연히 캐시에 반영되게 한다. `missa_onedrive`는 지연 임포트한다 — msal
    미설치 환경(로컬 동기화만 쓰는 PC)에서 이 경로를 아예 안 타면 import 자체가 필요
    없어야 하기 때문이다."""
    import missa_onedrive as od

    local_dir.mkdir(parents=True, exist_ok=True)
    for item in od.list_children(remote_path):
        name = item.get('name', '')
        if not name or name.startswith('~$'):
            continue
        if 'folder' in item:
            _mirror_onedrive_folder(f'{remote_path}/{name}', local_dir / name)
            continue
        dest = local_dir / name
        if not dest.exists():
            od.download_file(f'{remote_path}/{name}', dest)


def _resolve_onedrive_folder(local_key: str, remote_path_key: str, default_remote_path: str,
                              cache_subdir: str, popup_kwargs: dict) -> Path:
    """OneDrive 폴더 경로 해석 공통 로직 — 우선순위: (1) 로컬 동기화 폴더(config의
    local_key, 기존 방식·zero-risk 빠른 경로) → (2) Graph API 미러링(로컬 동기화가
    없는 배포 PC) → (3) 수동 폴더 선택 팝업(위 둘 다 실패했을 때의 최후 수단).

    (2)가 (3)보다 먼저인 이유: 배포 PC에는 애초에 선택할 로컬 OneDrive 동기화 폴더 자체가
    없는 경우가 정상 상태라, 먼저 팝업을 띄우면 사용자가 고를 게 없어 막힌다."""
    config = _load_config()

    path_str = config.get(local_key, '')
    if path_str and Path(path_str).is_dir():
        return Path(path_str)

    remote_path = config.get(remote_path_key, default_remote_path)
    local_cache = _SCRIPT_DIR / 'cache' / cache_subdir
    try:
        print(f'  [설정] 로컬 동기화 폴더가 없어 OneDrive(공용 계정)에서 "{remote_path}"를 '
              f'동기화합니다(최초 1회는 파일 수에 따라 시간이 걸릴 수 있습니다)...')
        _mirror_onedrive_folder(remote_path, local_cache)
        return local_cache
    except Exception as e:
        print(f'  [경고] OneDrive Graph API 동기화 실패({e}) — 폴더를 직접 선택해 주세요.')

    folder = _ask_onedrive_path_popup(**popup_kwargs)
    if not folder or not Path(folder).is_dir():
        raise RuntimeError('OneDrive 폴더 선택이 취소되었습니다.')
    config[local_key] = folder
    _save_config(config)
    print(f'  [설정] OneDrive 폴더 저장됨: {folder}')
    return Path(folder)


def get_onedrive_hymn_folder() -> Path:

    """OneDrive 성가(가톨릭성가) 폴더 경로 반환 — 우선순위는 `_resolve_onedrive_folder()`
    참고(로컬 동기화 → Graph API 미러링 → 수동 선택)."""

    return _resolve_onedrive_folder(
        local_key='onedrive_hymn_folder',
        remote_path_key='onedrive_hymn_path',
        default_remote_path=_DEFAULT_ONEDRIVE_HYMN_PATH,
        cache_subdir='가톨릭성가',
        popup_kwargs={},
    )


def get_onedrive_youth_hymn_root() -> Path:

    """OneDrive 청년미사 성가 폴더(2.성가 — 나주노 성가/야훼이레 성가 하위 폴더의 부모,
    PDF 원본도 이 하위 폴더에 있다) 경로 반환 — 우선순위는 `_resolve_onedrive_folder()`
    참고(로컬 동기화 → Graph API 미러링 → 수동 선택)."""

    return _resolve_onedrive_folder(
        local_key='onedrive_youth_hymn_folder',
        remote_path_key='onedrive_youth_hymn_path',
        default_remote_path=_DEFAULT_ONEDRIVE_YOUTH_HYMN_PATH,
        cache_subdir='청년미사_성가',
        popup_kwargs=dict(
            title='청년미사 성가 라이브러리 경로 설정',
            message='청년미사 성가(나주노/야훼이레) 폴더의 상위 폴더(2.성가)를 선택해 주세요.\n(최초 1회만 설정됩니다.)',
            dialog_title='청년미사 2.성가 폴더 선택',
        ),
    )


# ─────────────────────────────────────────────────────────────────────────────
# 청년미사 성가(가톨릭성가/나주노/야훼이레) 곡 단위 온디맨드 조회(2026-09-27, 신규).
#
# `get_onedrive_hymn_folder()`/`get_onedrive_youth_hymn_root()`는 폴더 전체를 미러링한다
# (사용자 실사용 관찰 — 성가 5곡만 필요한데 가톨릭성가 폴더 34곡이 전부 캐시로 내려옴).
# 청년미사는 이 절 아래 함수들로 "요청받은 곡 하나만" 찾아 받는다. 성인미사는
# get_onedrive_hymn_folder()를 그대로 계속 쓰므로(missa_to_ppt.py의 find_files()/main() —
# `mass_type == 'adult'`로 명시적으로 좁혀진 호출부만 그 함수를 부른다) 이 절의 변경은
# 성인미사 동작에 전혀 영향을 주지 않는다.
# ─────────────────────────────────────────────────────────────────────────────

def _youth_onedrive_local_sync_dir(local_key: str) -> Path | None:
    """config[local_key]가 유효한 로컬 폴더면 그 Path를, 아니면 None을 반환한다. "로컬
    동기화 폴더가 있는가"를 찾기·업로드-여부 판단 양쪽에서 재사용하는 단일 지점으로 둬서,
    두 판단이 서로 다른 경로로 갈라지지 않게 한다(§K그룹 리뷰에서 지적된 "같은 질문을
    두 곳에서 따로 판단하면 어긋날 수 있다"는 원칙을 여기서도 지킨다)."""
    config = _load_config()
    path_str = config.get(local_key, '')
    if path_str and Path(path_str).is_dir():
        return Path(path_str)
    return None


def _find_onedrive_file(remote_path: str, pattern: str) -> str | None:
    """remote_path 아래를 재귀적으로 `list_children()`만으로 탐색해(다운로드 없음 —
    메타데이터 조회만이라 저렴, §K1과 같은 원칙) 파일명이 `pattern`(정규식)에 매칭되는
    첫 항목의 전체 원격 경로를 반환한다. 없으면 None. 개별 폴더 조회 실패는 그 폴더
    아래만 건너뛴다(§J5/§K1과 동일한 "부분 실패해도 나머지는 계속 탐색" 원칙)."""
    import missa_onedrive as od
    for item in od.list_children(remote_path):
        name = item.get('name', '')
        if not name or name.startswith('~$'):
            continue
        item_remote = f'{remote_path}/{name}'
        if 'folder' in item:
            try:
                found = _find_onedrive_file(item_remote, pattern)
            except Exception:
                continue
            if found:
                return found
            continue
        if re.search(pattern, name):
            return item_remote
    return None


def find_youth_onedrive_hymn_file(
    remote_path_key: str, default_remote_path: str,
    subfolder: str | None, pattern: str, cache_subdir: str,
    local_kind: str | None = None,
) -> Path | None:
    """청년미사 성가 파일 하나를 찾아 로컬 경로로 반환한다 — 폴더 전체 미러링은 하지
    않는다. 우선순위: (1) 이미 받아둔 적 있는 로컬 캐시(`cache/{cache_subdir}/`)에서
    검색(성가 PPT는 한 번 만들어지면 바뀌지 않으므로 매번 재확인하지 않는다 — §K그룹과
    같은 전제) → (2) OneDrive를 `_find_onedrive_file()`로 재귀 검색해 매칭되는 파일
    하나만 다운로드. 어디서도 못 찾으면 None(폴더 자체가 없거나 조회에 실패해도 예외를
    삼키고 None — 호출부가 "없음"과 "조회 실패"를 구분할 필요가 없는 용도라, 실패 시
    경고만 출력하고 다음 단계(재생성 또는 오류)로 넘어가게 한다).

    로컬 모드(§3.3.1, `get_youth_ppt_mode() == 'local'`)이고 `local_kind`가 주어지면 위
    두 단계 대신 'PPT 문서' 하위의 해당 로컬 폴더(`youth_local_subdir(local_kind, subfolder)`)
    를 재귀 검색하는 것 하나뿐이다(캐시·Graph 모두 쓰지 않음)."""
    if local_kind and get_youth_ppt_mode() == 'local':
        local_dir = youth_local_subdir(local_kind, subfolder)
        if not local_dir.is_dir():
            return None
        for f in sorted(local_dir.rglob('*.pptx')):
            if not f.name.startswith('~$') and re.search(pattern, f.name):
                return f
        return None

    local_cache = _SCRIPT_DIR / 'cache' / cache_subdir
    if local_cache.is_dir():
        for f in sorted(local_cache.glob('*.pptx')):
            if re.search(pattern, f.name):
                return f

    config = _load_config()
    remote_root = config.get(remote_path_key, default_remote_path)
    remote_search = f'{remote_root}/{subfolder}' if subfolder else remote_root
    try:
        found_remote = _find_onedrive_file(remote_search, pattern)
    except Exception as e:
        print(f'  [경고] OneDrive 조회 실패({e}): {remote_search}')
        return None
    if not found_remote:
        return None

    import missa_onedrive as od
    local_cache.mkdir(parents=True, exist_ok=True)
    dest = local_cache / found_remote.rsplit('/', 1)[-1]
    od.download_file(found_remote, dest)
    return dest


def get_onedrive_youth_hymn_remote_path(root: Path) -> str:
    """`root`(호출부가 이번에 실제로 받은 `get_onedrive_youth_hymn_root()`의 반환값)가
    Graph API 미러링 캐시 폴더(`_resolve_onedrive_folder()`가 만드는
    `cache/청년미사_성가`)와 정확히 같을 때만 그 원격 상대경로(예: `PPT 문서/20.청년 미사/
    2.성가`)를 반환하고, 그 외에는 전부 None을 반환한다(로컬 동기화 폴더든, 테스트가
    `get_onedrive_youth_hymn_root()`만 스텁해서 넘겨준 임의 경로든 — 업로드 불필요).

    **`config.json`을 다시 읽어 독립적으로 판단하지 않는다** — 2026-09-18 실측으로 발견한
    버그: 이전 버전은 이 함수가 `config.json`을 스스로 다시 읽어 "로컬 동기화 중인가"를
    재판정했는데, `get_onedrive_youth_hymn_root()`가 몽키패치로 대체된 테스트에서는 두
    함수의 판단이 어긋나(root는 테스트 tmp_path인데 이 함수는 실제 config.json을 읽어
    "동기화 폴더 없음"으로 오판) 실제 `missa_onedrive.upload_file()`이 호출되고, 이 PC에
    캐시된 적 없는 MSAL 기기 코드 로그인이 대화형 입력을 기다리며 테스트가 영구히 멈췄다
    (`test_y2_H1`/`H2`/`C1`에서 재현). `root` 값 자체로 판단하면 두 함수가 절대 어긋날 수
    없다."""
    expected_cache_dir = _SCRIPT_DIR / 'cache' / '청년미사_성가'
    if root != expected_cache_dir:
        return None
    config = _load_config()
    return config.get('onedrive_youth_hymn_path', _DEFAULT_ONEDRIVE_YOUTH_HYMN_PATH)


# ─────────────────────────────────────────────────────────────────────────────
# 청년미사 OneDrive 커스텀 파일 브라우저 (요구사항 §A5)
#
# 성당 공용 계정(brokenbaykccppt)은 이 PC에 로컬 동기화가 안 돼 있어, 청년미사 입력
# 파일(참조 PPT/시작기도/공지사항/미사후기도)은 로컬 파일 다이얼로그로 고를 수 없다.
# Graph API로 원격 폴더를 Treeview로 보여주고, 선택한 파일을 이번 미사의 출력 폴더로
# 내려받는다. 트리 렌더링/이벤트 바인딩(Tkinter)과 목록 필터링·캐싱·다운로드(순수 함수)를
# 분리했다 — 후자는 GUI 없이 단위 테스트할 수 있다.
# ─────────────────────────────────────────────────────────────────────────────

_ONEDRIVE_BROWSER_ROOT = 'PPT 문서'


def _onedrive_visible_children(items: list) -> list:
    """OneDrive `list_children()` 원본 결과에서 브라우저에 보여줄 항목만 골라, 폴더를
    먼저·이름순으로 정렬해 반환한다. 숨김/캐시 파일(~$)과 .pptx가 아닌 파일은 제외한다 —
    이 브라우저는 참조/시작기도/공지사항/미사후기도 PPT를 고르는 용도라 다른 파일
    형식은 소음일 뿐이다."""
    def _keep(item):
        name = item.get('name', '')
        if not name or name.startswith('~$'):
            return False
        if 'folder' in item:
            return True
        return name.lower().endswith('.pptx')

    kept = [it for it in items if _keep(it)]
    kept.sort(key=lambda it: (0 if 'folder' in it else 1, it.get('name', '').lower()))
    return kept


# 코디네이터가 실제 Graph API로 직접 조회해 확정한 이름(J2) — 사용자가 구두로 말한 이름과
# 공백 유무가 미세하게 다를 수 있어(예: '11.공지사항 PPT문서' vs '11.공지사항PPT문서')
# `_onedrive_root_visible_children()`은 비교 시 공백을 제거해 정규화한다.
_ONEDRIVE_ROOT_ALLOWED_FOLDERS = ('11.공지사항 PPT문서', '13.기도문', '20.청년 미사')


def _onedrive_root_visible_children(items: list) -> list:
    """최상위 트리 레벨(§J2)에서만 위 3개 폴더로 제한한다 — 실제 'PPT 문서' 밑에는 이
    프로젝트와 무관한 폴더가 다수 섞여 있어 사용자가 매번 스크롤해서 찾아야 했다. 하위
    폴더 탐색(그 폴더를 편 뒤의 자식들)에는 이 필터를 적용하지 않는다 — 호출부
    (`_insert_children`)가 최상위 호출(parent_iid=='')에만 이 함수를 추가로 거친다."""
    allowed = {name.replace(' ', '') for name in _ONEDRIVE_ROOT_ALLOWED_FOLDERS}
    return [it for it in items if it.get('name', '').replace(' ', '') in allowed]


def _prefetch_onedrive_children(remote_paths, cache: dict) -> None:
    """§J5(선택 구현) → K1(2026-09-27, 재귀 확장) — 팝업이 열리는 시점에
    `_ONEDRIVE_ROOT_ALLOWED_FOLDERS`의 자식뿐 아니라 그 아래 모든 하위 폴더까지 재귀적으로
    캐시에 채운다. 반드시 백그라운드 스레드에서 호출해야 한다(호출부 참고) — 이 함수
    자체는 순차적으로 네트워크 호출을 하는 동기 함수라 GUI 메인 스레드에서 직접 부르면
    팝업이 그만큼 멈춘다.

    순차(비동시) 재귀를 택한 근거(실측, 2026-09-27): 실제 3개 허용 루트 폴더 트리를 이
    함수와 동일한 방식으로 순차 재귀 조회한 결과 총 API 호출 9회·하위폴더 6개·파일
    31개·최대 깊이 2·총 22.04초(평균 2.449초/호출, 네트워크 왕복이 대부분)였다 — 이 정도
    규모면 동시 요청 없이도 429/타임아웃 위험이 없고, 백그라운드 스레드에서 사용자가
    날짜·미사유형을 입력하는 동안 끝나기에 충분히 빠르다. 트리가 실제로 훨씬 커지면
    병렬화를 재검토해야 한다.

    개별 폴더 조회가 실패해도(네트워크 오류 등) 나머지 폴더/가지는 계속 시도한다 — 이건
    순수 성능 최적화라 실패해도 사용자가 실제로 그 폴더를 펼칠 때 기존 동기 조회 경로로
    정상 폴백되므로 예외를 삼켜도 안전하다. 실패한 경로는 캐시에 채우지 않으므로(원본
    `_cached_onedrive_children`이 예외를 던지면 그 경로는 cache에 아예 안 들어감) 그 아래
    가지로 재귀하지 않는다(존재하지 않는 children으로 잘못 재귀할 데이터가 없다)."""
    for remote_path in remote_paths:
        try:
            children = _cached_onedrive_children(remote_path, cache)
        except Exception:
            continue
        subfolders = [f"{remote_path}/{item['name']}" for item in children if 'folder' in item]
        if subfolders:
            _prefetch_onedrive_children(subfolders, cache)


def _start_onedrive_prefetch_thread(cache: dict, root_remote_path: str = _ONEDRIVE_BROWSER_ROOT):
    """§K1의 재귀 프리페치를 백그라운드 데몬 스레드로 **시작만** 하고 즉시 반환한다(호출부
    블로킹 없음). 이전에는 `_ask_onedrive_file_browser_popup()`이 열릴 때만 이 스레드를
    시작했는데, L1(2026-09-27)에서 `_ask_combined_input_popup()`이 뜨는 시점으로 트리거를
    앞당기기 위해 공유 헬퍼로 추출했다 — `_ask_onedrive_file_browser_popup()`도 여전히 이
    헬퍼를 호출한다(같은 `cache` dict를 공유하면 `_cached_onedrive_children()`의 "이미
    있으면 스킵" 계약 덕분에 두 트리거가 겹쳐도 중복 네트워크 호출이 최소화된다 — L1의
    조기 프리페치가 이미 채워둔 경로는 이 두 번째 트리거가 다시 조회하지 않는다).

    로그인이 안 돼 있거나 네트워크가 없어도 이 함수는 스레드 시작만 하고 반환하므로 호출부
    (Tk 메인 스레드)를 블로킹하지 않는다 — 실제 조회 실패는 `_prefetch_onedrive_children()`
    이 이미 경로 단위로 조용히 삼킨다(§J5). 반환하는 `Thread` 객체는 테스트가 `join()`으로
    완료를 기다릴 때만 쓰고, 실제 호출부는 반환값을 무시해도 안전하다(daemon=True — 프로세스
    종료를 막지 않음)."""
    import threading
    thread = threading.Thread(
        target=_prefetch_onedrive_children,
        args=([f'{root_remote_path}/{name}' for name in _ONEDRIVE_ROOT_ALLOWED_FOLDERS], cache),
        daemon=True,
    )
    thread.start()
    return thread


def _cached_onedrive_children(remote_path: str, cache: dict) -> list:
    """`cache`(호출부가 팝업 하나의 수명 동안 공유하는 dict)에 이미 있으면 재사용하고,
    없으면 `missa_onedrive.list_children()`으로 조회해 채운다 — 같은 폴더를 여러 행에서
    반복 조회하지 않기 위한 세션 내 메모리 캐시(요구사항 §A5)."""
    if remote_path not in cache:
        import missa_onedrive as od
        cache[remote_path] = od.list_children(remote_path)
    return cache[remote_path]


def _download_onedrive_selection(remote_path: str, dest_dir) -> Path:
    """선택한 OneDrive 파일(remote_path)을 dest_dir로 내려받고 로컬 경로를 반환한다.

    K2(2026-09-27) 이후 실제 팝업(`_on_choose`)은 이 함수 대신 캐시를 거치는
    `_resolve_onedrive_download()`를 호출한다 — 이 함수는 무조건 재다운로드하는 기존 동작
    그대로 남겨, 캐시를 원치 않는 다른 호출부나 테스트가 여전히 쓸 수 있게 한다."""
    import missa_onedrive as od
    dest = Path(dest_dir) / remote_path.rsplit('/', 1)[-1]
    return od.download_file(remote_path, dest)


# ─────────────────────────────────────────────────────────────────────────────
# K2(2026-09-27) — OneDrive 브라우저로 받는 파일의 로컬 캐시 재사용.
#
# 날짜 폴더(`output/{date}_youth/`)와 별개인 공유 캐시 폴더에 원본을 1부 보관해, 여러 주에
# 걸쳐 같은 파일(시작기도/미사후기도/성가 PPT 등)을 매번 재다운로드하지 않는다. 공지사항처럼
# 매주 실제로 바뀌는 파일은 이 로직이 정상 동작하는 결과로 자연히 매번 재다운로드될 뿐이라,
# 파일 종류로 분기하는 하드코딩은 두지 않는다 — 아래 함수들은 remote_path 하나만 보고
# 판단한다.
#
# 신선도 판단은 lastModifiedDateTime **정확히 일치**만 재사용 근거로 삼는다("더 최신인가"를
# 직접 비교하지 않는다) — 문자열 완전일치이므로 타임존/포맷 파싱 버그가 들어설 자리가 없고,
# "이름이 같다"·"캐시가 있다"만으로는 절대 재사용하지 않는다는 불변조건(코디네이터 명시)을
# 가장 단순하게 만족시킨다. 조회 자체가 불확실하면(예외) 캐시를 신뢰하지 않고 재다운로드로
# 안전하게 폴백한다.
# ─────────────────────────────────────────────────────────────────────────────

_ONEDRIVE_CACHE_ROOT = 'onedrive_cache'  # output/(날짜 폴더)와 별개 — 날짜를 넘나드는 재사용.


def _onedrive_cache_manifest_path() -> Path:
    return Path(_ONEDRIVE_CACHE_ROOT) / 'manifest.json'


def _load_onedrive_cache_manifest() -> dict:
    path = _onedrive_cache_manifest_path()
    if not path.is_file():
        return {}
    try:
        return json.loads(path.read_text(encoding='utf-8'))
    except (json.JSONDecodeError, OSError):
        # 손상된 매니페스트는 "기록 없음"으로 취급한다 — 판단이 불확실하면 재다운로드로
        # 안전하게 폴백하는 원칙과 같은 정신(사람이 파일을 잘못 건드렸어도 조용히 깨지지
        # 않고 다음 사용부터 정상 재구축된다).
        return {}


def _save_onedrive_cache_manifest(manifest: dict) -> None:
    path = _onedrive_cache_manifest_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')


def _onedrive_cache_blob_path(remote_path: str) -> Path:
    """remote_path 하나당 공유 캐시 폴더 안의 고정 파일 하나로 대응시킨다. 원본 경로
    문자열을 그대로 로컬 파일명에 쓰지 않고 해시로 변환한다 — OneDrive 경로에는 Windows
    파일명에 못 쓰는 문자가 없지만, 다른 폴더의 동명 파일(예: 서로 다른 연도 폴더의
    "시작기도.pptx")이 캐시 폴더 안에서 충돌하지 않도록 경로 전체를 키로 삼는다."""
    digest = hashlib.sha1(remote_path.encode('utf-8')).hexdigest()
    suffix = Path(remote_path).suffix
    return Path(_ONEDRIVE_CACHE_ROOT) / f'{digest}{suffix}'


def _ensure_onedrive_blob_fresh(remote_path: str, cache: dict = None) -> Path:
    """공유 캐시 폴더 안의 blob만 최신 상태로 만들어 그 경로를 반환한다 — **날짜 폴더
    (dest_dir)로의 커밋은 하지 않는다.** 로컬 캐시본이 있고 OneDrive 쪽
    lastModifiedDateTime이 캐시 기록 시점과 정확히 같을 때만 재다운로드를 스킵한다.
    조금이라도 다르면(더 최신이든, 기록/캐시본이 없든, 메타데이터 조회 자체가
    불확실하든) 무조건 새로 받는다 — "이름이 같다"는 절대 재사용 근거가 아니다(코디네이터
    명시 불변조건, `test_k2b_*`가 고정).

    `cache`(K1의 재귀 프리페치 결과)는 이 함수의 최종 판단에 **관여하지 않는다** — 리뷰
    라운드17에서 확정된 버그: 이전 구현은 `cache`에 이미 부모 폴더의 자식 목록이 있으면
    그 스냅샷을 "실시간 메타데이터"로 오인해 실제 `get_item_metadata()` 호출을 건너뛰었다.
    그 스냅샷은 프리페치(또는 사용자가 그 폴더를 한 번 펼친) 시점에 고정되고, 팝업이 열려
    있는 동안 전혀 갱신되지 않는다(`_cached_onedrive_children`의 "이미 있으면 재사용"
    계약 때문) — 그 결과 매니페스트와 K1 캐시가 "서로는" 일치해도 실제 OneDrive는 이미
    바뀐 상태일 수 있는데, 정확히 이 경우에 재다운로드를 건너뛰는 사고가 났다
    (`test_k2i_*`가 이 시나리오를 고정). 그래서 `cache` 파라미터는 이제 호출부 시그니처
    호환만을 위해 받고 실제로 쓰지 않는다 — 신선도 판단은 항상 이 함수 하나당 1회
    `missa_onedrive.get_item_metadata(remote_path)`로 실시간 재확인한다(다운로드 대상
    파일당 1회뿐이라 API 비용 증가는 무시할 수준).

    **L2 리뷰 라운드19 확정 버그로 이 함수가 `_resolve_onedrive_download()`에서
    분리됐다**: "blob을 최신화하는 것"과 "그 blob을 날짜 폴더의 최종 목적지에 커밋하는
    것"이 한 함수 안에 있으면, 같은 슬롯을 재선택했을 때(새 파일의 베이스네임이 이전
    선택과 같은 경우 — 이 앱의 실제 관례) 구세대(느린) 다운로드가 최종 목적지에 먼저
    또는 나중에 써버려 최신 선택을 조용히 덮어쓸 수 있다. 이 함수는 공유 캐시 폴더
    안의 blob만 다루므로, 여러 세대가 동시에 호출해도(각자 다른 remote_path를 다루는
    한) 서로의 최종 결과(날짜 폴더 안의 실제 파일)를 침범하지 않는다 — 커밋은 호출부
    (`_download_worker`)가 generation을 재확인한 뒤에만 별도로 수행한다."""
    blob_path = _onedrive_cache_blob_path(remote_path)
    manifest = _load_onedrive_cache_manifest()
    record = manifest.get(remote_path)

    import missa_onedrive as od
    try:
        remote_meta = od.get_item_metadata(remote_path)
    except Exception:
        remote_meta = None  # 조회 자체가 불확실하면 캐시를 신뢰하지 않는다.

    is_fresh = (
        record is not None and remote_meta is not None and blob_path.is_file()
        and record.get('lastModifiedDateTime') == remote_meta.get('lastModifiedDateTime')
    )
    if is_fresh:
        return blob_path

    # blob_path 상위 폴더 존재는 이 함수(캐시 폴더 구조의 소유자)가 직접 보장한다 —
    # `od.download_file()`의 실제 구현도 내부에서 mkdir하지만, 그건 그 모듈의 구현
    # 디테일이라 여기서 의존하지 않는다(모킹된 다른 구현으로 교체돼도 안전하게 동작).
    blob_path.parent.mkdir(parents=True, exist_ok=True)
    od.download_file(remote_path, blob_path)
    if remote_meta is not None:
        manifest[remote_path] = {
            'lastModifiedDateTime': remote_meta.get('lastModifiedDateTime'),
            'eTag': remote_meta.get('eTag'),
        }
        _save_onedrive_cache_manifest(manifest)
    return blob_path


def _commit_onedrive_blob_to_dest(blob_path: Path, dest: Path) -> Path:
    """캐시 폴더의 blob을 실제 목적지(날짜 폴더 안의 파일)로 복사한다 — 순수 파일 복사만
    하는 얇은 헬퍼(L2 리뷰 라운드19). 이 복사가 "언제" 일어나야 안전한지(예: generation
    재확인 후에만)는 호출부의 책임이다."""
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(blob_path, dest)
    return dest


def _resolve_onedrive_download(remote_path: str, dest_dir, cache: dict = None) -> Path:
    """OneDrive 파일을 캐시 인식하며 dest_dir로 준비해 반환한다 — `_ensure_onedrive_blob_
    fresh()`로 캐시를 최신화한 뒤 즉시 `_commit_onedrive_blob_to_dest()`로 커밋한다.
    세대(generation) 충돌 위험이 없는 단발성 호출(L2를 거치지 않는 경로)에 적합하다 —
    L2의 `_download_worker`처럼 재선택 가능한 슬롯에서는 이 함수를 직접 쓰지 않고
    `_ensure_onedrive_blob_fresh()` + generation 재확인 후 커밋을 따로 수행해야 한다
    (`_download_with_retries()`/`_download_worker` 참고)."""
    filename = remote_path.rsplit('/', 1)[-1]
    dest = Path(dest_dir) / filename
    blob_path = _ensure_onedrive_blob_fresh(remote_path, cache)
    return _commit_onedrive_blob_to_dest(blob_path, dest)


def _download_with_retries(remote_path: str, cache: dict, max_retries: int = 2) -> Path:
    """`_ensure_onedrive_blob_fresh()`(캐시 폴더 안의 blob만 최신화, 날짜 폴더 커밋은
    하지 않음)를 자동으로 최대 `max_retries`회 재시도한다(L2, 사용자 확정 결정 2 —
    "자동으로 1~2회 재시도, 그래도 실패하면 사용자에게 알림"). 성공하면 blob_path를
    반환하고, 마지막 시도까지 전부 실패하면 마지막 예외를 그대로 다시 던진다 — 호출부
    (백그라운드 워커)가 이걸 "재시도까지 다 실패함"으로 판단해 사용자에게 알린다.

    **L2 리뷰 라운드19 이후 dest_dir 파라미터를 받지 않는다** — 날짜 폴더로의 커밋은
    이 함수의 책임이 아니다. 재선택 가능한 슬롯(`_ask_combined_input_popup`의
    `_download_worker`)에서 이 함수가 커밋까지 해버리면, 구세대(느린) 호출이 최신 선택의
    커밋을 나중에 덮어쓸 수 있다(실제 재현: 같은 슬롯을 같은 베이스네임으로 재선택하는
    경우 — 이 앱에서 "시작기도.pptx" 등이 매주 다른 폴더에서 같은 이름으로 반복되는 게
    정상 관례라 실제로 발생 가능). 커밋은 호출부가 generation을 재확인한 뒤에만 별도로
    수행해야 한다.

    재시도 사이 지연(backoff)을 넣지 않는다 — 이 프로젝트의 실패 사례는 대부분 일시적
    네트워크 hiccup이라 즉시 재시도로 충분하고, 지연을 넣으면 실패 확정까지 걸리는 시간이
    늘어나 "확인" 클릭 시 대기가 그만큼 길어진다(사용자 확정 결정 3의 "무한정 대기 금지"
    정신과 상충)."""
    last_err = None
    for _ in range(max_retries + 1):
        try:
            return _ensure_onedrive_blob_fresh(remote_path, cache)
        except Exception as e:
            last_err = e
    raise last_err


def _onedrive_display_path(remote_path: str, root_remote_path: str = _ONEDRIVE_BROWSER_ROOT) -> str:
    """OneDrive에서 고른 파일의 Entry 표시용 문자열(버그4). 로컬 다운로드 경로
    (`output\\{date}_youth\\{파일명}`)는 모든 청년미사 파일에 공통으로 붙는 접두사라
    46자 폭 Entry에서 정작 파일을 구분하는 부분(파일명)을 밀어내 잘리게 만든다 —
    실사용자 스크린샷에서 확인. 대신 OneDrive 원본 위치를 보여준다: 모든 파일에 공통인
    `root_remote_path`('PPT 문서') 접두사는 정보가 없으므로 제거하고, 구분자는 이
    프로젝트의 다른 경로 표시 관례(Windows 백슬래시)와 맞춘다.

    **주의**: 이 반환값은 표시 전용이다 — 실제 파이프라인이 쓰는 로컬 경로는 별도
    dict(`_od_real_paths`, `_apply_onedrive_selection` 참고)에 저장한다. 이 함수의
    반환값을 파일시스템 경로로 사용하면 안 된다(그 경로는 존재하지 않는다)."""
    prefix = root_remote_path + '/'
    rel = remote_path[len(prefix):] if remote_path.startswith(prefix) else remote_path
    return rel.replace('/', '\\')


def _resolved_field_value(key: str, vars_: dict, od_real_paths: dict) -> str:
    """`_ask_combined_input_popup()`의 on_ok()가 각 필드의 최종 값을 읽을 때 반드시 거쳐야
    하는 단일 지점(버그4) — OneDrive에서 고른 파일은 `vars_[key]`(Entry에 보이는 값)가
    표시용 상대경로일 뿐이라, 실제 로컬 경로는 `od_real_paths`에서 가져와야 한다. 로컬
    파일 다이얼로그로 고른 경우엔 `od_real_paths`에 그 키가 없으므로(표시=실제값이 같음)
    기존 그대로 `vars_[key].get()`을 반환한다 — on_ok()의 각 필드를 개별적으로 풀어 쓰면
    이 우선순위를 한 곳에서 깜빡할 위험이 있어 헬퍼로 강제한다."""
    if key in od_real_paths:
        return od_real_paths[key]
    if key not in vars_:
        return ''
    return vars_[key].get()


def _apply_onedrive_selection(key: str, local_path, remote_path: str, vars_: dict,
                               od_real_paths: dict, entries: dict,
                               root_remote_path: str = _ONEDRIVE_BROWSER_ROOT) -> None:
    """`_ask_onedrive_file_browser_popup()`이 반환한 `(local_path, remote_path)`를 받아
    `_browse()` 클로저가 해야 하는 세 부작용을 한 곳에서 수행한다(버그4) — (1) 실제
    로컬 경로를 `od_real_paths[key]`에 저장(on_ok()가 이걸 우선 사용, `_resolved_field_
    value` 참고), (2) 표시용 StringVar(`vars_[key]`)엔 다듬은 OneDrive 상대경로만,
    (3) 그 Entry를 파일명 쪽(문자열 끝)으로 스크롤해 요구사항 "파일명 최소 절반 이상
    보이기"를 만족시킨다. `local_path`가 비어 있으면(취소) 아무 것도 건드리지 않는다 —
    기존 `if path: v.set(path)` 관례와 같은 정신."""
    if not local_path:
        return
    od_real_paths[key] = str(local_path)
    vars_[key].set(_onedrive_display_path(remote_path, root_remote_path))
    entry = entries.get(key)
    if entry is not None:
        entry.xview_moveto(1.0)


def _ask_onedrive_file_browser_popup(parent, dest_dir, root_remote_path: str = _ONEDRIVE_BROWSER_ROOT,
                                      cache: dict = None, start_subpath=None):
    """OneDrive 폴더를 Treeview로 탐색해 .pptx 파일 하나를 선택하면 `(local_path,
    remote_path)` 튜플을 반환한다(버그4 — 이전엔 local_path 문자열 하나만 반환했으나,
    호출부가 표시용 상대경로(`_onedrive_display_path`)를 만들려면 선택된 OneDrive 원본
    remote_path도 함께 필요해졌다). 취소하면 `('', '')`(로컬 `filedialog.askopenfilename()`
    의 취소 시 빈 문자열 관례를 튜플로 확장 — 호출부는 `if local_path:` 분기를 그대로
    재사용할 수 있다). 호출부는 `missa_gui.py` 안에 한 곳뿐이라(`_ask_combined_input_popup`)
    시그니처 변경이 안전하다.

    **L2(2026-09-27) 계약 변경 — 이 함수는 더 이상 실제로 다운로드하지 않는다.**
    `local_path`는 다운로드가 끝나면 저장될 위치를 결정적으로 계산한 값(`dest_dir` +
    OneDrive 파일명)일 뿐, 이 함수가 반환하는 시점에 그 파일이 실제로 로컬에 존재한다는
    보장은 없다 — 실제 다운로드(K2 캐시 확인 포함)는 호출부(`_browse()`)가
    `_download_with_retries()`를 백그라운드 스레드로 돌려 수행한다. 사용자가 파일을
    다음 파일 선택으로 즉시 넘어갈 수 있어야 한다는 요구(느린 다운로드가 브라우징을
    막으면 안 됨) 때문에 이 분리가 필요했다 — 브라우징(이 함수)과 다운로드(호출부)가
    서로 다른 속도 요구를 갖게 됐다.

    이 계약 변경이 안전한 이유: `_apply_onedrive_selection()`은 `local_path`를 그대로
    `od_real_paths[key]`에 저장할 뿐 파일 존재 여부를 확인하지 않고, 실제 파이프라인
    (`main()`)은 이 팝업이 완전히 끝난 뒤(`on_ok()`가 모든 백그라운드 다운로드를 기다린
    뒤)에야 실행되므로, 파이프라인이 이 경로를 읽는 시점에는 이미 다운로드가 끝나 있다.

    `start_subpath`(2026-10-03): `root_remote_path` 기준 하위 폴더 경로(튜플, 예: ('13.기도문',))가
    주어지면 그 폴더까지 펼쳐 보이게 시작한다(공백 무시 비교, 없으면 있는 곳까지만).

    시작 위치: `root_remote_path`('PPT 문서') 자체를 트리 노드로 보여주지 않고,
    `list_children(root_remote_path)` 결과를 트리 최상위 레벨로 바로 렌더링한다(사용자
    명시 지시 — 'PPT 문서' 폴더 한 칸을 더 펼치게 하지 않는다).

    폴더 lazy-load 트리거는 `<<TreeviewOpen>>`(화살표 클릭·더블클릭·키보드 등 펼치는
    방법과 무관하게 항상 발생하는 가상 이벤트)이다 — `<Double-1>`에만 걸어뒀던 이전
    구현은 ttk.Treeview의 확장 화살표 단일 클릭(네이티브 Tk 토글, `<Double-1>`과 무관)이
    걸리지 않아 "(불러오는 중...)" 플레이스홀더가 영구히 안 풀리는 버그가 있었다
    (2026-09-26 실사용자 실측). 더블클릭 핸들러는 그대로 유지하되 같은 로드 로직을
    공유하고, `already_loaded` 가드로 두 경로가 겹쳐 호출돼도 fetch는 최대 1회만
    일어난다."""
    import tkinter as tk
    from tkinter import ttk, messagebox

    if cache is None:
        cache = {}

    win = tk.Toplevel(parent)
    win.title('OneDrive에서 파일 선택')
    win.resizable(True, True)
    win.transient(parent)
    _set_window_icon(win)
    _apply_theme(win)

    frame = tk.Frame(win, padx=16, pady=12)
    frame.pack(fill='both', expand=True)

    tk.Label(frame, text='파일을 선택해 주세요',
              anchor='w', wraplength=380, justify='left').pack(anchor='w', pady=(0, 8))

    tree = ttk.Treeview(frame, show='tree', height=16)
    tree.pack(fill='both', expand=True)

    node_paths = {}  # iid -> (remote_path, is_folder)

    def _insert_children(parent_iid: str, remote_path: str):
        items = _onedrive_visible_children(_cached_onedrive_children(remote_path, cache))
        if parent_iid == '':
            # 최상위 레벨만 실측된 3개 폴더로 제한한다(§J2) — 하위 폴더 탐색(재귀 호출,
            # parent_iid가 트리 노드 iid로 비어있지 않음)에는 이 필터를 적용하지 않는다.
            items = _onedrive_root_visible_children(items)
        for item in items:
            name = item['name']
            child_path = f'{remote_path}/{name}'
            is_folder = 'folder' in item
            iid = tree.insert(parent_iid, 'end', text=('📁 ' if is_folder else '📄 ') + name)
            node_paths[iid] = (child_path, is_folder)
            if is_folder:
                # 실제 자식은 lazy-load 시점(<<TreeviewOpen>>)에 채운다 — 자리표시자
                # 하나를 미리 넣어 트리에 펼침 화살표가 보이게 한다(요구사항 §A5 "폴더는
                # lazy-load 확장").
                tree.insert(iid, 'end', text='(불러오는 중...)')

    try:
        _insert_children('', root_remote_path)
    except Exception as e:
        messagebox.showerror('OneDrive 조회 실패', str(e), parent=win)

    def _expand_start_path(subpath):
        """subpath의 폴더들을 위에서부터 차례로 펼치고 마지막 폴더를 보이게 한다(조회 실패는 삼킨다)."""
        parent_iid = ''
        for seg in subpath or ():
            norm = seg.replace(' ', '')
            target = None
            for iid in tree.get_children(parent_iid):
                info = node_paths.get(iid)
                if info and info[1] and info[0].rsplit('/', 1)[-1].replace(' ', '') == norm:
                    target = iid
                    break
            if target is None:
                return
            _ensure_folder_children_loaded(target)
            tree.item(target, open=True)
            parent_iid = target
        if parent_iid:
            tree.see(parent_iid)

    # §J5(선택 구현) → L1(2026-09-27, 공유 헬퍼로 추출) — 최상위가 정확히 3개 폴더로
    # 고정됐으니(§J2), 팝업이 열리는 시점에도 그 3개의 자식 목록을 백그라운드로 다시
    # 프리페치 트리거한다(이미 §K1 재귀 프리페치가 `_ask_combined_input_popup()` 표시
    # 직후부터 같은 `cache`를 채워두고 있었을 가능성이 높지만, 로그인 실패 등으로 그
    # 조기 프리페치가 못 끝났을 경우의 2차 기회로 유지 — `_cached_onedrive_children()`의
    # "이미 있으면 스킵" 계약 덕분에 중복 호출 비용은 최소화된다).
    _start_onedrive_prefetch_thread(cache, root_remote_path)

    def _ensure_folder_children_loaded(iid):
        """iid가 폴더 노드고 아직 실제 자식이 로드되지 않았으면(자리표시자만 있으면)
        list_children()으로 채운다 — idempotent(이미 로드됐으면 아무것도 안 함)라
        `<<TreeviewOpen>>`과 더블클릭 두 경로가 같은 폴더에 대해 겹쳐 호출돼도 안전하다."""
        info = node_paths.get(iid)
        if info is None:
            return
        remote_path, is_folder = info
        if not is_folder:
            return
        children = tree.get_children(iid)
        already_loaded = not (len(children) == 1 and children[0] not in node_paths)
        if already_loaded:
            return
        tree.delete(children[0])
        try:
            _insert_children(iid, remote_path)
        except Exception as e:
            messagebox.showerror('OneDrive 조회 실패', str(e), parent=win)

    def _on_double_click(event):
        iid = tree.identify_row(event.y)
        if not iid or iid not in node_paths:
            return
        remote_path, is_folder = node_paths[iid]
        if not is_folder:
            return
        _ensure_folder_children_loaded(iid)
        tree.item(iid, open=not tree.item(iid, 'open'))

    tree.bind('<Double-1>', _on_double_click)

    def _on_treeview_open(event):
        """ttk.Treeview의 확장 화살표(▶) 단일 클릭은 `<Double-1>`과 무관한 네이티브 Tk
        토글이라 위 `_on_double_click`이 전혀 호출되지 않는다 — 이것이 실사용자가 겪은
        "(불러오는 중...)"에서 영구히 멈추는 버그의 원인이었다(2026-09-26). `<<TreeviewOpen>>`
        은 화살표 클릭·더블클릭·키보드 등 펼치는 방법과 무관하게 항상 발생하는 가상
        이벤트라 lazy-load 트리거로 이걸 써야 한다. 이 이벤트가 발생한 시점엔 Tk가 이미
        열림 상태로 전환한 뒤이므로, 여기서는 콘텐츠만 채우고 `open` 속성을 다시 토글하지
        않는다(수동 토글은 네이티브 전환과 겹쳐 오히려 되닫힘을 유발할 수 있다).
        `tree.focus()`로 대상 iid를 얻는다 — 이 가상 이벤트는 이벤트 객체에 대상 정보를
        싣지 않으므로 이것이 Tkinter의 표준 관용구다."""
        iid = tree.focus()
        _ensure_folder_children_loaded(iid)

    tree.bind('<<TreeviewOpen>>', _on_treeview_open)

    if start_subpath:
        try:
            _expand_start_path(start_subpath)
        except Exception:
            pass

    selected = {'path': None}

    def _on_select(_event=None):
        sel = tree.selection()
        if not sel:
            selected['path'] = None
            return
        info = node_paths.get(sel[0])
        selected['path'] = info[0] if info and not info[1] else None

    tree.bind('<<TreeviewSelect>>', _on_select)

    result = [('', '')]

    def _on_choose():
        if not selected['path']:
            messagebox.showwarning('선택 오류', '파일을 선택해 주세요.', parent=win)
            return
        remote_path = selected['path']
        # L2(2026-09-27): 여기서는 실제로 다운로드하지 않는다 — 호출부(`_browse()`)가
        # 백그라운드로 다운로드를 넘겨받아 사용자가 즉시 다음 파일을 선택할 수 있게 한다.
        # 반환하는 local_path는 다운로드가 끝나면 실제로 저장될 위치를 결정적으로 계산한
        # 값일 뿐, 이 시점에 그 파일이 실제로 존재한다는 보장은 없다(호출부 docstring 참고).
        filename = remote_path.rsplit('/', 1)[-1]
        local_path = Path(dest_dir) / filename
        result[0] = (str(local_path), remote_path)
        win.destroy()

    def _on_cancel():
        win.destroy()

    btn_row = tk.Frame(win, pady=8)
    btn_row.pack()
    tk.Button(btn_row, text='선택', command=_on_choose, width=10).pack(side='left', padx=(0, 6))
    tk.Button(btn_row, text='취소', command=_on_cancel, width=10).pack(side='left')

    win.protocol('WM_DELETE_WINDOW', _on_cancel)
    win.update_idletasks()
    _w, _h = 440, 480
    _sw, _sh = win.winfo_screenwidth(), win.winfo_screenheight()
    win.geometry(f'{_w}x{_h}+{max(0,(_sw-_w)//2)}+{max(0,(_sh-_h)//2)}')
    win.grab_set()
    win.wait_window()
    return result[0]


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
    try:
        has_update, sha = mu.check_for_update(config)
    except mu.UpdateError as e:
        messagebox.showerror('업데이트 확인 실패', str(e), parent=parent)
        return
    if not has_update:
        messagebox.showinfo('업데이트 확인', '이미 최신 버전입니다.', parent=parent)
        return
    if not messagebox.askyesno('업데이트 확인', '새 버전이 있습니다. 지금 적용할까요?', parent=parent):
        return
    try:
        mu.apply_update(sha, _SCRIPT_DIR)
    except mu.UpdateError as e:
        messagebox.showerror('업데이트 적용 실패', str(e), parent=parent)
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

    mass_type='youth'(청년미사)는 두 가지가 성인미사와 다르다(요구사항 §A4/§A5):
    (1) 화답송 악보 PPT/사진 행 자체가 없다(청년미사 악보는 성가 선택 팝업에서 별도
        출처로 다룬다). (2) 나머지 파일 행의 '찾아보기'가 로컬 파일 다이얼로그 대신
        OneDrive Graph API 커스텀 파일 브라우저를 연다(성당 공용 계정은 이 PC에 로컬
        동기화가 안 돼 있어 로컬 다이얼로그가 무의미하기 때문).
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

    # 청년미사 로컬 모드(§3.3.1)면 OneDrive 브라우저·프리페치 대신 로컬 파일 선택창을 쓴다.
    # 모드는 호출부(_run_gui_mode)가 이미 결정해 둔 값을 읽기만 한다(여기서 선택창을 띄우지 않음).
    _youth_local = mass_type == 'youth' and get_youth_ppt_mode() == 'local'

    # 청년미사 OneDrive 파일 브라우저(§A5)의 세션 내 폴더 목록 캐시 — 이 팝업 하나가 열려
    # 있는 동안 여러 행('찾아보기')이 같은 폴더를 반복 조회하지 않도록 공유한다.
    _od_cache = {}

    # 버그4 — OneDrive에서 고른 파일은 Entry(`vars_[key]`)에 표시용 상대경로만 넣어두고
    # (`_onedrive_display_path`), 실제 로컬 다운로드 경로는 이 dict에 따로 저장한다.
    # on_ok()는 `_resolved_field_value()`로 이 dict를 우선 조회한다 — Entry의
    # textvariable이 곧 vars_[key]이므로, 표시 문자열을 그대로 파이프라인에 넘기면
    # 존재하지 않는 경로("13. 기도문\위령 성월 기도.pptx")를 쓰게 되는 회귀가 생긴다.
    _od_real_paths = {}

    # L2(2026-09-27) — 파일 선택 즉시 반환, 실제 다운로드는 백그라운드. `key`별 다운로드
    # 진행 상태를 추적한다: {'generation': int, 'status': 'pending'|'done'|'failed',
    # 'error': str|None, 'thread': Thread|None}. `generation`은 같은 슬롯을 재선택할
    # 때마다 증가시켜, 먼저 시작한(느린) 다운로드가 나중에 끝나도 최신 선택의 상태를
    # 덮어쓰지 못하게 한다(코디네이터 명시 우려 — "가장 마지막 선택만 유효해야 함").
    # 백그라운드 워커(`_download_worker`)는 이 dict와 `_download_lock`만 건지고 Tk
    # 위젯은 절대 건드리지 않는다(§K1/§K2와 동일한 스레드 안전 원칙).
    import threading
    _download_lock = threading.Lock()
    _download_state = {}

    def _download_worker(key, remote_path, dest_dir, generation):
        """L2 리뷰 라운드19 확정 버그 수정 — blob 최신화(느릴 수 있음, 락 밖)와 날짜
        폴더로의 실제 커밋(락 안, generation 재확인 후에만)을 분리한다. 이전 구현은
        `_download_with_retries()`가 그 자리에서 바로 dest_dir에 커밋해버려서, 같은
        슬롯을 같은 베이스네임으로 재선택했을 때(이 앱의 실제 관례 — "시작기도.pptx" 등)
        구세대(느린) 스레드가 generation 재확인 **이전에** 이미 최신 파일을 덮어쓴
        뒤였다 — 재확인 자체는 메모리 상태(`_download_state`)만 보호했을 뿐, 디스크
        커밋은 무조건 일어났다. 이제 커밋 자체를 락 안, generation 일치 확인 뒤로
        옮겨서 구세대 스레드가 최종 파일에 손댈 기회 자체를 없앤다."""
        try:
            blob_path = _download_with_retries(remote_path, _od_cache)
            status, error = 'done', None
        except Exception as e:
            blob_path = None
            status, error = 'failed', str(e)
        with _download_lock:
            entry = _download_state.get(key)
            if entry is None or entry.get('generation') != generation:
                return  # 이미 재선택으로 낡은 세대 — 아무것도 커밋하지 않는다.
            if status == 'done':
                try:
                    filename = remote_path.rsplit('/', 1)[-1]
                    dest = Path(dest_dir) / filename
                    _commit_onedrive_blob_to_dest(blob_path, dest)
                except Exception as e:
                    status, error = 'failed', str(e)
            entry['status'] = status
            entry['error'] = error

    def _wait_for_pending_downloads_or_report_failure() -> bool:
        """'확인' 클릭 시 호출한다. 아직 끝나지 않은 백그라운드 다운로드가 있으면 전부
        끝날 때까지 대기한다(사용자 확정 결정 3 — 별도 대기 표시는 하지 않음). 대기
        자체는 무한정이 아니다 — `_download_with_retries()`의 재시도 루프가 유한(기본
        최초 1회+2회 재시도)해서 각 워커 스레드는 반드시 끝난다. 재시도까지 모두 실패한
        항목이 있으면 그 시점에만 사용자에게 알리고 False를 반환한다(팝업은 유지 —
        불완전한 파일로 조용히 진행하지 않는다)."""
        with _download_lock:
            pending_threads = [
                entry['thread'] for entry in _download_state.values()
                if entry.get('status') == 'pending' and entry.get('thread') is not None
            ]
        for t in pending_threads:
            t.join()
        with _download_lock:
            failed_keys = [
                key for key, entry in _download_state.items() if entry.get('status') == 'failed'
            ]
        if failed_keys:
            key_to_label = {k: label for k, label, _req in file_rows}
            labels = ', '.join(key_to_label.get(k, k) for k in failed_keys)
            messagebox.showerror(
                '다운로드 실패',
                f'다음 파일을 OneDrive에서 받아오지 못했습니다: {labels}\n'
                '해당 항목을 다시 선택해 주세요.',
                parent=root,
            )
            return False
        return True

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

        if mass_type == 'youth' and not _youth_local:

            def _browse(v=var, entry=entry, key=key):
                root.focus_set()
                date_val = date_var.get().strip()
                if not re.match(r'^\d{8}$', date_val):
                    messagebox.showwarning(
                        '입력 오류', '먼저 미사 일자(YYYYMMDD)를 입력해 주세요.', parent=root,
                    )
                    return
                dest_dir = Path(OUTPUT_ROOT) / output_folder_key(date_val, mass_type)
                local_path, remote_path = _ask_onedrive_file_browser_popup(
                    root, dest_dir, cache=_od_cache,
                    start_subpath=browse_start_subpath(mass_type, key),
                )
                if not remote_path:
                    return  # 취소 — L2 이전과 동일하게 아무것도 건드리지 않는다.
                _apply_onedrive_selection(
                    key, local_path, remote_path, vars_, _od_real_paths, entries,
                )
                # L2(2026-09-27) — 실제 다운로드는 백그라운드로 넘기고 여기서는 즉시
                # 반환한다(사용자가 바로 다음 파일을 선택할 수 있어야 함). `generation`을
                # 올려서, 같은 슬롯을 재선택했을 때 이전(느린) 다운로드의 뒤늦은 완료가
                # 이번 선택의 상태를 덮어쓰지 못하게 한다.
                with _download_lock:
                    prev = _download_state.get(key, {})
                    generation = prev.get('generation', 0) + 1
                    _download_state[key] = {
                        'generation': generation, 'status': 'pending', 'error': None, 'thread': None,
                    }
                thread = threading.Thread(
                    target=_download_worker, args=(key, remote_path, dest_dir, generation),
                    daemon=True,
                )
                with _download_lock:
                    _download_state[key]['thread'] = thread
                thread.start()

        else:

            def _browse(v=var, types=row_types, key=key):

                root.focus_set()

                _kw = {}
                _start = browse_initial_dir(mass_type, key)
                if _start is not None:
                    _kw['initialdir'] = str(_start)
                path = filedialog.askopenfilename(parent=root, filetypes=types, **_kw)

                if path:

                    v.set(path)
                    if mass_type == 'adult' and key == '화답송_pptx':
                        remember_psalm_dir(path)

        tk.Button(frame, text='찾아보기', command=_browse, pady=0).grid(row=r, column=2, pady=6)



    result = [None]

    def _field_value(key: str) -> str:
        """버그4 — on_ok()의 모든 필드 조회는 반드시 이 함수를 거친다(`vars_[key].get()`을
        직접 부르면 OneDrive에서 고른 파일의 실제 경로 대신 표시용 상대경로를 쓰게 되는
        회귀가 생긴다). `_resolved_field_value()`(GUI 없이 단위 테스트됨)에 위임."""
        return _resolved_field_value(key, vars_, _od_real_paths)



    def on_ok():

        date_val = date_var.get().strip()

        if not re.match(r'^\d{8}$', date_val):

            messagebox.showwarning(

                '입력 오류',

                'YYYYMMDD 형식으로 날짜를 입력해 주세요.\n예: 20260628',

                parent=root,

            )

            return

        # L2(2026-09-27) — 아직 끝나지 않은 백그라운드 다운로드가 있으면 여기서 전부
        # 끝날 때까지 기다린다(사용자 확정 결정 1). 재시도까지 모두 실패한 항목이 있으면
        # 사용자에게 알리고 여기서 중단한다(불완전한 파일로 조용히 진행 금지 — 사용자
        # 확정 결정 2/3).
        if mass_type == 'youth' and not _wait_for_pending_downloads_or_report_failure():
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

    # L1(2026-09-27) — §K1의 재귀 프리페치를 "찾아보기" 클릭 시점이 아니라 이 첫 입력창이
    # 뜨는 시점으로 앞당긴다. 사용자가 날짜조차 입력하기 전부터 백그라운드로 캐시를
    # 채워두면, 실제로 '찾아보기'를 누를 즈음엔 이미 캐시가 준비돼 있을 가능성이 높다.
    # 청년미사만 트리거한다 — 성인미사는 이 OneDrive 브라우저 자체를 쓰지 않으므로
    # 프리페치가 아무 소용이 없는 API 호출 낭비가 된다. `_start_onedrive_prefetch_thread()`
    # 는 데몬 스레드를 시작만 하고 즉시 반환하므로, 로그인이 안 돼 있거나 네트워크가
    # 없어도(또는 MSAL 기기 코드 로그인 대기로 무한정 멈춰도) 이 창 자체는 항상 즉시
    # 뜬다 — 실패는 `_prefetch_onedrive_children()`이 경로 단위로 조용히 삼킨다.
    if mass_type == 'youth' and not _youth_local:
        _start_onedrive_prefetch_thread(_od_cache)

    root.after(50, date_entry.focus_set)

    root.mainloop()



    if result[0] is None:

        raise RuntimeError('입력이 취소됐습니다.')

    return result[0]




def _show_onedrive_device_code_window(flow: dict):
    """OneDrive 로그인 코드를 보여주는 비모달 안내 창을 만들어 표시하고 즉시 반환한다
    (mainloop()를 돌리지 않음 — 호출부가 로그인 완료 후 destroy()해야 함). 이 창을 띄운
    뒤 이어지는 MSAL의 블로킹 폴링(acquire_token_by_device_flow) 동안은 이벤트 루프가
    돌지 않아 Windows가 이 창을 '응답 없음'으로 표시할 수 있지만, 코드/URL은 이미
    렌더링돼 있어 사용자가 읽고 로그인을 진행하는 데는 지장이 없다."""
    import tkinter as tk
    import webbrowser

    win = tk.Tk()
    win.title('OneDrive 로그인 필요')
    win.resizable(False, False)
    _set_window_icon(win)
    _apply_theme(win)

    frame = tk.Frame(win, padx=24, pady=20)
    frame.pack()
    tk.Label(
        frame, text='청년미사 성가 자료를 내려받으려면 OneDrive 로그인이 필요합니다.',
        font=(_UI['font'], _UI['font_sz'], 'bold'), wraplength=340, justify='left',
    ).pack(anchor='w', pady=(0, 12))

    verification_uri = flow.get('verification_uri', 'https://microsoft.com/devicelogin')
    tk.Label(frame, text='1. 아래 링크를 클릭하거나, 브라우저 주소창에 붙여넣으세요:', anchor='w').pack(anchor='w')
    # tk.Label은 텍스트 드래그 선택을 지원하지 않는다(실사용자 보고 — 링크를 마우스로
    # 드래그해도 선택이 안 돼 복사할 수 없었음) — 읽기 전용 Entry로 바꿔 드래그 선택 +
    # Ctrl+C 복사를 네이티브로 지원하게 한다. 클릭-오픈 동작은 유지하되, 드래그로 텍스트를
    # 선택한 뒤 손을 뗀 경우(selection_present())는 브라우저를 열지 않는다 — 그렇지
    # 않으면 드래그 선택을 시작하는 순간(버튼을 누르는 시점)마다 매번 브라우저가 열려버린다.
    link_var = tk.StringVar(value=f'    {verification_uri}')
    link = tk.Entry(
        frame, textvariable=link_var, fg='#0563c1', readonlybackground=_UI['bg'],
        disabledforeground='#0563c1', relief='flat', bd=0, justify='left',
        font=(_UI['font'], _UI['font_sz']), state='readonly', cursor='hand2',
        width=len(link_var.get()) + 2,
    )
    link.pack(anchor='w', fill='x')

    def _open_link(_e=None):
        if not link.selection_present():
            webbrowser.open(verification_uri)

    link.bind('<ButtonRelease-1>', _open_link)

    tk.Label(frame, text='2. 아래 코드를 입력하세요:', anchor='w').pack(anchor='w', pady=(10, 2))
    code_row = tk.Frame(frame)
    code_row.pack(anchor='w', pady=(0, 12))
    user_code = flow.get('user_code', '')
    tk.Label(
        code_row, text=user_code, font=('Consolas', 22, 'bold'), fg=_UI['primary'],
    ).pack(side='left')

    def _copy_code():
        win.clipboard_clear()
        win.clipboard_append(user_code)
        win.update()
        copy_btn.config(text='복사됨!')
        win.after(1500, lambda: copy_btn.config(text='복사'))

    copy_btn = tk.Button(code_row, text='복사', command=_copy_code)
    copy_btn.pack(side='left', padx=(10, 0))

    tk.Label(
        frame, text='로그인이 완료되면 이 창은 자동으로 닫힙니다.', fg='#666666',
    ).pack(anchor='w')

    win.update_idletasks()
    _w, _h = 380, win.winfo_reqheight()
    _sw, _sh = win.winfo_screenwidth(), win.winfo_screenheight()
    win.geometry(f'{_w}x{_h}+{max(0,(_sw-_w)//2)}+{max(0,(_sh-_h)//2)}')
    win.deiconify()
    win.lift()
    win.attributes('-topmost', True)
    win.update()
    return win


def ensure_onedrive_login() -> None:
    """청년미사 워커 스레드를 시작하기 전, 메인 스레드에서 OneDrive 로그인을 미리
    끝내둔다.

    pythonw.exe(콘솔 없음) 배포 환경에서는 missa_onedrive.get_access_token()의 print()가
    전혀 보이지 않고, 실행 모드와 무관하게도 _run_with_progress_window()가 워커 스레드의
    stdout을 캡처해 완료 전까지는 화면에 보여주지 않는다 — 그래서 워커 스레드 안(성가
    교체 단계)에서 처음으로 로그인이 필요해지는 순간, 사용자에게는 이유를 알 수 없는
    무한 대기로만 보인다(2026-09-19 실사용자 테스트로 발견 — "성가 교체..." 88%에서 정지,
    msal_token_cache.bin 부재 + onedrive_youth_hymn_folder 로컬 미설정이 맞물려 재현).
    워커 스레드가 시작되기 **전**, 메인 스레드에서 캐시된 토큰이 있는지 조용히 확인하고,
    없으면 기기 코드를 GUI 창으로 보여주며 로그인을 여기서 끝내둔다 — 그러면 워커 스레드
    안에서 get_access_token()이 다시 호출돼도 캐시 히트라 조용히 통과한다.

    get_access_token() 자체(특히 app.acquire_token_by_device_flow()의 폴링)는 별도 스레드
    (_worker)에서 실행한다 — 이 함수를 메인 스레드에서 직접 블로킹 호출하면 그동안 Tk
    이벤트 루프가 멈춰, 창이 '응답 없음' 상태가 되어 사용자가 코드/URL을 마우스로 선택해
    복사할 수조차 없다(2026-09-19 실사용자 테스트로 재현 — 안내 창에 마우스를 올리자마자
    Not Responding). Tk 창 생성·update()는 반드시 메인 스레드에서만 해야 하므로(스레드
    안전성), on_device_code 콜백은 flow만 메인 스레드로 넘기고 창 생성은 메인 스레드의
    폴링 루프가 담당한다."""
    import threading

    import missa_onedrive as od

    flow_box = {}
    flow_ready = threading.Event()
    result_box = {}

    def _on_device_code(flow):
        flow_box['flow'] = flow
        flow_ready.set()

    def _worker():
        try:
            result_box['token'] = od.get_access_token(on_device_code=_on_device_code)
        except Exception as e:
            result_box['error'] = e

    t = threading.Thread(target=_worker, daemon=True)
    t.start()

    win = None
    while t.is_alive():
        if win is None and flow_ready.is_set():
            win = _show_onedrive_device_code_window(flow_box['flow'])
        if win is not None:
            try:
                win.update()
            except Exception:
                pass
        t.join(timeout=0.05)

    if win is not None:
        try:
            win.destroy()
        except Exception:
            pass

    if 'error' in result_box:
        raise result_box['error']




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

    root.mainloop()

    _progress_callback[0] = None

    return result[0], result[1]


# ─────────────────────────────────────────────────────────────────────────────
# 결과 창
# ─────────────────────────────────────────────────────────────────────────────


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

        tk.Button(

            btn_frame, text='파일 열기', command=open_file,

            bg=_UI['primary'], fg=_UI['fg_light'],

            font=(_UI['font'], _UI['font_sz'], 'bold'), width=12, relief='flat', cursor='hand2'

        ).pack(side=tk.LEFT, padx=6)

        tk.Button(

            btn_frame, text='폴더 열기', command=open_folder,

            bg='#546ea3', fg=_UI['fg_light'],

            font=(_UI['font'], _UI['font_sz'], 'bold'), width=12, relief='flat', cursor='hand2'

        ).pack(side=tk.LEFT, padx=6)



    tk.Button(

        btn_frame, text='닫기', command=root.destroy,

        bg=_UI['primary'], fg=_UI['fg_light'],

        font=(_UI['font'], _UI['font_sz'], 'bold'), width=14, relief='flat', cursor='hand2'

    ).pack(side=tk.LEFT, padx=6)

    # 2026-10-03: 기존 3개 버튼 아래, 한 사이즈 작은 '원드라이브로 복사' 버튼(성인·청년 공통).
    _mass = _last_output_mass[0]
    if output_path and date_str and _mass in ('adult', 'youth'):
        copy_frame = tk.Frame(root)
        copy_frame.pack(pady=(0, 10))
        tk.Button(
            copy_frame, text='원드라이브로 복사',
            command=lambda: _copy_result_with_ui(root, _mass, output_path, date_str),
            bg='#546ea3', fg=_UI['fg_light'],
            font=(_UI['font'], max(_UI['font_sz'] - 1, 8), 'bold'), width=16, relief='flat', cursor='hand2'
        ).pack()

    _center_window(root)

    root.deiconify()

    root.mainloop()
