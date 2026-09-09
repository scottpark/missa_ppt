"""
missa_to_ppt.py 모듈 분리 리팩토링 Phase 2 — Tkinter 팝업 + 날짜판단/설정.

is_sunday_mass()·config.json 로드/저장·get_onedrive_hymn_folder() 등은 논리적으로는
GUI가 아니지만, GUI 팝업(_ask_combined_input_popup 등)과 진입점(main, parse_args,
find_files) 양쪽에서 모두 호출된다. 진입점에 그대로 두면 entry↔gui 순환 임포트가
생기므로, 이 모듈에 흡수해 entry→gui 단방향 의존만 유지한다.
"""
from __future__ import annotations

import io
import json
import re
import sys
from pathlib import Path

OUTPUT_ROOT = 'output'  # 날짜 폴더(YYYYMMDD/)를 모아두는 상위 폴더.
# entry(missa_to_ppt.py find_files/get_json_data/main)와 gui(_show_result_window) 양쪽에서
# 쓰여서 여기에 둔다 — entry→gui 단방향 의존만 유지하기 위함(design doc §2.3 참고).

# ─── GUI 상태 변수 ───────────────────────────────────────────────────────────
_progress_callback = [None]        # 진행률 콜백 (진행 창 표시 중에만 설정)
_first_dialog_pos  = [None, None]  # 첫 번째 다이얼로그 위치 (x, y)
_last_output_path  = [None]        # main()이 저장한 최종 출력 경로
_last_date_str     = [None]        # main()이 사용한 날짜 문자열
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



def _ask_onedrive_path_popup() -> str:

    """tkinter 폴더 선택 다이얼로그로 OneDrive 성가 폴더 경로 입력."""

    import tkinter as tk

    from tkinter import filedialog, messagebox

    root = tk.Tk()

    root.withdraw()

    messagebox.showinfo(

        '성가 라이브러리 경로 설정',

        '성가 PPT 파일들이 저장된 OneDrive 폴더를 선택해 주세요.\n(최초 1회만 설정됩니다.)',

    )

    folder = filedialog.askdirectory(title='성가 OneDrive 폴더 선택')

    root.destroy()

    return folder



def get_onedrive_hymn_folder() -> Path:

    """저장된 OneDrive 성가 폴더 경로 반환. 미설정이거나 경로가 없으면 팝업으로 입력받아 저장."""

    config = _load_config()

    path_str = config.get('onedrive_hymn_folder', '')

    if path_str and Path(path_str).is_dir():

        return Path(path_str)

    print('  [설정] OneDrive 성가 폴더가 설정되지 않았습니다. 폴더를 선택해 주세요.')

    folder = _ask_onedrive_path_popup()

    if not folder or not Path(folder).is_dir():

        raise RuntimeError('OneDrive 성가 폴더 선택이 취소되었습니다.')

    config['onedrive_hymn_folder'] = folder

    _save_config(config)

    print(f'  [설정] OneDrive 성가 폴더 저장됨: {folder}')

    return Path(folder)


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




def _ask_combined_input_popup() -> tuple:

    """미사 일자 + 참조 파일 3개를 하나의 창에서 입력받는다.

    반환: (date_str, files_dict)

    """

    import tkinter as tk

    from tkinter import filedialog, messagebox



    DIALOG_TITLE = '미사 PPT 파일 자동화 - 브로큰베이교구 한인 천주교회'

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

    vars_ = {}

    for i, (key, label, _req) in enumerate(file_rows):

        r = i + 1

        tk.Label(frame, text=label, anchor='w', width=18).grid(

            row=r, column=0, sticky='w', pady=6

        )

        var = tk.StringVar()

        vars_[key] = var

        tk.Entry(frame, textvariable=var, width=46, state='readonly').grid(

            row=r, column=1, padx=(8, 6), pady=6

        )

        row_types = HWADAPSONG_TYPES if key == '화답송_pptx' else PPTX_TYPES

        def _browse(v=var, types=row_types):

            root.focus_set()

            path = filedialog.askopenfilename(parent=root, filetypes=types)

            if path:

                v.set(path)

        tk.Button(frame, text='찾아보기', command=_browse, pady=0).grid(row=r, column=2, pady=6)



    result = [None]



    def on_ok():

        date_val = date_var.get().strip()

        if not re.match(r'^\d{8}$', date_val):

            messagebox.showwarning(

                '입력 오류',

                'YYYYMMDD 형식으로 날짜를 입력해 주세요.\n예: 20260628',

                parent=root,

            )

            return

        ref = vars_['ref_pptx'].get()

        if not ref:

            messagebox.showwarning('입력 오류', '참조 미사 PPT를 선택해 주세요.', parent=root)

            return

        화답송 = vars_['화답송_pptx'].get()

        if not 화답송 and is_sunday_mass(date_val):

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
        공지사항 = vars_['공지사항'].get()

        if 공지사항 and Path(공지사항).suffix.lower() != '.pptx':

            messagebox.showwarning(

                '입력 오류',

                '공지사항 파일은 PowerPoint(.pptx) 파일만 선택할 수 있습니다.\n'

                '다른 파일을 선택해 주세요.',

                parent=root,

            )

            return  # 팝업 유지, 재선택 유도

        시작기도 = vars_['시작기도'].get()

        미사후기도 = vars_['미사후기도'].get()

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

        log_dir = Path(OUTPUT_ROOT) / date_str / 'log'

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

    _center_window(root)

    root.deiconify()

    root.mainloop()
