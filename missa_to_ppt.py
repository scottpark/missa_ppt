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

import copy

import io

import json

import os

import re

import subprocess

import sys

from pathlib import Path



from pptx import Presentation

from pptx.dml.color import RGBColor

from pptx.enum.shapes import MSO_SHAPE_TYPE

from pptx.oxml.ns import qn

from pptx.util import Emu, Pt



# python-pptx의 _next_slide_partname은 len(sldIdLst)+1을 사용하여

# 삭제 후 add_slide 시 기존 파트명과 충돌이 발생한다.

# 패키지 레벨의 next_partname을 사용하도록 패치.

from pptx.parts.presentation import PresentationPart

from pptx.opc.packuri import PackURI as _PackURI





@property  # type: ignore[misc]

def _safe_next_slide_partname(self):

    return self.package.next_partname('/ppt/slides/slide%d.xml')





PresentationPart._next_slide_partname = _safe_next_slide_partname



# ─────────────────────────────────────────────────────────────────────────────

# 상수

# ─────────────────────────────────────────────────────────────────────────────



CHARS_PER_LINE = 27   # 32pt 바탕체 24.8cm 텍스트박스 실측 기준 약 27자/줄

LINES_PER_SLIDE = 9

ORANGE = RGBColor(255, 192, 0)



HYMN_TYPES = ['입당', '봉헌', '성체', '2차봉헌', '파견']

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

    parser.add_argument('--test', action='store_true', help='팝업 없이 폴더 파일로 번호 자동 추론')

    args = parser.parse_args()



    numbers = {

        '입당': args.입당,

        '봉헌': args.봉헌,

        '성체': args.성체,

        '2차봉헌': args.차봉헌2,

        '파견': args.파견,

    }



    if any(v is None for v in numbers.values()):

        if args.test:

            numbers = _infer_hymn_numbers(numbers)

        else:

            numbers = _ask_numbers_popup(numbers)



    return args.date, numbers, args.화답송_pptx





def _ask_numbers_popup(defaults: dict) -> dict:

    import tkinter as tk



    root = tk.Tk()

    root.withdraw()

    root.title('성가 번호 입력')

    root.resizable(False, False)

    _set_window_icon(root)

    _apply_theme(root)

    frame = tk.Frame(root, padx=24, pady=16)

    frame.pack(fill='both', expand=True)



    labels = ['입당', '봉헌', '성체', '2차봉헌', '파견']

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

        for label in labels:

            val = entries[label].get().strip()

            try:

                result[label] = int(val)

            except ValueError:

                result[label] = None

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



    frame = tk.Frame(root, padx=20, pady=16)

    frame.pack(fill='both', expand=True)



    rows = [

        ('ref_pptx',    '참조 미사 PPT  *', True),

        ('시작기도',    '시작기도 PPT',     False),

        ('화답송_pptx', '화답송 악보 PPT *', True),

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

        def _browse(v=var):

            path = filedialog.askopenfilename(

                parent=root, filetypes=PPTX_TYPES

            )

            if path:

                v.set(path)

        tk.Button(frame, text='찾아보기', command=_browse, pady=0).grid(

            row=i, column=2, pady=6

        )



    result = [None]



    def on_ok():

        ref = vars_['ref_pptx'].get()

        화답송 = vars_['화답송_pptx'].get()

        if not ref:

            messagebox.showwarning('입력 오류', '참조 미사 PPT를 선택해 주세요.', parent=root)

            return

        if not 화답송:

            messagebox.showwarning('입력 오류', '화답송 악보 PPT를 선택해 주세요.', parent=root)

            return

        시작기도 = vars_['시작기도'].get()

        result[0] = {

            'ref_pptx':    Path(ref),

            '시작기도':    Path(시작기도) if 시작기도 else None,

            '화답송_pptx': Path(화답송),

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

        ('화답송_pptx', '화답송 악보 PPT *', True),

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

        def _browse(v=var):

            root.focus_set()

            path = filedialog.askopenfilename(parent=root, filetypes=PPTX_TYPES)

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

        화답송 = vars_['화답송_pptx'].get()

        if not ref:

            messagebox.showwarning('입력 오류', '참조 미사 PPT를 선택해 주세요.', parent=root)

            return

        if not 화답송:

            messagebox.showwarning('입력 오류', '화답송 악보 PPT를 선택해 주세요.', parent=root)

            return

        시작기도 = vars_['시작기도'].get()

        result[0] = (

            date_val,

            {

                'ref_pptx':    Path(ref),

                '시작기도':    Path(시작기도) if 시작기도 else None,

                '화답송_pptx': Path(화답송),

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



    log_io = io.StringIO()

    orig_out, orig_err = sys.stdout, sys.stderr

    sys.stdout = sys.stderr = log_io



    def worker():

        _progress_callback[0] = lambda p, lbl='': pq.put((p, lbl))

        try:

            main_func()

            txt = log_io.getvalue()

            sys.stdout, sys.stderr = orig_out, orig_err

            pq.put(('__done__', txt, None))

        except SystemExit as _e:

            txt = log_io.getvalue()

            sys.stdout, sys.stderr = orig_out, orig_err

            if _e.code and _e.code != 0:

                pq.put(('__done__', txt, f'종료 코드: {_e.code}'))

            else:

                pq.put(('__done__', txt, None))

        except Exception:

            import traceback as _tb

            tb_str = _tb.format_exc()

            txt = log_io.getvalue()

            sys.stdout, sys.stderr = orig_out, orig_err

            pq.put(('__done__', txt, tb_str))



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




def find_files(date_str: str, hymn_numbers: dict) -> dict:

    folder = Path(date_str)

    if not folder.is_dir():

        raise FileNotFoundError(f"폴더 없음: {folder}")



    files = {'ref_pptx': None, '시작기도': None, '화답송_pptx': None, '성가': {}}



    # 화답송 악보 PPT

    for f in folder.iterdir():

        if f.suffix.lower() == '.pptx' and '화답송 악보' in f.name and not f.name.startswith('~$'):

            files['화답송_pptx'] = f



    # 성가 PPT: OneDrive 폴더에서 우선 검색, 없으면 날짜 폴더 fallback

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

        if f.suffix.lower() == '.pptx' and '시작기도' in f.name:

            files['시작기도'] = f



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





# ─────────────────────────────────────────────────────────────────────────────

# JSON 생성

# ─────────────────────────────────────────────────────────────────────────────



def get_json_data(date_str: str) -> dict:

    json_path = Path(date_str) / f'missa_{date_str}.json'



    if not json_path.exists():

        print(f'[{date_str}] 미사 데이터 가져오는 중...')

        import missa_to_json as _m2j

        _m2j.run(date_str, output_root='.')



    with open(json_path, 'r', encoding='utf-8') as f:

        return json.load(f)





# ─────────────────────────────────────────────────────────────────────────────

# 슬라이드 XML 유틸

# ─────────────────────────────────────────────────────────────────────────────



def _slide_text(slide) -> str:

    parts = []

    for shape in slide.shapes:

        if shape.has_text_frame:

            parts.append(shape.text_frame.text)

    return ' '.join(parts)





def all_slide_texts(prs) -> list:

    return [_slide_text(s) for s in prs.slides]





def delete_slide(prs, idx: int):

    """슬라이드 삭제 (0-based 인덱스)."""

    sldIdLst = prs.slides._sldIdLst

    sldId = sldIdLst[idx]

    rId = sldId.get(qn('r:id'))

    sldIdLst.remove(sldId)

    try:

        prs.part.drop_rel(rId)

    except Exception:

        pass





def move_slide(prs, old_idx: int, new_idx: int):

    sldIdLst = prs.slides._sldIdLst

    el = sldIdLst[old_idx]

    sldIdLst.remove(el)

    sldIdLst.insert(new_idx, el)





def _blank_layout(prs):

    """빈 레이아웃 반환 (이름 우선, 없으면 마지막 레이아웃)."""

    for layout in prs.slide_layouts:

        if layout.name in ('Blank', '빈 화면', 'blank'):

            return layout

    return prs.slide_layouts[-1]





def _copy_spTree(src_slide, dst_slide):

    """src_slide의 shape 트리를 dst_slide로 복사."""

    dst_spTree = dst_slide.shapes._spTree

    # 기본 shape 제거

    for child in list(dst_spTree):

        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if tag in ('sp', 'pic', 'grpSp', 'cxnSp', 'graphicFrame'):

            dst_spTree.remove(child)



    src_spTree = src_slide.shapes._spTree

    for child in src_spTree:

        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if tag in ('sp', 'pic', 'grpSp', 'cxnSp', 'graphicFrame'):

            dst_spTree.append(copy.deepcopy(child))





_IMG_ID_COUNTER = [5000]  # 충돌 방지를 위해 큰 숫자부터 시작





def _copy_image_rels(src_slide, dst_slide):

    """이미지 관계를 src에서 dst로 유니크 이름으로 복사하고 rId 매핑 반환."""

    from pptx.parts.image import ImagePart

    from pptx.opc.packuri import PackURI



    rId_map = {}

    dst_pkg = dst_slide.part.package



    for rel in src_slide.part.rels.values():

        # image 타입과 hdphoto 타입 모두 복사 (a14:imgLayer는 hdphoto를 참조함)

        if 'image' not in rel.reltype and 'hdphoto' not in rel.reltype:

            continue

        orig_part = rel.target_part

        orig_name = str(orig_part.partname)

        ext = '.' + orig_name.rsplit('.', 1)[-1] if '.' in orig_name else '.png'



        # 유니크 파일명 생성 (충돌 방지)

        _IMG_ID_COUNTER[0] += 1

        new_partname = PackURI(f'/ppt/media/image{_IMG_ID_COUNTER[0]}{ext}')



        new_part = ImagePart(new_partname, orig_part.content_type, dst_pkg, orig_part.blob)

        new_rId = dst_slide.part.relate_to(new_part, rel.reltype)

        rId_map[rel.rId] = new_rId

    return rId_map





def _update_rId_in_spTree(dst_slide, rId_map: dict):

    """복사된 shape 트리의 rId를 새 rId로 업데이트."""

    for el in dst_slide.shapes._spTree.iter():

        for attr, val in list(el.attrib.items()):

            if val in rId_map:

                el.set(attr, rId_map[val])





def _effective_bg(src_slide):

    """슬라이드 → 레이아웃 → 마스터 순으로 effective 배경 요소와 소유 part를 반환.

    p:bg는 p:cSld 내부에 있으므로 재귀 탐색(.// )을 사용한다."""

    bg = src_slide.element.find('.//' + qn('p:bg'))

    if bg is not None:

        return bg, src_slide.part

    try:

        layout = src_slide.slide_layout

        bg = layout.element.find('.//' + qn('p:bg'))

        if bg is not None:

            return bg, layout.part

        master = layout.slide_master

        bg = master.element.find('.//' + qn('p:bg'))

        if bg is not None:

            return bg, master.part

    except Exception:

        pass

    return None, None





def _copy_bg_image_rels(bg_element, src_part, dst_slide) -> dict:

    """배경(p:bg) 요소 내 이미지 r:embed rId를 dst_slide로 복사하고 rId 매핑 반환."""

    from pptx.parts.image import ImagePart

    from pptx.opc.packuri import PackURI



    R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    rId_map = {}

    dst_pkg = dst_slide.part.package



    for el in bg_element.iter():

        for attr, val in list(el.attrib.items()):

            if attr == f'{{{R_NS}}}embed' and val not in rId_map:

                try:

                    rel = src_part.rels[val]

                    if 'image' not in rel.reltype:

                        continue

                    orig_part = rel.target_part

                    orig_name = str(orig_part.partname)

                    ext = '.' + orig_name.rsplit('.', 1)[-1] if '.' in orig_name else '.png'

                    _IMG_ID_COUNTER[0] += 1

                    new_partname = PackURI(f'/ppt/media/image{_IMG_ID_COUNTER[0]}{ext}')

                    new_part = ImagePart(new_partname, orig_part.content_type, dst_pkg, orig_part.blob)

                    rId_map[val] = dst_slide.part.relate_to(new_part, rel.reltype)

                except Exception:

                    pass

    return rId_map





def duplicate_slide(prs, src_idx: int) -> int:

    """src_idx 슬라이드를 복제해서 맨 끝에 추가. 새 인덱스 반환."""

    src_slide = prs.slides[src_idx]

    blank_layout = _blank_layout(prs)

    new_slide = prs.slides.add_slide(blank_layout)



    _copy_spTree(src_slide, new_slide)

    rId_map = _copy_image_rels(src_slide, new_slide)

    if rId_map:

        _update_rId_in_spTree(new_slide, rId_map)



    # 배경 복사 (p:bg는 p:cSld 내부에 위치)

    src_cSld = src_slide.element.find(qn('p:cSld'))

    src_bg = src_cSld.find(qn('p:bg')) if src_cSld is not None else None

    if src_bg is not None:

        new_cSld = new_slide.element.find(qn('p:cSld'))

        if new_cSld is not None:

            existing = new_cSld.find(qn('p:bg'))

            if existing is not None:

                new_cSld.remove(existing)

            new_cSld.insert(0, copy.deepcopy(src_bg))



    return len(prs.slides) - 1





def insert_slide_copy(prs, position: int, src_idx: int):

    """src_idx 슬라이드를 복제해서 position 위치에 삽입."""

    new_idx = duplicate_slide(prs, src_idx)

    move_slide(prs, new_idx, position)





def copy_slide_from_prs(target_prs, position: int, source_prs, source_idx: int):

    """다른 Presentation에서 슬라이드를 원본 서식(배경 포함)으로 복사해서 position에 삽입.

    소스 슬라이드의 레이아웃 이름으로 타겟에서 매칭 레이아웃을 찾아 사용한다 (원본 서식 유지).

    """

    src_slide = source_prs.slides[source_idx]



    # 소스 레이아웃 이름으로 타겟에서 매칭 레이아웃 탐색

    src_layout_name = src_slide.slide_layout.name

    target_layout = None

    for layout in target_prs.slide_layouts:

        if layout.name == src_layout_name:

            target_layout = layout

            break

    if target_layout is None:

        target_layout = _blank_layout(target_prs)



    new_slide = target_prs.slides.add_slide(target_layout)



    _copy_spTree(src_slide, new_slide)

    rId_map = _copy_image_rels(src_slide, new_slide)

    if rId_map:

        _update_rId_in_spTree(new_slide, rId_map)



    # 배경 복사: 슬라이드 자체 p:bg가 없으면 레이아웃/마스터에서 상속된 배경을 명시적으로 삽입

    # p:bg는 p:cSld 내부에 위치하므로 cSld를 통해 접근/삽입

    src_bg, bg_src_part = _effective_bg(src_slide)

    if src_bg is not None:

        bg_copy = copy.deepcopy(src_bg)

        # 마스터/레이아웃에서 가져온 배경에 이미지가 있으면 rId도 복사

        if bg_src_part is not src_slide.part:

            bg_rId_map = _copy_bg_image_rels(src_bg, bg_src_part, new_slide)

            if bg_rId_map:

                for el in bg_copy.iter():

                    for attr, val in list(el.attrib.items()):

                        if val in bg_rId_map:

                            el.set(attr, bg_rId_map[val])

        new_cSld = new_slide.element.find(qn('p:cSld'))

        if new_cSld is not None:

            existing = new_cSld.find(qn('p:bg'))

            if existing is not None:

                new_cSld.remove(existing)

            new_cSld.insert(0, bg_copy)



    new_idx = len(target_prs.slides) - 1

    move_slide(target_prs, new_idx, position)





# ─────────────────────────────────────────────────────────────────────────────

# 섹션 탐색

# ─────────────────────────────────────────────────────────────────────────────



def find_slide_with_text(prs, keyword: str, start: int = 0) -> int:

    for i in range(start, len(prs.slides)):

        if keyword in _slide_text(prs.slides[i]):

            return i

    return -1





def find_shape_exact_text(slide, exact: str) -> bool:

    for shape in slide.shapes:

        if shape.has_text_frame:

            for para in shape.text_frame.paragraphs:

                if para.text.strip() == exact:

                    return True

    return False





def find_content_range(prs, title_idx: int) -> tuple:

    """title_idx 다음 콘텐츠 슬라이드 범위 (start, end exclusive)."""

    NEXT_SECTION = [

        '화 답 송', '화답송', '제 2 독서', '제2독서',

        '복음 환호송', '복음환호송', '영성체송',

        '봉 헌', '성 체', '2차 봉헌', '파 견',

    ]

    start = title_idx + 1

    end = start

    n = len(prs.slides)

    for i in range(start, n):

        t = _slide_text(prs.slides[i]).strip()

        if not t:

            break

        if any(kw in t for kw in NEXT_SECTION):

            break

        end = i + 1

    return start, end





def find_복음_content_range(prs, title_idx: int) -> tuple:

    """복음 제목 슬라이드 다음 콘텐츠 슬라이드 범위."""

    STOP = ['영성체송', '봉 헌', '봉헌']

    start = title_idx + 1

    end = start

    n = len(prs.slides)

    for i in range(start, n):

        t = _slide_text(prs.slides[i]).strip()

        if not t:

            break

        if any(kw in t for kw in STOP):

            break

        end = i + 1

    return start, end





def find_sections(prs) -> dict:

    """PPT 내 모든 섹션 위치를 찾아 dict로 반환."""

    sections = {'title': 0}

    n = len(prs.slides)

    texts = all_slide_texts(prs)



    # 입당송

    i = find_slide_with_text(prs, '입당송')

    if i >= 0:

        sections['입당송'] = i



    # 시작기도: 입당송 직전 비어있지 않은 슬라이드들

    if '입당송' in sections:

        입당송_idx = sections['입당송']

        ptr = 입당송_idx - 1

        while ptr >= 0 and not texts[ptr].strip():

            ptr -= 1

        end_excl = ptr + 1

        start = ptr

        while start > 0 and texts[start - 1].strip():

            start -= 1

        if start < end_excl:

            sections['시작기도_start'] = start

            sections['시작기도_end'] = end_excl



    # 제1독서

    for i in range(n):

        t = texts[i]

        if '제 1 독서' in t or '제1독서' in t:

            sections['제1독서_title'] = i

            s, e = find_content_range(prs, i)

            sections['제1독서_start'] = s

            sections['제1독서_end'] = e

            break



    # 화답송

    for i in range(n):

        t = texts[i]

        if '화 답 송' in t or '화답송' in t:

            sections['화답송_start'] = i

            end = i + 1

            for j in range(i + 1, n):

                tj = texts[j]

                if not tj.strip():

                    break

                if '화 답 송' in tj or '화답송' in tj:

                    end = j + 1

                else:

                    break

            sections['화답송_end'] = end

            break



    # 제2독서

    for i in range(n):

        t = texts[i]

        if '제 2 독서' in t or '제2독서' in t:

            sections['제2독서_title'] = i

            s, e = find_content_range(prs, i)

            sections['제2독서_start'] = s

            sections['제2독서_end'] = e

            break



    # 복음환호송

    for i in range(n):

        t = texts[i]

        if '복음 환호송' in t or '복음환호송' in t:

            sections['복음환호송'] = i

            break



    # 복음 (제목 슬라이드: '복 음' 단독 shape)

    for i in range(n):

        slide = prs.slides[i]

        if find_shape_exact_text(slide, '복 음'):

            sections['복음_title'] = i

            s, e = find_복음_content_range(prs, i)

            sections['복음_start'] = s

            sections['복음_end'] = e

            break



    # 영성체송

    i = find_slide_with_text(prs, '영성체송')

    if i >= 0:

        sections['영성체송'] = i



    # 성가 섹션 (divider + content)

    HYMN_KEYWORDS = {

        '입당': '입 당', '봉헌': '봉 헌', '성체': '성 체',

        '2차봉헌': '2차 봉헌', '파견': '파 견',

    }

    for htype, kw in HYMN_KEYWORDS.items():

        for i in range(n):

            slide = prs.slides[i]

            if _is_hymn_divider(slide, kw):

                sections[f'{htype}_divider'] = i

                cs = i + 1

                ce = cs

                for j in range(cs, n):

                    if not texts[j].strip():

                        break

                    ce = j + 1

                sections[f'{htype}_content_start'] = cs

                sections[f'{htype}_content_end'] = ce

                break



    return sections





def _is_hymn_divider(slide, type_kw: str) -> bool:

    """성가 divider 슬라이드 판별: 타입 텍스트 단락 AND 순수 숫자 단락이 존재."""

    has_type = False

    has_number = False

    for shape in slide.shapes:

        if not shape.has_text_frame:

            continue

        for para in shape.text_frame.paragraphs:

            t = para.text.strip()

            if t == type_kw or t == type_kw.replace(' ', ''):

                has_type = True

            elif re.match(r'^\d+$', t):

                has_number = True

    return has_type and has_number





# ─────────────────────────────────────────────────────────────────────────────

# 텍스트 서식 유틸 (v3: clone 방식)

# ─────────────────────────────────────────────────────────────────────────────



def _para_append_run(p, new_r):

    """Run을 단락에 추가. <a:endParaRPr> 앞에 삽입해 OOXML 구조를 유지한다.

    endParaRPr 뒤에 run이 오면 PowerPoint가 해당 텍스트를 렌더링하지 않는다."""

    end_rpr = p.find(qn('a:endParaRPr'))

    if end_rpr is not None:

        end_rpr.addprevious(new_r)

    else:

        p.append(new_r)





def _replace_para_text_clone(para, new_text: str):

    """단락의 run을 제거하고 첫 run을 deepcopy해 new_text로 교체 (서식 보존)."""

    from pptx.oxml import parse_xml as pptx_parse_xml

    A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'



    p = para._p

    runs = p.findall(qn('a:r'))

    template_r = runs[0] if runs else None



    for r in runs:

        p.remove(r)

    for br in p.findall(qn('a:br')):

        p.remove(br)



    if template_r is not None:

        new_r = copy.deepcopy(template_r)

        t_el = new_r.find(qn('a:t'))

        if t_el is not None:

            t_el.text = new_text

        _para_append_run(p, new_r)

    else:

        _para_append_run(p, pptx_parse_xml(f'<a:r xmlns:a="{A_NS}"><a:t>{new_text}</a:t></a:r>'))





def _set_single_para_text(tf, text: str):

    """TextFrame을 단일 단락으로 설정 (서식 보존)."""

    from pptx.oxml import parse_xml as pptx_parse_xml

    A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'



    txBody = tf._txBody

    existing_paras = txBody.findall(qn('a:p'))



    template_r = None

    if existing_paras:

        for r in existing_paras[0].findall(qn('a:r')):

            template_r = r

            break



    for p in existing_paras[1:]:

        txBody.remove(p)



    if existing_paras:

        p = existing_paras[0]

        for r in p.findall(qn('a:r')): p.remove(r)

        for br in p.findall(qn('a:br')): p.remove(br)

        if template_r is not None:

            new_r = copy.deepcopy(template_r)

            t_el = new_r.find(qn('a:t'))

            if t_el is not None: t_el.text = text

            _para_append_run(p, new_r)

        else:

            _para_append_run(p, pptx_parse_xml(f'<a:r xmlns:a="{A_NS}"><a:t>{text}</a:t></a:r>'))





# ─────────────────────────────────────────────────────────────────────────────

# 성서 이름 유틸

# ─────────────────────────────────────────────────────────────────────────────



def _extract_book_name(json_title: str) -> str:

    """JSON title에서 성서 이름만 추출 (절 번호 제거).

    예: '열왕기 하권 4,8-11.14-16ㄴ' → '열왕기 하권'

    예: '마태오 10,37-42' → '마태오'

    """

    m = re.match(r'^(.+?)\s+\d', json_title.strip())

    return m.group(1).strip() if m else json_title.strip()





def _josa(word: str) -> str:

    """한글 주격 조사: 받침 없으면 '가', 있으면 '이'."""

    if not word:

        return '이'

    last = word[-1]

    code = ord(last)

    if 0xAC00 <= code <= 0xD7A3:

        return '이' if (code - 0xAC00) % 28 != 0 else '가'

    return '이'





def _update_book_name_after_br(para, new_book_name: str) -> bool:

    """<a:br/> 이후 첫 run에서 성서 이름만 교체하고 접미사('의 말씀입니다.') 보존."""

    p = para._p

    brs = p.findall(qn('a:br'))

    if not brs:

        return False



    br = brs[0]

    after_runs = []

    br_found = False

    for child in list(p):

        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if child is br:

            br_found = True

        elif br_found and tag == 'r':

            after_runs.append(child)



    if not after_runs:

        return False



    full_text = ''.join(

        (r.find(qn('a:t')).text or '') for r in after_runs

        if r.find(qn('a:t')) is not None

    )



    # 패턴 1: "[book]의 말씀입니다" (구약/사도행전 등)

    # 예: "신명기의 말씀입니다." → old_book="신명기"

    m1 = re.match(r'^(.+?)의 말씀입니다', full_text)

    if m1:

        old_book = m1.group(1)

        t_el = after_runs[0].find(qn('a:t'))

        if t_el is not None:

            old_run_text = t_el.text or ''

            if old_run_text.startswith(old_book):

                t_el.text = new_book_name + old_run_text[len(old_book):]

            else:

                t_el.text = new_book_name + '의 말씀입니다'

        return True



    # 패턴 2: "사도 [author]의 [book] 말씀입니다" (서신서)

    # 예: "사도 바오로의 코린토 1서 말씀입니다." → new: "사도 바오로의 로마서 말씀입니다."

    m2 = re.match(r'^((?:사도 )?[가-힣]+의 )(.+?)( 말씀입니다)', full_text)

    if m2:

        prefix = m2.group(1)

        suffix = m2.group(3)

        ending = '.' if full_text.rstrip().endswith('.') else ''

        new_text = prefix + new_book_name + suffix + ending

        t_el = after_runs[0].find(qn('a:t'))

        if t_el is not None:

            t_el.text = new_text

        # run이 여러 개로 분할된 경우 나머지 run 텍스트 제거

        for r in after_runs[1:]:

            t_el2 = r.find(qn('a:t'))

            if t_el2 is not None:

                t_el2.text = ''

        return True



    return False





# ─────────────────────────────────────────────────────────────────────────────

# 독서/복음 콘텐츠 파싱 & 슬라이드 분배

# ─────────────────────────────────────────────────────────────────────────────



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





def _visual_lines(text: str) -> int:
    """문자 수 기반 줄 수 추정 (영성체송 높이 조정 등 레이아웃 외 용도 전용)."""
    return max(1, (len(text) + CHARS_PER_LINE - 1) // CHARS_PER_LINE)


def _wrap_line_count(text: str) -> int:
    """단어 경계 word-wrap 시뮬레이션으로 줄 수 계산.

    layout_units_on_slides의 청킹 로직과 동일하므로 1 dl = 1줄이 보장된다.
    독서·복음 슬라이드 줄 수 계산에 사용한다.
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


def _page_visual_lines(page: list) -> int:
    """display_lines 리스트의 시각적 줄 수 — dl 1개 = 1줄."""
    return len(page)



def layout_units_on_slides(units: list) -> list:

    """절 단위를 슬라이드로 묶음.

    is_continuation=True인 unit은 이전 unit에 합쳐 하나의 논리 단락으로 처리.
    논리 단락은 CHARS_PER_LINE 기준 단어 경계에서 display_lines로 분해하여
    LINES_PER_SLIDE개씩 슬라이드로 묶음 (절 중간 분리 허용).

    반환: [[dl, ...], ...]
      dl: {'text', 'verse_num', 'extra_verses', 'new_para'}
        new_para=True  → 새 PPT paragraph 시작
        new_para=False → 이전 paragraph에 이어 붙임 (문장 연속, paragraph break 없음)
      extra_verses: [(pos_in_dl_text, verse_num), ...] — 오렌지색 절 번호 위치

    """

    # is_continuation 처리: 연속 절을 하나의 논리 단락으로 합침
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

    # 각 논리 단락을 display_lines로 분해 (CHARS_PER_LINE 기준 단어 경계 분리)
    display_lines = []

    for mu in merged:

        mu_text = mu['text']

        mu_extra = mu.get('extra_verses', [])

        mu_verse = mu['verse_num']

        pos = 0

        first_dl = True

        while pos < len(mu_text):

            end = pos + CHARS_PER_LINE

            if end >= len(mu_text):

                chunk = mu_text[pos:]

                next_pos = len(mu_text)

            else:

                # 단어 경계: end 이하에서 마지막 공백
                space = mu_text.rfind(' ', pos, end + 1)

                if space > pos:

                    chunk = mu_text[pos:space]

                    next_pos = space + 1

                else:

                    chunk = mu_text[pos:end]

                    next_pos = end

            if not chunk:

                pos = next_pos

                continue

            # 이 display_line에 해당하는 extra_verses (머지 텍스트 절대 위치 → 상대 위치)
            chunk_start = pos

            chunk_end = pos + len(chunk)

            chunk_extra = [

                (p - chunk_start, v)

                for p, v in mu_extra

                if chunk_start <= p < chunk_end

            ]

            display_lines.append({

                'text': chunk,

                'verse_num': mu_verse if first_dl else '',

                'extra_verses': chunk_extra,

                'new_para': first_dl,

            })

            pos = next_pos

            first_dl = False

    # dl 1개 = 시각적 1줄 (단어 경계 분리 기준)
    # _visual_lines 기반 추정 대신 dl 개수로 직접 분배
    slides = []
    current_slide = []

    for dl in display_lines:
        if len(current_slide) >= LINES_PER_SLIDE:
            slides.append(current_slide)
            current_slide = []
        current_slide.append(dl)

    if current_slide:
        slides.append(current_slide)

    return slides if slides else [[]]



def _verify_and_rebalance_pages(pages: list, label: str) -> list:
    """layout_units_on_slides() 결과를 검증하고 줄 수 이상 슬라이드를 재조정.

    - 비마지막 슬라이드가 LINES_PER_SLIDE 초과 → 마지막 dl들을 다음 슬라이드로 이동
    - 비마지막 슬라이드가 LINES_PER_SLIDE 미만 → 다음 슬라이드 앞 dl들을 흡수 시도
    - 마지막 슬라이드는 어떤 줄 수여도 건드리지 않음
    """
    i = 0
    while i < len(pages):
        lines = _page_visual_lines(pages[i])
        is_last = (i == len(pages) - 1)

        if lines > LINES_PER_SLIDE:
            if len(pages[i]) > 1:
                moved = []
                while _page_visual_lines(pages[i]) > LINES_PER_SLIDE and len(pages[i]) > 1:
                    moved.insert(0, pages[i].pop())
                if i + 1 < len(pages):
                    pages[i + 1] = moved + pages[i + 1]
                else:
                    pages.append(moved)
                new_lines = _page_visual_lines(pages[i])
                print(f'  [{label}] 슬라이드 {i+1}: {lines}줄 → {new_lines}줄 ({len(moved)}개 항목 다음 슬라이드로 이동)')
                continue  # 현재 슬라이드 재검사
            else:
                print(f'  [{label}] 경고 슬라이드 {i+1}: {lines}줄 (항목 1개라 분리 불가)')

        elif not is_last and lines < LINES_PER_SLIDE:
            absorbed = 0
            while i + 1 < len(pages) and pages[i + 1]:
                test = pages[i] + [pages[i + 1][0]]
                if _page_visual_lines(test) <= LINES_PER_SLIDE:
                    pages[i].append(pages[i + 1].pop(0))
                    absorbed += 1
                    if not pages[i + 1]:
                        pages.pop(i + 1)
                        break
                else:
                    break
            if absorbed:
                new_lines = _page_visual_lines(pages[i])
                print(f'  [{label}] 슬라이드 {i+1}: {lines}줄 → {new_lines}줄 ({absorbed}개 항목 흡수)')

        i += 1

    # 최종 검증 보고
    issues = []
    for i, page in enumerate(pages):
        lines = _page_visual_lines(page)
        is_last = (i == len(pages) - 1)
        if lines > LINES_PER_SLIDE:
            issues.append(f'슬라이드 {i+1}: {lines}줄 (분리 불가)')
        elif not is_last and lines < LINES_PER_SLIDE:
            issues.append(f'슬라이드 {i+1}: {lines}줄 (흡수 불가)')
    if issues:
        print(f'  [{label}] 미해결 이슈: {", ".join(issues)}')

    return pages



def _find_content_shape(slide):

    """슬라이드에서 본문 텍스트박스(가장 큰 것) 반환.

    '주님의 말씀입니다.' 등 고정 텍스트박스는 제외.

    """

    EXCLUDE_KW = ['전례문', 'Liturgy', 'Reading', '주님의 말씀입니다', '◎ 하느님', '◎ 그리스도님']

    best = None

    best_area = 0

    for shape in slide.shapes:

        if not shape.has_text_frame:

            continue

        t = shape.text_frame.text.strip()

        if not t or any(kw in t for kw in EXCLUDE_KW):

            continue

        area = shape.width * shape.height

        if area > best_area:

            best_area = area

            best = shape

    return best





def _has_ending_text(slide) -> bool:

    """슬라이드에 '주님의 말씀입니다.' 등 종료 텍스트가 있으면 True."""

    ENDING_KW = ('주님의 말씀입니다', '◎ 하느님', '◎ 그리스도님')

    for shape in slide.shapes:

        if not shape.has_text_frame:

            continue

        t = shape.text_frame.text.strip()

        if any(kw in t for kw in ENDING_KW):

            return True

    return False





def _clear_text_frame(tf):

    """텍스트 프레임의 내용을 비우고 빈 단락 하나를 남긴다."""

    from pptx.oxml import parse_xml as pptx_parse_xml

    A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

    txBody = tf._txBody

    for p in txBody.findall(qn('a:p')):

        txBody.remove(p)

    txBody.append(pptx_parse_xml(f'<a:p xmlns:a="{A_NS}"/>'))





def _set_reading_text(tf, units: list, line_spacing: float = None):

    """독서/복음 콘텐츠를 TextFrame에 설정 (절 번호 오렌지색, 서식 보존).

    extra_verses [(pos, verse_num)] 지원: 단락 내 여러 절 번호를 오렌지색으로 처리.
    line_spacing: 줄간격 배수 (예: 1.1 = 110%). None이면 템플릿 그대로.

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
                r = pptx_parse_xml(f'<a:r xmlns:a="{A_NS}"><a:t>{text}</a:t></a:r>')
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

            if line_spacing is not None:

                pPr = new_p.find(qn('a:pPr'))

                if pPr is None:

                    pPr = pptx_parse_xml(f'<a:pPr xmlns:a="{A_NS}"/>')

                    new_p.insert(0, pPr)

                for old_ln in pPr.findall(qn('a:lnSpc')):

                    pPr.remove(old_ln)

                val = int(line_spacing * 100000)

                pPr.append(pptx_parse_xml(f'<a:lnSpc xmlns:a="{A_NS}"><a:spcPct val="{val}"/></a:lnSpc>'))

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




def _count_slide_lines(slide) -> int:
    """슬라이드 본문 shape의 시각적 줄 수 계산 (word-wrap 시뮬레이션 기준)."""
    shape = _find_content_shape(slide)
    if shape is None:
        return 0
    return sum(
        _wrap_line_count(para.text)
        for para in shape.text_frame.paragraphs
        if para.text.strip()
    )


def _rebalance_reading_slides_post_write(prs, content_start: int, n_content: int, label: str) -> None:
    """독서/복음 본문 슬라이드 기록 후 실제 줄 수 검증 및 재조정.

    비마지막 슬라이드가 LINES_PER_SLIDE 미만(7·8줄) 또는 초과(10줄 이상)이면
    인접 슬라이드와 단락을 이동하여 LINES_PER_SLIDE에 맞춤.
    마지막 슬라이드는 건드리지 않음.
    """
    if n_content < 2:
        return

    def _content_paras(slide):
        shape = _find_content_shape(slide)
        if shape is None:
            return []
        return [p for p in shape.text_frame.paragraphs if p.text.strip()]

    def _get_txBody(slide):
        shape = _find_content_shape(slide)
        return shape.text_frame._txBody if shape else None

    changed = True
    while changed:
        changed = False
        for i in range(n_content - 1):  # 마지막 슬라이드 제외
            cur_slide = prs.slides[content_start + i]
            nxt_slide = prs.slides[content_start + i + 1]
            lines = _count_slide_lines(cur_slide)

            if lines == LINES_PER_SLIDE:
                continue

            if lines > LINES_PER_SLIDE:
                # 마지막 단락을 다음 슬라이드 앞으로 이동
                cur_paras = _content_paras(cur_slide)
                if len(cur_paras) <= 1:
                    print(f'  [{label}] 경고 슬라이드 {i+1}: {lines}줄 (단락 1개, 분리 불가)')
                    continue
                p_elem = cur_paras[-1]._p
                cur_txBody = _get_txBody(cur_slide)
                nxt_txBody = _get_txBody(nxt_slide)
                cur_txBody.remove(p_elem)
                first_nxt_p = nxt_txBody.find(qn('a:p'))
                if first_nxt_p is not None:
                    first_nxt_p.addprevious(p_elem)
                else:
                    nxt_txBody.append(p_elem)
                new_lines = _count_slide_lines(cur_slide)
                print(f'  [{label}] 슬라이드 {i+1}: {lines}줄 → {new_lines}줄 (마지막 단락 → 다음 슬라이드)')
                changed = True

            else:  # lines < LINES_PER_SLIDE
                # 다음 슬라이드의 첫 단락을 이 슬라이드로 흡수
                nxt_paras = _content_paras(nxt_slide)
                if len(nxt_paras) <= 1:
                    continue  # 다음 슬라이드가 빈 슬라이드가 됨 → 건드리지 않음
                first_para_text = nxt_paras[0].text
                if lines + _visual_lines(first_para_text) > LINES_PER_SLIDE:
                    continue  # 흡수 시 9줄 초과
                p_elem = nxt_paras[0]._p
                nxt_txBody = _get_txBody(nxt_slide)
                cur_txBody = _get_txBody(cur_slide)
                nxt_txBody.remove(p_elem)
                cur_all_p = cur_txBody.findall(qn('a:p'))
                if cur_all_p:
                    cur_all_p[-1].addnext(p_elem)
                else:
                    cur_txBody.append(p_elem)
                new_lines = _count_slide_lines(cur_slide)
                print(f'  [{label}] 슬라이드 {i+1}: {lines}줄 → {new_lines}줄 (다음 슬라이드 첫 단락 흡수)')
                changed = True



def replace_reading_slides(prs, content_start: int, content_end: int,

                            units_pages: list, template_idx: int,

                            line_spacing: float = None, merge_threshold: int = 5,

                            label: str = '') -> int:

    """

    독서/복음 콘텐츠 슬라이드를 교체.

    맨 뒤에 '주님의 말씀입니다.' 전용 슬라이드(TYPE B)가 있으면 보존하고

    본문 슬라이드는 그 앞에 삽입.

    units_pages: layout_units_on_slides() 결과

    template_idx: 새 슬라이드 복제 기준 슬라이드 인덱스

    반환: 슬라이드 수 변화 (양수 = 추가됨)

    """

    needed = len(units_pages)

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



    # 본문 슬라이드 수 조정

    if needed > n_usable:

        for k in range(needed - n_usable):

            insert_slide_copy(prs, insert_pos + k, template_idx)

            _set_slide_bg_black(prs.slides[insert_pos + k], prs)

    elif needed < n_usable:

        for _ in range(n_usable - needed):

            delete_slide(prs, content_start + needed)



    # 각 본문 슬라이드에 콘텐츠 설정 (종료 슬라이드는 건드리지 않음)

    for page_i, page_units in enumerate(units_pages):

        slide = prs.slides[content_start + page_i]

        shape = _find_content_shape(slide)

        if shape:

            _set_reading_text(shape.text_frame, page_units, line_spacing=line_spacing)



    # 기록 후 실제 줄 수 검증 및 재조정 (비마지막 슬라이드 7·8·10줄 → 9줄)
    if needed > 1:
        _rebalance_reading_slides_post_write(prs, content_start, needed, label)

    # 종료 슬라이드의 본문 텍스트박스 비우기 (참조 PPT 잔여 내용 제거)

    for i in range(content_start + needed, content_start + needed + n_ending):

        shape = _find_content_shape(prs.slides[i])

        if shape:

            _clear_text_frame(shape.text_frame)



    total_new = needed + n_ending

    # 마지막 본문 슬라이드의 줄 수가 merge_threshold 이하이면 ending shape를 본문 슬라이드로 통합
    if n_ending > 0 and units_pages and needed > 0:
        last_page_lines = _page_visual_lines(units_pages[-1])
        if last_page_lines <= merge_threshold:
            ENDING_KW = ('주님의 말씀입니다', '◎ 하느님', '◎ 그리스도님')
            ending_slide_idx = content_start + needed
            ending_slide = prs.slides[ending_slide_idx]
            last_slide = prs.slides[content_start + needed - 1]
            spTree = last_slide.shapes._spTree
            for shape in ending_slide.shapes:
                if shape.has_text_frame and any(kw in shape.text_frame.text for kw in ENDING_KW):
                    spTree.append(copy.deepcopy(shape._element))
            for i in range(content_start + needed + n_ending - 1, content_start + needed - 1, -1):
                delete_slide(prs, i)
            total_new -= n_ending

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



    for start_key, end_key in [

        ('제1독서_start', '제1독서_end'),

        ('제2독서_start', '제2독서_end'),

        ('복음_start', '복음_end'),

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

            # 실제 본문 줄 수 계산 (빈 슬라이드는 건너뜀)

            line_count = 0

            for para in content_shape.text_frame.paragraphs:

                text = para.text.strip()

                if text:

                    line_count += _wrap_line_count(text)

            if line_count == 0:

                continue

            # ending shape 탐색 및 위치 재조정

            for shape in slide.shapes:

                if shape.has_text_frame and any(kw in shape.text_frame.text for kw in ENDING_KW):

                    line_height = content_shape.height // LINES_PER_SLIDE

                    shape.top = content_shape.top + (line_count + 2) * line_height

                    break



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

            break





# ─────────────────────────────────────────────────────────────────────────────

# 독서 제목 슬라이드 업데이트

# ─────────────────────────────────────────────────────────────────────────────



def update_reading_title_slide(prs, title_idx: int, label_kw: str, json_title: str):

    """

    독서 제목 슬라이드에서 성서 이름만 교체 (절 번호 제외).

    json_title 예: '열왕기 하권 4,8-11.14-16ㄴ' → '열왕기 하권의 말씀입니다.'

    """

    book_name = _extract_book_name(json_title)

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

                    # <a:br/> 이후 성서 이름만 교체

                    if _update_book_name_after_br(para, book_name):

                        return

                elif para.text.strip() and label_kw.replace(' ', '') not in para.text.replace(' ', ''):

                    # br 없이 별도 단락인 경우

                    _replace_para_text_clone(para, book_name + '의 말씀입니다.')

                    return





def update_복음_title_slide(prs, title_idx: int, json_title: str):

    """복음 제목 슬라이드에서 복음사가 이름+조사만 교체.

    json_title 예: '마태오 10,37-42' → '마태오가 전한 거룩한 복음입니다.'

    """

    evangelist = _extract_book_name(json_title)

    particle = _josa(evangelist)

    new_name_with_particle = evangelist + particle



    slide = prs.slides[title_idx]

    for shape in slide.shapes:

        if not shape.has_text_frame:

            continue

        for para in shape.text_frame.paragraphs:

            t = para.text.strip()

            m = re.match(r'^(.+?)(가|이)(\s*전한\s*거룩한\s*복음입니다.*)', t)

            if m:

                rest = m.group(3)  # " 전한 거룩한 복음입니다."

                new_text = new_name_with_particle + rest

                p = para._p

                runs = p.findall(qn('a:r'))

                if runs:

                    template_r = runs[0]

                    new_r = copy.deepcopy(template_r)

                    t_el = new_r.find(qn('a:t'))

                    if t_el is not None:

                        t_el.text = new_text

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

        if '화 답 송' in t or '화답송' in t:

            para = shape.text_frame.paragraphs[0]

            _replace_para_text_clone(para, f'화 답 송   {new_title}')

            return





def update_화답송(prs, json_data: dict, sections: dict, 화답송_pptx_path):

    if '화답송_start' not in sections or not json_data.get('화답송'):

        return



    화 = json_data['화답송']

    title = 화['title']

    content = 화['content']



    # \n으로 분리; 첫 번째 항목(◎ 후렴)은 악보 슬라이드에 포함되므로 제외

    segments = [s.strip() for s in content.split('\n') if s.strip()]

    verses = segments[1:] if len(segments) > 1 else segments

    n = len(verses)



    start = sections['화답송_start']

    cur_end = sections['화답송_end']

    cur_total = cur_end - start



    # 필요 슬라이드: 악보(n+1) + 텍스트(n) = 2n+1

    needed = max(1, 2 * n + 1)



    # 슬라이드 수 조정 (텍스트 템플릿 = start+1)

    text_tmpl = start + 1 if cur_total > 1 else start

    if needed > cur_total:

        for k in range(needed - cur_total):

            tmpl = start if (cur_total + k) % 2 == 0 else text_tmpl

            insert_slide_copy(prs, cur_end + k, tmpl)

    elif needed < cur_total:

        for i in range(cur_end - 1, start + needed - 1, -1):

            delete_slide(prs, i)



    # 화답송 악보 PPT 로드

    악보_prs = None

    if 화답송_pptx_path and Path(화답송_pptx_path).exists():

        try:

            악보_prs = Presentation(str(화답송_pptx_path))

        except Exception as e:

            print(f'  [경고] 화답송 악보 PPT 로드 실패: {e}')



    # 각 슬라이드 업데이트

    for i in range(needed):

        idx = start + i



        if i % 2 == 0:

            # 악보 슬라이드: 화답송 악보 PPT에서 복사

            if 악보_prs:

                delete_slide(prs, idx)

                copy_slide_from_prs(prs, idx, 악보_prs, 0)

            _update_화답송_title_in_slide(prs.slides[idx], title)

        else:

            # 텍스트 슬라이드: 절 내용 업데이트

            _update_화답송_title_in_slide(prs.slides[idx], title)

            verse_idx = i // 2

            if verse_idx < n:

                verse_text = verses[verse_idx]

                # 참조 PPT 서식: ○ 다음 공백을 탭으로 교체하여 탭 스톱 정렬 적용

                if verse_text.startswith('○ '):

                    verse_text = '○\t' + verse_text[2:]

                for shape in prs.slides[idx].shapes:

                    if not shape.has_text_frame:

                        continue

                    t = shape.text_frame.text.strip()

                    # '화 답 송' 타이틀 shape만 제외 ('화' 단독 필터는 '평화' 등 오탐)

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

                        _para_append_run(new_p, pptx_parse_xml(f'<a:r xmlns:a="{A_NS}"><a:t>{line}</a:t></a:r>'))

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

                        _para_append_run(new_p, pptx_parse_xml(f'<a:r xmlns:a="{A_NS}"><a:t>{text}</a:t></a:r>'))

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



def _set_slide_bg_black(slide, prs=None):

    """슬라이드 배경을 검정으로 설정.

    p:bg를 black으로 설정하고, 레이아웃 shape들을 가리는

    전체 슬라이드 크기의 검정 사각형을 spTree 첫 번째로 삽입한다.

    """

    from pptx.oxml import parse_xml as pptx_parse_xml

    P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'

    A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'



    # 1. p:bg 검정으로 설정

    BLACK_BG_XML = (

        f'<p:bg xmlns:p="{P_NS}" xmlns:a="{A_NS}">'

        f'<p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill>'

        f'<a:effectLst/></p:bgPr></p:bg>'

    )

    cSld = slide.element.find(qn('p:cSld'))

    if cSld is None:

        return

    existing = cSld.find(qn('p:bg'))

    if existing is not None:

        cSld.remove(existing)

    cSld.insert(0, pptx_parse_xml(BLACK_BG_XML))



    # 2. 전체 슬라이드를 덮는 검정 사각형을 spTree 최하단(첫 번째)에 삽입

    # 레이아웃 shape들이 위에서 배경을 덮는 경우 이 사각형이 그것을 가린다

    if prs is not None:

        cx = prs.slide_width

        cy = prs.slide_height

    else:

        cx, cy = 9144000, 5143500  # 와이드스크린 기본값 (EMU)



    BLACK_RECT_XML = (

        f'<p:sp xmlns:p="{P_NS}" xmlns:a="{A_NS}">'

        f'<p:nvSpPr>'

        f'<p:cNvPr id="9999" name="BlackBg"/>'

        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'

        f'<p:nvPr/>'

        f'</p:nvSpPr>'

        f'<p:spPr>'

        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'

        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'

        f'<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'

        f'<a:ln><a:noFill/></a:ln>'

        f'</p:spPr>'

        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'

        f'</p:sp>'

    )

    spTree = cSld.find(qn('p:spTree'))

    if spTree is not None:

        # nvGrpSpPr, grpSpPr 다음 첫 번째 shape 위치에 삽입

        insert_pos = 2  # nvGrpSpPr(0), grpSpPr(1) 뒤

        spTree.insert(insert_pos, pptx_parse_xml(BLACK_RECT_XML))





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





def _update_성가_header(slide, expected_type: str, new_number: int):

    """성가 헤더 타입/번호 업데이트. run 서식(색상 등) 보존."""

    MATCH_PAT = r'^(\s*)(입당|봉헌|성체|2차봉헌|파견)(\s+)(\d+)'

    for shape in slide.shapes:

        if not shape.has_text_frame:

            continue

        for para in shape.text_frame.paragraphs:

            t = para.text

            m = re.match(MATCH_PAT, t)

            if not m:

                continue

            old_prefix = m.group(0)

            new_prefix = m.group(1) + expected_type + m.group(3) + str(new_number)

            if old_prefix != new_prefix:

                _update_prefix_in_runs(para, old_prefix, new_prefix)

            return









def replace_성가(prs, 성가_map: dict, hymn_numbers: dict):

    """5종 성가 슬라이드 교체. 각 타입마다 find_sections 재호출로 정확한 위치 파악."""



    for htype in HYMN_TYPES:

        if htype not in 성가_map:

            print(f'  [{htype}] 성가 PPT 없음, 건너뜀')

            continue



        # 매번 섹션 재탐색 (이전 교체로 인한 인덱스 변화 반영)

        sec = find_sections(prs)



        num = hymn_numbers.get(htype)

        pptx_path = 성가_map[htype]



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



        # 새 성가 PPT 슬라이드 로드

        src_prs = Presentation(str(pptx_path))

        n_src = len(src_prs.slides)



        # 기존 콘텐츠 슬라이드 삭제

        for i in range(ce - 1, cs - 1, -1):

            delete_slide(prs, i)



        # 새 슬라이드 삽입

        for i in range(n_src):

            copy_slide_from_prs(prs, cs + i, src_prs, i)



        # 헤더 업데이트

        if num:

            for i in range(n_src):

                _update_성가_header(prs.slides[cs + i], htype, num)



        print(f'  [{htype}] 성가 {num}: {n_existing}장 → {n_src}장')





# ─────────────────────────────────────────────────────────────────────────────

# 검증

# ─────────────────────────────────────────────────────────────────────────────



def validate(prs, json_data: dict) -> bool:

    texts = all_slide_texts(prs)

    errors = []

    warnings = []



    liturgy = json_data.get('liturgy', '')

    if not any(liturgy in t for t in texts):

        errors.append(f'liturgy 없음: {liturgy}')



    for kw in ['입당송', '화답송', '영성체송']:

        if not any(kw in t for t in texts):

            warnings.append(f'{kw} 슬라이드 없음')



    # 독서 절 번호 오렌지 확인 (샘플)

    if json_data.get('제1독서'):

        found_orange = False

        for slide in prs.slides:

            for shape in slide.shapes:

                if not shape.has_text_frame:

                    continue

                for para in shape.text_frame.paragraphs:

                    for run in para.runs:

                        try:

                            if run.font.color.rgb == ORANGE:

                                found_orange = True

                                break

                        except Exception:

                            pass

                    if found_orange:

                        break

                if found_orange:

                    break

        if not found_orange:

            warnings.append('독서 절 번호 오렌지색 미확인')



    print('\n=== 검증 결과 ===')

    if errors:

        for e in errors:

            print(f'  [오류] {e}')

    if warnings:

        for w in warnings:

            print(f'  [경고] {w}')

    if not errors and not warnings:

        print('  모든 검증 통과!')



    return len(errors) == 0





# ─────────────────────────────────────────────────────────────────────────────

# 메인

# ─────────────────────────────────────────────────────────────────────────────



def _strip_slide_xml(content: bytes) -> bytes:

    """슬라이드 XML에서 호환성 문제 요소 제거 (strip_ppt2007_incompatible 내부용)."""

    # <a:ext ...><a14:imgProps ...>...</a14:imgProps></a:ext> 블록 제거

    # (python-pptx가 효과 이미지를 복사 못 해 r:embed가 단절됨)

    content = re.sub(

        rb'<a:ext uri="[^"]+"><a14:imgProps\b.*?</a14:imgProps>\s*</a:ext>',

        b'',

        content,

        flags=re.DOTALL,

    )

    # imgProps 제거 후 남는 빈 <a:extLst></a:extLst> 제거

    content = re.sub(rb'<a:extLst>\s*</a:extLst>', b'', content)

    # <a:videoFile r:link="..."/> 제거

    content = re.sub(rb'<a:videoFile[^/]*/>\s*', b'', content)

    # <a:hlinkClick ... action="ppaction://media" .../> 제거

    content = re.sub(

        rb'<a:hlinkClick[^>]+action="ppaction://media"[^>]*/>\s*',

        b'',

        content,

    )

    return content





def strip_ppt2007_incompatible(pptx_path: str) -> None:

    """PPT 2007에서 에러를 유발하는 요소 제거.



    제거 대상:

    - ppt/changesInfos/ (PowerPoint 2016 공동 작성 추적 파일)

    - [Content_Types].xml 의 changesInfo Override

    - ppt/_rels/presentation.xml.rels 의 changesInfo Relationship

    - 슬라이드 내 외부 비디오 링크 (MP4 등 — PPT 2007 미지원)

    - <a14:imgProps> 블록 및 제거 후 남는 빈 <a:extLst>

    - 슬라이드 rels에서 고아 image/hdphoto Relationship 제거

    - app.xml 슬라이드 카운트 수정

    """

    import zipfile as _zip



    path = Path(pptx_path)

    data = path.read_bytes()



    # 1패스: 슬라이드 XML 전처리 후 사용 중인 rId 수집 (고아 rel 제거용)

    slide_used_rids: dict = {}

    with _zip.ZipFile(io.BytesIO(data), 'r') as _z:

        actual_slide_count = sum(

            1 for n in _z.namelist()

            if re.match(r'^ppt/slides/slide\d+\.xml$', n)

        )

        for name in _z.namelist():

            if re.match(r'ppt/slides/slide\d+\.xml$', name):

                stripped = _strip_slide_xml(_z.read(name))

                used = set(re.findall(rb'r:[a-zA-Z]+=[\'"](rId\d+)[\'"]', stripped))

                slide_used_rids[name] = used



    buf = io.BytesIO()

    with _zip.ZipFile(io.BytesIO(data), 'r') as zin:

        with _zip.ZipFile(buf, 'w', _zip.ZIP_DEFLATED) as zout:

            for name in zin.namelist():

                if 'changesInfo' in name:

                    continue

                content = zin.read(name)



                if name == '[Content_Types].xml':

                    content = re.sub(

                        rb'<Override[^>]*changesInfo[^>]*/>\s*',

                        b'',

                        content,

                    )

                elif name == 'ppt/_rels/presentation.xml.rels':

                    content = re.sub(

                        rb'<Relationship[^>]*changesInfo[^>]*/>\s*',

                        b'',

                        content,

                    )

                elif name == 'docProps/app.xml':

                    # Slides 카운트를 실제 슬라이드 수로 수정

                    content = re.sub(

                        rb'<Slides>\d+</Slides>',

                        f'<Slides>{actual_slide_count}</Slides>'.encode(),

                        content,

                    )

                    # 비디오 제거 후 MMClips도 0으로

                    content = re.sub(rb'<MMClips>\d+</MMClips>', b'<MMClips>0</MMClips>', content)

                elif re.match(r'ppt/slides/_rels/slide\d+\.xml\.rels', name):

                    # 슬라이드 rels에서 외부 비디오 Relationship 제거

                    content = re.sub(

                        rb'<Relationship[^>]+Type="[^"]*relationships/video[^"]*"[^>]*/>\s*',

                        b'',

                        content,

                    )

                    # 고아 image/hdphoto Relationship 제거

                    # (shape는 삭제됐으나 rel 항목이 남은 경우)

                    slide_name = name.replace('_rels/', '').replace('.rels', '')

                    used_rids = slide_used_rids.get(slide_name, set())



                    def _drop_orphan(m, _used=used_rids):

                        attrs = m.group(0)

                        rid_m = re.search(rb'Id="(rId\d+)"', attrs)

                        type_m = re.search(rb'Type="([^"]+)"', attrs)

                        if not rid_m or not type_m:

                            return attrs

                        rid = rid_m.group(1)

                        rtype = type_m.group(1)

                        if (b'image' in rtype or b'hdphoto' in rtype) and rid not in _used:

                            return b''

                        return attrs



                    content = re.sub(rb'<Relationship\s[^>]*/>', _drop_orphan, content)

                elif re.match(r'ppt/slides/slide\d+\.xml', name):

                    content = _strip_slide_xml(content)



                zout.writestr(name, content)



    path.write_bytes(buf.getvalue())





def main():

    if len(sys.argv) == 1:

        if _preloaded_inputs[0] is not None:

            # EXE GUI 흐름: 팝업에서 미리 수집한 값 사용

            _inp = _preloaded_inputs[0]

            _preloaded_inputs[0] = None

            date_str    = _inp['date_str']

            files       = _inp['files']

            hymn_numbers = _inp['hymn_numbers']

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

            Path(date_str).mkdir(exist_ok=True)

        else:

            # 인수 없이 실행: 대화형 모드 (팝업 순서: 날짜 → 파일 선택 → 성가번호)

            date_str = _ask_date_popup()

            files = _ask_input_files_popup()

            hymn_numbers = _ask_numbers_popup({})

            onedrive_folder = get_onedrive_hymn_folder()

            onedrive_pptxs = [f for f in onedrive_folder.rglob('*.pptx') if not f.name.startswith('~$')]

            for htype, num in hymn_numbers.items():

                if num is None:

                    continue

                for f in onedrive_pptxs:

                    if re.search(rf'성가 {num}(?!\d)', f.name):

                        files['성가'][htype] = f

                        break

            Path(date_str).mkdir(exist_ok=True)

    else:

        date_str, hymn_numbers, 화답송_override = parse_args()

        files = find_files(date_str, hymn_numbers)

        if 화답송_override:

            files['화답송_pptx'] = Path(화답송_override)



    print(f'날짜: {date_str}')

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

    print(f'  화답송 악보: {files["화답송_pptx"].name if files["화답송_pptx"] else "없음"}')

    for k, v in files['성가'].items():

        print(f'  성가({k}): {v.name}')



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

        update_화답송(prs, json_data, sec, files.get('화답송_pptx'))



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

    replace_성가(prs, files['성가'], hymn_numbers)



    # 7. 저장

    _report_progress(95, '파일 저장 중...')

    liturgy_safe = re.sub(r'[\\/:*?"<>|]', '_', json_data['liturgy'])

    output_name = f"{date_str}_{liturgy_safe}.pptx"

    output_path = Path(date_str) / output_name

    print(f'\n[7] 저장: {output_path}')

    prs.save(str(output_path))



    # 8. 검증

    prs2 = Presentation(str(output_path))

    validate(prs2, json_data)



    _last_output_path[0] = output_path

    print(f'\n완료! → {output_path}')





def _show_result_window(title: str, text: str, is_error: bool = False) -> None:

    import tkinter as tk

    from tkinter import scrolledtext as _st_mod

    import subprocess as _sp

    from datetime import datetime as _dt



    # 로그 저장

    date_str = _last_date_str[0]

    output_path = _last_output_path[0]

    if date_str and not is_error:

        log_dir = Path(date_str) / 'log'

        try:

            log_dir.mkdir(parents=True, exist_ok=True)

            ts = _dt.now().strftime('%H%M%S')

            log_file = log_dir / f'{date_str}_{ts}.txt'

            log_file.write_text(text.strip(), encoding='utf-8')

        except Exception:

            pass



    root = tk.Tk()

    root.withdraw()

    root.title(title)

    root.minsize(750, 480)

    root.resizable(True, True)

    _set_window_icon(root)

    _apply_theme(root)



    frame = tk.Frame(root)

    frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(12, 4))



    bg_color = '#fff0f0' if is_error else '#f9f9f9'

    text_widget = _st_mod.ScrolledText(

        frame, wrap=tk.WORD, font=('Consolas', 11), bg=bg_color, fg=_UI['fg']

    )

    text_widget.pack(fill=tk.BOTH, expand=True)

    text_widget.insert(tk.END, text.strip() if text.strip() else '(출력 없음)')

    text_widget.configure(state='disabled')



    btn_frame = tk.Frame(root)

    btn_frame.pack(pady=8)



    if output_path and not is_error:

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



    btn_color = '#c0392b' if is_error else _UI['primary']

    tk.Button(

        btn_frame, text='닫기', command=root.destroy,

        bg=btn_color, fg=_UI['fg_light'],

        font=(_UI['font'], _UI['font_sz'], 'bold'), width=14, relief='flat', cursor='hand2'

    ).pack(side=tk.LEFT, padx=6)

    _center_window(root)

    root.deiconify()

    root.mainloop()



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

            _hymn_numbers = _ask_numbers_popup({})

        except RuntimeError:

            sys.exit(0)



        _preloaded_inputs[0] = {

            'date_str':     _date_str,

            'files':        _files,

            'hymn_numbers': _hymn_numbers,

        }



        _out_text, _err_text = _run_with_progress_window(main)



        if _err_text:

            _show_result_window('오류 발생', (_out_text or '') + '\n\n' + _err_text, is_error=True)

        else:

            _show_result_window('완료', _out_text or '')

