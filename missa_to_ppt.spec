# -*- mode: python ; coding: utf-8 -*-
import PyInstaller.config
PyInstaller.config.CONF['distpath'] = SPECPATH

# 화답송 악보 템플릿(assets/화답송_악보_template.pptx)은 PyInstaller onefile datas로
# 묻지 않는다 — 실행 시점에 디스크로 추출되지 않는 문제가 실측 확인됨(바이너리/데이터
# 재분류 단계의 부작용으로 추정). config.json과 동일하게 exe 옆의 외부 assets/ 폴더로
# 배포한다(missa_psalm_score_image.py의 frozen 분기 참고). 그래서 datas는 비워둔다.

a = Analysis(
    ['missa_to_ppt.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'missa_to_json', 'bs4', 'lxml',
        'ppt_com_verify',
        'missa_psalm_score_image', 'numpy', 'PIL',
        'win32com.client', 'win32com.gen_py',
        'win32timezone', 'pythoncom', 'pywintypes', 'win32api',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='missa_to_ppt',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/missa_to_ppt.ico',
)
