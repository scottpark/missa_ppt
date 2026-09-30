"""미사 유형별 실행 바로가기(.lnk) 3종 생성.

§J9(2026-09-26) — `missa_to_ppt.exe`를 미사 유형마다 다른 아이콘/인자로 바로 실행할 수
있는 바로가기를 만든다. `--미사유형`(§J3, CLI 옵션)만 주고 date를 생략하면 곧바로 해당
미사 유형의 GUI 입력 흐름으로 들어가므로(`_dispatch_cli_or_gui()`), 바로가기 하나당
`Arguments='--미사유형 성인'`처럼 인자 하나만 다르면 된다.

바로가기의 "시작 위치"(WorkingDirectory)를 반드시 exe가 있는 폴더로 맞춘다 —
`build_dist_zip.py`가 설명하듯 exe는 "실행 파일 옆의 외부 assets/ 폴더"를 찾도록
설계돼 있어서, 바탕화면 등 다른 폴더에 바로가기를 둬도 실행 시 상대 경로 탐색이 exe
폴더 기준으로 이뤄지게 해야 한다.

사용법: python tools/build_shortcuts.py
(사전에 `pyinstaller missa_to_ppt.spec` 빌드 + `python tools/build_missa_icon.py`로
미사유형별 아이콘 생성 필요)
"""
from pathlib import Path

import win32com.client

BASE = Path(__file__).resolve().parent.parent
EXE = BASE / "missa_to_ppt.exe"
ASSETS_DIR = BASE / "assets"

# (미사유형 CLI 값, 바로가기 파일명, 아이콘 파일명)
SHORTCUTS = [
    ("성인", "성인미사 PPT.lnk", "missa_to_ppt_성인.ico"),
    ("청년", "청년미사 PPT.lnk", "missa_to_ppt_청년.ico"),
    ("어린이", "어린이미사 PPT.lnk", "missa_to_ppt_어린이.ico"),
]


def build(dest_dir: Path | None = None) -> list[Path]:
    """`dest_dir`(기본: 프로젝트 루트)에 바로가기 3개를 만들고 생성된 경로 목록을
    반환한다. exe/아이콘이 없으면 어느 것이 없는지 명시한 FileNotFoundError를 낸다 —
    빌드 순서(exe 빌드 → 아이콘 생성 → 바로가기 생성)를 지키지 않은 실수를 바로
    드러내기 위함이다."""
    if not EXE.exists():
        raise FileNotFoundError(
            f"{EXE} 없음 — 먼저 `pyinstaller missa_to_ppt.spec`으로 빌드하세요."
        )

    dest_dir = dest_dir or BASE
    dest_dir.mkdir(parents=True, exist_ok=True)

    shell = win32com.client.Dispatch("WScript.Shell")
    created = []
    for mass_type_kr, lnk_name, icon_name in SHORTCUTS:
        icon_path = ASSETS_DIR / icon_name
        if not icon_path.exists():
            raise FileNotFoundError(
                f"{icon_path} 없음 — 먼저 `python tools/build_missa_icon.py`로 "
                "미사유형별 아이콘을 생성하세요."
            )

        lnk_path = dest_dir / lnk_name
        shortcut = shell.CreateShortcut(str(lnk_path))
        shortcut.TargetPath = str(EXE)
        shortcut.Arguments = f"--미사유형 {mass_type_kr}"
        shortcut.WorkingDirectory = str(BASE)
        shortcut.IconLocation = str(icon_path)
        shortcut.Description = f"{mass_type_kr}미사 PPT 자동 생성"
        shortcut.save()
        created.append(lnk_path)

    return created


if __name__ == "__main__":
    paths = build()
    print(f"바로가기 {len(paths)}개 생성 완료:")
    for p in paths:
        print(f"  {p}")
