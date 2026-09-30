"""청년미사 PPT 배포판 설치 스크립트 — install.bat이 호출한다.

이 파일이 있는 폴더를 기준으로 "청년미사 PPT.lnk" 바로가기를 만든다(README_청년미사_
운영자.md가 안내하는 그 바로가기). 소스 zip 배포(exe 없음)라 tools/build_shortcuts.py
(exe 타겟)를 재사용할 수 없어, pythonw.exe + missa_to_ppt.py를 타겟으로 하는 이 전용
스크립트를 둔다.

.lnk는 실행한 PC의 실제 설치 경로를 그대로 담으므로(어디에 압축을 풀었는지는 PC마다
다르다) 배포 zip에 미리 만들어 넣지 않고, 설치 시점(install.bat 실행 시)에 매번 새로
생성한다 — config.json/msal_token_cache.bin이 배포에서 빠지는 것과 같은 이유."""
from pathlib import Path

import win32com.client

BASE = Path(__file__).resolve().parent


def build() -> Path:
    icon_path = BASE / "assets" / "missa_to_ppt_청년.ico"
    script_path = BASE / "missa_to_ppt.py"
    lnk_path = BASE / "청년미사 PPT.lnk"

    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(str(lnk_path))
    shortcut.TargetPath = "pythonw.exe"
    shortcut.Arguments = f'"{script_path}" --미사유형 청년'
    shortcut.WorkingDirectory = str(BASE)
    if icon_path.exists():
        shortcut.IconLocation = str(icon_path)
    shortcut.Description = "청년미사 PPT 자동 생성"
    shortcut.save()
    return lnk_path


if __name__ == "__main__":
    path = build()
    print(f"바로가기 생성 완료: {path}")
