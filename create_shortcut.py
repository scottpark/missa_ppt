"""미사 PPT 배포판 설치 스크립트 — install.bat이 호출한다.

이 파일이 있는 폴더를 기준으로 바로가기(.lnk)를 만든다. 어느 미사용 배포판인지는 폴더에 있는 실행 파일
(`run_missa_성인.bat`/`run_missa_청년.bat`)로 판단한다 — 성인·청년 zip이 이 스크립트를 공유하고 zip마다
해당 유형의 .bat 하나만 들어 있다. 소스 zip 배포(exe 없음)라 pythonw.exe + missa_to_ppt.py를 타겟으로 한다.

.lnk는 실행한 PC의 실제 설치 경로를 그대로 담으므로(어디에 압축을 풀었는지는 PC마다 다르다) 배포 zip에 미리
만들어 넣지 않고, 설치 시점(install.bat 실행 시)에 매번 새로 생성한다 — config.json/msal_token_cache.bin이
배포에서 빠지는 것과 같은 이유."""
from pathlib import Path

import win32com.client

BASE = Path(__file__).resolve().parent

# (실행 .bat, 미사유형 CLI 값, 바로가기 파일명, 아이콘 파일명, 설명)
KINDS = [
    ("run_missa_성인.bat", "성인", "성인미사 PPT.lnk", "missa_to_ppt_성인.ico", "성인미사 PPT 자동 생성"),
    ("run_missa_청년.bat", "청년", "청년미사 PPT.lnk", "missa_to_ppt_청년.ico", "청년미사 PPT 자동 생성"),
]


def build() -> list:
    script_path = BASE / "missa_to_ppt.py"
    shell = win32com.client.Dispatch("WScript.Shell")
    made = []
    for bat, mass_type, lnk_name, icon_name, desc in KINDS:
        if not (BASE / bat).exists():
            continue
        lnk_path = BASE / lnk_name
        icon_path = BASE / "assets" / icon_name
        shortcut = shell.CreateShortcut(str(lnk_path))
        shortcut.TargetPath = "pythonw.exe"
        shortcut.Arguments = f'"{script_path}" --미사유형 {mass_type}'
        shortcut.WorkingDirectory = str(BASE)
        if icon_path.exists():
            shortcut.IconLocation = str(icon_path)
        shortcut.Description = desc
        shortcut.save()
        made.append(lnk_path)
    return made


if __name__ == "__main__":
    for path in build():
        print(f"바로가기 생성 완료: {path}")
