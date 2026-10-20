"""미사 PPT 배포판 설치 스크립트 — install.bat이 호출한다.

이 파일이 있는 폴더를 기준으로 바로가기(.lnk)를 만든다. 어느 미사용 배포판인지는 폴더에 있는 실행 파일
(`run_missa_성인.bat`/`run_missa_청년.bat`)로 판단한다 — 성인·청년 zip이 이 스크립트를 공유하고 zip마다
해당 유형의 .bat 하나만 들어 있다. 소스 zip 배포(exe 없음)라 pythonw.exe + missa_to_ppt.py를 타겟으로 한다.

바로가기는 프로그램 폴더와 **바탕화면** 두 곳에 만든다(운영자가 바탕화면에서 찾는 경우가 많다). 바탕화면 경로는
OneDrive 백업으로 옮겨진 PC도 맞게 Windows에 물어 얻는다. pywin32가 없거나 실패하면 PowerShell로 다시 시도하고,
그래도 안 되면 원인을 화면에 출력한다(조용히 실패하지 않는다).

.lnk는 실행한 PC의 실제 설치 경로를 그대로 담으므로(어디에 압축을 풀었는지는 PC마다 다르다) 배포 zip에 미리
만들어 넣지 않고, 설치 시점(install.bat 실행 시)에 매번 새로 생성한다 — config.json이 배포에서 빠지는 것과 같은 이유."""
import subprocess
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent

# (실행 .bat, 미사유형 CLI 값, 바로가기 파일명, 아이콘 파일명, 설명)
KINDS = [
    ("run_missa_성인.bat", "성인", "성인미사 PPT.lnk", "missa_to_ppt_성인.ico", "성인미사 PPT 자동 생성"),
    ("run_missa_청년.bat", "청년", "청년미사 PPT.lnk", "missa_to_ppt_청년.ico", "청년미사 PPT 자동 생성"),
]


def _pythonw() -> str:
    """바로가기가 실행할 pythonw.exe. 설치된 Python의 전체 경로를 쓴다(PATH에만 의존하지 않는다)."""
    candidate = Path(sys.executable).with_name("pythonw.exe")
    return str(candidate) if candidate.exists() else "pythonw.exe"


def _desktop_via_com():
    import win32com.client
    shell = win32com.client.Dispatch("WScript.Shell")
    return Path(shell.SpecialFolders("Desktop"))


def _make_via_com(lnk_path: Path, target: str, arguments: str, workdir: str, icon: str, desc: str) -> None:
    import win32com.client
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(str(lnk_path))
    shortcut.TargetPath = target
    shortcut.Arguments = arguments
    shortcut.WorkingDirectory = workdir
    if icon:
        shortcut.IconLocation = icon
    shortcut.Description = desc
    shortcut.save()


def _ps_quote(text: str) -> str:
    return "'" + str(text).replace("'", "''") + "'"


def _make_via_powershell(lnk_path: Path, target: str, arguments: str, workdir: str, icon: str, desc: str) -> None:
    script = (
        "$s=(New-Object -ComObject WScript.Shell).CreateShortcut(%s);"
        "$s.TargetPath=%s;$s.Arguments=%s;$s.WorkingDirectory=%s;%s$s.Description=%s;$s.Save()"
    ) % (_ps_quote(lnk_path), _ps_quote(target), _ps_quote(arguments), _ps_quote(workdir),
         ("$s.IconLocation=%s;" % _ps_quote(icon)) if icon else "", _ps_quote(desc))
    subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
                   check=True, capture_output=True, timeout=60)


def _desktop_via_powershell() -> Path:
    out = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", "[Environment]::GetFolderPath('Desktop')"],
        check=True, capture_output=True, timeout=60)
    return Path(out.stdout.decode("utf-8", "replace").strip())


def _desktop_dir():
    for finder in (_desktop_via_com, _desktop_via_powershell):
        try:
            desktop = finder()
            if desktop.is_dir():
                return desktop
        except Exception:
            continue
    fallback = Path.home() / "Desktop"
    return fallback if fallback.is_dir() else None


def _make_lnk(lnk_path: Path, target: str, arguments: str, workdir: str, icon: str, desc: str) -> None:
    try:
        _make_via_com(lnk_path, target, arguments, workdir, icon, desc)
    except Exception as com_error:
        print(f"  (pywin32 방식 실패: {com_error} — PowerShell로 다시 시도합니다)")
        _make_via_powershell(lnk_path, target, arguments, workdir, icon, desc)
    if not lnk_path.exists():
        raise RuntimeError(f"바로가기 파일이 만들어지지 않았습니다: {lnk_path}")


def build() -> list:
    script_path = BASE / "missa_to_ppt.py"
    desktop = _desktop_dir()
    made = []
    for bat, mass_type, lnk_name, icon_name, desc in KINDS:
        if not (BASE / bat).exists():
            continue
        icon_path = BASE / "assets" / icon_name
        icon = str(icon_path) if icon_path.exists() else ""
        arguments = f'"{script_path}" --미사유형 {mass_type}'
        for folder in (BASE, desktop):
            if folder is None:
                continue
            lnk_path = folder / lnk_name
            _make_lnk(lnk_path, _pythonw(), arguments, str(BASE), icon, desc)
            made.append(lnk_path)
    return made


if __name__ == "__main__":
    try:
        paths = build()
    except Exception as e:
        print(f"[오류] 바로가기 생성 실패: {type(e).__name__}: {e}")
        sys.exit(1)
    if not paths:
        print("[오류] 실행 파일(run_missa_성인.bat / run_missa_청년.bat)을 찾지 못해 바로가기를 만들지 못했습니다.")
        sys.exit(1)
    for path in paths:
        print(f"바로가기 생성 완료: {path}")
