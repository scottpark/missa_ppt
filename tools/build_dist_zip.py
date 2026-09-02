"""missa_to_ppt.exe 배포용 zip 생성.

exe는 assets/화답송_악보_template.pptx를 PyInstaller onefile로 묻지 않고 "실행 파일 옆의
외부 assets/ 폴더"에서 찾도록 설계되어 있다(missa_psalm_score_image.py, missa_to_ppt.spec
참고 — onefile datas로 묻으면 실행 시점에 디스크로 추출되지 않는 문제가 실측 확인됨). exe만
단독으로 배포하면 화답송 악보 이미지 처리가 "Package not found" 에러로 실패한다.

config.json은 배포에 포함하지 않는다 — onedrive_hymn_folder가 사용자 PC마다 다른 개인화
설정이라, 없으면 GUI가 첫 실행 시 입력창을 띄워 자동 생성한다(missa_gui.py).

사용법: python tools/build_dist_zip.py  (사전에 `pyinstaller missa_to_ppt.spec` 빌드 필요)
"""
import sys
import zipfile
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE))
from missa_psalm_score_image import ASSET_TEMPLATE  # noqa: E402  (assets 경로 단일 출처)

EXE = BASE / "missa_to_ppt.exe"
ASSETS_DIR = BASE / "assets"
DEST_ZIP = BASE / "missa_to_ppt.zip"


def build() -> Path:
    if not EXE.exists():
        raise FileNotFoundError(
            f"{EXE} 없음 — 먼저 `pyinstaller missa_to_ppt.spec`으로 빌드하세요."
        )
    if not ASSET_TEMPLATE.exists():
        raise FileNotFoundError(
            f"{ASSET_TEMPLATE} 없음 — `python tools/build_화답송_template.py`로 먼저 생성하세요."
        )

    asset_files = sorted(p for p in ASSETS_DIR.rglob("*") if p.is_file())
    with zipfile.ZipFile(DEST_ZIP, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(EXE, EXE.name)
        for path in asset_files:
            zf.write(path, str(Path("assets") / path.relative_to(ASSETS_DIR)))

    return DEST_ZIP


if __name__ == "__main__":
    out = build()
    size_mb = out.stat().st_size / (1024 * 1024)
    print(f"배포 zip 생성 완료: {out} ({size_mb:.1f} MB)")
    print(f"  포함: {EXE.name}, assets/ 전체({ASSETS_DIR})")
    print("  미포함(의도적): config.json — 사용자 PC마다 다른 설정이라 첫 실행 시 자동 생성됨")
