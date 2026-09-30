"""청년미사 Python 소스 배포용 zip 생성.

청년미사는 아직 exe로 패키징돼 있지 않다(OneDrive 연동이 활발히 바뀌는 중이라, exe를 매번
다시 빌드·재배포하는 대신 운영자가 Python을 설치해 소스를 직접 실행하는 방식을 택함 —
`tools/build_dist_zip.py`는 exe 전용이라 이 용도로 재사용할 수 없다). 이 스크립트는 운영자
PC에서 그대로 실행 가능한 최소 소스 배포판을 만든다.

포함: 런타임에 필요한 .py 전부 + assets/(tessdata/ 제외) + requirements.txt + install.bat +
run_missa_청년.bat + create_shortcut.py("청년미사 PPT" 바로가기 설치 시 생성) + VERSION +
README_청년미사_운영자.md + 나주노/야훼이레 성가집 원본 PDF 2개(cache/청년미사_성가원본/,
2026-09-27부터 — `missa_youth_hymn_pdf._resolve_pdf_path()`가 더 이상 OneDrive 폴백 없이
이 경로 하나만 보므로, 배포 시점에 원본을 통째로 함께 담아야 한다). 나주노 PDF가 스캔
래스터라 약 80MB — 이 zip은 더 이상 "가벼운" 배포판이 아니다(의도적 트레이드오프: 곡 단위
OneDrive 조회 없이 항상 동작).

제외(의도적):
  - config.json, msal_token_cache.bin — PC/계정별 개인화 상태, 첫 실행 시 자동 생성/로그인됨
  - assets/tessdata/ — 개발용 Tesseract OCR 데이터, 런타임은 캐시 json만 읽음(Tesseract 미호출)
  - output/, docs/, _workspace/*, .claude/, .git/, test_*.py, tools/ — 운영자 실행에 불필요한
    개발/테스트/산출물 자산

사용법: python tools/build_youth_dist_zip.py
"""
import sys
import zipfile
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE))
from missa_youth_hymn_pdf import SOURCES as YOUTH_HYMN_PDF_SOURCES  # noqa: E402  (파일명 단일 출처)

DEST_ZIP = BASE / "dist" / "missa_ppt_청년미사.zip"
HYMN_PDF_PATHS = [Path(info["pdf"]) for info in YOUTH_HYMN_PDF_SOURCES.values()]

RUNTIME_PY_FILES = [
    "missa_to_ppt.py",
    "missa_to_json.py",
    "missa_ooxml_utils.py",
    "missa_gui.py",
    "missa_reading_layout.py",
    "missa_sections.py",
    "missa_content_updaters.py",
    "missa_psalm_score_image.py",
    "missa_youth_gospel.py",
    "missa_youth_hymn_pdf.py",
    "missa_onedrive.py",
    "missa_updater.py",
    "ppt_com_verify.py",
]

TOP_LEVEL_FILES = [
    "requirements.txt",
    "install.bat",
    "run_missa_청년.bat",
    "create_shortcut.py",
    "VERSION",
    "README_청년미사_운영자.md",
]

ASSETS_DIR = BASE / "assets"
ASSETS_EXCLUDE_DIRS = {"tessdata"}


def build() -> Path:
    missing = [f for f in RUNTIME_PY_FILES + TOP_LEVEL_FILES if not (BASE / f).exists()]
    if missing:
        raise FileNotFoundError(f"다음 파일이 없습니다: {missing}")
    if not ASSETS_DIR.exists():
        raise FileNotFoundError(f"{ASSETS_DIR} 없음")
    missing_pdfs = [p for p in HYMN_PDF_PATHS if not p.exists()]
    if missing_pdfs:
        raise FileNotFoundError(f"원본 PDF가 없습니다: {missing_pdfs}")

    asset_files = sorted(
        p for p in ASSETS_DIR.rglob("*")
        if p.is_file() and p.relative_to(ASSETS_DIR).parts[0] not in ASSETS_EXCLUDE_DIRS
    )

    with zipfile.ZipFile(DEST_ZIP, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in RUNTIME_PY_FILES + TOP_LEVEL_FILES:
            zf.write(BASE / name, name)
        for path in asset_files:
            zf.write(path, str(Path("assets") / path.relative_to(ASSETS_DIR)))
        for pdf_path in HYMN_PDF_PATHS:
            zf.write(pdf_path, str(Path("cache") / "청년미사_성가원본" / pdf_path.name))

    return DEST_ZIP


if __name__ == "__main__":
    out = build()
    size_mb = out.stat().st_size / (1024 * 1024)
    print(f"배포 zip 생성 완료: {out} ({size_mb:.1f} MB)")
    print(f"  포함: .py {len(RUNTIME_PY_FILES)}개, {TOP_LEVEL_FILES}, assets/(tessdata/ 제외), "
          f"성가집 원본 PDF {len(HYMN_PDF_PATHS)}개(cache/청년미사_성가원본/)")
    print("  미포함(의도적): config.json, msal_token_cache.bin (PC/계정별 개인화 상태)")
