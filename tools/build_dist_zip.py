"""성인·청년 미사 Python 소스 배포 zip 생성.

2026-10-05부터 성인미사도 청년미사와 같은 방식(Python 소스 zip + `.bat` 런처)으로 배포한다 — PyInstaller
exe는 폐기(코드가 바뀔 때마다 다시 빌드·재배포해야 하고, `.py`만 교체하는 자동 업데이트(`missa_updater.py`)가
exe에는 먹히지 않기 때문). 이 스크립트가 운영자 PC에서 그대로 실행 가능한 최소 소스 배포판을 만든다.

공통 포함: 런타임에 필요한 .py 13개 + assets/(tessdata/ 제외) + requirements.txt + install.bat +
create_shortcut.py(설치 시 바로가기 생성) + VERSION + 유형별 실행 .bat·운영자 안내서(`docs/README_*미사_운영자.md`를 zip 최상위로).
청년만 추가: 나주노/야훼이레 성가집 원본 PDF 2개(cache/청년미사_성가원본/, 나주노가 스캔 래스터라 약 80MB —
`missa_youth_hymn_pdf._resolve_pdf_path()`가 이 경로 하나만 본다). 성인 zip은 PDF가 없어 가볍다.

제외(의도적):
  - config.json, msal_token_cache.bin — PC/계정별 개인화 상태, 첫 실행 시 자동 생성/로그인됨
  - assets/tessdata/ — 개발용 Tesseract OCR 데이터, 런타임은 캐시 json만 읽음(Tesseract 미호출)
  - output/, docs/, _workspace/*, .claude/, .git/, test_*.py, tools/ — 개발/테스트/산출물 자산

사용법: python tools/build_dist_zip.py [성인|청년|all]   (기본 all)
"""
import sys
import zipfile
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE))
from missa_youth_hymn_pdf import SOURCES as YOUTH_HYMN_PDF_SOURCES  # noqa: E402  (파일명 단일 출처)

DIST_DIR = BASE / "dist"
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

COMMON_TOP_LEVEL_FILES = [
    "requirements.txt",
    "install.bat",
    "create_shortcut.py",
    "VERSION",
]

# 미사 유형별 구성: zip 파일명, 실행 .bat, 운영자 안내서(저장소에선 docs/ 아래, zip에는 최상위로 넣는다 — 운영자가
# 압축을 풀면 바로 보이게), 추가로 담을 파일(원본 경로 → zip 내 경로)
KINDS = {
    "성인": {
        "zip": "missa_ppt_성인미사.zip",
        "files": ["run_missa_성인.bat"],
        "readme": "docs/README_성인미사_운영자.md",
        "extra": [],
    },
    "청년": {
        "zip": "missa_ppt_청년미사.zip",
        "files": ["run_missa_청년.bat"],
        "readme": "docs/README_청년미사_운영자.md",
        "extra": [(p, Path("cache") / "청년미사_성가원본" / p.name) for p in HYMN_PDF_PATHS],
    },
}

ASSETS_DIR = BASE / "assets"
ASSETS_EXCLUDE_DIRS = {"tessdata"}


def _write_file(zf: zipfile.ZipFile, src: Path, arcname: str) -> None:
    """파일을 zip에 넣는다. `.bat`은 항상 CRLF로 정규화해서 넣는다 — LF만 있는 배치 파일은 한글(`chcp 65001`)과 함께
    쓰면 cmd가 줄 경계를 잘못 읽어 문장 조각을 명령어로 실행한다(2026-10-05 성인 install.bat 오류 'issa_ppt is not
    recognized…'). 작업 트리의 줄바꿈이 무엇이든(autocrlf·편집기) 배포물은 CRLF가 되게 한다."""
    if src.suffix.lower() == ".bat":
        data = src.read_bytes().replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")
        zf.writestr(arcname, data)
    else:
        zf.write(src, arcname)


def build(kind: str) -> Path:
    spec = KINDS[kind]
    top_files = COMMON_TOP_LEVEL_FILES + spec["files"]
    missing = [f for f in RUNTIME_PY_FILES + top_files + [spec["readme"]] if not (BASE / f).exists()]
    missing += [str(src) for src, _dst in spec["extra"] if not src.exists()]
    if missing:
        raise FileNotFoundError(f"다음 파일이 없습니다: {missing}")
    if not ASSETS_DIR.exists():
        raise FileNotFoundError(f"{ASSETS_DIR} 없음")

    asset_files = sorted(
        p for p in ASSETS_DIR.rglob("*")
        if p.is_file() and p.relative_to(ASSETS_DIR).parts[0] not in ASSETS_EXCLUDE_DIRS
    )

    DIST_DIR.mkdir(exist_ok=True)
    dest = DIST_DIR / spec["zip"]
    with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in RUNTIME_PY_FILES + top_files:
            _write_file(zf, BASE / name, name)
        _write_file(zf, BASE / spec["readme"], Path(spec["readme"]).name)
        for path in asset_files:
            _write_file(zf, path, str(Path("assets") / path.relative_to(ASSETS_DIR)))
        for src, dst in spec["extra"]:
            _write_file(zf, src, str(dst))
    return dest


if __name__ == "__main__":
    arg = sys.argv[1] if len(sys.argv) > 1 else "all"
    kinds = list(KINDS) if arg == "all" else [arg]
    unknown = [k for k in kinds if k not in KINDS]
    if unknown:
        raise SystemExit(f"알 수 없는 유형: {unknown} — 성인|청년|all 중 하나")
    for k in kinds:
        out = build(k)
        size_mb = out.stat().st_size / (1024 * 1024)
        print(f"[{k}] 배포 zip 생성 완료: {out} ({size_mb:.1f} MB)")
    print("  미포함(의도적): config.json, msal_token_cache.bin (PC/계정별 개인화 상태)")
