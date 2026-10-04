"""배포 zip 빌더(tools/build_dist_zip.py) — 배치 파일 CRLF 정규화와 유형별 구성."""
import sys
import zipfile
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))
import build_dist_zip as bd  # noqa: E402


def test_bat_files_are_written_with_crlf_even_when_source_is_lf(tmp_path):
    """LF만 있는 .bat는 한글+chcp 65001에서 cmd가 줄 경계를 잘못 읽는다 — zip에는 항상 CRLF로 넣는다."""
    src = tmp_path / "x.bat"
    src.write_bytes("@echo off\nchcp 65001 >nul\necho 한글\n".encode("utf-8"))
    out = tmp_path / "o.zip"
    with zipfile.ZipFile(out, "w") as zf:
        bd._write_file(zf, src, "x.bat")
    data = zipfile.ZipFile(out).read("x.bat")
    assert data.count(b"\r\n") == 3 and b"\n" not in data.replace(b"\r\n", b"")


def test_bat_already_crlf_is_not_doubled(tmp_path):
    src = tmp_path / "y.bat"
    src.write_bytes(b"@echo off\r\necho a\r\n")
    out = tmp_path / "o.zip"
    with zipfile.ZipFile(out, "w") as zf:
        bd._write_file(zf, src, "y.bat")
    assert zipfile.ZipFile(out).read("y.bat") == b"@echo off\r\necho a\r\n"


@pytest.mark.parametrize("kind,must,must_not", [
    ("성인", {"run_missa_성인.bat", "README_성인미사_운영자.md", "install.bat", "create_shortcut.py"}, {"cache/"}),
    ("청년", {"run_missa_청년.bat", "README_청년미사_운영자.md", "install.bat", "create_shortcut.py"}, set()),
])
def test_zip_contents_per_kind(kind, must, must_not, tmp_path, monkeypatch):
    monkeypatch.setattr(bd, "DIST_DIR", tmp_path)
    if kind == "청년" and not all(p.exists() for p in bd.HYMN_PDF_PATHS):
        pytest.skip("성가집 원본 PDF 없음(cache/, git 미포함)")
    out = bd.build(kind)
    names = set(zipfile.ZipFile(out).namelist())
    assert must <= names
    assert not any(n.startswith(p) for n in names for p in must_not)
    # 다른 유형의 실행 파일·안내서는 섞이지 않는다
    other = "청년" if kind == "성인" else "성인"
    assert f"run_missa_{other}.bat" not in names and f"README_{other}미사_운영자.md" not in names
    assert "config.json" not in names
    z = zipfile.ZipFile(out)
    for n in names:
        if n.endswith(".bat"):
            b = z.read(n)
            assert b"\n" not in b.replace(b"\r\n", b"")


@pytest.mark.parametrize("kind,shown,hidden", [("성인", "성인미사 PPT", "청년미사 PPT"), ("청년", "청년미사 PPT", "성인미사 PPT")])
def test_install_bat_names_only_own_mass_type_at_runtime(kind, shown, hidden, tmp_path, monkeypatch):
    """install.bat은 폴더의 run_missa_*.bat으로 유형을 판단해 안내 문구에 그 유형만 표시한다(고정 문구에 두 유형 병기 금지).
    실제 cmd로 실행해 출력 확인 — pip/바로가기 단계는 이미 설치된 환경에서 빠르게 끝난다."""
    import subprocess
    monkeypatch.setattr(bd, "DIST_DIR", tmp_path)
    if kind == "청년" and not all(p.exists() for p in bd.HYMN_PDF_PATHS):
        pytest.skip("성가집 원본 PDF 없음")
    out = bd.build(kind)
    dest = tmp_path / "x"
    zipfile.ZipFile(out).extractall(dest)
    r = subprocess.run(["cmd", "/c", str(dest / "install.bat")], cwd=dest, stdin=subprocess.DEVNULL,
                       capture_output=True, timeout=280)
    text = (r.stdout + r.stderr).decode("utf-8", errors="replace")
    assert "not recognized" not in text
    done = text[text.index("설치가 완료되었습니다"):]
    assert shown in done and hidden not in done
    assert shown in text[:text.index("설치가 완료되었습니다")]          # 바로가기 생성 안내도 해당 유형만
    assert hidden not in text
