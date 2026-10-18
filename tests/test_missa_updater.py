# -*- coding: utf-8 -*-
"""missa_updater.py 프로그레션 테스트 (청년미사 2단계 §1 — 배포 업데이터, U1/U2).

네트워크 호출(get_latest_commit_sha/apply_update의 다운로드 단계)은 monkeypatch로
대체하고, install_whitelisted()의 화이트리스트 로직은 순수 파일 연산이라 실제
tmp_path 디렉터리로 직접 검증한다(네트워크 모킹 불필요).
"""
from pathlib import Path

import pytest

import missa_updater as mu


class _FakeResponse:
    def __init__(self, json_data=None, content=b""):
        self._json = json_data
        self.content = content

    def raise_for_status(self):
        pass

    def json(self):
        return self._json


def test_U1_check_for_update_detects_new_commit(monkeypatch):
    """U1: config의 last_update_commit과 최신 SHA가 다르면 업데이트 있음으로 판정."""
    monkeypatch.setattr(mu, "get_latest_commit_sha", lambda: "newsha123")
    has_update, sha = mu.check_for_update({"last_update_commit": "oldsha000"})
    assert has_update is True
    assert sha == "newsha123"


def test_U1b_check_for_update_no_update_when_same(monkeypatch):
    monkeypatch.setattr(mu, "get_latest_commit_sha", lambda: "samesha")
    has_update, sha = mu.check_for_update({"last_update_commit": "samesha"})
    assert has_update is False


def test_U1c_check_for_update_first_run_no_stored_commit(monkeypatch):
    """last_update_commit 키 자체가 없는 최초 실행에서도 업데이트 있음으로 판정."""
    monkeypatch.setattr(mu, "get_latest_commit_sha", lambda: "anysha")
    has_update, _ = mu.check_for_update({})
    assert has_update is True


def test_U1d_get_latest_commit_sha_parses_api_response(monkeypatch):
    monkeypatch.setattr(
        mu.requests, "get", lambda *a, **k: _FakeResponse(json_data={"sha": "abc123"})
    )
    assert mu.get_latest_commit_sha() == "abc123"


def test_U1e_get_latest_commit_sha_wraps_network_error(monkeypatch):
    def _raise(*a, **k):
        raise mu.requests.RequestException("boom")

    monkeypatch.setattr(mu.requests, "get", _raise)
    with pytest.raises(mu.UpdateError):
        mu.get_latest_commit_sha()


def test_U2_install_whitelisted_copies_py_and_assets(tmp_path):
    """U2: *.py/*.spec/assets/만 복사되고, config.json 등은 전혀 건드리지 않는다."""
    src = tmp_path / "src_root"
    (src / "assets").mkdir(parents=True)
    (src / "missa_to_ppt.py").write_text("print('new code')", encoding="utf-8")
    (src / "missa_to_ppt.spec").write_text("# spec", encoding="utf-8")
    (src / "assets" / "template.pptx").write_bytes(b"fake-pptx-bytes")
    (src / "README.md").write_text("docs", encoding="utf-8")  # 화이트리스트 밖

    install_dir = tmp_path / "install"
    install_dir.mkdir()
    (install_dir / "config.json").write_text('{"onedrive_hymn_folder": "C:/x"}', encoding="utf-8")
    (install_dir / "output").mkdir()
    (install_dir / "output" / "20260913").mkdir()
    (install_dir / "output" / "20260913" / "result.pptx").write_bytes(b"user-data")

    mu.install_whitelisted(src, install_dir)

    assert (install_dir / "missa_to_ppt.py").read_text(encoding="utf-8") == "print('new code')"
    assert (install_dir / "missa_to_ppt.spec").exists()
    assert (install_dir / "assets" / "template.pptx").read_bytes() == b"fake-pptx-bytes"
    assert not (install_dir / "README.md").exists()  # 화이트리스트 밖은 복사되지 않음

    # 로컬 전용 상태는 절대 건드리지 않는다
    assert (install_dir / "config.json").read_text(encoding="utf-8") == '{"onedrive_hymn_folder": "C:/x"}'
    assert (install_dir / "output" / "20260913" / "result.pptx").read_bytes() == b"user-data"


def test_U2b_install_whitelisted_overwrites_existing_py(tmp_path):
    src = tmp_path / "src_root"
    src.mkdir()
    (src / "missa_gui.py").write_text("v2", encoding="utf-8")

    install_dir = tmp_path / "install"
    install_dir.mkdir()
    (install_dir / "missa_gui.py").write_text("v1-old", encoding="utf-8")

    mu.install_whitelisted(src, install_dir)
    assert (install_dir / "missa_gui.py").read_text(encoding="utf-8") == "v2"


def test_U2c_install_whitelisted_replaces_assets_dir_wholesale(tmp_path):
    """assets/는 통째로 교체(옛 파일이 새 버전엔 없으면 사라져야 함 — 자산 이름 변경 반영)."""
    src = tmp_path / "src_root"
    (src / "assets").mkdir(parents=True)
    (src / "assets" / "new_asset.pptx").write_bytes(b"new")

    install_dir = tmp_path / "install"
    (install_dir / "assets").mkdir(parents=True)
    (install_dir / "assets" / "old_asset.pptx").write_bytes(b"old")

    mu.install_whitelisted(src, install_dir)
    assert (install_dir / "assets" / "new_asset.pptx").exists()
    assert not (install_dir / "assets" / "old_asset.pptx").exists()


def test_apply_update_end_to_end_with_zip(monkeypatch, tmp_path):
    """apply_update()가 실제 zip 다운로드~압축해제~설치까지 이어지는지(다운로드만 모킹)."""
    import zipfile

    fake_repo_dir = tmp_path / "missa_ppt-abc123"
    fake_repo_dir.mkdir()
    (fake_repo_dir / "missa_to_ppt.py").write_text("code", encoding="utf-8")

    zip_path = tmp_path / "fake.zip"
    with zipfile.ZipFile(zip_path, "w") as zf:
        zf.write(fake_repo_dir / "missa_to_ppt.py", "missa_ppt-abc123/missa_to_ppt.py")

    def _fake_download(sha, dest_zip: Path, progress=None):
        dest_zip.write_bytes(zip_path.read_bytes())

    monkeypatch.setattr(mu, "_download_archive", _fake_download)

    install_dir = tmp_path / "install"
    install_dir.mkdir()
    mu.apply_update("abc123", install_dir)
    assert (install_dir / "missa_to_ppt.py").read_text(encoding="utf-8") == "code"


def test_U_readme_updated_only_when_installed_copy_exists(tmp_path):
    """운영자 안내서(docs/README_*_운영자.md)는 설치 폴더 최상위로 갱신된다 — 단, 설치 폴더에 이미 있는 이름만
    (성인 설치에 청년 안내서가 생기지 않게)."""
    src = tmp_path / "src_root"
    (src / "docs").mkdir(parents=True)
    (src / "docs" / "README_성인미사_운영자.md").write_text("새 성인 안내", encoding="utf-8")
    (src / "docs" / "README_청년미사_운영자.md").write_text("새 청년 안내", encoding="utf-8")
    (src / "docs" / "다른 문서.md").write_text("x", encoding="utf-8")

    install_dir = tmp_path / "install"
    install_dir.mkdir()
    (install_dir / "README_성인미사_운영자.md").write_text("옛 성인 안내", encoding="utf-8")

    copied = mu.install_whitelisted(src, install_dir)

    assert (install_dir / "README_성인미사_운영자.md").read_text(encoding="utf-8") == "새 성인 안내"
    assert not (install_dir / "README_청년미사_운영자.md").exists()
    assert not (install_dir / "다른 문서.md").exists() and not (install_dir / "docs").exists()
    assert Path("README_성인미사_운영자.md") in copied


# --- 2026-10-18 업데이트 진행 콜백 (SPEC_업데이트_진행바) --------------------------------
class _FakeResp:
    def __init__(self, data: bytes, with_length=True):
        self._data = data
        self.headers = {"Content-Length": str(len(data))} if with_length else {}

    def raise_for_status(self):
        pass

    def iter_content(self, n):
        for i in range(0, len(self._data), n):
            yield self._data[i:i + n]


def test_P1_download_reports_bytes_and_total(monkeypatch, tmp_path):
    data = b"x" * (mu._DOWNLOAD_CHUNK * 2 + 10)
    monkeypatch.setattr(mu.requests, "get", lambda *a, **k: _FakeResp(data))
    events = []
    mu._download_archive("sha", tmp_path / "a.zip", lambda st, d, t: events.append((st, d, t)))
    assert (tmp_path / "a.zip").read_bytes() == data
    assert events[0] == (mu.STAGE_DOWNLOAD, 0, len(data))
    assert events[-1] == (mu.STAGE_DOWNLOAD, len(data), len(data))
    assert [e[1] for e in events] == sorted(e[1] for e in events)  # 단조 증가


def test_P2_download_without_content_length_reports_unknown_total(monkeypatch, tmp_path):
    monkeypatch.setattr(mu.requests, "get", lambda *a, **k: _FakeResp(b"abc", with_length=False))
    events = []
    mu._download_archive("sha", tmp_path / "a.zip", lambda st, d, t: events.append((st, d, t)))
    assert all(t is None for _s, _d, t in events)
    assert events[-1][1] == 3


def test_P3_install_reports_file_counts(tmp_path):
    src = tmp_path / "src"
    (src / "assets").mkdir(parents=True)
    (src / "a.py").write_text("1", encoding="utf-8")
    (src / "b.py").write_text("2", encoding="utf-8")
    (src / "assets" / "x.png").write_bytes(b"1")
    install = tmp_path / "inst"
    install.mkdir()
    events = []
    mu.install_whitelisted(src, install, lambda st, d, t: events.append((st, d, t)))
    assert events[0] == (mu.STAGE_INSTALL, 0, 3)
    assert events[-1] == (mu.STAGE_INSTALL, 3, 3)
    assert (install / "assets" / "x.png").exists()


def test_P4_apply_update_emits_all_stages_in_order(monkeypatch, tmp_path):
    import zipfile
    zip_path = tmp_path / "fake.zip"
    with zipfile.ZipFile(zip_path, "w") as zf:
        zf.writestr("repo-abc/missa_to_ppt.py", "code")
    monkeypatch.setattr(mu.requests, "get", lambda *a, **k: _FakeResp(zip_path.read_bytes()))
    install = tmp_path / "inst"
    install.mkdir()
    seen = []
    mu.apply_update("abc", install, lambda st, d, t: seen.append(st) if not seen or seen[-1] != st else None)
    assert seen == [mu.STAGE_DOWNLOAD, mu.STAGE_EXTRACT, mu.STAGE_INSTALL, mu.STAGE_DONE]
