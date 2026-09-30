# -*- coding: utf-8 -*-
"""missa_gui.py의 OneDrive 폴더 해석(_resolve_onedrive_folder/_mirror_onedrive_folder)
프로그레션 테스트 (청년미사 2단계 §2.4 — 로컬 동기화 폴더 → Graph API 미러링 → 수동 선택
우선순위 전환).

실제 네트워크는 발생시키지 않는다 — `missa_onedrive.list_children`/`download_file`을
가짜 파일시스템 딕셔너리로 monkeypatch한다.
"""
from pathlib import Path

import pytest

import missa_gui as gui


@pytest.fixture(autouse=True)
def _isolate_script_dir(monkeypatch, tmp_path):
    monkeypatch.setattr(gui, '_SCRIPT_DIR', tmp_path)


class _FakeOneDriveFS:
    """{'folder/sub': ['a.pptx', 'b.pptx']} 형태로 원격 트리를 흉내 낸다."""

    def __init__(self, tree: dict, contents: dict = None):
        self.tree = tree  # path -> list of {'name', 'folder': {}} 또는 {'name', 'file': {}}
        self.contents = contents or {}  # 'path/name' -> bytes
        self.download_calls = []

    def list_children(self, path):
        return self.tree.get(path, [])

    def download_file(self, path, dest):
        self.download_calls.append(path)
        dest = Path(dest)
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(self.contents.get(path, b'stub'))
        return dest


def test_G1_local_path_skips_graph_entirely(monkeypatch, tmp_path):
    """G1: 로컬 동기화 경로가 이미 유효하면 Graph API(미러링)를 아예 건드리지 않는다."""
    local_dir = tmp_path / 'synced'
    local_dir.mkdir()
    monkeypatch.setattr(gui, '_load_config', lambda: {'onedrive_hymn_folder': str(local_dir)})

    def _boom(*a, **k):
        raise AssertionError('로컬 경로가 있으면 미러링을 호출하면 안 됨')

    monkeypatch.setattr(gui, '_mirror_onedrive_folder', _boom)
    result = gui.get_onedrive_hymn_folder()
    assert result == local_dir


def test_G2_missing_local_path_falls_back_to_graph_mirror(monkeypatch, tmp_path):
    """G2: 로컬 경로가 없으면 Graph API로 미러링한 캐시 폴더를 반환한다."""
    monkeypatch.setattr(gui, '_load_config', lambda: {})
    calls = []

    def _fake_mirror(remote_path, local_dir):
        calls.append((remote_path, local_dir))
        local_dir.mkdir(parents=True, exist_ok=True)

    monkeypatch.setattr(gui, '_mirror_onedrive_folder', _fake_mirror)
    result = gui.get_onedrive_hymn_folder()
    assert calls[0][0] == gui._DEFAULT_ONEDRIVE_HYMN_PATH
    assert result == calls[0][1]
    assert result.is_dir()


def test_G3_graph_failure_falls_back_to_manual_popup(monkeypatch, tmp_path):
    """G3: Graph 미러링이 실패하면(네트워크/로그인 불가 등) 수동 폴더 선택 팝업으로 폴백한다."""
    monkeypatch.setattr(gui, '_load_config', lambda: {})
    saved = {}
    monkeypatch.setattr(gui, '_save_config', lambda c: saved.update(c))

    def _boom(remote_path, local_dir):
        raise RuntimeError('simulated: 네트워크 불가')

    monkeypatch.setattr(gui, '_mirror_onedrive_folder', _boom)

    picked = tmp_path / 'manually_picked'
    picked.mkdir()
    monkeypatch.setattr(gui, '_ask_onedrive_path_popup', lambda **k: str(picked))

    result = gui.get_onedrive_hymn_folder()
    assert result == picked
    assert saved['onedrive_hymn_folder'] == str(picked)


def test_G4_youth_hymn_root_uses_youth_default_path(monkeypatch, tmp_path):
    """G4: 청년미사 성가 경로는 별도 기본 원격 경로/캐시 폴더를 쓴다(가톨릭성가와 섞이지 않음)."""
    monkeypatch.setattr(gui, '_load_config', lambda: {})
    calls = []

    def _fake_mirror(remote_path, local_dir):
        calls.append((remote_path, local_dir))
        local_dir.mkdir(parents=True, exist_ok=True)

    monkeypatch.setattr(gui, '_mirror_onedrive_folder', _fake_mirror)
    result = gui.get_onedrive_youth_hymn_root()
    assert calls[0][0] == gui._DEFAULT_ONEDRIVE_YOUTH_HYMN_PATH
    assert '청년미사_성가' in str(result)


def test_G5_mirror_downloads_missing_files_recursively(monkeypatch, tmp_path):
    """G5: 하위 폴더까지 재귀적으로 순회하며 파일을 내려받는다."""
    fake = _FakeOneDriveFS({
        'root': [
            {'name': '나주노 성가', 'folder': {}},
            {'name': 'readme.txt', 'file': {}},
        ],
        'root/나주노 성가': [
            {'name': '나주노 성가 447 제목.pptx', 'file': {}},
        ],
    })
    import sys
    monkeypatch.setitem(sys.modules, 'missa_onedrive', fake)

    dest = tmp_path / 'mirror'
    gui._mirror_onedrive_folder('root', dest)

    assert (dest / 'readme.txt').exists()
    assert (dest / '나주노 성가' / '나주노 성가 447 제목.pptx').exists()
    assert set(fake.download_calls) == {
        'root/readme.txt', 'root/나주노 성가/나주노 성가 447 제목.pptx',
    }


def test_G6_mirror_skips_already_cached_files(monkeypatch, tmp_path):
    """G6: 로컬에 이미 있는 파일은 다시 내려받지 않는다(증분 캐시)."""
    fake = _FakeOneDriveFS({'root': [{'name': 'a.pptx', 'file': {}}]})
    import sys
    monkeypatch.setitem(sys.modules, 'missa_onedrive', fake)

    dest = tmp_path / 'mirror'
    dest.mkdir()
    (dest / 'a.pptx').write_bytes(b'already-here')

    gui._mirror_onedrive_folder('root', dest)
    assert fake.download_calls == []
    assert (dest / 'a.pptx').read_bytes() == b'already-here'


def test_G7_mirror_skips_lock_files(monkeypatch, tmp_path):
    """G7: PowerPoint 임시 잠금 파일(~$로 시작)은 미러링 대상에서 제외한다."""
    fake = _FakeOneDriveFS({'root': [
        {'name': '~$a.pptx', 'file': {}},
        {'name': 'b.pptx', 'file': {}},
    ]})
    import sys
    monkeypatch.setitem(sys.modules, 'missa_onedrive', fake)

    dest = tmp_path / 'mirror'
    gui._mirror_onedrive_folder('root', dest)
    assert fake.download_calls == ['root/b.pptx']


def test_G8_remote_path_none_when_root_is_not_mirror_cache(monkeypatch, tmp_path):
    """G8: root가 미러링 캐시 폴더(cache/청년미사_성가)가 아니면(로컬 동기화 폴더든,
    테스트가 get_onedrive_youth_hymn_root()만 스텁해 넘겨준 임의 tmp_path든) None을
    반환해야 한다 — 실제 업로드(및 그 과정에서의 MSAL 로그인 시도)를 트리거하면 안 된다.
    2026-09-18 실측 버그(H1/H2/C1 테스트가 config.json을 독립적으로 재판정하다가 실제
    MSAL 기기 코드 로그인 대기로 영구히 멈춘 사건)의 회귀 가드."""
    monkeypatch.setattr(gui, '_load_config', lambda: {})  # config.json에 아무 키도 없어도
    arbitrary_root = tmp_path / 'some_test_stub_dir'
    assert gui.get_onedrive_youth_hymn_remote_path(arbitrary_root) is None


def test_G9_remote_path_returned_only_for_exact_mirror_cache_dir(monkeypatch, tmp_path):
    """G9: root가 _resolve_onedrive_folder()가 실제로 만드는 미러링 캐시 경로와 정확히
    같을 때만 원격 상대경로를 반환한다."""
    monkeypatch.setattr(gui, '_load_config', lambda: {})
    mirror_dir = tmp_path / 'cache' / '청년미사_성가'
    assert gui.get_onedrive_youth_hymn_remote_path(mirror_dir) == gui._DEFAULT_ONEDRIVE_YOUTH_HYMN_PATH


def test_G10_remote_path_respects_config_override(monkeypatch, tmp_path):
    """G10: config.json에 onedrive_youth_hymn_path가 있으면 기본값 대신 그 값을 쓴다."""
    monkeypatch.setattr(gui, '_load_config', lambda: {'onedrive_youth_hymn_path': 'Custom/Path'})
    mirror_dir = tmp_path / 'cache' / '청년미사_성가'
    assert gui.get_onedrive_youth_hymn_remote_path(mirror_dir) == 'Custom/Path'
