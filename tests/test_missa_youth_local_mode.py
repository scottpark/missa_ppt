"""청년미사 OneDrive 접속 모드(로컬 'PPT 문서' 폴더 우선 / 로그인 폴백) 테스트 — 요구사항 §3.3.1."""
import json
import sys
import types
from pathlib import Path

import pytest

import missa_gui as gui


@pytest.fixture
def ppt_root(tmp_path):
    root = tmp_path / "PPT 문서"
    (root / "20.청년 미사" / "2.성가" / "나주노 성가").mkdir(parents=True)
    (root / "20.청년 미사" / "2.성가" / "야훼이레 성가").mkdir(parents=True)
    (root / "09.가톨릭 성가" / "성가-악보버전").mkdir(parents=True)
    return root


@pytest.fixture
def cfg(tmp_path, monkeypatch):
    path = tmp_path / "config.json"
    monkeypatch.setattr(gui, "CONFIG_FILE", path)
    gui._reset_youth_ppt_mode()
    return path


def _write_cfg(cfg, **kw):
    cfg.write_text(json.dumps(kw, ensure_ascii=False), encoding="utf-8")


def test_valid_folder_requires_youth_subfolder(tmp_path, ppt_root):
    assert gui._is_valid_youth_ppt_folder(ppt_root)
    other = tmp_path / "other"
    other.mkdir()
    assert not gui._is_valid_youth_ppt_folder(other)
    assert not gui._is_valid_youth_ppt_folder(tmp_path / "nope")


def test_saved_valid_folder_gives_local_without_popup(cfg, ppt_root, monkeypatch):
    _write_cfg(cfg, onedrive_ppt_folder=str(ppt_root))
    monkeypatch.setattr(gui, "_ask_onedrive_path_popup",
                        lambda **k: pytest.fail("선택창이 뜨면 안 됨"))
    assert gui.get_youth_ppt_mode(ask_if_missing=True) == "local"
    assert gui.youth_local_subdir("start_dir") == ppt_root


def test_missing_setting_without_ask_is_fallback_and_not_cached(cfg, ppt_root, monkeypatch):
    assert gui.get_youth_ppt_mode() == "fallback"
    _write_cfg(cfg, onedrive_ppt_folder=str(ppt_root))
    assert gui.get_youth_ppt_mode() == "local"


def test_saved_folder_without_youth_subfolder_is_treated_as_missing(cfg, tmp_path):
    bad = tmp_path / "bad"
    bad.mkdir()
    _write_cfg(cfg, onedrive_ppt_folder=str(bad))
    assert gui.get_youth_ppt_mode() == "fallback"


def test_ask_selection_is_saved_and_local(cfg, ppt_root, monkeypatch):
    monkeypatch.setattr(gui, "_ask_onedrive_path_popup", lambda **k: str(ppt_root))
    assert gui.get_youth_ppt_mode(ask_if_missing=True) == "local"
    assert json.loads(cfg.read_text(encoding="utf-8"))["onedrive_ppt_folder"] == str(ppt_root)


def test_invalid_selection_reasks_then_accepts(cfg, tmp_path, ppt_root, monkeypatch):
    bad = tmp_path / "bad"
    bad.mkdir()
    answers = iter([str(bad), str(ppt_root)])
    messages = []

    def fake(**k):
        messages.append(k["message"])
        return next(answers)

    monkeypatch.setattr(gui, "_ask_onedrive_path_popup", fake)
    assert gui.get_youth_ppt_mode(ask_if_missing=True) == "local"
    assert len(messages) == 2 and "20.청년 미사" in messages[1]


def test_cancel_gives_fallback_not_saved_and_cached(cfg, monkeypatch):
    calls = []
    monkeypatch.setattr(gui, "_ask_onedrive_path_popup", lambda **k: calls.append(1) or "")
    assert gui.get_youth_ppt_mode(ask_if_missing=True) == "fallback"
    assert not cfg.exists()
    assert gui.get_youth_ppt_mode(ask_if_missing=True) == "fallback"
    assert len(calls) == 1  # 같은 실행에서는 다시 묻지 않는다


def test_local_subdir_paths(cfg, ppt_root):
    gui._YOUTH_PPT_STATE.update(mode="local", folder=ppt_root)
    assert gui.youth_local_subdir("catholic_hymn") == ppt_root / "09.가톨릭 성가" / "성가-악보버전"
    assert gui.youth_local_subdir("youth_hymn", "나주노 성가") == \
        ppt_root / "20.청년 미사" / "2.성가" / "나주노 성가"
    assert gui.youth_local_subdir("output", year="2026") == ppt_root / "20.청년 미사" / "2026"
    with pytest.raises(ValueError):
        gui.youth_local_subdir("bogus")


def test_local_subdir_requires_local_mode(cfg):
    with pytest.raises(RuntimeError):
        gui.youth_local_subdir("start_dir")


def _no_graph(monkeypatch):
    """로컬 모드에서 Graph 모듈이 import/호출되면 실패하게 한다."""
    fake = types.ModuleType("missa_onedrive")

    def _boom(*a, **k):
        pytest.fail("로컬 모드에서 Graph API 호출 금지")

    for name in ("list_children", "download_file", "upload_file", "ensure_folder", "get_item_metadata"):
        setattr(fake, name, _boom)
    monkeypatch.setitem(sys.modules, "missa_onedrive", fake)


def test_local_hymn_lookup_finds_file_without_graph(cfg, ppt_root, monkeypatch):
    gui._YOUTH_PPT_STATE.update(mode="local", folder=ppt_root)
    _no_graph(monkeypatch)
    f = ppt_root / "20.청년 미사" / "2.성가" / "나주노 성가" / "나주노 성가 172 나의 고백.pptx"
    f.write_bytes(b"x")
    (f.parent / "~$나주노 성가 171 임시.pptx").write_bytes(b"x")
    found = gui.find_youth_onedrive_hymn_file(
        "onedrive_youth_hymn_path", gui._DEFAULT_ONEDRIVE_YOUTH_HYMN_PATH, "나주노 성가",
        r"^나주노 성가 172 ", "청년미사_성가/나주노 성가", local_kind="youth_hymn",
    )
    assert found == f
    missing = gui.find_youth_onedrive_hymn_file(
        "onedrive_youth_hymn_path", gui._DEFAULT_ONEDRIVE_YOUTH_HYMN_PATH, "나주노 성가",
        r"^나주노 성가 999 ", "청년미사_성가/나주노 성가", local_kind="youth_hymn",
    )
    assert missing is None


def test_local_catholic_hymn_lookup(cfg, ppt_root, monkeypatch):
    gui._YOUTH_PPT_STATE.update(mode="local", folder=ppt_root)
    _no_graph(monkeypatch)
    f = ppt_root / "09.가톨릭 성가" / "성가-악보버전" / "성가 96 하느님 약속하신.pptx"
    f.write_bytes(b"x")
    assert gui.find_youth_onedrive_hymn_file(
        "onedrive_hymn_path", gui._DEFAULT_ONEDRIVE_HYMN_PATH, None,
        r"성가 96(?!\d)", "청년_가톨릭성가", local_kind="catholic_hymn",
    ) == f


def test_youth_output_is_no_longer_auto_copied_or_uploaded():
    """2026-10-03: 청년 결과 PPT는 output/에만 생성한다 — 자동 복사·업로드 함수는 삭제됐다."""
    import missa_to_ppt as m
    assert not hasattr(m, "_upload_youth_output_to_onedrive")


# ── 결과창 '원드라이브로 복사'(2026-10-03) ──

def test_copy_result_youth_local_goes_to_year_folder_under_youth_root(cfg, ppt_root, tmp_path):
    gui._YOUTH_PPT_STATE.update(mode="local", folder=ppt_root)
    out = tmp_path / "결과.pptx"
    out.write_bytes(b"pptx")
    dest = gui.copy_result_to_onedrive("youth", out, "20261003")
    assert Path(dest) == ppt_root / "20.청년 미사" / "2026"
    assert (ppt_root / "20.청년 미사" / "2026" / "결과.pptx").read_bytes() == b"pptx"
    assert out.exists()


def test_copy_result_adult_goes_to_year_folder_directly_under_root(cfg, ppt_root, tmp_path):
    gui._ADULT_PPT_STATE["folder"] = ppt_root
    out = tmp_path / "성인.pptx"
    out.write_bytes(b"a")
    dest = gui.copy_result_to_onedrive("adult", out, "20261004")
    assert Path(dest) == ppt_root / "2026"
    assert (ppt_root / "2026" / "성인.pptx").read_bytes() == b"a"


def test_copy_result_exists_detects_same_name(cfg, ppt_root, tmp_path):
    gui._YOUTH_PPT_STATE.update(mode="local", folder=ppt_root)
    out = tmp_path / "x.pptx"
    out.write_bytes(b"1")
    assert not gui.result_copy_exists("youth", out, "20261003")
    gui.copy_result_to_onedrive("youth", out, "20261003")
    assert gui.result_copy_exists("youth", out, "20261003")


def test_copy_result_youth_fallback_uploads_via_graph_to_same_remote_path(cfg, tmp_path, monkeypatch):
    calls = {"ensure": [], "upload": []}

    class _OD:
        @staticmethod
        def ensure_folder(path):
            calls["ensure"].append(path)

        @staticmethod
        def upload_file(path, local):
            calls["upload"].append((path, local))

    monkeypatch.setitem(sys.modules, "missa_onedrive", _OD)
    out = tmp_path / "청년.pptx"
    out.write_bytes(b"1")
    dest = gui.copy_result_to_onedrive("youth", out, "20261003")
    assert calls["ensure"] == ["PPT 문서/20.청년 미사/2026"]
    assert calls["upload"] == [("PPT 문서/20.청년 미사/2026/청년.pptx", out)]
    assert "20.청년 미사" in dest and "2026" in dest


def test_copy_result_adult_without_root_raises(cfg, tmp_path):
    out = tmp_path / "a.pptx"
    out.write_bytes(b"1")
    with pytest.raises(RuntimeError):
        gui.copy_result_to_onedrive("adult", out, "20261004")


def _stub_messagebox(monkeypatch, answer=True):
    from tkinter import messagebox
    shown = []
    monkeypatch.setattr(messagebox, "showinfo", lambda t, m, **k: shown.append(("info", m)))
    monkeypatch.setattr(messagebox, "showerror", lambda t, m, **k: shown.append(("error", m)))
    monkeypatch.setattr(messagebox, "showwarning", lambda t, m, **k: shown.append(("warn", m)))
    monkeypatch.setattr(messagebox, "askyesno", lambda t, m, **k: (shown.append(("ask", m)), answer)[1])
    return shown


def test_copy_ui_announces_destination_folder(cfg, ppt_root, tmp_path, monkeypatch):
    gui._YOUTH_PPT_STATE.update(mode="local", folder=ppt_root)
    shown = _stub_messagebox(monkeypatch)
    out = tmp_path / "r.pptx"
    out.write_bytes(b"1")
    gui._copy_result_with_ui(None, "youth", out, "20261003")
    assert shown[-1][0] == "info"
    assert str(ppt_root / "20.청년 미사" / "2026") in shown[-1][1]


def test_copy_ui_asks_before_overwrite_and_respects_no(cfg, ppt_root, tmp_path, monkeypatch):
    gui._YOUTH_PPT_STATE.update(mode="local", folder=ppt_root)
    out = tmp_path / "r.pptx"
    out.write_bytes(b"new")
    dest_file = ppt_root / "20.청년 미사" / "2026" / "r.pptx"
    dest_file.parent.mkdir(parents=True)
    dest_file.write_bytes(b"old")
    shown = _stub_messagebox(monkeypatch, answer=False)
    gui._copy_result_with_ui(None, "youth", out, "20261003")
    assert [k for k, _ in shown] == ["ask"]
    assert dest_file.read_bytes() == b"old"


def test_copy_ui_reports_failure(cfg, ppt_root, tmp_path, monkeypatch):
    gui._YOUTH_PPT_STATE.update(mode="local", folder=ppt_root)
    shown = _stub_messagebox(monkeypatch)
    gui._copy_result_with_ui(None, "youth", tmp_path / "없는파일.pptx", "20261003")
    assert shown[-1][0] == "error"


def test_copy_ui_adult_without_root_asks_picker_then_cancel_warns(cfg, tmp_path, monkeypatch):
    monkeypatch.setattr(gui, "_ask_onedrive_path_popup", lambda **k: "")
    shown = _stub_messagebox(monkeypatch)
    out = tmp_path / "a.pptx"
    out.write_bytes(b"1")
    gui._copy_result_with_ui(None, "adult", out, "20261004")
    assert shown[-1][0] == "warn"


def test_result_window_has_copy_button_below_three_buttons():
    import inspect
    src = inspect.getsource(gui._show_result_window)
    assert "원드라이브로 복사" in src
    assert src.index("text='닫기'") < src.index("원드라이브로 복사")
    assert "_copy_result_with_ui" in src


# ── 성인 'PPT 문서' 폴더(2026-10-03) ──

def test_adult_valid_folder_requires_catholic_hymn_subfolder(tmp_path, ppt_root):
    assert gui._is_valid_adult_ppt_folder(ppt_root)
    other = tmp_path / "o"
    other.mkdir()
    assert not gui._is_valid_adult_ppt_folder(other)


def test_adult_picker_saves_ppt_folder_and_hymn_folder(cfg, ppt_root, monkeypatch):
    monkeypatch.setattr(gui, "_ask_onedrive_path_popup", lambda **k: str(ppt_root))
    got = gui.get_adult_ppt_folder(ask_if_missing=True)
    assert got == ppt_root
    saved = json.loads(cfg.read_text(encoding="utf-8"))
    assert saved["onedrive_ppt_folder"] == str(ppt_root)
    assert Path(saved["onedrive_hymn_folder"]) == ppt_root / "09.가톨릭 성가" / "성가-악보버전"


def test_adult_picker_reasks_on_invalid_folder_and_none_on_cancel(cfg, tmp_path, ppt_root, monkeypatch):
    bad = tmp_path / "bad"
    bad.mkdir()
    answers = iter([str(bad), ""])
    msgs = []

    def _popup(**k):
        msgs.append(k["message"])
        return next(answers)

    monkeypatch.setattr(gui, "_ask_onedrive_path_popup", _popup)
    assert gui.get_adult_ppt_folder(ask_if_missing=True) is None
    assert len(msgs) == 2 and "09.가톨릭 성가" in msgs[1]


def test_adult_without_ask_never_opens_popup(cfg, monkeypatch):
    monkeypatch.setattr(gui, "_ask_onedrive_path_popup", lambda **k: pytest.fail("팝업 금지"))
    assert gui.get_adult_ppt_folder() is None


def test_adult_derives_root_from_existing_hymn_folder_setting(cfg, ppt_root, monkeypatch):
    _write_cfg(cfg, onedrive_hymn_folder=str(ppt_root / "09.가톨릭 성가" / "성가-악보버전"))
    monkeypatch.setattr(gui, "_ask_onedrive_path_popup", lambda **k: pytest.fail("팝업 금지"))
    assert gui.get_adult_ppt_folder(ask_if_missing=True) == ppt_root


def test_run_gui_mode_adult_asks_ppt_folder_before_combined_popup(monkeypatch):
    import missa_to_ppt as m
    calls = []
    monkeypatch.setattr(m, "get_adult_ppt_folder", lambda ask_if_missing=False: calls.append(("adult", ask_if_missing)))
    monkeypatch.setattr(m, "_ask_combined_input_popup",
                        lambda mass_type: (calls.append("combined"), ("20261004", {}))[1])
    monkeypatch.setattr(m, "_ask_numbers_popup", lambda n, s: {})
    monkeypatch.setattr(m, "_run_with_progress_window", lambda fn: ("", ""))
    monkeypatch.setattr(m, "_show_result_window", lambda *a, **k: None)
    m._preloaded_inputs[0] = None
    try:
        m._run_gui_mode("성인")
    finally:
        m._preloaded_inputs[0] = None
    assert calls[:2] == [("adult", True), "combined"]


# ── 찾아보기 시작 폴더(2026-10-03) ──

@pytest.fixture
def full_root(ppt_root):
    for parts in (("13.기도문",), ("11.공지사항 PPT문서",), ("10.기타 PPT문서", "PPT 자동화", "1.Template"),
                  ("20.청년 미사", "1.Template"), ("14.화답송 악보",)):
        (ppt_root.joinpath(*parts)).mkdir(parents=True, exist_ok=True)
    return ppt_root


def test_start_subpaths_per_mass_type():
    f = gui.browse_start_subpath
    for mt in ("adult", "youth"):
        assert f(mt, "시작기도") == ("13.기도문",)
        assert f(mt, "미사후기도") == ("13.기도문",)
        assert f(mt, "공지사항") == ("11.공지사항 PPT문서",)
    assert f("adult", "ref_pptx") == ("10.기타 PPT문서", "PPT 자동화", "1.Template")
    assert f("youth", "ref_pptx") == ("20.청년 미사", "1.Template")
    assert f("adult", "화답송_pptx") is None


def test_initial_dir_adult_and_youth(cfg, full_root):
    gui._ADULT_PPT_STATE["folder"] = full_root
    assert gui.browse_initial_dir("adult", "시작기도") == full_root / "13.기도문"
    assert gui.browse_initial_dir("adult", "공지사항") == full_root / "11.공지사항 PPT문서"
    assert gui.browse_initial_dir("adult", "ref_pptx") == full_root / "10.기타 PPT문서" / "PPT 자동화" / "1.Template"
    gui._YOUTH_PPT_STATE.update(mode="local", folder=full_root)
    assert gui.browse_initial_dir("youth", "ref_pptx") == full_root / "20.청년 미사" / "1.Template"
    assert gui.browse_initial_dir("youth", "미사후기도") == full_root / "13.기도문"


def test_initial_dir_ignores_space_differences_and_falls_back_to_deepest_existing(cfg, ppt_root):
    (ppt_root / "13. 기도문").mkdir()          # 공백이 다른 이름
    gui._ADULT_PPT_STATE["folder"] = ppt_root
    assert gui.browse_initial_dir("adult", "시작기도") == ppt_root / "13. 기도문"
    # 하위 폴더가 아예 없으면 'PPT 문서' 자체
    assert gui.browse_initial_dir("adult", "공지사항") == ppt_root


def test_initial_dir_none_without_root(cfg):
    assert gui.browse_initial_dir("adult", "시작기도") is None
    assert gui.browse_initial_dir("youth", "ref_pptx") is None   # 로그인 폴백


def test_psalm_dir_first_root_then_remembered(cfg, full_root):
    gui._ADULT_PPT_STATE["folder"] = full_root
    assert gui.browse_initial_dir("adult", "화답송_pptx") == full_root
    picked = full_root / "14.화답송 악보" / "연중 제27주일.pptx"
    gui.remember_psalm_dir(picked)
    assert json.loads(cfg.read_text(encoding="utf-8"))["last_psalm_dir"] == str(picked.parent)
    assert gui.browse_initial_dir("adult", "화답송_pptx") == picked.parent


def test_psalm_remembered_dir_that_vanished_falls_back_to_root(cfg, full_root):
    gui._ADULT_PPT_STATE["folder"] = full_root
    _write_cfg(cfg, last_psalm_dir=str(full_root / "사라진 폴더"))
    assert gui.browse_initial_dir("adult", "화답송_pptx") == full_root


def test_graph_browser_accepts_start_subpath_and_youth_fallback_passes_it():
    import inspect
    assert "start_subpath" in inspect.signature(gui._ask_onedrive_file_browser_popup).parameters
    assert "start_subpath=browse_start_subpath(mass_type, key)" in inspect.getsource(gui._ask_combined_input_popup)
    assert "13.기도문" in gui._ONEDRIVE_ROOT_ALLOWED_FOLDERS


def test_resolve_new_hymn_saved_into_local_folder_without_upload(cfg, ppt_root, monkeypatch):
    import missa_content_updaters as cu
    import missa_youth_hymn_pdf as yh
    from pptx import Presentation
    gui._YOUTH_PPT_STATE.update(mode="local", folder=ppt_root)
    _no_graph(monkeypatch)
    monkeypatch.setattr(yh, "find_song_title", lambda src, n: "나의 고백")
    monkeypatch.setattr(yh, "build_hymn_pptx", lambda *a, **k: Presentation())
    monkeypatch.setattr(yh, "build_header_runs", lambda *a, **k: None)
    prs, title = cu.resolve_youth_hymn_pptx({"출처": "나주노", "번호": 172, "제목": None}, "입당")
    assert title == "나의 고백"
    saved = ppt_root / "20.청년 미사" / "2.성가" / "나주노 성가" / "나주노 성가 172 나의 고백.pptx"
    assert saved.exists()
