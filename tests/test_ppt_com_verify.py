"""
ppt_com_verify.py 단위/스모크 테스트

목적
----
Pillow 폭 추정 기반 줄 수 계산의 오프바이원 등 오차를 실제 PowerPoint COM
자동화로 검증하는 ppt_com_verify 모듈의 계약(contract)을 확인한다.

- is_available(): pywin32/Dispatch 실패 시 예외 없이 False를 반환하는가.
- count_slide_lines(): 내부에서 무슨 예외가 나든 ComVerificationUnavailable
  하나로 통일해서 올리는가.
- shutdown(): 어떤 상황에서도 예외를 밖으로 내보내지 않는가.
- 실제 PowerPoint COM이 있는 환경에서는, 한 줄짜리 텍스트가 실제로 1로
  측정되는가(스모크 테스트).

실행
----
    pytest test_ppt_com_verify.py -v
"""
from __future__ import annotations

import pytest

import ppt_com_verify as m


@pytest.fixture(autouse=True)
def _reset_module_state(monkeypatch):
    """각 테스트 전에 모듈 전역 상태(_app/_com_initialized/_available_cache)를
    초기값으로 되돌린다. monkeypatch를 사용해 테스트 종료 시 자동으로 이전
    값으로 복원되도록 한다."""
    monkeypatch.setattr(m, "_app", None)
    monkeypatch.setattr(m, "_com_initialized", False)
    monkeypatch.setattr(m, "_available_cache", None)
    yield


@pytest.fixture(scope="module", autouse=True)
def _cleanup_real_com_at_end():
    """이 파일의 실제 COM 스모크 테스트가 PowerPoint를 띄웠다면, 전체 테스트
    종료 후 반드시 정리해서 PowerPoint.exe 프로세스가 남지 않도록 한다."""
    yield
    m.shutdown()


class TestIsAvailable:
    def test_false_when_dispatch_fails(self, monkeypatch):
        def _boom():
            raise RuntimeError("simulated: pywin32/Dispatch 실패")

        monkeypatch.setattr(m, "_ensure_app", _boom)
        assert m.is_available() is False

    def test_never_raises_on_failure(self, monkeypatch):
        def _boom():
            raise OSError("simulated: 어떤 종류의 COM 오류든")

        monkeypatch.setattr(m, "_ensure_app", _boom)
        # 예외가 밖으로 나오면 이 호출 자체가 실패한다.
        m.is_available()

    def test_result_is_memoized(self, monkeypatch):
        calls = {"n": 0}

        def _boom():
            calls["n"] += 1
            raise RuntimeError("simulated failure")

        monkeypatch.setattr(m, "_ensure_app", _boom)
        first = m.is_available()
        second = m.is_available()
        assert first is False
        assert second is False
        assert calls["n"] == 1, "메모이즈되지 않고 매번 재시도하고 있음"


class TestCountSlideLines:
    def test_wraps_any_failure_as_com_verification_unavailable(self, monkeypatch):
        def _boom():
            raise RuntimeError("simulated: Dispatch/Open 실패")

        monkeypatch.setattr(m, "_ensure_app", _boom)
        with pytest.raises(m.ComVerificationUnavailable):
            m.count_slide_lines("does_not_matter.pptx", 1)

    def test_wraps_open_failure_from_a_real_looking_app(self, monkeypatch):
        class _FakePresentations:
            def Open(self, *args, **kwargs):
                raise RuntimeError("simulated: 파일을 열 수 없음")

        class _FakeApp:
            Presentations = _FakePresentations()

        monkeypatch.setattr(m, "_ensure_app", lambda: _FakeApp())
        with pytest.raises(m.ComVerificationUnavailable):
            m.count_slide_lines("bad_path.pptx", 1)

    def test_closes_presentation_even_when_measurement_fails(self, monkeypatch):
        close_calls = {"n": 0}

        class _FakeShapes:
            def __call__(self, idx):
                raise RuntimeError("simulated: 도형 인덱스 오류")

        class _FakeSlide:
            Shapes = _FakeShapes()

        class _FakeSlides:
            def __call__(self, idx):
                return _FakeSlide()

        class _FakePresentation:
            Slides = _FakeSlides()

            def Close(self):
                close_calls["n"] += 1

        class _FakePresentations:
            def Open(self, *args, **kwargs):
                return _FakePresentation()

        class _FakeApp:
            Presentations = _FakePresentations()

        monkeypatch.setattr(m, "_ensure_app", lambda: _FakeApp())
        with pytest.raises(m.ComVerificationUnavailable):
            m.count_slide_lines("some_path.pptx", 1)
        # count_slide_lines()는 첫 시도 실패 시 연결을 버리고 한 번 재시도하므로
        # Open된 프레젠테이션마다(2회) Close()가 호출돼야 한다 — 측정 실패 시에도
        # 열었던 프레젠테이션을 닫지 않고 방치하면 안 된다.
        assert close_calls["n"] == 2, "측정 실패 시에도 매 시도마다 Close()가 호출돼야 함"


class TestShutdown:
    def test_never_raises_with_no_app_ever_created(self):
        m.shutdown()  # 예외가 나오면 이 호출 자체가 실패한다.

    def test_never_raises_when_quit_throws(self, monkeypatch):
        class _FakeApp:
            def Quit(self):
                raise RuntimeError("simulated: Quit() 실패")

        monkeypatch.setattr(m, "_app", _FakeApp())
        monkeypatch.setattr(m, "_com_initialized", True)
        m.shutdown()  # 예외가 나오면 이 호출 자체가 실패한다.

    def test_resets_internal_state(self, monkeypatch):
        class _FakeApp:
            def Quit(self):
                pass

        monkeypatch.setattr(m, "_app", _FakeApp())
        monkeypatch.setattr(m, "_com_initialized", True)
        m.shutdown()
        assert m._app is None


class _CountApp:
    """Presentations.Count와 Quit() 호출만 흉내 내는 가짜 Application."""

    def __init__(self, count=0, count_raises=False):
        self.quit_calls = 0
        self.Visible = False

        class _Pres:
            Count = 0

            def Add(self, **kwargs):
                return object()

        pres = _Pres()
        if count_raises:
            class _Raising(_Pres):
                @property
                def Count(self):
                    raise RuntimeError("simulated: Count 조회 실패")
            pres = _Raising()
        else:
            pres.Count = count
        self.Presentations = pres

    def Quit(self):
        self.quit_calls += 1


class TestDiscardAppProtectsUserPowerPoint:
    """2026-11-18: Dispatch는 사용자가 이미 띄운 PowerPoint에 붙을 수 있다 — 우리가 띄운 인스턴스이고 열린
    프레젠테이션이 0개임이 확실할 때만 Quit한다(불확실하면 Quit 안 함)."""

    @staticmethod
    def _run(monkeypatch, app, launched):
        monkeypatch.setattr(m, "_app", app)
        monkeypatch.setattr(m, "_keepalive_presentation", None)
        monkeypatch.setattr(m, "_we_launched_app", launched, raising=False)
        m._discard_app()
        assert m._app is None

    def test_quits_when_we_launched_it_and_no_presentation_is_open(self, monkeypatch):
        app = _CountApp(count=0)
        self._run(monkeypatch, app, launched=True)
        assert app.quit_calls == 1

    def test_does_not_quit_when_other_presentations_are_open(self, monkeypatch):
        app = _CountApp(count=1)
        self._run(monkeypatch, app, launched=True)
        assert app.quit_calls == 0

    def test_does_not_quit_when_count_lookup_fails(self, monkeypatch):
        app = _CountApp(count_raises=True)
        self._run(monkeypatch, app, launched=True)
        assert app.quit_calls == 0

    def test_does_not_quit_an_instance_that_was_already_running(self, monkeypatch):
        app = _CountApp(count=0)
        self._run(monkeypatch, app, launched=False)
        assert app.quit_calls == 0


class _StuckPresentation:
    """Close()가 실패해(예: 0x800401fd) PowerPoint 쪽에는 계속 열려 있는 우리 프레젠테이션."""

    def Close(self):
        raise RuntimeError("simulated: Close 실패")


class TestOurOwnPresentationsAreNotUserDocuments:
    """우리가 만든 keepalive/probe가 닫히지 않고 Count에 남아도 '사용자 문서'로 세지 않는다 — 안 그러면 우리가
    띄운 PowerPoint가 좀비로 남는다(2026-11-18 전체 pytest 후 POWERPNT.EXE 1개 잔존)."""

    @staticmethod
    def _run(monkeypatch, app, keepalive, launched=True, unclosed=0):
        monkeypatch.setattr(m, "_app", app)
        monkeypatch.setattr(m, "_keepalive_presentation", keepalive)
        monkeypatch.setattr(m, "_we_launched_app", launched, raising=False)
        monkeypatch.setattr(m, "_unclosed_ours", unclosed, raising=False)
        m._discard_app()

    def test_quits_when_only_our_unclosable_keepalive_remains(self, monkeypatch):
        app = _CountApp(count=1)
        self._run(monkeypatch, app, _StuckPresentation())
        assert app.quit_calls == 1

    def test_does_not_quit_when_a_user_document_is_also_open(self, monkeypatch):
        app = _CountApp(count=2)
        self._run(monkeypatch, app, _StuckPresentation())
        assert app.quit_calls == 0

    def test_unclosed_probe_presentations_are_ours_too(self, monkeypatch):
        app = _CountApp(count=2)
        self._run(monkeypatch, app, None, unclosed=2)
        assert app.quit_calls == 1

    def test_attached_user_instance_is_never_quit_even_with_only_ours_left(self, monkeypatch):
        app = _CountApp(count=1)
        self._run(monkeypatch, app, _StuckPresentation(), launched=False)
        assert app.quit_calls == 0

    def test_failed_probe_close_is_tracked(self, monkeypatch):
        monkeypatch.setattr(m, "_unclosed_ours", 0, raising=False)
        m._close_ours(_StuckPresentation())
        assert m._unclosed_ours == 1


class TestEnsureAppRecordsWhoLaunched:
    @staticmethod
    def _install(monkeypatch, get_active, dispatched):
        import sys
        import types
        wc = types.ModuleType("win32com.client")
        wc.Dispatch = lambda name: dispatched
        wc.GetActiveObject = get_active
        w32 = types.ModuleType("win32com")
        w32.client = wc
        pc = types.ModuleType("pythoncom")
        pc.CoInitialize = lambda: None
        monkeypatch.setitem(sys.modules, "win32com", w32)
        monkeypatch.setitem(sys.modules, "win32com.client", wc)
        monkeypatch.setitem(sys.modules, "pythoncom", pc)
        monkeypatch.setattr(m, "_keepalive_presentation", None)
        monkeypatch.setattr(m, "_we_launched_app", False, raising=False)

    def test_already_running_instance_is_not_ours(self, monkeypatch):
        existing = _CountApp(count=2)
        self._install(monkeypatch, lambda name: existing, dispatched=_CountApp())
        assert m._ensure_app() is existing
        assert m._we_launched_app is False

    def test_not_running_means_we_launched_it(self, monkeypatch):
        import pywintypes

        def not_running(name):
            raise pywintypes.com_error(-2147221021, "unavailable", None, None)

        ours = _CountApp()
        self._install(monkeypatch, not_running, dispatched=ours)
        assert m._ensure_app() is ours
        assert m._we_launched_app is True

    def test_unknown_lookup_error_is_treated_as_not_ours(self, monkeypatch):
        def weird(name):
            raise RuntimeError("simulated: 알 수 없는 실패")

        self._install(monkeypatch, weird, dispatched=_CountApp())
        m._ensure_app()
        assert m._we_launched_app is False


class TestRealComSmoke:
    def test_real_com_counts_one_line_text(self, tmp_path):
        if not m.is_available():
            pytest.skip("이 환경에서 PowerPoint COM을 사용할 수 없어 건너뜀")

        from pptx import Presentation
        from pptx.util import Inches, Pt

        prs = Presentation()
        blank_layout = prs.slide_layouts[6]  # 자리표시자 없는 빈 레이아웃
        slide = prs.slides.add_slide(blank_layout)
        textbox = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.5), Inches(9), Inches(1.2)
        )
        tf = textbox.text_frame
        tf.word_wrap = True
        tf.text = "Hello world"
        tf.paragraphs[0].runs[0].font.size = Pt(24)

        pptx_path = tmp_path / "smoke_one_line.pptx"
        prs.save(str(pptx_path))

        # 빈 레이아웃 슬라이드에 도형을 하나만 추가했으므로 1번(1-based)이 그 도형.
        count = m.count_slide_lines(str(pptx_path), 1)
        assert count == 1
