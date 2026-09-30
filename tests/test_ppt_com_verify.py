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
