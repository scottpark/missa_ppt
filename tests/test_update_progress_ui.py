"""업데이트 진행 창(SPEC_업데이트_진행바) GUI 로직 테스트.

Tk 창은 띄우지 않는다(규칙 20·21) — 순수 함수, 큐 상태 전이, 모킹한 `_on_check_update` 흐름만 검증한다.
"""
import queue
import sys

import pytest

import missa_gui as gui
import missa_updater as mu


# ── format_update_progress ──────────────────────────────────────────────

def test_check_stage_is_indeterminate():
    text, frac = gui.format_update_progress(mu.STAGE_CHECK, 0, None)
    assert '확인' in text and frac is None


def test_download_with_total_shows_mb_and_percent():
    text, frac = gui.format_update_progress(mu.STAGE_DOWNLOAD, 5 * 1024 * 1024, 10 * 1024 * 1024)
    assert '5.0 MB / 10.0 MB (50%)' in text
    assert frac == pytest.approx(0.5)


def test_download_without_total_is_indeterminate_with_received_mb():
    text, frac = gui.format_update_progress(mu.STAGE_DOWNLOAD, 3 * 1024 * 1024, None)
    assert '3.0 MB' in text and '/' not in text.split('\n', 1)[1]
    assert frac is None


def test_download_zero_total_treated_as_unknown():
    _, frac = gui.format_update_progress(mu.STAGE_DOWNLOAD, 10, 0)
    assert frac is None


def test_extract_and_install_show_count_and_fraction():
    text, frac = gui.format_update_progress(mu.STAGE_EXTRACT, 3, 12)
    assert '3 / 12' in text and frac == pytest.approx(0.25)
    text, frac = gui.format_update_progress(mu.STAGE_INSTALL, 4, 8)
    assert '4 / 8' in text and frac == pytest.approx(0.5)


def test_fraction_clamped_and_done_is_full():
    _, frac = gui.format_update_progress(mu.STAGE_INSTALL, 20, 8)
    assert frac == 1.0
    text, frac = gui.format_update_progress(mu.STAGE_DONE, 1, 1)
    assert frac == 1.0 and '완료' in text


def test_install_zero_total_is_indeterminate():
    _, frac = gui.format_update_progress(mu.STAGE_INSTALL, 0, 0)
    assert frac is None


# ── 작업 스레드 + 큐 ───────────────────────────────────────────────────

def test_worker_posts_progress_then_result():
    q = queue.Queue()

    def work(progress):
        progress(mu.STAGE_DOWNLOAD, 1, 2)
        return 'ok'

    gui._update_worker(work, q)
    assert q.get_nowait() == ('progress', mu.STAGE_DOWNLOAD, 1, 2)
    assert q.get_nowait() == ('result', 'ok')


def test_worker_posts_exception():
    q = queue.Queue()
    err = mu.UpdateError('x')

    def work(progress):
        raise err

    gui._update_worker(work, q)
    assert q.get_nowait() == ('error', err)


def test_drain_applies_only_last_progress_and_not_finished():
    q = queue.Queue()
    for i in range(5):
        q.put(('progress', mu.STAGE_DOWNLOAD, i, 10))
    seen = []
    finished = gui._drain_update_queue(q, lambda *a: seen.append(a), lambda v: 1 / 0, lambda e: 1 / 0)
    assert finished is False
    assert seen == [(mu.STAGE_DOWNLOAD, 4, 10)]


def test_drain_result_and_error_finish():
    q = queue.Queue()
    q.put(('progress', mu.STAGE_CHECK, 0, None))
    q.put(('result', 42))
    got = {}
    assert gui._drain_update_queue(q, lambda *a: None, lambda v: got.update(r=v), lambda e: got.update(e=e))
    assert got == {'r': 42}
    q.put(('error', ValueError('b')))
    assert gui._drain_update_queue(q, lambda *a: None, lambda v: got.update(r=v), lambda e: got.update(e=e))
    assert isinstance(got['e'], ValueError)


def test_drain_empty_queue_not_finished():
    assert gui._drain_update_queue(queue.Queue(), lambda *a: None, lambda v: None, lambda e: None) is False


# ── _on_check_update 흐름 (진행 창은 모킹) ─────────────────────────────

class _Parent:
    destroyed = False

    def destroy(self):
        self.destroyed = True


@pytest.fixture
def flow(monkeypatch):
    from tkinter import messagebox
    log = {'boxes': [], 'saved': None, 'runs': []}
    monkeypatch.setattr(gui, '_load_config', lambda: {})
    monkeypatch.setattr(gui, '_save_config', lambda c: log.update(saved=dict(c)))
    for name in ('showinfo', 'showerror'):
        monkeypatch.setattr(messagebox, name, lambda t, m, **k: log['boxes'].append((t, m)))
    monkeypatch.setattr(messagebox, 'askyesno', lambda *a, **k: log.get('yes', True))
    return log


def _script_runs(monkeypatch, log, outcomes):
    outcomes = list(outcomes)

    def fake(parent, title, func):
        log['runs'].append(title)
        return outcomes.pop(0)

    monkeypatch.setattr(gui, '_run_with_progress', fake)


def test_up_to_date_shows_info_only(monkeypatch, flow):
    _script_runs(monkeypatch, flow, [(True, (False, 'sha'))])
    gui._on_check_update(_Parent())
    assert flow['boxes'] == [('업데이트 확인', '이미 최신 버전입니다.')]
    assert len(flow['runs']) == 1


def test_check_error_shows_error_box(monkeypatch, flow):
    _script_runs(monkeypatch, flow, [(False, mu.UpdateError('net'))])
    gui._on_check_update(_Parent())
    assert flow['boxes'] == [('업데이트 확인 실패', 'net')]


def test_declined_does_not_apply(monkeypatch, flow):
    flow['yes'] = False
    _script_runs(monkeypatch, flow, [(True, (True, 'sha'))])
    gui._on_check_update(_Parent())
    assert len(flow['runs']) == 1 and flow['saved'] is None


def test_apply_error_shows_error_box(monkeypatch, flow):
    _script_runs(monkeypatch, flow, [(True, (True, 'sha')), (False, mu.UpdateError('boom'))])
    gui._on_check_update(_Parent())
    assert flow['boxes'] == [('업데이트 적용 실패', 'boom')]
    assert flow['saved'] is None


def test_apply_success_saves_config_and_exits(monkeypatch, flow):
    _script_runs(monkeypatch, flow, [(True, (True, 'abc')), (True, ['a.py'])])
    p = _Parent()
    with pytest.raises(SystemExit):
        gui._on_check_update(p)
    assert flow['saved'] == {'last_update_commit': 'abc'}
    assert flow['boxes'][-1][0] == '업데이트 완료'
    assert p.destroyed


def test_non_update_error_in_worker_is_reraised(monkeypatch, flow):
    _script_runs(monkeypatch, flow, [(False, RuntimeError('bug'))])
    with pytest.raises(RuntimeError):
        gui._on_check_update(_Parent())
