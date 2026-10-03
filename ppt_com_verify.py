"""
ppt_com_verify.py

PowerPoint COM 자동화로 슬라이드의 실제 렌더링 줄 수(TextRange.Lines.Count)를
측정하는 "그라운드 트루스" 검증 모듈.

배경
----
`missa_to_ppt.py`의 `_get_slide_render_params()`는 Pillow로 폰트 폭을 측정해
PowerPoint가 실제로 몇 줄로 렌더링할지 추정한다. 이 추정에는 TTC 폰트 인덱스
오해, libraqm(커닝) 미지원 등으로 오프바이원 같은 오차가 생길 수 있다
(자세한 배경은 CLAUDE.md "한글 줄 수 계산: TTC 폰트 인덱스와 Pillow 커닝 한계"
참조). 이 모듈은 실제 PowerPoint 인스턴스를 통해 그 추정값을 검증하는 용도로
쓰인다 — Pillow 추정을 대체하는 것이 아니라, 의심스러운 케이스를 실제
PowerPoint 렌더링과 대조하는 보조 수단이다.

설계 제약 (변경 금지)
--------------------
- late-binding만 사용한다: `win32com.client.Dispatch("PowerPoint.Application")`.
  `win32com.client.gencache.EnsureDispatch`는 절대 쓰지 않는다 — 이건 최초
  실행한 PowerPoint 버전에 고정된 typelib 래퍼를 캐싱하는데, 빌드된 exe가
  다른 PowerPoint 버전(2007~365)의 최종 사용자 PC에서 실행될 때 깨진다.
- 프로세스 내에서 `count_slide_lines()`를 여러 번 호출해도 PowerPoint
  Application 인스턴스 하나를 재사용한다(모듈 전역 상태). 호출마다 새로
  PowerPoint를 띄우면 이미 감수하기로 한 실행당 10~80초 오버헤드가 분 단위로
  불어난다.
- `app.Visible = True`로 둔다 — PowerPoint COM 자동화는 모든 버전에서
  안정적으로 완전히 숨겨서 돌릴 수 없다(버그가 아니라 받아들인 제약). 다만
  이건 **Application 자체**의 얘기고, **개별 프레젠테이션 창**은 별개다 —
  `Presentations.Add`/`Presentations.Open` 둘 다 `WithWindow=False`로 열어서
  검증 대상 프레젠테이션마다 창이 뜨고 닫히는 것(사용자 보고: 화면이 떴다
  닫혔다 반복되는 플리커, 2026-09-18)을 막는다. `Application.Visible=True`는
  유지되므로 최초 Dispatch 시점에 PowerPoint 프로세스/작업표시줄 아이콘이
  나타날 수 있다 — 그래서 `_ensure_app()`이 최초 1회 콘솔에 안내 메시지를
  출력해, 처음 쓰는 사용자가 놀라지 않게 한다.
- 여는 프레젠테이션은 항상 읽기 전용으로 열고(`ReadOnly=True`), 저장은 절대
  하지 않는다.
- `Presentations.Open(...)`으로 연 프레젠테이션은 측정 성공/실패와 무관하게
  반드시 `finally`에서 `.Close()`한다.
"""
from __future__ import annotations

# 재사용되는 단일 PowerPoint.Application COM 객체. None이면 아직 생성되지 않음.
_app = None

# PowerPoint는 열려 있는 프레젠테이션이 0개가 되면 프레임(메인) 창 자체가
# 사라진다 — 그 상태에서 Presentations.Open을 호출하면 "The PowerPoint Frame
# window does not exist" 오류가 난다(실측으로 확인됨). count_slide_lines()가
# 매번 열고-닫는 measurement용 프레젠테이션과 별개로, 빈 프레젠테이션 하나를
# Application 생성 시점부터 shutdown()까지 계속 열어 둬서 프레임이 항상
# 존재하도록 보장한다.
_keepalive_presentation = None

# 이 프로세스에서 CoInitialize()를 이미 호출했는지 여부.
_com_initialized = False

# is_available()의 메모이즈된 결과. None이면 아직 판정하지 않음.
_available_cache = None


class ComVerificationUnavailable(Exception):
    """pywin32가 없거나, Dispatch 실패, 혹은 COM 사용 중 발생한 모든 오류.

    호출자는 이 예외 하나만 잡으면 된다 — 원인이 ImportError든 com_error든
    다른 무엇이든 여기로 통일해서 올라온다.
    """


def _ensure_app():
    """모듈 전역 `_app`을 lazily 생성해서 반환한다.

    이미 생성돼 있으면 그대로 재사용한다. 이 함수는 실패 시 원래 예외를 그대로
    올린다(래핑은 호출부인 is_available()/count_slide_lines()의 책임).
    """
    global _app, _com_initialized, _keepalive_presentation

    import pythoncom
    import win32com.client

    if not _com_initialized:
        pythoncom.CoInitialize()
        _com_initialized = True

    if _app is None:
        # §J7(2026-09-26) 실측: EnumWindows로 PP12FrameClass 프레임 창을 직접 찾아
        # GetForegroundWindow()와 비교(Dispatch 직후·Presentations.Open() 측정 도중 각각
        # 2회 반복) — Application.Visible=True로 프레임 창이 실제로 IsWindowVisible=1
        # (작업표시줄에 나타남) 상태가 되지만, 두 시점 모두 포그라운드는 바뀌지 않았다.
        # 창이 전혀 뜨지 않는다는 예전 주장은 이 실측과 어긋나므로 쓰지 않는다 — 창은
        # 실제로 나타나지만 지금 쓰고 있는 창의 포커스를 가져가지는 않는다는 사실만 말한다.
        print('  [알림] 줄 수를 정확히 확인하려고 PowerPoint를 백그라운드에서 실행합니다 '
              '(작업표시줄에 잠깐 나타날 수 있지만, 지금 쓰고 있는 창의 포커스를 '
              '가져가지는 않습니다).')
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = True
        _app = app

    if _keepalive_presentation is None:
        # WithWindow=False — 프레임을 살려두는 용도일 뿐 사용자에게 보여줄 필요가
        # 없다(위 "설계 제약" 참고, 2026-09-18 플리커 수정).
        _keepalive_presentation = _app.Presentations.Add(WithWindow=False)

    return _app


def is_available() -> bool:
    """pywin32를 임포트할 수 있고 PowerPoint.Application을 Dispatch할 수
    있으면 True. 어떤 이유로든 실패하면(임포트 실패, com_error, 그 밖의 모든
    예외) False. 예외를 절대 밖으로 내보내지 않는다. 결과는 메모이즈된다.
    """
    global _available_cache

    if _available_cache is not None:
        return _available_cache

    try:
        _ensure_app()
        _available_cache = True
    except Exception:
        _available_cache = False

    return _available_cache


def _open_and_measure(temp_pptx_path: str, shape_index: int, want_bounds: bool = False,
                      line_starts_max: int = None):
    app = _ensure_app()
    presentation = app.Presentations.Open(
        temp_pptx_path,
        ReadOnly=True,
        Untitled=False,
        WithWindow=False,
    )
    try:
        shape = presentation.Slides(1).Shapes(shape_index)
        # TextRange.Lines는 PowerPoint 객체 모델에서 선택적 인자(Start, Length)를
        # 받는 파라미터화된 속성이라, late-bound COM에서는 일반 속성이 아니라
        # 인자 없이 호출해야 하는 메서드로 노출된다(`Lines()` → TextRange 컬렉션).
        tr = shape.TextFrame.TextRange
        lines = int(tr.Lines().Count)
        if line_starts_max is not None:
            # Lines(i,1).Start는 전체 텍스트 기준 1-based 문자 위치(문단 구분 '\r'은 앞 줄 Length에 포함)라
            # 줄들이 텍스트를 빈틈 없이 분할한다 — 슬라이드 단위 채우기가 "10번째 줄이 시작하는 오프셋"에서 자른다.
            k = min(lines, line_starts_max)
            starts, texts = [], []
            for i in range(1, k + 1):
                ln = tr.Lines(i, 1)
                starts.append(int(ln.Start) - 1)
                texts.append(str(ln.Text))
            return {'lines': lines, 'starts': starts, 'texts': texts}
        if want_bounds:
            # BoundHeight(pt)는 실제 렌더된 텍스트 전체의 세로 폭 — Pillow 글리프
            # 메트릭(ascent+descent)은 이보다 14~18% 작게 나오는 것으로 이미 확인돼
            # 있다(CLAUDE.md). 1줄 오차가 바로 시각적 결함(종료 텍스트박스가 본문과
            # 거의 겹침)으로 드러나는 measure_text_metrics() 용도로 함께 반환한다.
            return {'lines': lines, 'bound_height_pt': float(tr.BoundHeight)}
        return lines
    finally:
        presentation.Close()


def _discard_app():
    """캐시된 Application/keepalive 상태를 버린다(Quit 시도는 하되 실패는 무시).

    count_slide_lines()가 "The PowerPoint Frame window does not exist"류의
    드문 COM 상태 이상을 만났을 때, 원인을 정확히 진단하는 대신 캐시된 연결을
    통째로 버리고 다음 시도에서 완전히 새로 Dispatch하기 위한 것이다."""
    global _app, _keepalive_presentation

    if _keepalive_presentation is not None:
        try:
            _keepalive_presentation.Close()
        except Exception:
            pass
    _keepalive_presentation = None

    if _app is not None:
        try:
            _app.Quit()
        except Exception:
            pass
    _app = None


def count_slide_lines(temp_pptx_path: str, shape_index: int) -> int:
    """temp_pptx_path를 읽기 전용으로 열어 슬라이드 1의 Shapes(shape_index)
    (1-based, COM 관례) `.TextFrame.TextRange.Lines.Count`를 반환한다.

    프레젠테이션은 절대 저장하지 않고, 측정 성공/실패와 무관하게 항상 닫는다.
    어떤 예외가 나든 ComVerificationUnavailable 하나로 통일해서 올린다.

    첫 시도가 COM 상태 이상으로 실패하면, 캐시된 Application을 버리고 완전히
    새로 Dispatch해서 한 번만 재시도한다(원인 진단보다 복구를 우선한다).
    """
    try:
        return _open_and_measure(temp_pptx_path, shape_index)
    except Exception:
        pass

    try:
        _discard_app()
        return _open_and_measure(temp_pptx_path, shape_index)
    except Exception as exc:
        raise ComVerificationUnavailable(
            f"PowerPoint COM 줄 수 측정 실패: {exc}"
        ) from exc


def measure_text_metrics(temp_pptx_path: str, shape_index: int) -> dict:
    """temp_pptx_path를 읽기 전용으로 열어 슬라이드 1의 Shapes(shape_index)
    TextRange의 실제 줄 수와 세로 범위(BoundHeight, pt)를 함께 반환한다.

    {'lines': int, 'bound_height_pt': float} — (bound_height_pt/lines)가 실측
    1줄 높이(pt)다. count_slide_lines()와 동일한 재시도/예외 변환 규칙을 따른다
    (한 번 실패하면 캐시된 Application을 버리고 1회 재시도, 그래도 실패하면
    ComVerificationUnavailable 하나로 통일)."""
    try:
        return _open_and_measure(temp_pptx_path, shape_index, want_bounds=True)
    except Exception:
        pass

    try:
        _discard_app()
        return _open_and_measure(temp_pptx_path, shape_index, want_bounds=True)
    except Exception as exc:
        raise ComVerificationUnavailable(
            f"PowerPoint COM 텍스트 메트릭 측정 실패: {exc}"
        ) from exc


def measure_line_starts(temp_pptx_path: str, shape_index: int, max_lines: int = 10) -> dict:
    """슬라이드 1의 Shapes(shape_index) TextRange의 총 줄 수와 앞쪽 max_lines줄의 시작 문자 오프셋(0-based)·줄
    텍스트를 한 번의 프레젠테이션 열기로 돌려준다: {'lines': int, 'starts': [int], 'texts': [str]}.
    재시도/예외 변환 규칙은 count_slide_lines()와 동일."""
    try:
        return _open_and_measure(temp_pptx_path, shape_index, line_starts_max=max_lines)
    except Exception:
        pass

    try:
        _discard_app()
        return _open_and_measure(temp_pptx_path, shape_index, line_starts_max=max_lines)
    except Exception as exc:
        raise ComVerificationUnavailable(
            f"PowerPoint COM 줄 시작 측정 실패: {exc}"
        ) from exc


def shutdown() -> None:
    """생성했던 Application이 있으면 Quit()하고, COM을 uninitialize하고,
    내부 상태를 초기화한다. cleanup/finally 경로에서 호출되므로 어떤 예외도
    절대 밖으로 내보내지 않는다.
    """
    global _com_initialized, _available_cache

    _discard_app()

    if _com_initialized:
        try:
            import pythoncom

            pythoncom.CoUninitialize()
        except Exception:
            pass
    _com_initialized = False

    _available_cache = None
