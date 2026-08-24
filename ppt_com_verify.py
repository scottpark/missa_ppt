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
  안정적으로 완전히 숨겨서 돌릴 수 없다(버그가 아니라 받아들인 제약).
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
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = True
        _app = app

    if _keepalive_presentation is None:
        _keepalive_presentation = _app.Presentations.Add()

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


def _open_and_measure(temp_pptx_path: str, shape_index: int) -> int:
    app = _ensure_app()
    presentation = app.Presentations.Open(
        temp_pptx_path,
        ReadOnly=True,
        Untitled=False,
        WithWindow=True,
    )
    try:
        shape = presentation.Slides(1).Shapes(shape_index)
        # TextRange.Lines는 PowerPoint 객체 모델에서 선택적 인자(Start, Length)를
        # 받는 파라미터화된 속성이라, late-bound COM에서는 일반 속성이 아니라
        # 인자 없이 호출해야 하는 메서드로 노출된다(`Lines()` → TextRange 컬렉션).
        return int(shape.TextFrame.TextRange.Lines().Count)
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
