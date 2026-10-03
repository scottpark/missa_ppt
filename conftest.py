"""tests/ 안의 테스트 파일에서 저장소 루트의 애플리케이션 모듈(missa_gui 등)을 바로
import할 수 있도록 저장소 루트를 sys.path에 추가한다.

tests/에는 `__init__.py`가 없어(rootless 레이아웃) pytest의 기본 "prepend" import
모드가 각 테스트 파일 자신의 디렉터리(tests/)만 sys.path에 넣는다 — `python -m pytest`로
저장소 루트에서 실행하면 파이썬 자체가 CWD를 추가해주므로 우연히 동작하지만, 바로 `pytest`
명령을 쓰면(CLAUDE.md에 문서화된 실행 방식) `import missa_gui`가 실패한다. pytest는
루트의 이 conftest.py를 항상 먼저 로드하며, 로드 과정에서 이 파일의 디렉터리(저장소 루트)를
sys.path에 넣어주므로 호출 방식과 무관하게 항상 안전하다."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))


import pytest


@pytest.fixture(autouse=True)
def _youth_ppt_mode_isolated():
    """청년미사 OneDrive 접속 모드(`missa_gui._YOUTH_PPT_STATE`)는 프로세스 전역 상태다.
    모든 테스트를 실제 config.json과 무관하게 '폴백'(기존 Graph 동작)으로 시작시키고, 끝나면
    비워서 한 테스트의 로컬 모드 설정이 다른 테스트로 새지 않게 한다. 로컬 모드를 검증하는
    테스트는 이 상태를 직접 덮어쓴다."""
    import missa_gui
    missa_gui._YOUTH_PPT_STATE['mode'] = 'fallback'
    missa_gui._YOUTH_PPT_STATE['folder'] = None
    missa_gui._reset_adult_ppt_folder()
    missa_gui._last_output_mass[0] = None
    yield
    missa_gui._reset_youth_ppt_mode()
    missa_gui._reset_adult_ppt_folder()
    missa_gui._last_output_mass[0] = None
