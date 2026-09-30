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
