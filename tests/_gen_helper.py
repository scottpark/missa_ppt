"""성인미사 회귀 픽스처(20260624/20260705/20260712)의 결과 PPT를 **임시 폴더**에 생성한다.

`output/{date}/`에는 입력 파일(참조 템플릿·JSON·화답송 악보·시작기도)만 git으로 추적하고, 결과 PPT와 log/는
추적하지 않는다. 테스트가 저장소의 `output/`을 덮어쓰면 추적 대상 바이너리가 매번 바뀌기 때문이다(2026-10-04).
여기서는 입력 파일만 임시 `output/{date}/`로 복사하고, `cwd`를 임시 폴더로 하여 `missa_to_ppt.py`를 실행한다
(프로그램의 `OUTPUT_ROOT='output'`이 cwd 기준 상대경로라 결과가 임시 폴더에만 생긴다).

같은 날짜는 프로세스 안에서 한 번만 생성해 재사용한다(테스트 파일들이 이 모듈을 같은 이름으로 import하므로
캐시가 공유된다). 임시 폴더는 프로세스 종료 시 지운다.
"""
import atexit
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent

# 날짜별 성가번호(회귀 기준 입력). 20260705는 성수축복 주일, 20260624는 평일(2차봉헌 없음).
FIXTURE_HYMNS = {
    "20260712": {"입당": "329", "봉헌": "221", "성체": "156", "2차봉헌": "221", "파견": "25"},
    "20260624": {"입당": "329", "봉헌": "221", "성체": "156", "파견": "25"},
    "20260705": {"입당": "287", "봉헌": "217", "성체": "152", "2차봉헌": "47", "파견": "286"},
}

_ROOT = [None]
_CACHE = {}  # date_str -> (결과 Path, stdout)


def _temp_root() -> Path:
    if _ROOT[0] is None:
        _ROOT[0] = Path(tempfile.mkdtemp(prefix="missa_fixture_"))
        atexit.register(shutil.rmtree, str(_ROOT[0]), ignore_errors=True)
    return _ROOT[0]


def generate_case(date_str: str, hymns: dict = None):
    """date_str의 결과 PPT를 임시 폴더에 생성하고 (결과 경로, stdout)을 반환한다(캐시됨)."""
    if date_str in _CACHE:
        return _CACHE[date_str]
    hymns = hymns if hymns is not None else FIXTURE_HYMNS[date_str]

    src = REPO_ROOT / "output" / date_str
    root = _temp_root()
    dst = root / "output" / date_str
    dst.mkdir(parents=True, exist_ok=True)
    for f in src.iterdir():
        if not f.is_file() or f.name.startswith("~$"):
            continue
        if f.suffix.lower() == ".pptx" and f.name.startswith(f"{date_str}_"):
            continue  # 예전 결과 PPT(추적 제외 대상)는 입력이 아니다
        shutil.copy2(f, dst / f.name)

    args = [sys.executable, str(REPO_ROOT / "missa_to_ppt.py"), date_str]
    for k, v in hymns.items():
        args += [f"--{k}", v]
    result = subprocess.run(args, cwd=root, capture_output=True, text=True, encoding="utf-8")
    assert result.returncode == 0, (
        f"{date_str} 생성 실패 (exit={result.returncode})\n"
        f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
    )
    assert "모든 검증 통과!" in result.stdout, (
        f"{date_str} 내부 검증(validate) 실패\nstdout:\n{result.stdout}"
    )
    outs = [f for f in dst.glob(f"{date_str}_*.pptx") if not f.name.startswith("~$")]
    assert outs, f"{date_str} 출력 PPT를 찾을 수 없음"
    path = max(outs, key=lambda f: f.stat().st_mtime)
    _CACHE[date_str] = (path, result.stdout)
    return _CACHE[date_str]


def generated_path(date_str: str) -> Path:
    """이미 생성한 결과 PPT 경로(생성 전이면 생성한다)."""
    return generate_case(date_str)[0]
