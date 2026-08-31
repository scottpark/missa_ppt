"""
missa_to_ppt.py 모듈 분리 리팩토링 전/후 출력이 100% 동일한지 확인하는 골든 마스터 diff 도구.

순수 파일 분리 리팩토링(로직 무변경)이므로, 각 단계 이후에도 동일 입력에 대해 출력 pptx가
완전히 동일해야 한다. 회귀/프로그레션 테스트는 "이미 알려진 버그 패턴"만 확인하므로, 이 도구는
"체크리스트에 없는" 의도치 않은 변경까지 전부 잡기 위한 보조 안전망이다.

사용법
------
    python tools/golden_diff.py snapshot   # 현재 코드로 출력 생성 → _golden/ 에 저장 (기준선)
    python tools/golden_diff.py check      # 다시 생성 후 _golden/ 과 비교, 차이나면 실패

_golden/ 은 로컬 전용 산출물이라 git으로 추적하지 않는다(.gitignore 참고).
"""
from __future__ import annotations

import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
GOLDEN_DIR = REPO_ROOT / "_golden"

# 회귀 테스트(test_missa_regression.py)의 MASS_CASES에 성수축복(20260705)을 더한 3케이스.
# 20260705는 주일미사와 구조는 같지만 실제 문구 치환이 반영되는지까지 함께 확인한다.
CASES = [
    {
        "id": "sunday",
        "date": "20260712",
        "hymns": {"입당": "329", "봉헌": "221", "성체": "156", "2차봉헌": "221", "파견": "25"},
    },
    {
        "id": "weekday",
        "date": "20260624",
        "hymns": {"입당": "329", "봉헌": "221", "성체": "156", "파견": "25"},
    },
    {
        "id": "holy_water_blessing",
        "date": "20260705",
        "hymns": {"입당": "287", "봉헌": "217", "성체": "152", "2차봉헌": "47", "파견": "286"},
    },
]


def _find_output(date_str: str) -> Path:
    folder = REPO_ROOT / "output" / date_str
    candidates = [
        f for f in folder.glob(f"{date_str}_*.pptx")
        if not f.name.startswith("~$")
    ]
    if not candidates:
        raise FileNotFoundError(f"{date_str} 출력 PPT를 찾을 수 없음")
    return max(candidates, key=lambda f: f.stat().st_mtime)


def _generate(case: dict) -> Path:
    date_str = case["date"]
    folder = REPO_ROOT / "output" / date_str

    locked = list(folder.glob("~$*.pptx"))
    if locked:
        raise RuntimeError(f"{locked[0].name} 이(가) PowerPoint에서 열려 있어 생성을 건너뜀")

    args = [sys.executable, str(REPO_ROOT / "missa_to_ppt.py"), date_str]
    for k, v in case["hymns"].items():
        args += [f"--{k}", v]

    result = subprocess.run(
        args, cwd=REPO_ROOT, capture_output=True, text=True, encoding="utf-8",
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"{date_str} 생성 실패 (exit={result.returncode})\n"
            f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
        )
    return _find_output(date_str)


def _extract(pptx_path: Path, dest_dir: Path) -> None:
    if dest_dir.exists():
        shutil.rmtree(dest_dir)
    dest_dir.mkdir(parents=True)
    with zipfile.ZipFile(pptx_path) as zf:
        zf.extractall(dest_dir)


def _diff_trees(golden_dir: Path, current_dir: Path) -> list[str]:
    golden_files = {p.relative_to(golden_dir).as_posix() for p in golden_dir.rglob("*") if p.is_file()}
    current_files = {p.relative_to(current_dir).as_posix() for p in current_dir.rglob("*") if p.is_file()}

    problems = []
    for rel in sorted(golden_files - current_files):
        problems.append(f"골든에는 있으나 현재 출력에 없음: {rel}")
    for rel in sorted(current_files - golden_files):
        problems.append(f"현재 출력에만 있고 골든에는 없음: {rel}")
    for rel in sorted(golden_files & current_files):
        a = (golden_dir / rel).read_bytes()
        b = (current_dir / rel).read_bytes()
        if a != b:
            problems.append(f"내용 다름: {rel} (golden {len(a)} bytes vs current {len(b)} bytes)")
    return problems


def snapshot() -> None:
    if GOLDEN_DIR.exists():
        shutil.rmtree(GOLDEN_DIR)
    GOLDEN_DIR.mkdir(parents=True)

    for case in CASES:
        print(f"[snapshot] {case['id']} ({case['date']}) 생성 중...")
        out = _generate(case)
        _extract(out, GOLDEN_DIR / case["id"])
        print(f"[snapshot] {case['id']} → {GOLDEN_DIR / case['id']}")

    print("\n골든 마스터 스냅샷 생성 완료.")


def check() -> bool:
    if not GOLDEN_DIR.exists():
        print("골든 마스터가 없습니다. 먼저 'python tools/golden_diff.py snapshot'을 실행하세요.", file=sys.stderr)
        return False

    all_ok = True
    for case in CASES:
        print(f"[check] {case['id']} ({case['date']}) 재생성 중...")
        out = _generate(case)
        current_dir = REPO_ROOT / "_golden_check_tmp" / case["id"]
        _extract(out, current_dir)

        problems = _diff_trees(GOLDEN_DIR / case["id"], current_dir)
        if problems:
            all_ok = False
            print(f"[check] {case['id']}: 차이 발견 ({len(problems)}건)")
            for p in problems:
                print(f"    - {p}")
        else:
            print(f"[check] {case['id']}: 골든 마스터와 완전히 동일")

    shutil.rmtree(REPO_ROOT / "_golden_check_tmp", ignore_errors=True)

    if all_ok:
        print("\n모든 케이스가 골든 마스터와 동일합니다.")
    else:
        print("\n차이가 발견됐습니다 — 순수 이동 리팩토링이라면 이는 버그입니다.")
    return all_ok


def main() -> None:
    if len(sys.argv) != 2 or sys.argv[1] not in ("snapshot", "check"):
        print(__doc__)
        sys.exit(1)

    if sys.argv[1] == "snapshot":
        snapshot()
    else:
        ok = check()
        sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
