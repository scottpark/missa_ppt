# -*- coding: utf-8 -*-
"""청년미사 2단계 — GitHub zip 다운로드 기반 자동 업데이트.

배포 형태를 PyInstaller exe에서 "Python 인터프리터 + run_missa.bat 런처"로 전환하면서,
`.py` 파일 교체만으로 즉시 최신화가 가능해졌다(재빌드·재배포 과정이 사라짐 — 구현 계획
`docs/청년미사 2단계 구현 계획.md` §1 참고). 대상 PC(성당 사무실 PC·담당자 개인 노트북)에
git 설치를 요구하지 않기 위해, git CLI 대신 GitHub REST API + 브랜치 zip 아카이브
다운로드만으로 구현한다(저장소가 public이라 인증 불필요).

완전 신규 leaf 모듈 — 다른 프로젝트 모듈을 import하지 않는다. `config.json`의
`last_update_commit` 읽기/쓰기는 호출부(missa_gui)가 `_load_config`/`_save_config`로
처리하고, 이 모듈은 dict를 그대로 주고받는다(파일 I/O를 이 모듈에 두지 않아 테스트가
config.json 실체에 의존하지 않게 한다).
"""
import shutil
import tempfile
import zipfile
from pathlib import Path

import requests

REPO = "scottpark/missa_ppt"
BRANCH = "master"
_API_COMMIT_URL = f"https://api.github.com/repos/{REPO}/commits/{BRANCH}"
_ARCHIVE_URL_TMPL = f"https://github.com/{REPO}/archive/{{sha}}.zip"

# apply_update()가 덮어쓸 대상(화이트리스트) — 이 패턴에 해당하는 파일/폴더만 설치 위치에
# 복사한다. 화이트리스트에 없는 모든 것(config.json, output/, reference/, cache/, .git/ 등
# 로컬 전용 상태)은 자동으로 보존된다 — 블랙리스트가 아니라 화이트리스트 방식이라 새로
# 추가되는 로컬 전용 폴더를 깜빡하고 블랙리스트에 빠뜨릴 위험이 없다.
# requirements.txt는 포함한다 — 새 의존 패키지가 생겨도 업데이트 후 install.bat을 다시 돌리면 설치된다(2026-10-21 PyMuPDF 누락 사례).
# install.bat은 일부러 뺀다: 저장소의 LF 줄바꿈 그대로 덮어쓰면 한글+chcp 65001에서 깨진다(zip 빌더가 CRLF로 정규화).
_UPDATE_FILE_GLOBS = ["*.py", "*.spec", "VERSION", "requirements.txt"]
_UPDATE_DIRS = ["assets"]
# 운영자 안내서: 저장소에서는 docs/ 아래에 있지만 배포 zip에서는 설치 폴더 최상위에 놓인다(tools/build_dist_zip.py).
# 설치 폴더에 **이미 같은 이름의 파일이 있을 때만** 갱신한다 — 성인 설치에 청년 안내서가 생기는 일이 없게(2026-10-05).
_UPDATE_README_GLOB = "docs/README_*_운영자.md"


class UpdateError(Exception):
    """네트워크 오류·API 응답 이상 등 업데이트 확인/적용 실패 공통 예외."""


def get_latest_commit_sha() -> str:
    """GitHub API로 BRANCH 최신 커밋 SHA 조회. public repo라 인증 불필요."""
    try:
        resp = requests.get(
            _API_COMMIT_URL,
            headers={"Accept": "application/vnd.github+json", "User-Agent": "missa_ppt-updater"},
            timeout=15,
        )
        resp.raise_for_status()
    except requests.RequestException as e:
        raise UpdateError(f"GitHub 업데이트 확인 실패: {e}") from e
    return resp.json()["sha"]


def check_for_update(config: dict) -> tuple:
    """config(dict)의 last_update_commit과 최신 SHA를 비교.
    반환: (업데이트 있음 여부, 최신 SHA)."""
    latest = get_latest_commit_sha()
    current = config.get("last_update_commit")
    return (current != latest, latest)


# 진행 콜백: progress(stage, done, total). stage는 아래 단계 상수, total이 None이면 전체를 알 수
# 없는 단계(애니메이션 막대로 표시). 콜백은 작업 스레드에서 불리므로 화면 갱신은 호출부가 책임진다.
STAGE_CHECK = "check"
STAGE_DOWNLOAD = "download"
STAGE_EXTRACT = "extract"
STAGE_INSTALL = "install"
STAGE_DONE = "done"

_DOWNLOAD_CHUNK = 64 * 1024


def _report(progress, stage, done, total):
    if progress is not None:
        progress(stage, done, total)


def _download_archive(sha: str, dest_zip: Path, progress=None) -> None:
    url = _ARCHIVE_URL_TMPL.format(sha=sha)
    try:
        resp = requests.get(url, timeout=60, stream=True)
        resp.raise_for_status()
        total = int(resp.headers.get("Content-Length") or 0) or None
        done = 0
        _report(progress, STAGE_DOWNLOAD, 0, total)
        with open(dest_zip, "wb") as f:
            for chunk in resp.iter_content(_DOWNLOAD_CHUNK):
                if not chunk:
                    continue
                f.write(chunk)
                done += len(chunk)
                _report(progress, STAGE_DOWNLOAD, done, total)
    except requests.RequestException as e:
        raise UpdateError(f"업데이트 다운로드 실패: {e}") from e


def _extract_archive_root(zip_path: Path, extract_to: Path, progress=None) -> Path:
    """zip을 풀고, GitHub 아카이브 특유의 단일 최상위 폴더({repo}-{sha}/)를 반환한다."""
    with zipfile.ZipFile(zip_path) as zf:
        members = zf.infolist()
        total = len(members)
        for i, m in enumerate(members, 1):
            zf.extract(m, extract_to)
            _report(progress, STAGE_EXTRACT, i, total)
    roots = [p for p in extract_to.iterdir() if p.is_dir()]
    if len(roots) != 1:
        raise UpdateError(f"예상치 못한 압축 구조(최상위 폴더 {len(roots)}개): {extract_to}")
    return roots[0]


def install_whitelisted(src_root: Path, install_dir: Path, progress=None) -> list:
    """src_root(압축 해제된 저장소 루트)에서 화이트리스트 파일/폴더만 install_dir로 복사.
    반환: 실제로 복사된 대상의 install_dir 기준 상대경로 목록(테스트·로그용).

    네트워크와 완전히 분리된 순수 파일 연산이라, 다운로드를 모킹하지 않고도 임의의
    src_root/install_dir 조합으로 직접 테스트할 수 있다. `progress`가 있으면 복사할 파일 수를
    미리 세어 파일 하나를 복사할 때마다 (STAGE_INSTALL, 복사한 수, 전체 수)를 알린다."""
    glob_files = [f for pattern in _UPDATE_FILE_GLOBS for f in src_root.rglob(pattern)]
    readmes = [f for f in src_root.glob(_UPDATE_README_GLOB) if (install_dir / f.name).exists()]
    dirs = [d for d in _UPDATE_DIRS if (src_root / d).is_dir()]
    total = (len(glob_files) + len(readmes)
             + sum(1 for d in dirs for f in (src_root / d).rglob("*") if f.is_file()))
    state = {"n": 0}

    def _tick():
        state["n"] += 1
        _report(progress, STAGE_INSTALL, state["n"], total)

    _report(progress, STAGE_INSTALL, 0, total)
    copied = []
    for f in glob_files:
        rel = f.relative_to(src_root)
        dest = install_dir / rel
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(f, dest)
        copied.append(rel)
        _tick()
    for f in readmes:
        shutil.copy2(f, install_dir / f.name)
        copied.append(Path(f.name))
        _tick()
    for d in dirs:
        dest_dir = install_dir / d
        if dest_dir.exists():
            shutil.rmtree(dest_dir)

        def _copy(src, dst, _tick=_tick):
            shutil.copy2(src, dst)
            _tick()

        shutil.copytree(src_root / d, dest_dir, copy_function=_copy)
        copied.append(Path(d))
    return copied


def apply_update(sha: str, install_dir: Path, progress=None) -> list:
    """{REPO}/archive/{sha}.zip을 받아 install_dir 위에 화이트리스트 파일만 덮어쓴다.
    반환: install_whitelisted()와 동일(복사된 상대경로 목록). `progress`는 위 진행 콜백."""
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        zip_path = tmp_path / "update.zip"
        _download_archive(sha, zip_path, progress)
        src_root = _extract_archive_root(zip_path, tmp_path / "extracted", progress)
        copied = install_whitelisted(src_root, install_dir, progress)
    _report(progress, STAGE_DONE, 1, 1)
    return copied
