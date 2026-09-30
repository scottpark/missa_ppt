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
_UPDATE_FILE_GLOBS = ["*.py", "*.spec", "VERSION"]
_UPDATE_DIRS = ["assets"]


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


def _download_archive(sha: str, dest_zip: Path) -> None:
    url = _ARCHIVE_URL_TMPL.format(sha=sha)
    try:
        resp = requests.get(url, timeout=60)
        resp.raise_for_status()
    except requests.RequestException as e:
        raise UpdateError(f"업데이트 다운로드 실패: {e}") from e
    dest_zip.write_bytes(resp.content)


def _extract_archive_root(zip_path: Path, extract_to: Path) -> Path:
    """zip을 풀고, GitHub 아카이브 특유의 단일 최상위 폴더({repo}-{sha}/)를 반환한다."""
    with zipfile.ZipFile(zip_path) as zf:
        zf.extractall(extract_to)
    roots = [p for p in extract_to.iterdir() if p.is_dir()]
    if len(roots) != 1:
        raise UpdateError(f"예상치 못한 압축 구조(최상위 폴더 {len(roots)}개): {extract_to}")
    return roots[0]


def install_whitelisted(src_root: Path, install_dir: Path) -> list:
    """src_root(압축 해제된 저장소 루트)에서 화이트리스트 파일/폴더만 install_dir로 복사.
    반환: 실제로 복사된 대상의 install_dir 기준 상대경로 목록(테스트·로그용).

    네트워크와 완전히 분리된 순수 파일 연산이라, 다운로드를 모킹하지 않고도 임의의
    src_root/install_dir 조합으로 직접 테스트할 수 있다."""
    copied = []
    for pattern in _UPDATE_FILE_GLOBS:
        for f in src_root.rglob(pattern):
            rel = f.relative_to(src_root)
            dest = install_dir / rel
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(f, dest)
            copied.append(rel)
    for d in _UPDATE_DIRS:
        src_dir = src_root / d
        if not src_dir.is_dir():
            continue
        dest_dir = install_dir / d
        if dest_dir.exists():
            shutil.rmtree(dest_dir)
        shutil.copytree(src_dir, dest_dir)
        copied.append(Path(d))
    return copied


def apply_update(sha: str, install_dir: Path) -> list:
    """{REPO}/archive/{sha}.zip을 받아 install_dir 위에 화이트리스트 파일만 덮어쓴다.
    반환: install_whitelisted()와 동일(복사된 상대경로 목록)."""
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        zip_path = tmp_path / "update.zip"
        _download_archive(sha, zip_path)
        src_root = _extract_archive_root(zip_path, tmp_path / "extracted")
        return install_whitelisted(src_root, install_dir)
