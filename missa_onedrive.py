# -*- coding: utf-8 -*-
"""청년미사 2단계 — OneDrive 접근 (Microsoft Graph API, Device Code Flow).

성당 공용 계정(brokenbaykccppt@gmail.com) 하나로 로그인한다 — 개인별 계정 분기가 없다
(구현 계획 `docs/청년미사 2단계 구현 계획.md` §2.1). Public client(secret 불필요)로 Azure AD에
등록된 앱을 사용하며, CLIENT_ID는 비밀이 아니라 공개 저장소에 있어도 안전하다.

완전 신규 leaf 모듈 — 다른 프로젝트 모듈을 import하지 않는다(missa_gui.py/
missa_youth_hymn_pdf.py가 이 모듈을 호출하는 방향만 있다). 토큰 캐시 파일 I/O는 이 모듈이
직접 수행한다(GitHub 업데이터와 달리, MSAL의 SerializableTokenCache 자체가 파일 스키마를
가지고 있어 호출부에 dict로 넘겨줄 이유가 없다).
"""
import sys
import urllib.parse
from pathlib import Path

import msal
import requests

CLIENT_ID = "21b7d1a4-87bb-4148-8ae8-8b57231c125b"
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["Files.ReadWrite"]
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"

_SCRIPT_DIR = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
TOKEN_CACHE_FILE = _SCRIPT_DIR / 'msal_token_cache.bin'


class OneDriveError(Exception):
    """인증 실패·API 응답 이상 등 OneDrive 접근 공통 예외."""


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_FILE.exists():
        cache.deserialize(TOKEN_CACHE_FILE.read_text(encoding='utf-8'))
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    # has_state_changed는 acquire_token_silent가 캐시된 토큰을 그대로 재사용한 경우 False라,
    # 매번 디스크에 쓰지 않아도 된다(불필요한 파일 I/O 회피).
    if cache.has_state_changed:
        TOKEN_CACHE_FILE.write_text(cache.serialize(), encoding='utf-8')


# [app, cache] 프로세스 수명 캐시 (2026-09-26, 버그2). msal.PublicClientApplication(...)
# 생성 자체가 매번 약 0.9~1.0초 걸린다(AUTHORITY="https://login.microsoftonline.com/consumers"
# 가 네트워크 기반 authority라 생성자 안에서 매번 instance discovery를 다시 하는 것으로
# 보임 — 코디네이터 실측). 반면 그 뒤의 acquire_token_silent()는 0초. 이 앱을 매 HTTP
# 요청마다 새로 만들면(`_headers()`가 매번 `get_access_token()`을 부르므로) OneDrive
# 커스텀 파일 브라우저에서 폴더 하나 여는 것만으로도 초 단위 지연이 누적됐다(실측 3~6초).
# MSAL 공식 권장 사용 패턴대로 프로세스 수명 동안 한 번만 생성해 재사용한다.
_APP_CACHE = [None, None]


def _get_app():
    """캐시된 (app, cache) 쌍을 반환한다. 최초 호출에서만 디스크의 msal_token_cache.bin을
    읽어(`_load_cache()`) PublicClientApplication을 생성하고, 이후 호출은 그 인스턴스를
    그대로 재사용한다."""
    if _APP_CACHE[0] is None:
        cache = _load_cache()
        app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
        _APP_CACHE[0] = app
        _APP_CACHE[1] = cache
    return _APP_CACHE[0], _APP_CACHE[1]


def get_access_token(on_device_code=None) -> str:
    """캐시된 토큰을 우선 쓰고, 없거나 만료됐으면 device code flow로 새로 로그인한다.

    device code flow는 콘솔에 로그인 안내를 출력하고 담당자의 로그인을 기다린다 — 성당
    공용 계정으로 사람이 직접 로그인해야 하므로 완전 자동화할 수 없다(§2.1). 토큰 캐시
    덕분에 최초 1회 로그인 후에는 refresh_token으로 조용히 갱신된다.

    app(`msal.PublicClientApplication`)과 그 token_cache는 `_get_app()`을 통해 프로세스
    수명 동안 재사용된다(2026-09-26, 버그2) — 매 호출마다 새로 만드는 것 자체가 병목이었다.

    on_device_code: flow(dict)를 받는 콜백(선택). device code flow에 진입할 때(캐시된
    토큰이 없을 때) 호출된다. pythonw.exe처럼 콘솔이 없는 배포 환경에서는 print()만으로는
    사용자가 로그인 코드를 볼 수 없어 무한 대기처럼 보이므로, GUI 알림이 필요한 호출부
    (missa_gui.ensure_onedrive_login())가 이 콜백으로 코드를 화면에 띄운다. 콜백이
    없으면(CLI/테스트 등) 기존처럼 print()만 한다."""
    app, cache = _get_app()

    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if 'user_code' not in flow:
            raise OneDriveError(f'device code flow 시작 실패: {flow}')
        print(f"  [OneDrive 로그인 필요] {flow['message']}")
        if on_device_code is not None:
            on_device_code(flow)
        result = app.acquire_token_by_device_flow(flow)

    _save_cache(cache)

    if not result or 'access_token' not in result:
        detail = (result or {}).get('error_description', result)
        raise OneDriveError(f'OneDrive 로그인 실패: {detail}')
    return result['access_token']


def _headers(extra: dict = None) -> dict:
    headers = {'Authorization': f'Bearer {get_access_token()}'}
    if extra:
        headers.update(extra)
    return headers


# 2026-09-26 실측 발견: 성당 공용 계정에서 "PPT 문서"는 진짜 폴더가 아니라 다른 드라이브를
# 가리키는 공유 바로가기(remoteItem)다. 경로 기반 주소(`root:/PPT 문서/...:`)는 이 바로가기를
# 통과하지 못해 하위 경로 조회가 404/422로 실패한다 — remoteItem이면
# `/drives/{driveId}/items/{itemId}`로 주소 체계를 전환해야 실제 하위 폴더에 닿는다.
# 세그먼트 이름 → (driveId, itemId)|None(=remoteItem 아님) 캐시. 프로세스 수명 동안만
# 유효하면 충분하다(매 프로세스 시작마다 새로 조회해도 비용은 API 호출 1회뿐).
_REMOTE_ITEM_CACHE = {}


def _resolve_remote_base(first_segment: str):
    """OneDrive 최상위 경로 세그먼트 하나가 remoteItem(공유 바로가기)인지 실제로 GET해서
    확인한다. remoteItem이면 (driveId, itemId)를, 아니면(평범한 폴더 또는 존재하지 않는
    경로) None을 반환하고 결과를 _REMOTE_ITEM_CACHE에 캐시한다.

    404를 여기서 raise_for_status()로 터뜨리지 않는다 — "경로가 아예 없다"는 진짜 에러
    메시지(파일명 포함)는 호출부(list_children 등)의 기존 404 처리가 담당해야 하므로,
    여기서는 "remoteItem 여부만" 판단하고 그 외 상태는 전부 None(=기존 root:/...: 경로를
    그대로 쓰라는 신호)으로 넘긴다."""
    if first_segment in _REMOTE_ITEM_CACHE:
        return _REMOTE_ITEM_CACHE[first_segment]

    encoded = urllib.parse.quote(first_segment, safe='')
    url = f'{GRAPH_ROOT}/me/drive/root:/{encoded}:'
    resp = requests.get(url, headers=_headers())

    result = None
    if resp.status_code != 404:
        resp.raise_for_status()
        remote = resp.json().get('remoteItem')
        if remote:
            result = (remote['parentReference']['driveId'], remote['id'])

    _REMOTE_ITEM_CACHE[first_segment] = result
    return result


def _ensure_base_resolved(path: str) -> None:
    """path의 첫 세그먼트가 remoteItem인지 미리 확인해 _REMOTE_ITEM_CACHE를 채운다.

    `_item_url()`은 순수 함수(네트워크 호출 없음)로 유지한다 — 호출부(list_children 등,
    실제 I/O를 하는 공개 함수)는 URL을 만들기 **전에** 반드시 이 함수를 먼저 호출해야
    `_item_url()`이 remoteItem 경로로 전환된다."""
    first = path.strip('/').split('/')[0] if path.strip('/') else ''
    if first:
        _resolve_remote_base(first)


def _item_url(path: str, suffix: str = '') -> str:
    """path를 Graph의 path-based addressing URL로 변환한다. 순수 함수(네트워크 호출 없음).

    path='' (또는 '/')는 드라이브 루트를 가리킨다 — 이때는 콜론 구문(`root:/...:`) 자체를
    쓰지 않는다(Graph는 루트에 빈 path-segment를 붙이는 걸 허용하지 않는다).

    첫 세그먼트가 `_REMOTE_ITEM_CACHE`에 remoteItem으로 이미 해석돼 있으면(호출부가 미리
    `_ensure_base_resolved()`를 불렀다는 뜻) `/drives/{driveId}/items/{itemId}...` 로
    주소를 전환한다. 캐시에 없거나(아직 해석 안 됨) None(remoteItem 아님)이면 기존
    `/me/drive/root:/...` 를 그대로 쓴다."""
    path = path.strip('/')
    if not path:
        return f'{GRAPH_ROOT}/me/drive/root{suffix}'

    segments = path.split('/')
    base = _REMOTE_ITEM_CACHE.get(segments[0])
    if base is None:
        encoded = urllib.parse.quote(path, safe='/')
        return f'{GRAPH_ROOT}/me/drive/root:/{encoded}:{suffix}'

    drive_id, item_id = base
    rest = segments[1:]
    if not rest:
        return f'{GRAPH_ROOT}/drives/{drive_id}/items/{item_id}{suffix}'
    encoded_rest = urllib.parse.quote('/'.join(rest), safe='/')
    return f'{GRAPH_ROOT}/drives/{drive_id}/items/{item_id}:/{encoded_rest}:{suffix}'


def list_children(path: str) -> list:
    """path 폴더의 자식 항목(Graph driveItem dict) 목록을 반환한다.

    결과가 여러 페이지로 나뉘면(@odata.nextLink) 전부 모아서 반환한다 — 나주노/야훼이레
    성가 폴더처럼 파일이 계속 쌓이는 폴더에서 일부만 보고 누락되는 것을 막기 위함."""
    _ensure_base_resolved(path)
    url = _item_url(path, '/children')
    items = []
    while url:
        resp = requests.get(url, headers=_headers())
        if resp.status_code == 404:
            raise FileNotFoundError(f'OneDrive 경로를 찾을 수 없습니다: {path}')
        resp.raise_for_status()
        data = resp.json()
        items.extend(data.get('value', []))
        url = data.get('@odata.nextLink')
    return items


def get_item_metadata(path: str) -> dict:
    """path의 OneDrive 항목 메타데이터(Graph driveItem dict, children 목록은 제외)를
    조회한다. K2(로컬 캐시 재사용, 2026-09-27)가 `lastModifiedDateTime`/`eTag`로 원격 변경
    여부를 판단할 때 쓴다 — `list_children()`이 이미 그 폴더를 조회해 캐시에 갖고 있으면
    호출부는 이 함수 대신 그 캐시된 항목 목록에서 이름으로 찾아 재사용해야 한다(이 함수는
    캐시에 없을 때의 폴백 경로)."""
    _ensure_base_resolved(path)
    url = _item_url(path)
    resp = requests.get(url, headers=_headers())
    if resp.status_code == 404:
        raise FileNotFoundError(f'OneDrive 경로를 찾을 수 없습니다: {path}')
    resp.raise_for_status()
    return resp.json()


def download_file(path: str, dest: Path) -> Path:
    """path의 OneDrive 파일을 dest(로컬 경로)로 내려받는다. dest의 상위 폴더가 없으면
    만든다. 반환값은 dest 그대로(호출부 체이닝 편의)."""
    _ensure_base_resolved(path)
    url = _item_url(path, '/content')
    resp = requests.get(url, headers=_headers(), stream=True)
    if resp.status_code == 404:
        raise FileNotFoundError(f'OneDrive 파일을 찾을 수 없습니다: {path}')
    resp.raise_for_status()
    dest = Path(dest)
    dest.parent.mkdir(parents=True, exist_ok=True)
    with open(dest, 'wb') as f:
        for chunk in resp.iter_content(chunk_size=1 << 16):
            if chunk:
                f.write(chunk)
    return dest


def upload_file(path: str, local_path: Path) -> None:
    """local_path 파일을 OneDrive의 path 위치에 업로드한다(단순 업로드 — Graph 문서 기준
    250MB 이하에서 충분, 이 프로젝트의 pptx 산출물 크기에 넉넉히 해당).

    상위 폴더가 없으면 실패한다 — Graph의 path-based addressing은 중간 폴더를 자동으로
    만들어주지 않으므로, 호출부가 먼저 ensure_folder()로 경로를 준비해야 한다."""
    _ensure_base_resolved(path)
    url = _item_url(path, '/content')
    data = Path(local_path).read_bytes()
    resp = requests.put(
        url, headers=_headers({'Content-Type': 'application/octet-stream'}), data=data,
    )
    resp.raise_for_status()


def ensure_folder(path: str) -> None:
    """path의 모든 경로 세그먼트를 위에서부터 순회하며, 없는 폴더만 생성한다.

    세그먼트별로 존재 여부를 먼저 확인(GET)한 뒤 없을 때만 생성(POST)한다 — 매번 무조건
    생성을 시도하면 이미 있는 폴더(대부분의 경우)에도 매번 conflictBehavior 처리가 걸려
    비효율적이고, 동시 실행 경쟁 시 409를 정상 케이스로 다뤄야 하는 부담이 커진다."""
    _ensure_base_resolved(path)
    segments = [s for s in path.strip('/').split('/') if s]
    built = ''
    for seg in segments:
        parent = built
        built = f'{built}/{seg}' if built else seg

        resp = requests.get(_item_url(built), headers=_headers())
        if resp.status_code == 200:
            continue
        if resp.status_code != 404:
            resp.raise_for_status()

        create_url = _item_url(parent, '/children')
        body = {'name': seg, 'folder': {}, '@microsoft.graph.conflictBehavior': 'fail'}
        resp = requests.post(
            create_url, headers=_headers({'Content-Type': 'application/json'}), json=body,
        )
        if resp.status_code in (200, 201):
            continue
        if resp.status_code == 409:
            # 폴더 존재 확인(GET)과 생성(POST) 사이의 경쟁 — 이미 다른 프로세스가 만들었다면
            # 정상 케이스로 취급하고 다음 세그먼트로 진행한다.
            continue
        resp.raise_for_status()
