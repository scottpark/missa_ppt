# -*- coding: utf-8 -*-
"""missa_onedrive.py 프로그레션 테스트 (청년미사 2단계 §2 — OneDrive Graph API 연동).

실제 네트워크/MSAL 로그인은 절대 발생시키지 않는다 — requests.get/put/post와
msal.PublicClientApplication을 전부 monkeypatch로 대체해서 순수 로직(URL 조합, 페이지네이션,
화이트리스트식 폴더 생성, 에러 래핑)만 검증한다.
"""
from pathlib import Path

import pytest

import missa_onedrive as od


class _FakeResponse:
    def __init__(self, status_code=200, json_data=None, content=b""):
        self.status_code = status_code
        self._json = json_data
        self.content = content

    def raise_for_status(self):
        if self.status_code >= 400:
            raise od.requests.HTTPError(f"{self.status_code} error")

    def json(self):
        return self._json

    def iter_content(self, chunk_size=1):
        yield self.content


@pytest.fixture(autouse=True)
def _clear_msal_app_cache(monkeypatch):
    """G2(2026-09-26): get_access_token()이 이제 msal.PublicClientApplication을 프로세스
    수명 동안 재사용하는 `_APP_CACHE`를 쓴다 — 이 모듈 레벨 캐시를 그대로 두면, 한 테스트가
    monkeypatch한 가짜 msal.PublicClientApplication으로 만든 app이 캐시에 남아 **다음
    테스트가 자기 자신의 monkeypatch를 걸어도 무시되고 이전 테스트의 가짜 app이 그대로
    재사용**되는 오염이 생긴다(실측: 이 fixture 없이 O2b/O2c가 O2 실행 순서에 의존해 실패함).
    매 테스트 시작 전 `[None, None]`으로 monkeypatch(자동 원복)해 격리한다."""
    monkeypatch.setattr(od, '_APP_CACHE', [None, None])


def test_O1_item_url_root_has_no_colon_segment():
    """O1: 빈 경로/루트는 콜론 구문 없이 root 엔드포인트를 가리켜야 한다."""
    assert od._item_url('', '/children') == f'{od.GRAPH_ROOT}/me/drive/root/children'
    assert od._item_url('/', '/children') == f'{od.GRAPH_ROOT}/me/drive/root/children'


def test_O1b_item_url_encodes_spaces_and_korean():
    url = od._item_url('PPT 문서/20.청년 미사', '/children')
    assert url.startswith(f'{od.GRAPH_ROOT}/me/drive/root:/')
    assert url.endswith(':/children')
    assert ' ' not in url  # 공백은 인코딩돼야 함
    assert '%20' in url
    assert '%2520' not in url  # 이중 인코딩은 없어야 함


def test_O1c_item_url_preserves_slash_separators():
    url = od._item_url('a/b/c')
    assert url == f'{od.GRAPH_ROOT}/me/drive/root:/a/b/c:'


def test_O2_get_access_token_uses_silent_cache_first(monkeypatch, tmp_path):
    """O2: 캐시된 계정이 있으면 device flow를 아예 시작하지 않는다."""
    monkeypatch.setattr(od, 'TOKEN_CACHE_FILE', tmp_path / 'cache.bin')

    class _FakeApp:
        def __init__(self, *a, **k):
            pass

        def get_accounts(self):
            return [{'username': 'brokenbaykccppt@gmail.com'}]

        def acquire_token_silent(self, scopes, account):
            return {'access_token': 'silent-token-123'}

        def initiate_device_flow(self, scopes):
            raise AssertionError('device flow는 호출되면 안 됨(캐시 히트 시나리오)')

    monkeypatch.setattr(od.msal, 'PublicClientApplication', _FakeApp)
    assert od.get_access_token() == 'silent-token-123'


def test_O2b_get_access_token_falls_back_to_device_flow(monkeypatch, tmp_path, capsys):
    """O2b: 캐시 미스면 device code flow로 진행하고, 안내 메시지를 출력한다."""
    monkeypatch.setattr(od, 'TOKEN_CACHE_FILE', tmp_path / 'cache.bin')

    class _FakeApp:
        def __init__(self, *a, **k):
            pass

        def get_accounts(self):
            return []

        def initiate_device_flow(self, scopes):
            return {'user_code': 'ABC123', 'message': '코드 ABC123을 입력하세요'}

        def acquire_token_by_device_flow(self, flow):
            return {'access_token': 'device-token-456'}

    monkeypatch.setattr(od.msal, 'PublicClientApplication', _FakeApp)
    token = od.get_access_token()
    assert token == 'device-token-456'
    assert 'ABC123' in capsys.readouterr().out


def test_O2c_get_access_token_raises_on_failure(monkeypatch, tmp_path):
    """O2c: 로그인 실패(access_token 없는 응답)는 OneDriveError로 래핑된다."""
    monkeypatch.setattr(od, 'TOKEN_CACHE_FILE', tmp_path / 'cache.bin')

    class _FakeApp:
        def __init__(self, *a, **k):
            pass

        def get_accounts(self):
            return []

        def initiate_device_flow(self, scopes):
            return {'user_code': 'X', 'message': 'm'}

        def acquire_token_by_device_flow(self, flow):
            return {'error': 'authorization_declined', 'error_description': '사용자가 거부함'}

    monkeypatch.setattr(od.msal, 'PublicClientApplication', _FakeApp)
    with pytest.raises(od.OneDriveError):
        od.get_access_token()


def test_O3_list_children_follows_pagination(monkeypatch):
    """O3: @odata.nextLink가 있으면 계속 따라가 전부 모은다."""
    monkeypatch.setattr(od, 'get_access_token', lambda: 'tok')
    monkeypatch.setitem(od._REMOTE_ITEM_CACHE, '나주노 성가', None)  # remoteItem 해석 스킵
    calls = []

    def _fake_get(url, headers=None, **kwargs):
        calls.append(url)
        if len(calls) == 1:
            return _FakeResponse(json_data={'value': [{'name': 'a.pptx'}], '@odata.nextLink': 'http://next'})
        return _FakeResponse(json_data={'value': [{'name': 'b.pptx'}]})

    monkeypatch.setattr(od.requests, 'get', _fake_get)
    items = od.list_children('나주노 성가')
    assert [i['name'] for i in items] == ['a.pptx', 'b.pptx']
    assert len(calls) == 2


def test_O3b_list_children_missing_path_raises_file_not_found(monkeypatch):
    monkeypatch.setattr(od, 'get_access_token', lambda: 'tok')
    monkeypatch.setattr(od.requests, 'get', lambda *a, **k: _FakeResponse(status_code=404))
    with pytest.raises(FileNotFoundError):
        od.list_children('없는 폴더')


def test_O4_download_file_writes_content_and_makes_parent_dirs(monkeypatch, tmp_path):
    monkeypatch.setattr(od, 'get_access_token', lambda: 'tok')
    monkeypatch.setitem(od._REMOTE_ITEM_CACHE, '나주노 성가', None)  # remoteItem 해석 스킵
    monkeypatch.setattr(od.requests, 'get', lambda *a, **k: _FakeResponse(content=b'pptx-bytes'))
    dest = tmp_path / 'nested' / 'out.pptx'
    result = od.download_file('나주노 성가/447.pptx', dest)
    assert result == dest
    assert dest.read_bytes() == b'pptx-bytes'


def test_O4b_download_file_missing_raises_file_not_found(monkeypatch, tmp_path):
    monkeypatch.setattr(od, 'get_access_token', lambda: 'tok')
    monkeypatch.setattr(od.requests, 'get', lambda *a, **k: _FakeResponse(status_code=404))
    with pytest.raises(FileNotFoundError):
        od.download_file('없음.pptx', tmp_path / 'x.pptx')


def test_O5_upload_file_puts_raw_bytes(monkeypatch, tmp_path):
    monkeypatch.setattr(od, 'get_access_token', lambda: 'tok')
    monkeypatch.setitem(od._REMOTE_ITEM_CACHE, '20.청년 미사', None)  # remoteItem 해석 스킵
    src = tmp_path / 'local.pptx'
    src.write_bytes(b'local-bytes')

    captured = {}

    def _fake_put(url, headers=None, data=None, **kwargs):
        captured['url'] = url
        captured['data'] = data
        captured['headers'] = headers
        return _FakeResponse()

    monkeypatch.setattr(od.requests, 'put', _fake_put)
    od.upload_file('20.청년 미사/2026/결과.pptx', src)
    assert captured['data'] == b'local-bytes'
    assert captured['headers']['Content-Type'] == 'application/octet-stream'


def test_O6_ensure_folder_skips_existing_segments(monkeypatch):
    """O6: 이미 존재하는(GET 200) 세그먼트는 생성(POST) 시도를 하지 않는다."""
    monkeypatch.setattr(od, 'get_access_token', lambda: 'tok')
    monkeypatch.setitem(od._REMOTE_ITEM_CACHE, 'a', None)  # remoteItem 해석 스킵
    get_calls = []
    post_calls = []

    def _fake_get(url, headers=None, **kwargs):
        get_calls.append(url)
        return _FakeResponse(status_code=200)  # 전부 이미 존재

    def _fake_post(url, headers=None, json=None, **kwargs):
        post_calls.append((url, json))
        return _FakeResponse(status_code=201)

    monkeypatch.setattr(od.requests, 'get', _fake_get)
    monkeypatch.setattr(od.requests, 'post', _fake_post)
    od.ensure_folder('a/b/c')
    assert len(get_calls) == 3
    assert post_calls == []


def test_O6b_ensure_folder_creates_missing_segments_only(monkeypatch):
    """O6b: 'a'는 있고 'a/b'부터 없으면, b/c만 순서대로 생성한다."""
    monkeypatch.setattr(od, 'get_access_token', lambda: 'tok')
    monkeypatch.setitem(od._REMOTE_ITEM_CACHE, 'a', None)  # remoteItem 해석 스킵
    post_calls = []

    def _fake_get(url, headers=None, **kwargs):
        if url == od._item_url('a'):
            return _FakeResponse(status_code=200)
        return _FakeResponse(status_code=404)

    def _fake_post(url, headers=None, json=None, **kwargs):
        post_calls.append((url, json['name']))
        return _FakeResponse(status_code=201)

    monkeypatch.setattr(od.requests, 'get', _fake_get)
    monkeypatch.setattr(od.requests, 'post', _fake_post)
    od.ensure_folder('a/b/c')
    assert [name for _, name in post_calls] == ['b', 'c']
    # b는 a 밑에, c는 a/b 밑에 생성돼야 한다(부모 경로가 맞아야 폴더 트리가 꼬이지 않음)
    assert post_calls[0][0] == od._item_url('a', '/children')
    assert post_calls[1][0] == od._item_url('a/b', '/children')


def test_O6c_ensure_folder_treats_409_as_already_exists(monkeypatch):
    """O6c: GET-then-POST 사이 경쟁으로 409가 나도 예외를 던지지 않고 계속 진행한다."""
    monkeypatch.setattr(od, 'get_access_token', lambda: 'tok')
    monkeypatch.setattr(od.requests, 'get', lambda *a, **k: _FakeResponse(status_code=404))
    monkeypatch.setattr(od.requests, 'post', lambda *a, **k: _FakeResponse(status_code=409))
    od.ensure_folder('x/y')  # 예외 없이 끝나야 함


# ---------------------------------------------------------------------------
# 2026-09-26 F그룹 후속(coordinator 실측) — "PPT 문서"가 실제로는 진짜 폴더가 아니라
# 다른 드라이브를 가리키는 공유 바로가기(remoteItem)라, 경로 기반 주소
# (`root:/PPT 문서/...:`)로는 하위 경로를 못 찾는다(404/422). remoteItem이면
# `/drives/{driveId}/items/{itemId}`로 주소 체계를 전환해야 한다.
#
# 설계: _item_url()은 순수 함수로 유지한다(기존 O1/O1b/O1c가 네트워크 없이 직접 호출하는
# 계약을 깨지 않기 위함) — 실제 네트워크 GET으로 remoteItem 여부를 확인하는 것은 별도
# 함수 _resolve_remote_base()가 맡고, 그 결과를 _REMOTE_ITEM_CACHE(첫 세그먼트 이름 키)에
# 저장한다. _item_url()은 이 캐시를 "이미 채워져 있으면" 읽기만 한다 — 캐시가 비어있으면
# (=_resolve_remote_base()가 아직 호출 안 됨) 기존 root:/... 경로를 그대로 쓴다. 공개 I/O
# 함수(list_children 등)는 URL을 만들기 전에 반드시 _resolve_remote_base()를 먼저 호출해야
# 한다.
# ---------------------------------------------------------------------------


@pytest.fixture(autouse=True)
def _clear_remote_item_cache(monkeypatch):
    """모듈 레벨 _REMOTE_ITEM_CACHE는 프로세스 수명 캐시라 테스트 간 오염 위험이 있다 —
    매 테스트 시작 전 빈 dict로 monkeypatch(자동 원복)해 격리한다."""
    monkeypatch.setattr(od, '_REMOTE_ITEM_CACHE', {})


def test_O7_item_url_uses_remote_base_when_cached(monkeypatch):
    """O7: 첫 세그먼트가 캐시에 remoteItem으로 등록돼 있으면 /drives/{id}/items/{id}:/...: 로
    전환된다."""
    monkeypatch.setitem(od._REMOTE_ITEM_CACHE, 'PPT 문서', ('eef896ec7e438428', 'EEF896EC7E438428!581'))
    url = od._item_url('PPT 문서/20.청년 미사', '/children')
    assert url == (
        f'{od.GRAPH_ROOT}/drives/eef896ec7e438428/items/EEF896EC7E438428!581'
        ':/20.%EC%B2%AD%EB%85%84%20%EB%AF%B8%EC%82%AC:/children'
    )


def test_O7b_item_url_no_rest_path_uses_bare_drives_item_url(monkeypatch):
    """O7b: remoteItem 자신을 가리키는 경로(세그먼트 1개)는 콜론 구문 없이
    /drives/{id}/items/{id}{suffix}가 된다."""
    monkeypatch.setitem(od._REMOTE_ITEM_CACHE, 'PPT 문서', ('drv1', 'item1'))
    url = od._item_url('PPT 문서', '/children')
    assert url == f'{od.GRAPH_ROOT}/drives/drv1/items/item1/children'


def test_O7c_item_url_falls_back_to_root_path_when_not_cached_or_not_remote(monkeypatch):
    """O7c(회귀 가드): 캐시가 비어있거나(미해석) None(remoteItem 아님)이면 기존 동작
    그대로 — O1b/O1c가 계속 통과해야 한다."""
    # 캐시 비어있음(기본 fixture 상태) — 기존 동작
    url = od._item_url('PPT 문서/20.청년 미사', '/children')
    assert url.startswith(f'{od.GRAPH_ROOT}/me/drive/root:/')

    # 명시적으로 "remoteItem 아님"으로 해석된 경우도 동일
    monkeypatch.setitem(od._REMOTE_ITEM_CACHE, '나주노 성가', None)
    url2 = od._item_url('나주노 성가/447.pptx')
    assert url2 == f'{od.GRAPH_ROOT}/me/drive/root:/%EB%82%98%EC%A3%BC%EB%85%B8%20%EC%84%B1%EA%B0%80/447.pptx:'


def test_O8_resolve_remote_base_detects_remoteitem_and_caches(monkeypatch):
    """O8: remoteItem 파셋이 있는 응답이면 (driveId, itemId)를 반환하고 캐시에 저장,
    두 번째 호출은 네트워크를 다시 타지 않는다."""
    monkeypatch.setattr(od, 'get_access_token', lambda: 'tok')
    calls = []

    def _fake_get(url, headers=None, **kwargs):
        calls.append(url)
        return _FakeResponse(json_data={
            'remoteItem': {
                'id': 'EEF896EC7E438428!581',
                'folder': {'childCount': 16},
                'parentReference': {'driveType': 'personal', 'driveId': 'eef896ec7e438428'},
            },
        })

    monkeypatch.setattr(od.requests, 'get', _fake_get)
    result = od._resolve_remote_base('PPT 문서')
    assert result == ('eef896ec7e438428', 'EEF896EC7E438428!581')
    assert len(calls) == 1
    assert calls[0] == f'{od.GRAPH_ROOT}/me/drive/root:/PPT%20%EB%AC%B8%EC%84%9C:'

    result2 = od._resolve_remote_base('PPT 문서')
    assert result2 == ('eef896ec7e438428', 'EEF896EC7E438428!581')
    assert len(calls) == 1  # 캐시 히트 — 추가 네트워크 호출 없음


def test_O8b_resolve_remote_base_returns_none_for_normal_folder(monkeypatch):
    """O8b(회귀 가드): folder 파셋만 있고 remoteItem이 없는 평범한 폴더는 None(=
    remoteItem 아님)을 반환해야 한다 — 기존 성가 폴더 경로들이 계속 root:/...: 를 써야
    하므로."""
    monkeypatch.setattr(od, 'get_access_token', lambda: 'tok')
    monkeypatch.setattr(
        od.requests, 'get',
        lambda *a, **k: _FakeResponse(json_data={'id': 'x', 'folder': {'childCount': 3}}),
    )
    assert od._resolve_remote_base('나주노 성가') is None


def test_O8c_resolve_remote_base_treats_404_as_not_remote_without_raising(monkeypatch):
    """O8c: 첫 세그먼트 자체가 없는 경로(404)는 예외를 올리지 않고 None을 반환한다 —
    실제 "경로 없음" 에러 메시지는 호출부(list_children 등)의 기존 404 처리가 맡아야
    하므로, 여기서 먼저 raise_for_status()로 터뜨리면 안 된다."""
    monkeypatch.setattr(od, 'get_access_token', lambda: 'tok')
    monkeypatch.setattr(od.requests, 'get', lambda *a, **k: _FakeResponse(status_code=404))
    assert od._resolve_remote_base('없는 폴더') is None


def test_O9_list_children_routes_through_remote_base_for_shortcut_folder(monkeypatch):
    """O9: 'PPT 문서'가 remoteItem이면 list_children이 /drives/{id}/items/{id}:/...:/children
    를 실제로 호출해야 한다(경로 기반 root:/PPT 문서/...:는 404/422가 나는 실측 버그)."""
    monkeypatch.setattr(od, 'get_access_token', lambda: 'tok')
    calls = []

    def _fake_get(url, headers=None, **kwargs):
        calls.append(url)
        if url == f'{od.GRAPH_ROOT}/me/drive/root:/PPT%20%EB%AC%B8%EC%84%9C:':
            return _FakeResponse(json_data={
                'remoteItem': {
                    'id': 'ITEM1',
                    'parentReference': {'driveId': 'DRV1'},
                },
            })
        assert url == (
            f'{od.GRAPH_ROOT}/drives/DRV1/items/ITEM1'
            ':/20.%EC%B2%AD%EB%85%84%20%EB%AF%B8%EC%82%AC:/children'
        )
        return _FakeResponse(json_data={'value': [{'name': '2026'}]})

    monkeypatch.setattr(od.requests, 'get', _fake_get)
    items = od.list_children('PPT 문서/20.청년 미사')
    assert [i['name'] for i in items] == ['2026']
    assert len(calls) == 2  # remoteItem 해석 1회 + children 조회 1회


def test_O10_ensure_folder_creates_missing_segment_under_remote_base(monkeypatch):
    """O10: 'PPT 문서'가 remoteItem일 때, 그 아래 없는 하위 폴더를 생성하면 POST가
    /drives/{id}/items/{id}:/기존경로:/children 로 나가야 한다."""
    monkeypatch.setattr(od, 'get_access_token', lambda: 'tok')
    get_calls = []
    post_calls = []

    def _fake_get(url, headers=None, **kwargs):
        get_calls.append(url)
        if url == f'{od.GRAPH_ROOT}/me/drive/root:/PPT%20%EB%AC%B8%EC%84%9C:':
            return _FakeResponse(json_data={
                'remoteItem': {'id': 'ITEM1', 'parentReference': {'driveId': 'DRV1'}},
            })
        if url == f'{od.GRAPH_ROOT}/drives/DRV1/items/ITEM1':
            return _FakeResponse(status_code=200)  # 'PPT 문서' 자신은 이미 있음
        return _FakeResponse(status_code=404)  # '2026' 하위 폴더는 없음

    def _fake_post(url, headers=None, json=None, **kwargs):
        post_calls.append((url, json['name']))
        return _FakeResponse(status_code=201)

    monkeypatch.setattr(od.requests, 'get', _fake_get)
    monkeypatch.setattr(od.requests, 'post', _fake_post)
    od.ensure_folder('PPT 문서/2026')
    assert len(post_calls) == 1
    assert post_calls[0] == (f'{od.GRAPH_ROOT}/drives/DRV1/items/ITEM1/children', '2026')


def test_O11_upload_file_routes_through_remote_base(monkeypatch, tmp_path):
    """O11: upload_file도 동일한 remoteItem 해석 경로를 거쳐야 한다."""
    monkeypatch.setattr(od, 'get_access_token', lambda: 'tok')
    src = tmp_path / 'local.pptx'
    src.write_bytes(b'data')

    def _fake_get(url, headers=None, **kwargs):
        return _FakeResponse(json_data={
            'remoteItem': {'id': 'ITEM1', 'parentReference': {'driveId': 'DRV1'}},
        })

    captured = {}

    def _fake_put(url, headers=None, data=None, **kwargs):
        captured['url'] = url
        return _FakeResponse()

    monkeypatch.setattr(od.requests, 'get', _fake_get)
    monkeypatch.setattr(od.requests, 'put', _fake_put)
    od.upload_file('PPT 문서/20.청년 미사/2026/결과.pptx', src)
    assert captured['url'] == (
        f'{od.GRAPH_ROOT}/drives/DRV1/items/ITEM1'
        ':/20.%EC%B2%AD%EB%85%84%20%EB%AF%B8%EC%82%AC/2026/%EA%B2%B0%EA%B3%BC.pptx:/content'
    )
