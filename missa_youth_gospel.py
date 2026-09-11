# -*- coding: utf-8 -*-
"""청년미사 영문 복음 조회 (universalis.com).

날짜(YYYYMMDD)를 받아 `https://universalis.com/{date}/mass.htm`에서 Gospel 섹션을
추출해 JSON dict으로 반환한다. 완전 신규 leaf 모듈로, 프로젝트 내부 어떤 모듈도 import하지
않는다(설계서 §1). missa_to_json.py 파서는 cbck.or.kr 전용이라 재사용하지 않는다.

핵심 함정(설계서 §2.2, 요구사항 §오류 처리) — 리다이렉트를 따라가면 안 된다:
- 과거로 먼 날짜(예: 20200101)는 301 → /mass.htm(날짜 없는 오늘자)으로 리다이렉트되어,
  따라가면 **조용히 오늘의 복음**이 반환된다(가장 위험).
- 미래로 먼 날짜(예: 20270101)는 301 → /n-otherdates.htm(안내 페이지)로 리다이렉트된다.
따라서 리다이렉트를 차단(follow 금지)하고, 응답이 3xx면 즉시 GospelFetchError를 던진다.
날짜 계산으로 범위를 추정하지 않는다(서버 응답만으로 판정).
"""
import html as _html
import json
import sys
import urllib.error
import urllib.parse
import urllib.request
from html.parser import HTMLParser

GOSPEL_URL = "https://universalis.com/{yyyymmdd}/mass.htm"

# 본문 단락으로 수집할 div class 토큰. class 속성은 "v gb" / "p gb"처럼 복수 토큰일 수
# 있으므로(실측 20260908 족보), 정확 문자열이 아니라 토큰 교집합으로 판정한다.
_BODY_CLASSES = {"p", "pi", "v", "vi"}
_REDIRECT_CODES = (301, 302, 303, 307, 308)


class GospelFetchError(Exception):
    """리다이렉트/범위초과/HTTP오류/파싱실패 공통 예외."""


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    """리다이렉트를 자동 추적하지 않는다 — redirect_request가 None을 반환하면 urllib은
    3xx 응답을 HTTPError로 올린다(설계서 §2.2 (A))."""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def fetch_gospel_html(date_str: str) -> str:
    """{date}/mass.htm을 요청. 리다이렉트(3xx) 또는 최종 경로 불일치면 GospelFetchError.

    200 + 최종 경로가 정확히 /{date}/mass.htm일 때만 HTML을 반환한다."""
    url = GOSPEL_URL.format(yyyymmdd=date_str)
    opener = urllib.request.build_opener(_NoRedirect)
    try:
        resp = opener.open(url, timeout=30)
    except urllib.error.HTTPError as e:
        if e.code in _REDIRECT_CODES:
            loc = e.headers.get("Location")
            raise GospelFetchError(
                f"{date_str} 복음을 가져올 수 없습니다 — 리다이렉트({e.code} → {loc}). "
                f"날짜가 제공 범위를 벗어났을 수 있습니다."
            ) from e
        raise GospelFetchError(f"{date_str} 복음 요청 실패 (HTTP {e.code}).") from e
    except urllib.error.URLError as e:
        raise GospelFetchError(f"{date_str} 복음 요청 실패 (네트워크 오류: {e.reason}).") from e

    # (B) 최종 URL 경로 대조(보강). _NoRedirect로 이미 추적을 막았지만, 경로가 다르면
    # 무조건 에러로 확정한다(설계서 §2.2 (B) — 가장 위험한 조용한 오늘자 반환 차단).
    final_path = urllib.parse.urlparse(resp.geturl()).path
    if final_path != f"/{date_str}/mass.htm":
        raise GospelFetchError(
            f"{date_str} 복음을 가져올 수 없습니다 — 최종 경로 불일치({final_path})."
        )
    return resp.read().decode("utf-8", "replace")


class _GospelParser(HTMLParser):
    """이벤트 스트림 수집기. 정규식 단독 파싱은 중첩/엔티티에 취약하므로 stdlib
    html.parser로 (start/data/end) 이벤트를 선형 기록한 뒤 extract 단계에서 처리한다
    (설계서 §2.3)."""

    def __init__(self):
        super().__init__(convert_charrefs=True)  # 엔티티는 여기서 유니코드로 디코딩됨
        self.events = []

    def handle_starttag(self, tag, attrs):
        self.events.append(("start", tag, dict(attrs)))

    def handle_startendtag(self, tag, attrs):
        self.events.append(("startend", tag, dict(attrs)))

    def handle_endtag(self, tag):
        self.events.append(("end", tag, None))

    def handle_data(self, data):
        self.events.append(("data", data, None))


def _class_tokens(attrs: dict) -> set:
    return set((attrs.get("class") or "").split())


# 미사 섹션(웹사이트 th 텍스트 그대로 — 영어 키). "Or:"(대체 환호송)는 최상위 키가 아니라
# 직전 "Gospel Acclamation" 하위 "or" 필드로 붙인다(그 자체로 의미 있는 최상위 키가 아님).
_KNOWN_SECTIONS = ("First reading", "Responsorial Psalm", "Second reading",
                   "Gospel Acclamation", "Gospel")
_THEME_SECTIONS = {"First reading", "Second reading", "Gospel"}  # h4 주제 있는 섹션
_NO_CONTENT_SECTIONS = {"Responsorial Psalm"}  # 사이트가 저작권상 시편 본문 미제공 → reference만


def _events(html_text: str) -> list:
    p = _GospelParser()
    p.feed(html_text)
    return p.events


def _text_after(ev, start_idx, tag):
    """start_idx 이후 첫 <tag> 시작의 텍스트를 중첩 고려해 이어붙여 반환. (start_idx, text)."""
    i = start_idx
    while i < len(ev):
        if ev[i][0] == "start" and ev[i][1] == tag:
            j = i + 1
            buf = []
            depth = 1
            while j < len(ev) and depth > 0:
                kind, a, _ = ev[j]
                if kind == "start" and a == tag:
                    depth += 1
                elif kind == "end" and a == tag:
                    depth -= 1
                elif kind == "data" and depth > 0:
                    buf.append(a)
                j += 1
            return i, "".join(buf).strip()
        i += 1
    return None, ""


def parse_mass(html_text: str) -> dict:
    """HTML에서 미사 전 섹션(First reading/Responsorial Psalm/Second reading/Gospel
    Acclamation/Gospel)을 추출해 영어 키 dict으로 반환(한글 키 없음).

    - 섹션 마커: align="left" & 텍스트가 KNOWN(또는 "Or:")인 <th>. universalis는 제목 th를
      align=left, 참조 th를 align=right로 구분한다(실측).
    - 각 섹션: reference = 마커 뒤 첫 align="right" th, theme = (해당 섹션만) 그 뒤 첫 <h4>,
      content = 본문 div({p,pi,v,vi}) 를 \\n\\n으로 이어붙임. 종료 경계 = **다음 섹션 마커 th**
      / <hr class="shortrule"> / <div class="audioclip"> 중 먼저 오는 것(다음 섹션 마커를 반드시
      경계로 둬야 Acclamation 본문이 "Or:"/다음 Gospel로 새지 않는다).
    - "Or:" 대체 환호송은 직전 Gospel Acclamation 하위 "or"(reference/content)로 중첩.
    - Responsorial Psalm은 reference만(theme/content 키 없음)."""
    ev = _events(html_text)
    known = set(_KNOWN_SECTIONS) | {"Or:"}

    markers = []  # (name, event_idx) — 문서 순서
    for i in range(len(ev)):
        if ev[i][0] == "start" and ev[i][1] == "th" and ev[i][2].get("align") == "left":
            _, txt = _text_after(ev, i, "th")
            if txt in known:
                markers.append((txt, i))
    if not any(name == "Gospel" for name, _ in markers):
        raise GospelFetchError("Gospel 섹션을 찾지 못했습니다 (구조 변경 가능).")

    def _first_right_th(start, stop):
        j = start
        while j < stop:
            if ev[j][0] == "start" and ev[j][1] == "th" and ev[j][2].get("align") == "right":
                return _text_after(ev, j, "th")
            j += 1
        return None, ""

    def _first_h4(start, stop):
        j = start
        while j < stop:
            if ev[j][0] == "start" and ev[j][1] == "h4":
                return _text_after(ev, j, "h4")[1]
            j += 1
        return ""

    def _collect_body(start, stop):
        paras = []
        i = start
        while i < stop:
            kind, a, attrs = ev[i]
            if a == "hr" and kind in ("start", "startend") and "shortrule" in _class_tokens(attrs):
                break
            if kind == "start" and a == "div":
                toks = _class_tokens(attrs)
                if "audioclip" in toks:
                    break
                if toks & _BODY_CLASSES:
                    _, txt = _text_after(ev, i, "div")
                    if txt:
                        paras.append(txt)
            i += 1
        return "\n\n".join(paras)

    result = {}
    accl = None
    for k, (name, idx) in enumerate(markers):
        stop = markers[k + 1][1] if k + 1 < len(markers) else len(ev)
        ref_i, ref = _first_right_th(idx + 1, stop)
        scan = (ref_i if ref_i is not None else idx) + 1
        entry = {"reference": _html.unescape(ref)}
        if name in _THEME_SECTIONS:
            theme = _first_h4(scan, stop)
            if theme:
                entry["theme"] = _html.unescape(theme)
        if name not in _NO_CONTENT_SECTIONS:
            entry["content"] = _collect_body(scan, stop)
        if name == "Or:":
            if accl is not None:
                accl["or"] = entry
            continue
        result[name] = entry
        if name == "Gospel Acclamation":
            accl = entry
    return result


def parse_gospel(html_text: str) -> dict:
    """(하위호환) Gospel 섹션만 {reference, theme, content}로 반환. parse_mass 래퍼."""
    gospel = parse_mass(html_text).get("Gospel", {})
    return {
        "reference": gospel.get("reference", ""),
        "theme": gospel.get("theme", ""),
        "content": gospel.get("content", ""),
    }


def get_youth_mass(date_str: str) -> dict:
    """fetch + parse_mass 조합. {"date", <영어 섹션 키>…} 반환(한글 키 없음).

    빈 본문 예외 판정은 Gospel 기준으로만 한다 — Responsorial Psalm은 사이트가 본문을 안 주는
    것이 정상이라 빈 것이 오류가 아니다."""
    mass = parse_mass(fetch_gospel_html(date_str))
    if not (mass.get("Gospel") or {}).get("content"):
        raise GospelFetchError(f"{date_str} 복음 본문이 비어 있습니다 (파싱 실패 가능).")
    return {"date": date_str, **mass}


def get_youth_gospel(date_str: str) -> dict:
    """(하위호환) {"date", "복음":{reference, theme, content}} 반환."""
    html_text = fetch_gospel_html(date_str)
    gospel = parse_gospel(html_text)
    if not gospel.get("content"):
        raise GospelFetchError(f"{date_str} 복음 본문이 비어 있습니다 (파싱 실패 가능).")
    return {"date": date_str, "복음": gospel}


def main(argv=None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if not argv:
        print("usage: python missa_youth_gospel.py YYYYMMDD [out.json]", file=sys.stderr)
        return 2
    date_str = argv[0]
    try:
        out = get_youth_mass(date_str)
    except GospelFetchError as e:
        print(f"[오류] {e}", file=sys.stderr)
        return 1
    text = json.dumps(out, ensure_ascii=False, indent=2)
    # 한글 미사 missa_YYYYMMDD.json과 구분되게 영문 미사 데이터는 missa_en_ 접두(리더 확정).
    path = argv[1] if len(argv) >= 2 else f"missa_en_{date_str}.json"
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    print(f"저장: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
