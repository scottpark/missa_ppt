# -*- coding: utf-8 -*-
"""가톨릭 성가(1~528) 웹 악보 → 성가 PPT 일회성 생성기 (maria.catholic.or.kr).

- 제목: sungga_search_player 목록(POST curpage=N)에서 수집
- 악보: /files/mp3/sungga/img/2012/2012040NNN.jpg (580px 스캔 JPG, 상단 '구분/제목/번호' 머리글 포함 → 잘라냄)
- 템플릿: 기존 '성가 294 …pptx' 1번 슬라이드(헤더 + '가톨릭 성가 © 한국천주교중앙협의회' 박스)
- 구분: 기존 폴더의 같은 번호 파일 헤더 첫 run, 없으면 '봉헌'
- 악보→슬라이드 변환은 청년미사(missa_youth_hymn_pdf)의 시스템 검출·패킹·배치를 재사용

사용: python tools/make_catholic_hymns.py --test | --all | 번호...
"""
import argparse
import copy
import html
import io
import json
import re
import sys
import urllib.parse
import urllib.request
from pathlib import Path

import numpy as np
from PIL import Image
from pptx import Presentation
from pptx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import missa_ooxml_utils as ou  # noqa: E402
import missa_youth_hymn_pdf as yh  # noqa: E402

BASE = Path(r"C:\Users\Scott\OneDrive\PPT 문서\00.공용 PPT\09.가톨릭 성가\성가-악보버전")
OUT_DIR = BASE / "New"
TEMPLATE = BASE / "성가 294 모든 성인 성녀시여.pptx"
LIST_URL = "https://maria.catholic.or.kr/sungga/search/sungga_search_player.asp"
IMG_URL = "https://maria.catholic.or.kr/files/mp3/sungga/img/2012/2012040{n:03d}.jpg"
DEFAULT_구분 = "봉헌"
FILE_PREFIX = "New_"   # 테스트 기간 임시 접두어(비교 편의)
UPSCALE = 3   # 검출 임계값이 300dpi 기준이라 580px 원본을 3배 확대해 분석(크롭은 원본에서)
TEST = [434, 294, 50, 513, 203, 456, 40, 410, 400, 504, 506, 202]
CACHE = Path.home() / ".cache_catholic_hymns"


def _get(url, data=None):
    req = urllib.request.Request(url, data, headers={"User-Agent": "Mozilla/5.0"})
    return urllib.request.urlopen(req, timeout=60).read()


def fetch_titles() -> dict:
    CACHE.mkdir(exist_ok=True)
    f = CACHE / "titles.json"
    if f.exists():
        return {int(k): v for k, v in json.loads(f.read_text(encoding="utf-8")).items()}
    out = {}
    for p in range(1, 30):
        t = _get(LIST_URL, urllib.parse.urlencode({"curpage": p}).encode()).decode("utf-8", "replace")
        n = 0
        for r in re.findall(r"<tr[^>]*>(.*?)</tr>", t, re.S):
            m = re.search(r"<td>(?:<a[^>]*>)?(\d+)(?:</a>)?</td>", r)
            h = re.search(r'<input type="hidden" VALUE="([^"]*)"', r)
            if m and h:
                out[int(m.group(1))] = re.sub(r"\s+", " ", html.unescape(h.group(1))).strip()
                n += 1
        if n == 0:
            break
    f.write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
    return out


def fetch_image(n: int) -> Image.Image:
    CACHE.mkdir(exist_ok=True)
    f = CACHE / f"{n:03d}.jpg"
    if not f.exists():
        f.write_bytes(_get(IMG_URL.format(n=n)))
    return Image.open(f).convert("RGB")


def existing_구분_map() -> dict:
    """기존 폴더 파일(번호→구분). 슬라이드1 헤더 첫 run 텍스트."""
    res = {}
    for f in BASE.glob("성가 *.pptx"):
        m = re.match(r"성가 (\d+)\s", f.name)
        if not m:
            continue
        try:
            prs = Presentation(str(f))
            for sh in prs.slides[0].shapes:
                if sh.name == "Rectangle 11":
                    num = str(int(m.group(1)))
                    parts = []
                    for r in sh.text_frame.paragraphs[0].runs:
                        if r.text.strip() == num:
                            break
                        parts.append(r.text)
                    res[int(m.group(1))] = "".join(parts).strip()
                    break
        except Exception as e:  # 손상 파일 등은 건너뜀(기본값 사용)
            print("  [경고] 구분 읽기 실패:", f.name, e)
    return res


# ---------------------------------------------------------------------------
# 헤더: 구분·번호는 그대로, 제목만 자간 → 폰트 순으로 줄임
# ---------------------------------------------------------------------------
_TAB_EMU = 914400  # 기본 탭 간격(1인치)


def _width(r):
    return yh._measure_run_emu(r)


def _header_total(runs):
    """[구분][ ][번호][  \t][제목]. 탭은 다음 1인치 배수까지 진행."""
    prefix = sum(_width(r) for r in runs[:3]) + _width_notab(runs[3])
    prefix = ((int(prefix) // _TAB_EMU) + 1) * _TAB_EMU
    return prefix + _width(runs[4])


def _width_notab(r):
    t = r.find(qn("a:t"))
    txt = (t.text or "").replace("\t", "")
    _rPr, spc, sz = yh._run_spc_sz(r)
    return (yh._gungsuh_font(sz / 100.0).getlength(txt) + spc / 100.0 * len(txt)) * yh._EMU_PER_PT


def fill_header(slide, 구분, number, title):
    hdr = yh._find_header(slide)
    p = hdr.text_frame.paragraphs[0]._p
    runs = p.findall(qn("a:r"))
    assert len(runs) == 5, f"템플릿 헤더 run {len(runs)}개"
    yh._set_run_text(runs[0], 구분)
    yh._set_run_text(runs[2], str(number))
    yh._set_run_text(runs[4], title)
    limit = yh._usable_header_emu(hdr) * yh._HEADER_SAFETY
    tr = runs[4].find(qn("a:rPr"))
    note = ""
    if _header_total(runs) > limit:
        spc = int(tr.get("spc", "0"))
        while spc > 0 and _header_total(runs) > limit:   # 1) 제목 자간만 축소(0까지)
            spc -= 50
            tr.set("spc", str(max(spc, 0)))
        note = "자간축소"
        sz = int(tr.get("sz", "2800"))
        while sz > yh._MIN_TITLE_PT * 100 and _header_total(runs) > limit:   # 2) 제목 폰트 축소
            sz -= 100
            tr.set("sz", str(sz))
            note = f"폰트{sz // 100}pt"
        if _header_total(runs) > limit:
            note += "(여전히 넘침!)"
    return note


# ---------------------------------------------------------------------------
# 악보 → 슬라이드 그룹
# ---------------------------------------------------------------------------
def _blank_runs_list(blank):
    """불리언 배열에서 True가 이어지는 구간 [(start, end_inclusive)]."""
    runs, i, n = [], 0, len(blank)
    while i < n:
        if blank[i]:
            j = i
            while j + 1 < n and blank[j + 1]:
                j += 1
            runs.append((i, j))
            i = j + 1
        else:
            i += 1
    return runs


def _longest_run(row) -> int:
    """한 행에서 잉크가 끊김(≤2px 틈 허용) 없이 이어지는 최장 구간 길이. 오선은 기울어도 길고 가사는 짧다."""
    best = cur = gap = 0
    for v in row:
        if v:
            cur += 1 + gap
            gap = 0
            best = max(best, cur)
        elif cur and gap < 2:
            gap += 1
        else:
            cur = gap = 0
    return best


def detect_systems(gray: np.ndarray):
    """웹 악보(580px) 전용 시스템 검출: 오선 → 시스템 묶음 → 시스템 사이 가장 넓은 빈 구간에서 분할.

    상단(빠르기말·코드)은 오선 위로 올라가다 15px 이상 빈 구간을 만나면 멈춘다(머리글 제외).
    반환: [(top, bottom)] 원본 픽셀 좌표(포함)."""
    h, w = gray.shape
    ink = gray < yh._INK
    left, right, _t, _b = yh.ps.detect_content_bounds(gray)
    cw = max(1, right - left + 1)
    fill = ink[:, left:right + 1].sum(axis=1) / cw
    row_has_ink = ink.any(axis=1)
    # 오선 선 행(굵기 1~3px) → 선 단위로 묶기
    light = gray < 170   # 얇고 흐린 오선도 잡히도록 선 검출만 임계값을 완화
    lfill = light[:, left:right + 1].mean(axis=1)
    def _find_staves(thr):
        out, cur = [], []
        for y in range(h):
            if lfill[y] > thr:   # 오선 행(가사 행은 0.5를 넘지 않는다 — 12곡 실측)
                if cur and y - cur[-1] > 22:
                    out.append(cur)
                    cur = []
                cur.append(y)
        if cur:
            out.append(cur)
        return [(st[0], st[-1]) for st in out if len(st) >= 3]

    def _find_staves_struct(thr=0.3):
        """연한 오선 보완: 얇은(≤4행) 선 4개 이상이 5~16px 간격으로 늘어선 묶음을 오선으로 본다."""
        runs, y = [], 0
        lf = lfill
        while y < h:
            if lf[y] > thr:
                y0 = y
                while y < h and lf[y] > thr:
                    y += 1
                if y - y0 <= 4:
                    runs.append((y0, y - 1))
            else:
                y += 1
        out, cur = [], []
        for ru in runs:
            c = (ru[0] + ru[1]) / 2
            if cur and not (5 <= c - (cur[-1][0] + cur[-1][1]) / 2 <= 16):
                if len(cur) >= 4:
                    out.append((cur[0][0], cur[-1][1]))
                cur = []
            cur.append(ru)
        if len(cur) >= 4:
            out.append((cur[0][0], cur[-1][1]))
        return out

    staves = _find_staves(0.5)
    if not staves:   # 55·56번처럼 음표가 많아 선이 끊겨 보이는 악보: 오선을 전혀 못 찾을 때만 임계 완화
        staves = _find_staves(0.4)
    # 240번처럼 오선이 연해 일부만 잡힌 악보: 구조 기반으로 찾은 오선 중 기존과 겹치지 않는 것을 추가
    for st in _find_staves_struct():
        if not any(st[0] <= e[1] + 4 and st[1] >= e[0] - 4 for e in staves):
            staves.append(st)
    staves.sort()
    if not staves:
        return []
    GAP = 15       # 아래쪽 탐색용
    GAP_TOP = 8    # 위쪽: 머리글(제목/번호)과 빠르기말 사이가 50·513번처럼 8~9px뿐이라 8px로 끊는다
    # 위쪽: 첫 오선에서 위로, GAP_TOP 이상 빈 구간 전까지
    top = staves[0][0]
    blank = 0
    y = top - 1
    while y >= 0:
        if row_has_ink[y]:
            blank = 0
            top = y
        else:
            blank += 1
            if blank >= GAP_TOP:
                break
        y -= 1
    # 아래쪽: 마지막 오선 아래의 [오선 근처 잉크] → [오선↔가사 빈 구간] → [가사 블록(9px 미만 간격은 같은 블록)].
    # 예전에는 "15px 이상 빈 구간에서 멈춤"이라 오선↔가사 간격(203번 마지막 시스템 ≈20px)에서 멈춰
    # 마지막 줄 가사가 통째로 잘렸다(2026-11-06 원인 분석). 안쪽 시스템과 같은 규칙을 쓴다.
    bottom = staves[-1][1]
    y = bottom + 1
    while y < h and row_has_ink[y]:          # 1) 오선에 붙은 잉크(낮은 음표·꼬리)
        bottom = y
        y += 1
    while y < h and not row_has_ink[y]:      # 2) 오선↔가사 빈 구간(길이 무관)
        y += 1
    blank = 0
    while y < h:                             # 3) 가사 블록
        if row_has_ink[y]:
            bottom, blank = y, 0
        else:
            blank += 1
            if blank >= 14:   # 185번: 한국어 가사↔라틴어 가사 간격이 9px를 넘음
                break
        y += 1
    # 시스템 경계: 인접 오선 사이의 빈 구간(6px 이상) 중 **두 번째**가 가사 끝/다음 시스템 시작이다.
    # 첫 번째는 오선↔가사 사이 간격, 두 번째는 가사↔다음 시스템 위쪽 표시(빠르기말·코드) 간격이다.
    # 구간이 하나뿐이면(가사 없음) 그 구간을 쓴다. 가사 줄 간격은 3~5px라 6px 미만으로 같은 블록이다.
    MIN_RUN = 6
    bands = []
    cur_top = top
    for i in range(len(staves) - 1):
        lo, hi = staves[i][1] + 1, staves[i + 1][0] - 1
        runs = [(a_ + lo, b_ + lo) for a_, b_ in _blank_runs_list(~row_has_ink[lo:hi + 1]) if b_ - a_ + 1 >= MIN_RUN]
        if len(runs) >= 2:
            # 첫 구간(오선↔가사) 이후 가장 넓은 구간이 가사 끝/다음 시스템 시작. 185·194번처럼 가사 아래
            # 라틴어 등 보조 가사 줄이 한 줄 더 있으면 그 줄 사이(≈8px)가 아니라 마지막 줄 아래(≈20px)가 경계다.
            s_, e_ = max(runs[1:], key=lambda r: (r[1] - r[0], -r[0]))
        elif runs:
            s_, e_ = runs[0]
        else:
            s_ = e_ = (lo + hi) // 2
        bands.append((cur_top, s_ - 1))
        cur_top = e_ + 1
    bands.append((cur_top, bottom))
    return bands


def place_scaled(slide, pil_img, pct):
    """가로폭 꽉 채움 배율의 pct(%)로 배치. 밴드를 넘으면 헤더 바로 아래(여백 12만 EMU)부터 쌓는다."""
    scale = yh._SCORE_BAND_W / pil_img.width * pct / 100.0
    w, h = int(round(pil_img.width * scale)), int(round(pil_img.height * scale))
    left = yh._SCORE_BAND_LEFT + (yh._SCORE_BAND_W - w) // 2
    top = max(yh.HEADER_H + 120000, yh._SCORE_BAND_TOP + (yh._SCORE_BAND_H - h) // 2)
    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    buf.seek(0)
    slide.shapes.add_picture(buf, left, top, w, h)


SHRINK_STEPS = (92, 82)  # 100%로 안 들어가는 2줄 이상 묶음에 순서대로 시도하는 축소 배율(사용자 지정, 2026-11-06)
_AVAIL_H = 4800000 - (yh.HEADER_H + 120000)   # 축소 배치 시 쓸 수 있는 높이(헤더 아래 12만 EMU ~ 저작권 위)


def _group_scale(gray, group):
    """묶음 배율: 1줄=100, 100%로 들어가면 100, 92%→82% 순으로 들어가는 첫 배율, 아니면 None(불가)."""
    if len(group) == 1:
        return 100
    if yh._group_fits_band(gray, group, 1.0):
        return 100
    l, t, r, b = yh._group_content_bbox(gray, group)
    full_h = (b - t) * (yh._SCORE_BAND_W / max(1, r - l))
    for pct in SHRINK_STEPS:
        if full_h * pct / 100.0 <= _AVAIL_H:
            return pct
    return None


def lost_ink_rows(gray, bands, min_rows=4, min_px=20):
    """첫 시스템 윗단 아래에서 어느 밴드에도 속하지 않은 잉크 블록(가사 누락 의심)을 [(y0, y1, px)]로 반환.
    머리글은 첫 밴드 위라 제외된다. 작은 얼룩(min_rows 미만 또는 min_px 미만)은 무시."""
    ink = gray < yh._INK
    covered = np.zeros(len(gray), dtype=bool)
    for t, b in bands:
        covered[t:b + 1] = True
    res, y, h = [], bands[0][0], len(gray)
    while y < h:
        if not covered[y] and ink[y].any():
            y0 = y
            while y < h and not covered[y] and (ink[y].any() or y - y0 < 0):
                y += 1
            px = int(ink[y0:y].sum())
            if y - y0 >= min_rows and px >= min_px:
                res.append((y0, y - 1, px))
        else:
            y += 1
    return res


def pack_with_shrink(bands, gray, max_per_slide=3):
    """슬라이드 수 최소 → 최대 묶음 최소 → 앞쪽에 큰 묶음(청년미사 pack_systems와 같은 기준)."""
    n = len(bands)
    ok = {}
    for st in range(n):
        for size in range(1, max_per_slide + 1):
            if st + size <= n:
                ok[(st, size)] = _group_scale(gray, bands[st:st + size]) is not None
    # 뒤에서부터 DP: best[i] = i번째 밴드부터의 최선 (슬라이드 수, 최대 묶음, -크기들)
    best = {n: (0, 0, ())}
    for st in range(n - 1, -1, -1):
        cand = None
        for size in range(1, max_per_slide + 1):
            if st + size > n or not ok[(st, size)] or (st + size) not in best:
                continue
            c, m, neg = best[st + size]
            key = (c + 1, max(m, size), (-size,) + neg)
            if cand is None or key < cand:
                cand = key
        if cand is not None:
            best[st] = cand
    if 0 not in best:
        raise RuntimeError("묶음을 만들 수 없음")
    best = (None, [-x for x in best[0][2]])
    groups, i = [], 0
    for sz in best[1]:
        groups.append(bands[i:i + sz])
        i += sz
    return groups


def add_text_bands(gray, bands, para_gap=12):
    """어느 밴드에도 속하지 않은 잉크 블록(악보 아래 2~7절·라틴어 가사·해설 등)을 문단 단위 밴드로 추가.
    줄 사이 간격이 para_gap 미만이면 한 문단. 한 슬라이드에 못 들어가는 문단은 추가하지 않는다(누락 의심으로 남음)."""
    lost = lost_ink_rows(gray, bands)
    paras = []
    for y0, y1, _px in lost:
        if paras and y0 - paras[-1][1] < para_gap:
            paras[-1] = (paras[-1][0], y1)
        else:
            paras.append((y0, y1))
    extra = [pb for pb in paras if _group_scale(gray, [pb]) is not None]
    return sorted(list(bands) + extra), set(extra)


def score_groups(img: Image.Image, sizes=None, scale_override=None):
    """원본 이미지 → (slide별 crop, 묶음 크기, 배율). 머리글은 첫 오선 위쪽 탐색에서 제외된다."""
    gray = np.asarray(img.convert("L"))
    bands = detect_systems(gray)
    if not bands:
        raise RuntimeError("오선 미검출")
    bands, text_bands = add_text_bands(gray, bands)
    page_l, page_r, _pt, _pb = yh.ps.detect_content_bounds(gray)
    if sizes:
        assert sum(sizes) == len(bands), f"묶음 합 {sum(sizes)} != 시스템 {len(bands)}"
        groups, i = [], 0
        for sz in sizes:
            groups.append(bands[i:i + sz])
            i += sz
    else:
        groups = pack_with_shrink(bands, gray)
    crops, scales = [], []
    for g in groups:
        l, t, r, b = yh._group_content_bbox(gray, g)
        if all(x in text_bands for x in g):   # 문단만 있는 슬라이드: 페이지 본문 폭 기준으로 잘라 글자 크기를 악보 슬라이드와 맞춘다
            l, r = min(l, page_l), max(r, page_r)
        crops.append(img.crop((max(0, l - 1), max(0, t - 1), min(img.width, r + 1), min(img.height, b + 1))))
        scales.append(scale_override if (scale_override and len(g) > 1) else (_group_scale(gray, g) or 100))
    warn = lost_ink_rows(gray, bands)
    ink = gray < yh._INK
    kept = sum(int((np.asarray(c.convert("L")) < yh._INK).sum()) for c in crops)
    total = int(ink[bands[0][0]:].sum())
    ratio = kept / max(1, total)
    if warn or ratio < 0.995:
        LOSS_WARN.append((warn, round(ratio, 4)))
    return crops, [len(g) for g in groups], scales


LOSS_WARN = []   # 가사/악보 누락 의심 기록 [(잃은 잉크 블록 목록, 보존 잉크 비율)] — build마다 비운다


def build(number, title, 구분, img, sizes_override=None, scale_override=None):
    LOSS_WARN.clear()
    crops, sizes, scales = score_groups(img, sizes_override, scale_override)
    prs = Presentation(str(TEMPLATE))
    # 1장만 남기고 그림 제거(+rel 정리)
    ou_delete = ou.delete_slide
    while len(prs.slides) > 1:
        ou_delete(prs, len(prs.slides) - 1)
    s0 = prs.slides[0]
    rids = set()
    for sh in list(s0.shapes):
        if sh.shape_type == 13:
            rids.add(sh._element.find(".//" + qn("a:blip")).get(qn("r:embed")))
            sh._element.getparent().remove(sh._element)
    for rid in rids:
        if rid in s0.part.rels:
            s0.part.rels.pop(rid)
    for _ in range(len(crops) - 1):
        ou.duplicate_slide(prs, 0)
    note = ""
    for si, crop in enumerate(crops):
        s = prs.slides[si]
        n = fill_header(s, 구분, number, title)
        note = note or n
        if scales[si] != 100:
            place_scaled(s, crop, scales[si])
        else:
            yh._place_score(s, crop)
    return prs, sizes, note, scales


def safe_name(s):
    return re.sub(r'[\\/:*?"<>|]', "", s)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("numbers", nargs="*", type=int)
    ap.add_argument("--test", action="store_true")
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--out", default=str(OUT_DIR))
    ap.add_argument("--scale", type=float, default=None, help="--sizes와 함께: 2줄 이상 슬라이드에 강제 적용할 배율(%%)")
    ap.add_argument("--sizes", default=None, help="슬라이드별 시스템 수 수동 지정(예: 2,2,1), 번호 1개일 때만")
    ns = ap.parse_args()
    nums = TEST if ns.test else list(range(1, 529)) if ns.all else ns.numbers
    out = Path(ns.out)
    out.mkdir(parents=True, exist_ok=True)
    titles = fetch_titles()
    구분s = existing_구분_map()
    missing, fails = [], []
    for n in nums:
        if n not in titles:
            missing.append(n)
            continue
        try:
            구분 = 구분s.get(n, DEFAULT_구분)
            prs, sizes, note, scales = build(n, titles[n], 구분, fetch_image(n), [int(x) for x in ns.sizes.split(",")] if ns.sizes else None, ns.scale)
            lw = f" ★가사누락의심 {LOSS_WARN[0]}" if LOSS_WARN else ""
            fn = out / f"{FILE_PREFIX}성가 {n} {safe_name(titles[n])}.pptx"
            prs.save(str(fn))
            print(f"OK {n} {titles[n]} [{구분}{'' if n in 구분s else '(기본)'}] 슬라이드{len(prs.slides)} 묶음{sizes} 배율{scales} {note}{lw}")
        except Exception as e:
            fails.append((n, repr(e)))
            print(f"FAIL {n}: {e!r}")
    print("목록에 없음:", missing, "실패:", fails)


if __name__ == "__main__":
    main()
