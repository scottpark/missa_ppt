# -*- coding: utf-8 -*-
"""청년미사 PDF 성가집 → 성가 슬라이드 PPT 생성 (1단계).

성가 구분(입당/봉헌/…)·출처(나주노/야훼 이레)·번호를 받아, 샘플 PPT
(reference/청년미사/나주노 성가 447 ….pptx)와 동일한 포맷의 슬라이드 여러 장짜리
Presentation을 만든다. 완전 신규 모듈로 missa_to_ppt.py 파이프라인을 import하지 않는다
(설계서 §1). missa_ooxml_utils(leaf)와 missa_psalm_score_image의 이미지/헤더 유틸 패턴만
재사용한다.

두 PDF의 근본 차이(요구사항 §두 PDF의 근본적인 차이):
- 나주노: 스캔 래스터(텍스트 레이어 없음). 번호→페이지는 TOC(get_toc)로, 저작권은 개발자
  OCR로 1회 만든 캐시(assets/나주노_copyright_bbox_cache.json)로 처리(런타임 OCR 없음).
- 야훼 이레: 텍스트 레이어 있음. 번호→페이지·저작권 모두 search_for로 실시간 처리.
  한 페이지에 여러 곡이 있을 수 있어(809/810/811) 곡 y-범위를 텍스트 검색으로 잘라낸다.
"""
import copy
import functools
import io
import json
import logging
import re
import sys
from pathlib import Path

import fitz  # PyMuPDF
import numpy as np
from PIL import Image
from pptx import Presentation
from pptx.oxml.ns import qn

import missa_ooxml_utils as ou
import missa_psalm_score_image as ps

logger = logging.getLogger(__name__)

# --- 이미지 분석 임계값 (설계서 §3.4) ---
_INK = 128
# 오선(staff) 판정은 **콘텐츠 폭 기준** row_fill로 한다. 나주노는 페이지 전체를 렌더하는데
# 좌우 여백이 넓어 페이지 폭 기준으로는 스캔이 흐린 곡(362)의 오선이 0.5를 못 넘긴다(실측:
# 페이지폭 기준 max 0.486 → 오선 0개). 콘텐츠 좌우 경계(detect_content_bounds)로 트리밍한
# 폭 기준이면 362도 max 0.558로 오선이 검출된다(5곡 실측).
# 0.5 엄격 초과는 362 마지막 시스템(오선 max_fill이 정확히 0.500)을 탈락시켰다(실측). 반면
# 제목/작곡가/저작권 텍스트 세그먼트의 max_fill은 5곡 전부 ≤0.346이라, 0.40이 시스템(≥0.50)과
# 텍스트(≤0.346) 사이 안전한 중앙이다.
_STAFF_ROW_FILL = 0.40
_BLANK_ROW_FILL = 0.005       # 이 비율 미만 잉크 행 = 공백(콘텐츠 폭 기준)
# 151번은 위 5곡 실측 범위를 벗어난다: 저작권 두 줄(영문, 글자가 조밀)의 row_fill이 최대
# 0.438까지 올라가 _STAFF_ROW_FILL(0.40)을 5개 행에서 넘는다 — _MIN_STAFF_ROWS=3이면 이
# 저작권 줄이 그대로 '시스템'으로 오검출돼 스퓨리어스 그룹(=슬라이드)이 생긴다(2026-09-24
# 실사용자 보고, output/20260919 산출물 슬라이드 17이 악보 대신 저작권 텍스트를 보여줌).
# _STAFF_ROW_FILL 자체를 올리는 방법(예: 0.5)은 362/447처럼 약한 진짜 시스템(오선이 얇거나
# 스캔이 흐려 max_fill이 0.5~0.56에 불과한 시스템)을 탈락시키는 회귀를 실측으로 확인해
# 채택하지 않았다. 대신 "행 개수" 축으로 분리한다 — 7개 곡 40개 실제 시스템 밴드를 전수
# 조사한 결과 row_fill>0.40인 행이 가장 적은 시스템도 12개였고(447), 151 저작권 줄은 5개뿐
# 이었다(margin 12 vs 5, _STAFF_ROW_FILL의 margin 0.5 vs 0.438보다 훨씬 넓어 더 안전).
_MIN_STAFF_ROWS = 8           # 세그먼트에 오선행이 이만큼 있으면 '시스템'(5선 중 일부라도)
# 시스템 사이 여백은 시스템 내부 여백(코드/오선/가사 간)보다 크다. 이 값보다 짧은 공백은
# 한 시스템 내부로 보고 병합한다. 페이지 높이 비례(스캔 해상도 불변) + 하한.
_SYSTEM_GAP_RATIO = 0.017
_SYSTEM_GAP_MIN = 20

# 실행 파일 옆 자산 폴더(missa_psalm_score_image._BASE와 동일 관용, PyInstaller onefile 회피).
_BASE = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent

# 배포 zip이 원본 PDF 2개를 이 경로에 그대로 포함해 배포한다(tools/build_youth_dist_zip.py).
# 개발 PC에서도 같은 경로(cache/청년미사_성가원본/)에 원본을 두고 쓴다 — reference/나
# OneDrive 조회 폴백은 없다(2026-09-27, "성가 원본은 매번 OneDrive에서 곡 단위로 받아오는
# 것이 아니라 배포 시점에 통째로 포함한다"는 결정 — 상세는 CLAUDE.md 참고).
_CACHE_DIR = _BASE / "cache" / "청년미사_성가원본"

SOURCES = {
    "나주노": {"pdf": str(_CACHE_DIR / "N Hymns_Songs 나는 주님께 노래하리라.pdf"),
              "kind": "scan", "onedrive_subfolder": "나주노 성가"},
    "야훼 이레": {"pdf": str(_CACHE_DIR / "Y Hymns_Songs 야훼이레.pdf"),
                "kind": "text", "onedrive_subfolder": "야훼이레 성가"},
}


def _resolve_pdf_path(source: str) -> Path:
    """SOURCES[source]의 PDF 원본 경로를 반환한다 — 항상 `_CACHE_DIR`(배포 zip이 원본을
    미리 담아두는 고정 위치) 하나만 확인한다. 없으면 배포가 잘못된 것이므로 즉시 명확한
    오류로 중단한다(조용한 OneDrive 폴백 없음)."""
    pdf_path = Path(SOURCES[source]["pdf"])
    if not pdf_path.exists():
        raise FileNotFoundError(
            f"{source} 성가집 PDF를 찾을 수 없습니다: {pdf_path}\n"
            "배포 zip에 원본 PDF가 빠졌거나 삭제된 것 같습니다 — 개발 담당자에게 문의해 주세요."
        )
    return pdf_path

# 헤더 표시명 매핑: SOURCES 키/CLI/find_song_* 내부 조회는 원본 키("야훼 이레")를 유지하되
# 헤더에 렌더되는 문자열만 공백 없는 표시명으로 바꾼다(요청 항목 3). SOURCES 키를 바꾸면
# find_song_page 등 조회가 전부 깨지므로 표시 단계에서만 치환한다.
SOURCE_DISPLAY = {"야훼 이레": "야훼이레"}

COPYRIGHT_TRIGGERS = ["Administered by", "Adm. by", "Adm.by"]  # 처음 매치되는 것을 크롭 경계로
RENDER_DPI = 300

# 이 스캔·300dpi에서 마침표(".")는 폭 3~4px, 한글/라틴 글자는 ≥10px. 8px가 둘 사이 안전
# 경계다(글자로 끝나면 마지막 잉크 런 폭>8, 꼬리 마침표만 남으면 ≤8). 저작권 크롭 우측
# 꼬리 마침표 트림(_trim_trailing_period_px)과 캐시 도구가 공유하는 임계.
_PERIOD_MAX_W_PX = 8

# 야훼 제목 번호는 좌측 여백(x0≈65)에서 시작한다. 가사/화음 중간의 숫자는 x0가 크므로
# 이 임계값으로 "줄머리 번호."만 제목으로 인정한다(설계서 §3.3).
_TITLE_LEFT_MARGIN_X = 110

# 헤더 자동 축소용 폰트 메트릭(요청 항목 2). 헤더 폰트는 궁서(Gungsuh)이며 batang.ttc의
# index 2다(TTC 인덱스는 실측 확인 — CLAUDE.md TTC 인덱스 원칙: 0=Batang,1=BatangChe,
# 2=Gungsuh,3=GungsuhChe). 폭 측정은 getlength(pt)=px 가정으로 pt를 얻고, 자간 spc는
# 1/100pt/글자로 더한 뒤 ×12700로 EMU 환산한다(설계표 실측과 0.5% 이내 일치 검증).
_HEADER_FONT_PATH = r"C:\Windows\Fonts\batang.ttc"
_HEADER_FONT_INDEX = 2
_HEADER_SAFETY = 0.98        # wrap 직전 여유(측정 오차·궁서 커닝 흡수)
_MIN_TITLE_PT = 18           # 제목 폰트 축소 하한
_EMU_PER_PT = 12700


# ---------------------------------------------------------------------------
# 번호 → 페이지/제목 조회
# ---------------------------------------------------------------------------

def _open(source: str) -> fitz.Document:
    if source not in SOURCES:
        raise ValueError(f"알 수 없는 출처: {source!r} (지원: {list(SOURCES)})")
    return fitz.open(str(_resolve_pdf_path(source)))


def _naju_toc_entry(number: int):
    """나주노 TOC에서 번호로 시작하는 항목 (title, page_1based) 반환. 없으면 None."""
    pat = re.compile(rf"^\s*{number}\b")
    with _open("나주노") as d:
        for _lvl, title, page in d.get_toc():
            if pat.match(title):
                return title, page
    return None


def find_song_page(source: str, number: int) -> int:
    """번호 → 0-based 페이지 인덱스.

    나주노=get_toc(1-based page → index로 -1), 야훼 이레=search_for로 좌측여백 제목 스캔."""
    if source == "나주노":
        entry = _naju_toc_entry(number)
        if entry is None:
            raise ValueError(f"나주노 TOC에 성가 {number} 없음")
        return entry[1] - 1
    # 야훼 이레: 텍스트 검색
    needle = f"{number}."
    with _open(source) as d:
        for pi in range(d.page_count):
            for h in d[pi].search_for(needle):
                if h.x0 < _TITLE_LEFT_MARGIN_X:
                    return pi
    raise ValueError(f"{source}에서 성가 {number} 페이지 미검출")


def find_song_title(source: str, number: int) -> str:
    """헤더용 곡 제목(번호 접두 제거)."""
    if source == "나주노":
        entry = _naju_toc_entry(number)
        if entry is None:
            raise ValueError(f"나주노 TOC에 성가 {number} 없음")
        return re.sub(rf"^\s*{number}\s*", "", entry[0]).strip()
    # 야훼 이레: 제목 줄에서 번호 뒤 텍스트 추출
    pi = find_song_page(source, number)
    with _open(source) as d:
        page = d[pi]
        for line in page.get_text("text").splitlines():
            m = re.match(rf"^\s*{number}\.\s*(.+)", line)
            if m:
                return _strip_english_subtitle(m.group(1).strip())
    return ""


def _strip_english_subtitle(title: str) -> str:
    """야훼 이레 제목의 영문 부제를 제거해 한글 제목만 남긴다(리더 결정 B, 2026-09-10).

    야훼 PDF는 제목 줄에 한글 제목과 영문 부제를 한 줄로 병기한다
    (810: '주의 자비가 내려와 Mercy is falling'). 헤더를 샘플(447)처럼 한글 단일 줄로 통일하기
    위해, **첫 라틴 알파벳([A-Za-z]) 등장 지점에서 자르고 그 앞 공백도 제거**한다.
    라틴 문자가 없으면 원문 그대로 둔다.

    일반화 리스크(구현 노트 '남은 리스크' 참고): 한글 제목 자체에 영문 단어가 섞인 곡이 있으면
    잘못 잘린다. 검증 5곡 중 야훼는 810뿐이고 이 규칙으로 정확히 '주의 자비가 내려와'가 된다."""
    m = re.search(r"[A-Za-z]", title)
    if m:
        return title[: m.start()].rstrip()
    return title


def _title_hits_below(page, number, y_floor=None):
    """페이지에서 '{number}.' 좌측여백 제목 히트들(y_floor 아래로 제한 가능)."""
    hits = [h for h in page.search_for(f"{number}.") if h.x0 < _TITLE_LEFT_MARGIN_X]
    if y_floor is not None:
        hits = [h for h in hits if h.y0 > y_floor]
    return hits


def _other_song_numbers_on_page(page):
    """페이지 텍스트에서 줄머리 '숫자.' 패턴으로 곡 번호 후보 수집(설계서 §3.3)."""
    nums = set()
    for line in page.get_text("text").splitlines():
        m = re.match(r"^\s*(\d{1,4})\.", line)
        if m:
            nums.add(int(m.group(1)))
    return nums


def _first_trigger_y_below(page, y_floor):
    """y_floor 아래에서 처음 등장하는 저작권 트리거 줄의 y0. 없으면 None.

    'Copyright' 줄 자체가 곡의 하단 경계 후보다. 트리거(Administered/Adm.)가 없어도
    'Copyright' 줄이 있으면 그 y를 쓴다."""
    cands = []
    for term in ["Copyright"] + COPYRIGHT_TRIGGERS:
        for h in page.search_for(term):
            if h.y0 > y_floor:
                cands.append(h.y0)
    return min(cands) if cands else None


def find_song_y_range(page, source: str, number: int):
    """곡의 세로 크롭 범위 (top_y, bottom_y) in PDF pt (설계서 §3.3, 야훼 다곡 경계).

    top = 이 곡 제목 아래(제목 제외), bottom = (다음 곡 제목 top, 저작권 top, 페이지 하단)
    중 가장 위. 809/810/811 실측으로 검증."""
    hits = _title_hits_below(page, number)
    if not hits:
        raise ValueError(f"{source} p{page.number}: 성가 {number} 제목 미검출")
    title_top = min(h.y0 for h in hits)
    title_bottom = max(h.y1 for h in hits)

    next_tops = []
    for other in _other_song_numbers_on_page(page):
        if other == number:
            continue
        for h in _title_hits_below(page, other, y_floor=title_bottom):
            next_tops.append(h.y0)
    next_song_top = min(next_tops) if next_tops else None

    copyright_top = _first_trigger_y_below(page, title_bottom)

    candidates = [y for y in (next_song_top, copyright_top, page.rect.height) if y is not None]
    bottom = min(candidates)
    return title_bottom, bottom


# ---------------------------------------------------------------------------
# 이미지 렌더
# ---------------------------------------------------------------------------

def render_region(page, y_range, dpi: int = RENDER_DPI) -> Image.Image:
    """곡 영역(y_range=(top,bottom) pt, None이면 페이지 전체)을 렌더한 PIL 이미지.

    나주노 스캔 페이지는 텍스트 레이어가 없어도 get_pixmap이 래스터를 그대로 렌더한다."""
    clip = None
    if y_range is not None:
        top, bottom = y_range
        clip = fitz.Rect(0, top, page.rect.width, bottom)
    pix = page.get_pixmap(dpi=dpi, clip=clip)
    mode = "RGBA" if pix.alpha else "RGB"
    img = Image.frombytes(mode, (pix.width, pix.height), pix.samples)
    return img.convert("RGB")


# ---------------------------------------------------------------------------
# 오선보 시스템 밴드 탐지 (설계서 §3.4)
# ---------------------------------------------------------------------------

def _blank_runs(blank: np.ndarray):
    """blank(bool 1D)에서 연속 True 구간을 (start, length)로 반환."""
    runs = []
    i = 0
    n = len(blank)
    while i < n:
        if blank[i]:
            s = i
            while i < n and blank[i]:
                i += 1
            runs.append((s, i - s))
        else:
            i += 1
    return runs


def _nonblank_runs(row_ink: np.ndarray):
    """row_ink(1D)에서 잉크가 있는(>0) 연속 구간을 (start, end) inclusive로 반환."""
    runs = []
    i = 0
    n = len(row_ink)
    while i < n:
        if row_ink[i] > 0:
            s = i
            while i < n and row_ink[i] > 0:
                i += 1
            runs.append((s, i - 1))
        else:
            i += 1
    return runs


def _trim_top_above_chords(row_ink: np.ndarray, t: int, staff_top: int) -> int:
    """세그먼트 상단에서 오선 위 텍스트(작곡가/제목)를 제외한 crop 상단을 반환.

    오선 바로 위 화음(코드) 줄은 크롭에 포함해야 하지만(요구사항), 그보다 더 위의 작곡가명·
    제목은 제외해야 한다. 첫 시스템은 작곡가명이 코드 줄과 min_gap 미만으로 붙어 같은 세그먼트에
    병합된다(실측: 173 'Claire Cloninger / Don Moen', 362 'Paullo Roberto'). 오선 위 영역
    [t, staff_top)을 공백행으로 블록 분할해, 오선에 가장 가까운(마지막) 블록=코드 줄만 남기고
    그 위 블록(작곡가/제목)은 잘라낸다. 블록이 하나뿐이면(코드가 오선에 인접) 그대로 둔다.
    """
    if staff_top <= t:
        return t
    blocks = _nonblank_runs(row_ink[t:staff_top])
    if len(blocks) <= 1:
        return t
    return t + blocks[-1][0]


def detect_system_bands(gray: np.ndarray) -> list:
    """오선보 시스템별 세로 밴드 [(top, bottom), …]를 반환(상단 제목/작곡가·하단 저작권 제외).

    1) 콘텐츠 좌우 경계로 트리밍한 폭 기준 행별 잉크 비율(row_fill)을 구한다.
    2) 공백 밴드(row_fill<0.005)가 시스템 간격(_SYSTEM_GAP_MIN 이상)이면 세그먼트 경계로 삼아
       콘텐츠를 세그먼트로 분할한다(그보다 짧은 공백은 시스템 내부로 병합).
    3) 세그먼트 중 오선행(row_fill>0.5)이 _MIN_STAFF_ROWS 이상인 것만 '시스템'으로 인정.
       오선이 없는 세그먼트(상단 제목/작곡가 줄, 하단 저작권 줄)는 자연히 제외된다 — 번호가
       좌/우 어디 있든 동일(146 제목·번호 좌우 반대여도 동작, 요구사항 §추가 검증).
    """
    ink = gray < _INK
    h, w = gray.shape
    left, right, _t, _b = ps.detect_content_bounds(gray)
    cw = max(1, right - left + 1)
    row_ink = ink[:, left:right + 1].sum(axis=1)
    row_fill = row_ink / cw
    blank = row_fill < _BLANK_ROW_FILL

    min_gap = max(_SYSTEM_GAP_MIN, round(h * _SYSTEM_GAP_RATIO))
    is_gap = np.zeros(h, dtype=bool)
    for s, length in _blank_runs(blank):
        if length >= min_gap:
            is_gap[s:s + length] = True

    systems = []
    i = 0
    while i < h:
        if not is_gap[i]:
            s = i
            while i < h and not is_gap[i]:
                i += 1
            seg_top, seg_bot = s, i - 1
            # 세그먼트 안의 실제 잉크 범위로 상하 여백 트림
            seg_rows = np.where(row_ink[seg_top:seg_bot + 1] > 0)[0]
            if not len(seg_rows):
                continue
            t = seg_top + int(seg_rows.min())
            b = seg_top + int(seg_rows.max())
            staff_line_rows = np.where(row_fill[t:b + 1] > _STAFF_ROW_FILL)[0]
            if len(staff_line_rows) >= _MIN_STAFF_ROWS:
                staff_top = t + int(staff_line_rows.min())
                t = _trim_top_above_chords(row_ink, t, staff_top)
                systems.append((t, b))
        else:
            i += 1
    return systems


def group_systems(bands: list, per_slide: int = 2) -> list:
    """시스템 밴드를 per_slide개씩 묶는다(마지막 홀수는 1개). ceil(n/per_slide)개 그룹.

    고정 per_slide 패킹(참조·단순 카운트용). 실제 슬라이드 조립은 pack_systems(동적 패킹)를
    쓴다."""
    return [bands[i:i + per_slide] for i in range(0, len(bands), per_slide)]


def _group_content_bbox(gray: np.ndarray, group: list) -> tuple:
    """그룹(연속 시스템 밴드들)의 **행 범위 안에서만** 잉크 열의 좌우 경계를 잡아 크롭 박스
    (left, top, right, bottom)를 반환한다. right/bottom은 PIL crop 관례상 exclusive.

    전역 ps.detect_content_bounds는 저작권/제목 줄까지 폭에 섞어(447 전역 우측=1764px가 악보
    실제 우측 ~1590px보다 넓음) 각 크롭 우측에 ~10% 흰 여백을 남긴다(요청 항목 6 버그). 그룹의
    실제 행 범위 [group[0][0] .. group[-1][1]](사이 시스템 간격 포함) 안에서만 잉크 열을 보면
    저작권 줄이 폭에 섞이지 않아 샘플처럼 그룹별로 타이트하게 잘린다."""
    top = group[0][0]
    bottom = group[-1][1]
    band = gray[top:bottom + 1] < _INK
    cols = np.where(band.any(axis=0))[0]
    if not len(cols):
        h, w = gray.shape
        return (0, top, w, bottom + 1)
    return (int(cols.min()), top, int(cols.max()) + 1, bottom + 1)


# 동적 패킹(요청 항목 5): "같은 배율로 겹침 없이 들어가면" 최대 3개까지 한 슬라이드에 묶되,
# 실제로 슬라이드 수를 줄이거나(min-slide) 균형이 나아질 때만 3-pack을 채택한다(순수 greedy는
# 810형 곡을 97% 빡빡 [3,1]로 악화시키므로 비채택). safety는 밴드 세로 채움 상한(빡빡 방지).
_PACK_SAFETY = 0.95
_PACK_MAX_PER_SLIDE = 3


def _group_fits_band(gray: np.ndarray, group: list, safety: float) -> bool:
    """그룹을 밴드 폭에 꽉 채우는 배율(width-fill)로 축소했을 때 세로가 밴드 안에 들어가나.

    width-limited이면 scaled_h = ch × (bandW/cw) ≤ bandH ⇔ ch/cw ≤ bandH/bandW = band_aspect.
    safety로 상한을 낮춰(밴드 세로의 safety까지만 허용) 810형 97% 빡빡 3-pack을 배제한다.
    폭은 반드시 per-group 크롭 폭(_group_content_bbox)으로 — 전역 폭을 쓰면 폭이 과대해져
    종횡비가 과소평가되고 3-pack이 오판된다(항목 6 헬퍼 공유)."""
    band_aspect = _SCORE_BAND_H / _SCORE_BAND_W
    left, top, right, bottom = _group_content_bbox(gray, group)
    cw = max(1, right - left)
    ch = max(1, bottom - top)
    return (ch / cw) <= band_aspect * safety


def pack_systems(bands: list, gray: np.ndarray,
                 max_per_slide: int = _PACK_MAX_PER_SLIDE,
                 safety: float = _PACK_SAFETY) -> list:
    """시스템 밴드를 슬라이드 그룹으로 동적 패킹. min-slide 균형 패커.

    선택 기준(사전식 최소화): (1) 슬라이드 수 최소, (2) 최대 그룹 크기 최소(균형), (3) 앞쪽에
    큰 그룹을 몰아 홀수 나머지 1개를 뒤로(현재 [2,2,…,1] 관례 재현). 크기 1은 항상 허용(항상
    해가 존재하도록), 크기 2 이상은 _group_fits_band를 만족해야 한다. n(≤10)이 작아 완전 열거."""
    n = len(bands)
    best = None

    def rec(start, sizes):
        nonlocal best
        if start == n:
            key = (len(sizes), max(sizes), tuple(-s for s in sizes))
            if best is None or key < best[0]:
                best = (key, list(sizes))
            return
        for size in range(1, max_per_slide + 1):
            if start + size > n:
                break
            group = bands[start:start + size]
            if size == 1 or _group_fits_band(gray, group, safety):
                sizes.append(size)
                rec(start + size, sizes)
                sizes.pop()

    rec(0, [])
    if best is None:
        return group_systems(bands, 2)
    groups, i = [], 0
    for size in best[1]:
        groups.append(bands[i:i + size])
        i += size
    return groups


# ---------------------------------------------------------------------------
# 저작권 크롭 (설계서 §3.7, §3.8)
# ---------------------------------------------------------------------------

NAJU_COPYRIGHT_CACHE = _BASE / "assets" / "나주노_copyright_bbox_cache.json"


def _load_naju_cache() -> dict:
    try:
        return json.loads(NAJU_COPYRIGHT_CACHE.read_text(encoding="utf-8"))
    except FileNotFoundError:
        logger.warning("나주노 저작권 캐시 파일 없음: %s", NAJU_COPYRIGHT_CACHE)
        return {}


def _render_pt_rect(page, rect, dpi: int = RENDER_DPI) -> Image.Image:
    pix = page.get_pixmap(dpi=dpi, clip=rect)
    mode = "RGBA" if pix.alpha else "RGB"
    return Image.frombytes(mode, (pix.width, pix.height), pix.samples).convert("RGB")


def yahwe_copyright_bounds(page, number: int):
    """야훼 이레 저작권 줄의 크롭 경계 (left_x, right_x, y0, y1) in PDF pt. 없으면 None.

    left = 'Copyright' 단어 좌측, right = 같은 줄에서 처음 매치되는 트리거(Administered/Adm.)
    직전(설계서 §3.7). 트리거가 없으면 그 줄의 마지막 단어 우측까지(요구사항 §Copyright 변형).
    테스트가 경계를 직접 검증할 수 있도록 별도 함수로 노출한다."""
    top, bottom = find_song_y_range(page, "야훼 이레", number)
    cw_hits = page.search_for("Copyright")
    # 이 곡의 저작권 줄 = y-범위 하단 경계(bottom≈저작권 top) 부근의 Copyright
    cw_hits = [h for h in cw_hits if top <= h.y0 <= page.rect.height]
    if not cw_hits:
        return None
    cr = min(cw_hits, key=lambda h: abs(h.y0 - bottom))
    line_y0, line_y1, left_x = cr.y0, cr.y1, cr.x0

    right_x = None
    for trig in COPYRIGHT_TRIGGERS:
        for h in page.search_for(trig):
            if abs(h.y0 - line_y0) < 5 and h.x0 > left_x:
                right_x = h.x0 if right_x is None else min(right_x, h.x0)
    if right_x is None:
        # 트리거 없음 → 같은 줄 마지막 단어 우측까지
        line_words = [w for w in page.get_text("words")
                      if abs(w[1] - line_y0) < 5 and w[0] >= left_x]
        right_x = max((w[2] for w in line_words), default=page.rect.width)
    return left_x, right_x, line_y0, line_y1


def _trim_trailing_period_px(pil_img: Image.Image) -> Image.Image:
    """렌더된 저작권 크롭의 우측 꼬리 마침표를 세로 잉크 열 런 분석으로 제거한다(항목 4-1).

    트리거 유무·경로(트리거/폴백)와 무관한 공통 후처리다. 마지막 잉크 런이 마침표 폭
    (≤_PERIOD_MAX_W_PX @300dpi)이고 앞에 다른 런이 있으면 그 앞 런 끝까지로 크롭을 좁힌다.
    이름 내부 마침표(267식 '천태혁. 진경')는 마지막 런이 아니라 그대로 보존된다."""
    col = (np.asarray(pil_img.convert("L")) < _INK).sum(axis=0)
    runs = []
    i, n = 0, len(col)
    while i < n:
        if col[i] > 0:
            s = i
            while i < n and col[i] > 0:
                i += 1
            runs.append((s, i - 1))
        else:
            i += 1
    if len(runs) >= 2:
        lx0, lx1 = runs[-1]
        if (lx1 - lx0 + 1) <= _PERIOD_MAX_W_PX:
            return pil_img.crop((0, 0, runs[-2][1] + 1, pil_img.height))
    return pil_img


def _yahwe_copyright_crop(number: int, page) -> Image.Image:
    b = yahwe_copyright_bounds(page, number)
    if b is None:
        return None
    left_x, right_x, y0, y1 = b
    # 트리거 첫 글자 침범 방지로 우측 2pt 후퇴, 상하 소폭 여유
    rect = fitz.Rect(left_x - 1, y0 - 2, max(left_x + 1, right_x - 2), y1 + 2)
    return _trim_trailing_period_px(_render_pt_rect(page, rect))


def _naju_copyright_crop(number: int, page) -> Image.Image:
    entry = _load_naju_cache().get(str(number))
    if entry is None:
        logger.warning("[경고] 성가 %s: Copyright 자동 인식 실패 — 수동 확인 필요", number)
        return None
    if not entry.get("has_copyright"):
        return None
    x0, y0, x1, y1 = entry["bbox"]
    img = render_region(page, None, RENDER_DPI)  # 나주노는 페이지 전체 렌더
    w, h = img.size
    return img.crop((round(x0 * w), round(y0 * h), round(x1 * w), round(y1 * h)))


def resolve_copyright_crop(source: str, number: int, page) -> Image.Image:
    """저작권 크롭 PIL 이미지. 저작권 없음/캐시 미스면 None(요소 생략, 잘못된 크롭 금지)."""
    if source == "나주노":
        return _naju_copyright_crop(number, page)
    return _yahwe_copyright_crop(number, page)


# ---------------------------------------------------------------------------
# 슬라이드 조립 (설계서 §3.6)
# ---------------------------------------------------------------------------

TEMPLATE = _BASE / "assets" / "청년미사_성가_template.pptx"

# 슬라이드/배치 상수(EMU). 샘플 실측: 슬라이드 9144000×5143500, 헤더 높이 523220,
# 저작권 (7020272,4855468) 크기 (1490253,288032)(우하단, 하단이 슬라이드 바닥에 붙음).
SLIDE_W, SLIDE_H = 9144000, 5143500
HEADER_H = 523220
# 악보 배치 밴드: 헤더 아래 ~ 저작권 위. 저작권 유무와 무관하게 동일 밴드를 쓴다(단순·일관).
_SCORE_BAND_LEFT = 300000
_SCORE_BAND_W = SLIDE_W - 2 * _SCORE_BAND_LEFT
_SCORE_BAND_TOP = HEADER_H + 220000
_SCORE_BAND_BOTTOM = 4700000
_SCORE_BAND_H = _SCORE_BAND_BOTTOM - _SCORE_BAND_TOP
# 저작권 배치: 높이 고정, 폭은 각 곡 자신의 크롭 종횡비 유지로 비례 조정(우하단 정렬).
# 요청 항목 4-2: 샘플 실측 288032의 0.9배로 전곡 통일(폭은 곡마다 글자 수가 달라 다른 것이 정상).
_COPY_H = round(288032 * 0.9)  # = 259229
_COPY_RIGHT = 7020272 + 1490253       # 샘플 우측 끝 = 8510525
_COPY_BOTTOM = 4855468 + 288032       # = 5143500 (슬라이드 바닥)


def _find_header(slide):
    """헤더 도형을 이름으로 찾는다(텍스트 부분매칭 금지 — CLAUDE.md 부분문자열 충돌 함정)."""
    for sh in slide.shapes:
        if sh.name == "Rectangle 11":
            return sh
    raise RuntimeError("헤더 'Rectangle 11'를 찾지 못함")


def _set_run_text(r, text):
    t = r.find(qn("a:t"))
    if t is None:
        from pptx.oxml import parse_xml
        A = "http://schemas.openxmlformats.org/drawingml/2006/main"
        t = parse_xml(f'<a:t xmlns:a="{A}"/>')
        r.append(t)
    t.text = text


def build_header_runs(slide, 구분: str, 출처: str, number, title: str) -> None:
    """헤더 run을 [구분(흰색)][출처][번호][제목](출처·번호·제목 FFC000)으로 채운다.

    템플릿 헤더의 run0(흰색 라벨)·run1(FFC000)의 rPr을 그대로 쓰거나 deepcopy해 텍스트만
    채우므로 색/폰트/자간(spc300)이 보존된다(설계서 §3.6, run 개수 가정·endParaRPr append
    함정 회피 — _insert_run_before_end_para_rpr 재사용). 색을 모르는 기존 run을 편집하는
    _update_prefix_in_runs 경로가 아니라, 서식을 아는 run을 새로 채우는 경로라 길이-오염
    위험이 없다."""
    if 구분 not in ou.HYMN_TYPES:
        raise ValueError(f"알 수 없는 성가 구분: {구분!r} (지원: {ou.HYMN_TYPES})")
    hdr = _find_header(slide)
    p = hdr.text_frame.paragraphs[0]._p
    runs = p.findall(qn("a:r"))
    if len(runs) < 2:
        raise RuntimeError("헤더 템플릿 run이 2개 미만 — 템플릿 자산 손상")
    label_r, color_r = runs[0], runs[1]
    _set_run_text(label_r, f"{구분} ")
    _set_run_text(color_r, SOURCE_DISPLAY.get(str(출처), str(출처)))  # 표시명만 치환(항목 3)
    for extra in runs[2:]:
        p.remove(extra)
    title_runs = []
    for seg in (" ", str(number), " ", str(title)):
        new_r = copy.deepcopy(color_r)
        _set_run_text(new_r, seg)
        ps._insert_run_before_end_para_rpr(p, new_r)
        if seg == str(title):
            title_runs.append(new_r)
    _autosize_header(hdr, p.findall(qn("a:r")), title_runs)


@functools.lru_cache(maxsize=None)
def _gungsuh_font(size_pt: int):
    from PIL import ImageFont
    return ImageFont.truetype(_HEADER_FONT_PATH, int(round(size_pt)), index=_HEADER_FONT_INDEX)


def _run_spc_sz(r):
    rPr = r.find(qn("a:rPr"))
    spc = int(rPr.get("spc")) if rPr is not None and rPr.get("spc") else 0
    sz = int(rPr.get("sz")) if rPr is not None and rPr.get("sz") else 1800
    return rPr, spc, sz


def _measure_run_emu(r) -> float:
    """run 한 개의 렌더 폭(EMU). getlength(pt)=px 가정 + 자간(spc는 1/100pt/글자) ×12700."""
    _rPr, spc, sz = _run_spc_sz(r)
    t = r.find(qn("a:t"))
    text = (t.text or "") if t is not None else ""
    adv_pt = _gungsuh_font(sz / 100.0).getlength(text)
    spc_pt = spc / 100.0 * len(text)
    return (adv_pt + spc_pt) * _EMU_PER_PT


def _header_total_emu(runs) -> float:
    """헤더 run 전체의 렌더 폭 합(EMU)."""
    return sum(_measure_run_emu(r) for r in runs)


def _usable_header_emu(hdr) -> int:
    """헤더 도형의 텍스트 사용가능 폭(EMU). bodyPr lIns/rIns를 실제로 읽어 계산(하드코딩 X)."""
    body = hdr.text_frame._txBody.find(qn("a:bodyPr"))
    lIns = int(body.get("lIns")) if body is not None and body.get("lIns") else 91440
    rIns = int(body.get("rIns")) if body is not None and body.get("rIns") else 91440
    return hdr.width - lIns - rIns


def _autosize_header(hdr, all_runs, title_runs) -> None:
    """헤더 제목이 도형 폭을 넘치면 2단계로 축소: (1) 전 run 자간 spc→0, (2) 제목 run 폰트
    크기만 축소(구분/출처/번호는 28pt 유지). 넘치지 않으면 무동작(요청 항목 2).

    현재 6곡은 무발동이 정상(가장 긴 447도 여유 있음) — 미래의 더 긴 제목/영문부제 미제거 곡을
    위한 방어 기능. 폰트 미존재 환경에서는 측정 불가라 무동작(폴백)."""
    try:
        limit = _usable_header_emu(hdr) * _HEADER_SAFETY

        def total():
            return _header_total_emu(all_runs)

        if total() <= limit:
            return
        for r in all_runs:                       # 1) 자간 0(모든 run 일괄 — 부분만 바꾸면 불균형)
            rPr = r.find(qn("a:rPr"))
            if rPr is not None:
                rPr.set("spc", "0")
        if total() <= limit:
            return
        for tr in title_runs:                    # 2) 제목 run만 폰트 축소
            tr_rPr = tr.find(qn("a:rPr"))
            if tr_rPr is None:
                continue
            sz = int(tr_rPr.get("sz") or "2800")
            while sz > _MIN_TITLE_PT * 100 and total() > limit:
                sz -= 100
                tr_rPr.set("sz", str(sz))
    except (OSError, ImportError):
        return


def _content_crop_box(region_img, gray, group):
    """그룹(연속 밴드들)을 하나의 악보 이미지로 크롭할 박스(left,top,right,bottom).

    좌우를 전역 콘텐츠 경계가 아니라 **그룹 행 범위 안의 잉크**로 잡는다(항목 6): 전역 경계는
    저작권/제목 줄이 폭을 오염시켜 각 크롭 우측에 흰 여백을 남긴다."""
    return _group_content_bbox(gray, group)


def _place_score(slide, pil_img):
    """악보 이미지를 밴드에 종횡비 유지 축소·가운데 정렬로 배치(슬라이드별 독립 스케일)."""
    iw, ih = pil_img.width, pil_img.height
    scale = min(_SCORE_BAND_W / iw, _SCORE_BAND_H / ih)
    w = int(round(iw * scale))
    h = int(round(ih * scale))
    left = _SCORE_BAND_LEFT + (_SCORE_BAND_W - w) // 2
    top = _SCORE_BAND_TOP + (_SCORE_BAND_H - h) // 2
    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    buf.seek(0)
    slide.shapes.add_picture(buf, left, top, w, h)


def _place_copyright(slide, pil_img):
    """저작권 이미지를 우하단에 배치(높이 고정, 폭은 종횡비 유지). 샘플과 동일 우측끝/바닥."""
    aspect = pil_img.width / pil_img.height
    h = _COPY_H
    w = int(round(h * aspect))
    left = _COPY_RIGHT - w
    top = _COPY_BOTTOM - h
    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    buf.seek(0)
    slide.shapes.add_picture(buf, left, top, w, h)


def build_hymn_pptx(구분: str, 출처: str, number: int, title: str = None) -> Presentation:
    """성가 구분·출처·번호 → 샘플과 동일 포맷의 슬라이드 여러 장짜리 Presentation.

    부작용 없음(파일 저장은 호출자 몫). 템플릿을 duplicate_slide로 필요한 장수만큼 복제한 뒤
    각 슬라이드에 헤더·악보·저작권을 채운다. 복제 후 _set_slide_bg_black류 배경 재설정을 하지
    않는다(템플릿 배경이 이미 정답 — CLAUDE.md '슬라이드 복사 시 배경 재설정 금지')."""
    if 출처 not in SOURCES:
        raise ValueError(f"알 수 없는 출처: {출처!r}")
    if 구분 not in ou.HYMN_TYPES:
        raise ValueError(f"알 수 없는 성가 구분: {구분!r}")

    if title is None:
        title = find_song_title(출처, number)

    pi = find_song_page(출처, number)
    doc = fitz.open(str(_resolve_pdf_path(출처)))
    try:
        page = doc[pi]
        y_range = None if SOURCES[출처]["kind"] == "scan" else find_song_y_range(page, 출처, number)
        region_img = render_region(page, y_range, RENDER_DPI)
        gray = np.asarray(region_img.convert("L"))
        bands = detect_system_bands(gray)
        if not bands:
            raise RuntimeError(f"{출처} {number}: 시스템 밴드 미검출")
        groups = pack_systems(bands, gray)
        copyright_img = resolve_copyright_crop(출처, number, page)
    finally:
        doc.close()

    prs = Presentation(str(TEMPLATE))
    # 템플릿은 1장 → 필요한 만큼 복제(맨 끝에 추가되므로 인덱스 0..N-1)
    for _ in range(len(groups) - 1):
        ou.duplicate_slide(prs, 0)

    for si, group in enumerate(groups):
        slide = prs.slides[si]
        build_header_runs(slide, 구분, 출처, number, title)
        box = _content_crop_box(region_img, gray, group)
        _place_score(slide, region_img.crop(box))
        if copyright_img is not None:
            _place_copyright(slide, copyright_img)
    return prs


def main(argv=None) -> int:
    import argparse
    argv = sys.argv[1:] if argv is None else argv
    ap = argparse.ArgumentParser()
    ap.add_argument("--출처", required=True, choices=list(SOURCES))
    ap.add_argument("--번호", required=True, type=int)
    ap.add_argument("--구분", required=True, choices=ou.HYMN_TYPES)
    ap.add_argument("--제목", default=None)
    ap.add_argument("--out", default=None)
    ns = ap.parse_args(argv)
    prs = build_hymn_pptx(ns.구분, getattr(ns, "출처"), getattr(ns, "번호"), ns.제목)
    out = ns.out or f"청년성가_{getattr(ns,'출처')}_{getattr(ns,'번호')}.pptx"
    prs.save(out)
    print("저장:", out, "슬라이드", len(prs.slides))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
