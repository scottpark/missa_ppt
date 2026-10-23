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

try:
    import pymupdf as fitz  # PyMuPDF 1.24.3+. `import fitz`는 "fitz API is deprecated" 경고를 낸다(2026-10-25).
except ImportError:  # 옛 PyMuPDF
    import fitz
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

# 오선 시스템 판정 보강(2026-10-18, 나주노 697곡 전수 실측). 스캔이 기울거나 흐린 곡(479/123/127/
# 495 등)은 오선이 여러 행에 번져 row_fill이 0.4를 못 넘어 진짜 시스템이 _MIN_STAFF_ROWS 미만으로
# 탈락했다(479는 5줄 중 3줄만 생성). row_fill 대신 **열 연속 커버리지**(세그먼트 안에서 잉크가
# 끊김 없이 이어지는 가장 긴 가로 구간 / 콘텐츠 폭)로 오선을 판정한다: 오선은 기울어도 전폭에
# 걸쳐 이어지지만(실측 시스템 99% 이상 ≥0.89) 제목·작곡가·저작권 텍스트는 단어 사이 공백 때문에
# 낮다(실측 저작권 ≤0.889, 대부분 ≤0.78, 제목 ≤0.52). 저작권(≤85px)과 겹치는 0.85~0.89 구간은
# 높이로 가른다: 흐린 진짜 시스템은 높이 ≥190px, 저작권 줄은 ≤85px.
_SYSTEM_MIN_COVERAGE = 0.80
_SYSTEM_MIN_HEIGHT = 150
_SYSTEM_MIN_PEAK_FILL = 0.30   # 높이 경로의 오선 밀도 하한(가사 줄 row_fill은 이보다 낮다)
_TRAILING_TEXT_MAX_HEIGHT = 200  # 흡수할 보충줄 블록의 최대 높이(가사 3~4줄 ≈127px, 359번). 다음 곡 제목은 곡 영역 절단으로 이미 제외
_COVERAGE_BRIDGE_PX = 8       # 이보다 짧은 잉크 끊김(스캔 드롭아웃)은 이어진 것으로 본다
# 시스템 바로 아래에 붙은 가사 보충줄(2절 가사 "2.예--수" 등, 공백행 ≥_SYSTEM_GAP_MIN로 분리돼
# 별도 세그먼트가 된 것)을 앞 시스템에 흡수하는 최대 간격(px@300dpi). 468번 2절 가사가 잘렸던
# 원인. 페이지 하단 _COPYRIGHT_ZONE_RATIO 아래(저작권 줄)는 흡수하지 않는다.
_TRAILING_TEXT_MAX_GAP = 120
_COPYRIGHT_ZONE_RATIO = 0.90
_COPYRIGHT_LINE_MAX_H = 45      # 저작권 한 줄의 최대 높이(px@300dpi) — 넘으면 2줄 블록으로 본다

# 실행 파일 옆 자산 폴더(missa_psalm_score_image._BASE와 동일 관용, PyInstaller onefile 회피).
_BASE = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent

# 배포 zip이 원본 PDF 2개를 이 경로에 그대로 포함해 배포한다(tools/build_dist_zip.py).
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


# 시스템(오선 묶음) 사이 간격은 시스템 내부 오선 5줄 사이 간격보다 훨씬 크다 — 7개 곡
# (나주노 151/362/173/146/267/172, 야훼 이레 810) 오선행(row_fill>_STAFF_ROW_FILL) 군집
# 실측: 시스템 내부 최대 간격 21px(151) vs 시스템 사이 최소 간격 60px(362). 40은 이 범위
# 중앙에 가깝고 양쪽에 충분한 여유(19px/20px)가 있다. _split_merged_systems()가 쓴다.
_SYSTEM_LINE_MERGE_GAP = 40


def _staff_threshold(row_fill: np.ndarray, t: int, b: int) -> float:
    """세그먼트 [t,b]의 오선행 임계. 기본 _STAFF_ROW_FILL이되, 오선이 번진 흐린 스캔은 그
    세그먼트 최대 row_fill의 60%까지 낮춘다(479 시스템4·5는 max 0.44라 0.40 기준으로는 2~4행)."""
    return min(_STAFF_ROW_FILL, 0.6 * float(row_fill[t:b + 1].max()))


_CHORD_MAX_GAP = 35   # 다음 시스템 오선 바로 위 화음 줄로 볼 수 있는 최대 간격(px@300dpi)
_THIN_BLOCK_H = 8      # 이 높이 이하의 글줄 블록은 글자가 아니라 가로선(볼타 괄호 등)으로 본다
_BLOCK_MIN_INK = 5    # 이 이상 잉크가 있는 행만 '글줄'로 센다(스캔 얼룩 행 ≤7px 무시용 하한)


def _system_boundary(row_ink: np.ndarray, lo: int, hi: int) -> int:
    """오선 군집 사이 구간 (lo, hi)에서 두 시스템의 경계 행.

    사이 구간을 글줄 블록(행 잉크 ≥ _BLOCK_MIN_INK의 연속 구간)으로 나눈다. 마지막 블록이 다음
    시스템 오선 바로 위(≤ _CHORD_MAX_GAP)면 그것은 다음 시스템의 화음 줄이므로 그 위에서 자르고
    (앞 블록들=가사 1·2째 줄 등은 윗 시스템 몫), 아니면 모든 블록이 윗 시스템 가사이므로 마지막
    블록 아래에서 자른다. 블록이 없으면 중간점.

    이전 방식(군집 중간점 → 692 가사 2째 줄 잘림, 가장 긴 공백 런 → 212 가사줄이 아랫 시스템으로
    넘어감)은 줄 간격이 곡마다 달라 어긋났다. 화음 줄은 항상 오선에 바짝 붙는다는 구조적 단서를 쓴다."""
    blocks = _nonblank_runs((row_ink[lo + 1:hi] >= _BLOCK_MIN_INK).astype(int))
    if not blocks:
        return (lo + hi) // 2
    ls, le = blocks[-1]
    last_start, last_end = lo + 1 + ls, lo + 1 + le
    if hi - last_end <= _CHORD_MAX_GAP:
        # 화음 줄 바로 위의 얇은 선(반복 종지 괄호 '1.'/'2.'의 가로선, ≤_THIN_BLOCK_H)도 다음 시스템 몫
        k = len(blocks) - 1
        while k > 0 and blocks[k - 1][1] - blocks[k - 1][0] + 1 <= _THIN_BLOCK_H                 and blocks[k][0] - blocks[k - 1][1] <= _CHORD_MAX_GAP:
            k -= 1
            last_start = lo + 1 + blocks[k][0]
        return max(lo + 1, last_start - 1)
    return min(hi - 1, last_end + 1)


def _split_merged_systems(row_fill: np.ndarray, row_ink: np.ndarray, t: int, b: int,
                          ink: np.ndarray = None, left: int = 0, right: int = 0) -> list:
    """[t, b] 세그먼트가 실제로는 여러 시스템인데 공백행 연속 길이 기반 1차 분할
    (_page_segments의 _SYSTEM_GAP_MIN/RATIO)이 하나로 합쳐버린 경우, 오선행
    (row_fill>_staff_threshold) 밀도 군집으로 다시 쪼갠다.

    1차 분할이 이미 시스템 하나만 올바르게 담고 있으면(대다수 곡) 오선행 군집도 하나뿐이라
    [(t, b)] 그대로(입력과 동일) 반환해 기존 동작과 완전히 같다 — 여러 시스템이 실제로
    합쳐진 경우에만 분기한다.

    **발견 경위 (2026-10-04):** 나주노 172번("나의 고백") 페이지에서 시스템 2~9 사이의
    여백 구간에 가사 꼬리 등 미세한 잉크가 산발적으로 섞여 있어, 공백행(row_fill<0.005)
    연속 런이 전부 _SYSTEM_GAP_MIN(45px) 밑으로 끊겼다(실측 최대 36px) — 그 결과
    시스템 8개가 세그먼트 하나로 합쳐져, 생성된 PPT가 슬라이드 1장에
    시스템 1개(정상), 슬라이드 2장째에 8개를 욱여넣은(가독 불가) 결과를 냈다. 오선행
    밀도만 보는 군집화는 가사/화음 텍스트의 잉크 유무에 흔들리지 않는다(시스템 내부 간격
    21px 이하 vs 시스템 사이 간격 60px 이상 — 10배 이상 margin, 실측 상세는
    docs/ooxml-pitfalls-log.md 참고).

    **경계 보정 (2026-10-18):** 시스템 사이 경계는 군집 중간점이 아니라 두 군집 사이 가장 긴
    화음 줄 단서로 정한다(`_system_boundary`; 692번 가사 2째 줄 잘림, 212번 가사줄 이동)."""
    thr = _staff_threshold(row_fill, t, b)
    staff_rows = np.where(row_fill[t:b + 1] > thr)[0] + t
    if not len(staff_rows):
        return [(t, b)]

    groups = []
    s = p = int(staff_rows[0])
    for r in staff_rows[1:]:
        r = int(r)
        if r - p > _SYSTEM_LINE_MERGE_GAP:
            groups.append((s, p))
            s = r
        p = r
    groups.append((s, p))

    # 각 군집의 실제 오선행 개수(span이 아니라 row_fill>threshold인 행 수)가
    # _MIN_STAFF_ROWS 미만이면 노이즈(점 하나 등)이므로 버린다 — 151 저작권 줄 오검출을
    # 막던 것과 동일한 기준(_MIN_STAFF_ROWS)을 재사용한다.
    accepted = [
        (gs, ge) for gs, ge in groups
        if int(np.sum(row_fill[gs:ge + 1] > thr)) >= _MIN_STAFF_ROWS
        # 굵은 제목 숫자 행 군집이 오선으로 세어지지 않게 열 커버리지도 확인한다
        and (ink is None or _column_coverage(ink, left, right, gs, ge) >= _SYSTEM_MIN_COVERAGE)
    ]

    if len(accepted) <= 1:
        return [(t, b)]

    # 첫/마지막 시스템은 원래 세그먼트 경계(t/b)까지 확장한 뒤, 각 구간 안의 실제 잉크
    # 범위로 다시 트림한다(1차 분할의 seg_rows 트림과 동일한 원칙).
    bounds = (
        [t]
        + [_system_boundary(row_ink, accepted[i][1], accepted[i + 1][0])
           for i in range(len(accepted) - 1)]
        + [b]
    )

    result = []
    for i in range(len(accepted)):
        win_t = bounds[i] + (1 if i else 0)   # 경계 행은 윗 시스템 몫 — 인접 밴드가 한 행도 겹치지 않게
        win_b = bounds[i + 1]
        ink_rows = np.where(row_ink[win_t:win_b + 1] > 0)[0]
        if not len(ink_rows):
            continue
        result.append((win_t + int(ink_rows.min()), win_t + int(ink_rows.max())))
    return result


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


def _system_top(ink: np.ndarray, row_ink: np.ndarray, row_fill: np.ndarray,
                left: int, right: int, t: int, b: int, first: bool = True) -> int:
    """시스템 밴드 [t,b]의 crop 상단(오선 위 작곡가/제목 제외, 화음 줄은 포함).

    [t,b]를 잉크 없는 행 기준 블록으로 나눠 **오선을 담은 첫 블록**(열 커버리지 ≥
    _SYSTEM_MIN_COVERAGE, 높이 ≥ _SYSTEM_MIN_HEIGHT//5)을 찾고, 그 블록 **자신의** row_fill 최대값
    기준 임계(_staff_threshold)로 첫 오선행(staff_top)을 잡은 뒤 _trim_top_above_chords로 위쪽
    블록을 정리한다. 임계를 세그먼트 전체가 아니라 오선 블록 안에서만 계산하므로 굵은 제목 숫자
    행(122)이 오선으로 오인되지 않고, 오선에 화음이 붙은 블록 앞의 작곡가 블록(173·362·146)도
    예전처럼 잘려 나간다. 오선 블록을 못 찾으면 t.

    **첫 시스템에서만** 쓴다(`first`): 작곡가·제목은 곡 맨 위에만 있다. 둘째 이후 시스템은 시스템 경계
    (`_system_boundary`)나 세그먼트 첫 잉크 행이 이미 상단이므로, 여기서 위 블록을 또 버리면 반복 종지
    괄호(1./2. 볼타 선)가 잘려 나간다(독립 재리뷰 지적: 107·332·255 등 약 30곡)."""
    if not first:
        return t
    blocks = _nonblank_runs(row_ink[t:b + 1])
    for bs, be in blocks:
        if be - bs + 1 < _SYSTEM_MIN_HEIGHT // 5:
            continue
        if _column_coverage(ink, left, right, t + bs, t + be) < _SYSTEM_MIN_COVERAGE:
            continue
        top, bot = t + bs, t + be
        thr = _staff_threshold(row_fill, top, bot)
        rows = np.where(row_fill[top:bot + 1] > thr)[0]
        staff_top = top + int(rows.min()) if len(rows) else top
        return _trim_top_above_chords(row_ink, t, staff_top)
    return t


def _column_coverage(ink: np.ndarray, left: int, right: int, t: int, b: int) -> float:
    """[t,b] 행 범위에서 잉크가 끊김 없이(<=_COVERAGE_BRIDGE_PX 틈은 이어 붙임) 이어지는 가장
    긴 가로 구간 / 콘텐츠 폭. 오선은 기울거나 흐려도 전폭에 걸쳐 이어져 ~1.0, 텍스트 줄은 단어
    사이 공백 때문에 낮다."""
    col = ink[t:b + 1, left:right + 1].any(axis=0)
    best = 0
    run_start = last_end = None
    for s, e in _nonblank_runs(col.astype(int)):
        if run_start is not None and s - last_end - 1 <= _COVERAGE_BRIDGE_PX:
            last_end = e
        else:
            run_start = s
            last_end = e
        best = max(best, last_end - run_start + 1)
    return best / max(1, len(col))


def _page_segments(gray: np.ndarray) -> dict:
    """페이지(또는 곡 영역) 그레이 이미지를 공백 간격으로 세그먼트 분할하고 각 세그먼트가
    오선 시스템인지 판정한다. detect_system_bands와 저작권 폴백 탐지가 공유한다.

    반환: {'segments': [(t, b, is_system), ...], 'row_ink', 'row_fill', 'left', 'right', 'ink'}"""
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

    segments = []
    i = 0
    while i < h:
        if is_gap[i]:
            i += 1
            continue
        s = i
        while i < h and not is_gap[i]:
            i += 1
        seg_rows = np.where(row_ink[s:i] > 0)[0]
        if not len(seg_rows):
            continue
        t, b = s + int(seg_rows.min()), s + int(seg_rows.max())
        staff_rows = int(np.sum(row_fill[t:b + 1] > _STAFF_ROW_FILL))
        coverage = _column_coverage(ink, left, right, t, b)
        # 높이 경로(흐린 오선)는 오선 밀도 하한도 요구한다 — 가사만 여러 줄 쌓인 블록(커버리지는
        # 줄 수가 늘면 올라간다)이 시스템으로 오인되지 않게(독립 리뷰 지적, 2026-10-18).
        is_system = coverage >= _SYSTEM_MIN_COVERAGE and (
            staff_rows >= _MIN_STAFF_ROWS
            or ((b - t + 1) >= _SYSTEM_MIN_HEIGHT
                and float(row_fill[t:b + 1].max()) >= _SYSTEM_MIN_PEAK_FILL))
        segments.append((t, b, is_system))
    return {"segments": segments, "row_ink": row_ink, "row_fill": row_fill,
            "left": left, "right": right, "ink": ink}


def detect_system_bands(gray: np.ndarray, copyright_zone: bool = True) -> list:
    """오선보 시스템별 세로 밴드 [(top, bottom), …]를 반환(상단 제목/작곡가·하단 저작권 제외).

    1) 콘텐츠 좌우 경계로 트리밍한 폭 기준 행별 잉크 비율(row_fill)을 구한다.
    2) 공백 밴드(row_fill<0.005)가 시스템 간격(_SYSTEM_GAP_MIN 이상)이면 세그먼트 경계로 삼아
       콘텐츠를 세그먼트로 분할한다(그보다 짧은 공백은 시스템 내부로 병합).
    3) 열 연속 커버리지(_column_coverage)로 오선 세그먼트만 '시스템'으로 인정한다
       (_page_segments). 오선이 없는 세그먼트(상단 제목/작곡가 줄, 하단 저작권 줄)는 자연히
       제외된다 — 번호가 좌/우 어디 있든 동일(146 제목·번호 좌우 반대여도 동작).
    4) 시스템 바로 아래(≤_TRAILING_TEXT_MAX_GAP)의 비-시스템 세그먼트(2절 가사 등)는 앞
       시스템에 흡수한다. `copyright_zone`이면 페이지 하단 저작권 영역은 흡수하지 않는다 — 곡 영역이
       이미 저작권 위에서 잘려 있는 경우(`naju_song_bands`, 캐시 등록 곡)에는 False로 끈다. 잘린 영역의
       높이 기준 비율을 쓰면 마지막 가사줄이 '저작권 영역'으로 오인돼 빠진다(358·359·68·78, 35번).
    """
    h = gray.shape[0]
    info = _page_segments(gray)
    segs = info["segments"]
    row_fill, row_ink = info["row_fill"], info["row_ink"]

    systems = []          # [top, bottom] — 흡수 단계에서 bottom을 늘리므로 list
    owner = {}            # 세그먼트 인덱스 → systems 내 마지막 시스템 인덱스
    for si, (t, b, is_system) in enumerate(segs):
        if not is_system:
            continue
        # 공백행 연속 길이 기반 분할이 여러 시스템을 세그먼트 하나로 합쳤을 수 있으므로
        # (나주노 172), 오선행 밀도 군집으로 한 번 더 쪼갠다. 세그먼트가 이미 시스템 하나만
        # 담고 있으면 [(t, b)] 그대로 돌아와 기존 동작과 같다.
        for sub_t, sub_b in _split_merged_systems(row_fill, row_ink, t, b, info["ink"], info["left"], info["right"]):
            sub_t = _system_top(info["ink"], row_ink, row_fill, info["left"], info["right"], sub_t, sub_b,
                                first=not systems)
            systems.append([sub_t, sub_b])
            owner[si] = len(systems) - 1

    # 비-시스템 세그먼트를 바로 앞 시스템 세그먼트에 흡수(가사 보충줄). 저작권 영역 제외.
    for si, (t, b, is_system) in enumerate(segs):
        if is_system or (copyright_zone and b >= h * _COPYRIGHT_ZONE_RATIO):
            continue
        prev_si = si - 1
        if prev_si in owner:
            k = owner[prev_si]
            if 0 < t - systems[k][1] <= _TRAILING_TEXT_MAX_GAP and (b - t + 1) <= _TRAILING_TEXT_MAX_HEIGHT:
                systems[k][1] = b
                owner[si] = k   # 연속 보충줄도 같은 시스템에 이어 붙인다
    return [(t, b) for t, b in systems]


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
# ch/cw <= band_aspect*safety 식이므로 safety를 "낮추면" 상한이 낮아져 오히려 더 엄격해진다
# (3-pack을 더 쉽게 떨어뜨림) — 완화하려면 올려야 한다. 1.0을 넘기면 band_aspect 자체보다도
# 느슨해져(겹치지 않는 선에서는 꽉 채워도 허용) 172번(3-group 종횡비 0.449, band_aspect*1.0=
# 0.463)처럼 "근소하게 초과" 수준의 3-pack을 통과시킨다.
# 2026-10-04: 0.95→1.05로 완화(코디네이터 요청) — 172번이 [2,2,2,2,1](5장)이 아니라
# [3,3,3](3장)으로 묶이도록. 기존 회귀 고정곡 362/173/146/810/267은 safety 0.95~1.10 전
# 구간에서 그룹 크기 불변(합성곡 테스트 S7~S10도 동일) — 810형 "97% 빡빡 3-pack" 케이스는
# 종횡비 자체가 아니라 균형 타이브레이크([2,2]가 [3,1]과 슬라이드 수가 같으면 균형 우선)로
# 여전히 보호되므로 안전.
# 2026-10-18: 1.05→1.069 — 허용 한도 ch/cw를 0.486→약 0.495로 완화(사용자 지시, 나주노 603번 3줄 묶음).
_PACK_SAFETY = 1.069
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


def naju_page_is_shared(page_index: int) -> bool:
    """TOC상 이 페이지(0-based)에 곡이 둘 이상 실려 있나(예: 14·15번). 공유 페이지는 저작권이
    곡마다 페이지 중간/하단에 따로 있어 하단 기하 탐지로 곡을 구분할 수 없다."""
    n = 0
    with _open("나주노") as d:
        for _lvl, title, page in d.get_toc():
            if page - 1 == page_index and re.match(r"^\s*\d+\b", title):
                n += 1
    return n > 1


def _naju_copyright_fallback(page):
    """캐시 없이 나주노 페이지 하단의 저작권 첫 줄을 크롭한다. 없으면 None.

    _page_segments의 비-시스템 세그먼트 중 페이지 하단(_COPYRIGHT_ZONE_RATIO 아래)에 걸친
    마지막 것을 저작권 블록으로 보고, 그 안의 첫 잉크 행 런(=첫 줄)을 좌우 잉크 경계로 자른다."""
    if naju_page_is_shared(page.number):
        return None   # 공유 페이지: 어느 곡의 저작권인지 기하만으로 알 수 없다(수동 등록 필요)
    img = render_region(page, None, RENDER_DPI)
    gray = np.asarray(img.convert("L"))
    h = gray.shape[0]
    info = _page_segments(gray)
    cands = [(t, b) for t, b, is_sys in info["segments"]
             if not is_sys and b >= h * _COPYRIGHT_ZONE_RATIO]
    if not cands:
        return None
    t, b = cands[-1]
    runs = _nonblank_runs(info["row_ink"][t:b + 1])
    if not runs:
        return None
    y0, y1 = t + runs[0][0], t + runs[0][1]
    # 줄 사이에 빈 행이 없을 만큼 붙은 2줄 블록(479·692)은 행 잉크 골짜기에서 첫 줄만 남긴다.
    if y1 - y0 + 1 > _COPYRIGHT_LINE_MAX_H:
        prof = info["row_ink"][y0:y1 + 1].astype(float)
        lo, hi = int(len(prof) * 0.3), int(len(prof) * 0.7) + 1
        valley = lo + int(np.argmin(prof[lo:hi]))
        if prof[valley] <= 0.25 * prof.max():
            y1 = y0 + valley - 1
    cols = np.where((gray[y0:y1 + 1] < _INK).any(axis=0))[0]
    if not len(cols):
        return None
    return _trim_trailing_period_px(img.crop((max(0, int(cols.min()) - 3), max(0, y0 - 2),
                                              int(cols.max()) + 4, y1 + 3)))


def _naju_copyright_crop(number: int, page) -> Image.Image:
    entry = _load_naju_cache().get(str(number))
    if entry is None:
        # 캐시 미등록 곡(캐시는 개발자 OCR로 곡마다 채우는 방식이라 새 곡이 누락될 수 있다 —
        # 2026-10-18 479/519/692 저작권이 통째로 빠진 원인). 요소를 조용히 생략하는 대신
        # 하단 텍스트 세그먼트의 첫 줄을 기하학적으로 잘라 쓴다(OCR 없음, 트리거 이후 문구까지
        # 포함될 수 있으므로 경고를 남겨 캐시 등록을 유도).
        crop = _naju_copyright_fallback(page)
        logger.warning("[경고] 성가 %s: 저작권 캐시 미등록 — %s (tools/build_나주노_copyright_cache.py로 등록 권장)",
                       number, "하단 첫 줄 폴백 크롭 사용" if crop is not None else "하단 텍스트 없음, 생략")
        return crop
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


_MIN_REGION_H = 200   # 곡 영역 최소 높이(px@300dpi) — 시스템 1개(~200px)보다 작으면 잘못된 절단


def naju_song_region_image(page, number: int) -> Image.Image:
    """나주노 곡 하나의 악보 영역 이미지(300dpi). 페이지 전체가 아니라 **이 곡만** 담는다.

    - 캐시 항목에 `y_range`(페이지 대비 0~1 세로 구간)가 있으면 그 구간만 잘라 쓴다 — 한 페이지에
      곡이 2~3개 실린 공유 페이지(14·15번 등)에서 14번을 요청해도 15번 악보가 섞이지 않게 한다.
    - 캐시에 저작권 bbox가 있으면 그 상단 위까지만 쓴다 — 저작권 줄이 마지막 시스템의 가사 보충줄로
      오인돼 악보 크롭에 들어가는 일을 막는다(14번 등).
    - 공유 페이지인데 `y_range`가 없으면 다른 곡 악보가 섞이므로 조용히 진행하지 않고 예외를 낸다."""
    full = render_region(page, None, RENDER_DPI)
    entry = _load_naju_cache().get(str(number)) or {}
    W, H = full.size
    top, bottom = 0, H
    yr = entry.get("y_range")
    if yr:
        top, bottom = int(round(yr[0] * H)), int(round(yr[1] * H))
    elif naju_page_is_shared(page.number):
        raise RuntimeError(
            f"성가 {number}: 한 페이지에 곡이 여러 개인데 곡별 영역(y_range)이 캐시에 없습니다 — "
            "tools/build_나주노_copyright_cache.py --shared 로 등록해 주세요.")
    if entry.get("has_copyright") and entry.get("bbox"):
        cut = int(round(entry["bbox"][1] * H)) - 3
        if cut - top >= _MIN_REGION_H:   # bbox가 영역 밖/너무 위면 무시(1px 영역으로 무너지지 않게)
            bottom = min(bottom, cut)
    if (top, bottom) == (0, H):
        return full
    return full.crop((0, top, W, bottom))


def naju_song_bands(page, number: int):
    """(곡 영역 이미지, 시스템 밴드 목록). `build_hymn_pptx`와 테스트가 같은 경로를 쓴다.
    캐시에 이 곡 항목이 있으면(저작권 위치·곡 영역이 이미 확정) 저작권 영역 흡수 제외는 끈다."""
    region = naju_song_region_image(page, number)
    gray = np.asarray(region.convert("L"))
    known = str(number) in _load_naju_cache()
    return region, detect_system_bands(gray, copyright_zone=not known)


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
        if SOURCES[출처]["kind"] == "scan":
            region_img, bands = naju_song_bands(page, number)
            gray = np.asarray(region_img.convert("L"))
        else:
            region_img = render_region(page, find_song_y_range(page, 출처, number), RENDER_DPI)
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
