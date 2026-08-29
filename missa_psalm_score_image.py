# -*- coding: utf-8 -*-
"""화답송 악보 원본 이미지(PNG/JPG) → 슬라이드 변환 (1차 마일스톤).

원본 사진(오선보 1줄 + 좌측 "화답송" 워터마크)에서 워터마크를 크롭하고, 마디 경계
(바 라인)에서 시각적으로 균형 잡힌 지점을 골라 상/하 2줄로 나눈 뒤, 고정 템플릿 자산
(assets/화답송_악보_template.pptx)의 제목/라벨/저작권/잔재 도형을 보존한 채 악보 2장을
고정 밴드에 종횡비 유지·공통 배율로 앉힌 슬라이드 1장짜리 in-memory Presentation을 만든다.
가로 정렬은 상단 줄 왼쪽·하단 줄 오른쪽이다(사용자 지시).

설계서: _workspace/01_architect_design.md. 이 모듈은 순수 함수 묶음으로, missa_to_ppt.py를
import 하지 않는다(단방향 의존 — 설계서 §1.4). Pillow/numpy/python-pptx만 의존한다.

임계값(0.85·0.30/0.70·0.5·col≤10)은 절대 픽셀이 아니라 이미지 상대값이므로 스케일이 절반인
816(1954×201)에서도 그대로 동작한다(설계서 §2.4, 실측 검증).
"""
import copy
import io
import logging
import sys
from pathlib import Path

import numpy as np
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

logger = logging.getLogger(__name__)

# PyInstaller onefile 빌드에서는 datas로 묻은 파일이 실행 시점에 디스크로 추출되지 않는
# 경우가 있어(바이너리/데이터 재분류 단계의 알 수 없는 부작용, 실측 확인) config.json과
# 동일하게 "실행 파일 옆의 외부 자산 폴더"로 취급한다(missa_to_ppt.py의 _SCRIPT_DIR과 동일
# 패턴). 소스에서 직접 실행할 때는 이 파일의 위치를 그대로 쓴다.
_BASE = Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent
ASSET_TEMPLATE = _BASE / "assets" / "화답송_악보_template.pptx"

# --- 이미지 분석 임계값 (설계서 §2, §3) ---
_INK_THRESHOLD = 128          # gray < 128 = 잉크
_BARLINE_FILL = 0.85          # 오선 밴드에서 이 비율 이상 채운 컬럼 = 바 라인 후보
_STAFF_ROW_FILL = 0.5         # 폭의 이 비율 이상 잉크인 행 = 오선 밴드
_CENTER_LO, _CENTER_HI = 0.30, 0.70   # 중앙 탐색창 (clef/조표/끝 겹세로줄 배제)
_WATERMARK_SEARCH = 0.30      # 워터마크 분리 공백 띠를 찾는 좌측 폭 비율
_LEFT_MARGIN = 10             # 이 컬럼 이하에서 시작하는 zero-run = 좌측 여백(워터마크 아님)
_BARLINE_GAP = 3              # 바라인 후보 컬럼 그룹핑 허용 간격(px)

# 노트헤드/스템/플래그 인접 배제 필터(설계서 §2.4) 임계값. col_fill>0.85만으로는 밴드 전체를
# 관통하는 온음표/스템+플래그가 바라인으로 오검출된다(20260823 col=2304가 실사용자 확인 사례 —
# 8분음표 기둥+깃발). 판별 원리: 후보 바로 옆(offset 1~skip)에 다른 음표가 "가까이 있는 것"은
# 정상이지만(모든 바라인 옆엔 음표가 있다), 그 잉크가 skip을 지나서도(offset skip+1~
# skip+check_width) 계속 이어지면 후보 자체가 그 음표(스템/플래그/부점머리)에 융합된 획이다.
# 3세트(809/816/823) 전 후보에 대해 실측 검증 완료 — 이 임계값으로 확인된 페이크(628 박자표,
# 2304/2920 등 스템+플래그, 999/1234/1865 부점머리 인접)와 진짜(913/1408/1971/2647/3084/
# 3674/3875, 809의 2146, 816의 1201)가 정확히 분리된다.
_ADJACENT_SKIP_MIN = 3         # 오선 밴드 높이 비례 최소 스킵 폭(안티에일리어싱/정상 인접 허용)
_ADJACENT_SKIP_RATIO = 0.02    # 스킵 폭 = max(_ADJACENT_SKIP_MIN, round(band_h*비율))
_ADJACENT_INK_THRESHOLD = 0.05  # 스킵 이후 구간에서 이 이상 잉크가 지속되면 융합된 것으로 간주

# --- 슬라이드 배치 밴드 (EMU, 설계서 §5.3) ---
SLIDE_W, SLIDE_H = 12192000, 6858000
_BAND_LEFT = 400000
_BAND_W = 11392000
_BAND_H = 2500000
_TOP_BAND_TOP = 1000000
_LOW_BAND_TOP = 3650000

# 악보 그림으로 간주할 최소 크기(EMU). 잔재 도형(12×7·41×17)은 이 아래라 자동 배제.
_SCORE_MIN_EMU = 1_000_000


# ---------------------------------------------------------------------------
# 이미지 로드 (RGBA 알파 평탄화 — 설계서 §0.1)
# ---------------------------------------------------------------------------

def load_flattened_image(image_path) -> Image.Image:
    """이미지를 열어 RGB로 평탄화한다.

    RGBA/LA/P 모드는 투명 픽셀이 그레이스케일 변환 시 검정(0=잉크)으로 오인될 수 있으므로
    반드시 흰 배경 위에 합성한 뒤 반환한다(설계서 §0.1, test_07).
    """
    im = Image.open(image_path)
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
        im = Image.alpha_composite(bg, im)
    return im.convert("RGB")


def load_gray_array(image_path) -> np.ndarray:
    """이미지를 평탄화 후 그레이스케일 2D ndarray(L, uint8)로 반환한다."""
    return np.asarray(load_flattened_image(image_path).convert("L"))


# ---------------------------------------------------------------------------
# 워터마크 크롭 (설계서 §3)
# ---------------------------------------------------------------------------

def detect_watermark_crop_x(gray: np.ndarray) -> int:
    """워터마크("화답송")와 오선보 사이의 분리 공백 띠 우측 끝+1(=크롭 시작 col)을 반환.

    좌측 여백(col≤10에서 시작하는 zero-run)은 워터마크가 아니라 촬영 여백이므로 통째로
    제외한다 — 좌측 여백을 col 11부터로 잘라 쓰면 실제 분리 띠와 폭이 같아져(816에서 둘 다
    37px) 오검출된다(설계서 §3.2 실측 근거).

    좌측 30% 안에 유효한 zero-run이 없으면 미크롭(0) + 경고 — 음표 손실보다 워터마크
    잔존이 안전하다(설계서 §3.3 폴백).
    """
    ink = gray < _INK_THRESHOLD
    h, w = gray.shape
    col_sum = ink.sum(axis=0)
    zero = col_sum == 0
    search_limit = int(w * _WATERMARK_SEARCH)

    best = None  # (right_end, width)
    c = 0
    while c < w:
        if zero[c]:
            s = c
            while c < w and zero[c]:
                c += 1
            e = c - 1
            # 좌측 여백(col≤10에서 시작)은 통째로 제외, 좌측 30% 안의 zero-run만 후보
            if s > _LEFT_MARGIN and s < search_limit:
                width = e - s + 1
                if best is None or width > best[1]:
                    best = (e, width)
        else:
            c += 1

    if best is None:
        logger.warning("워터마크 분리 공백 띠 미검출 — 미크롭(crop_left=0)으로 진행")
        return 0
    return best[0] + 1


# ---------------------------------------------------------------------------
# 오선 밴드 / 콘텐츠 경계 (설계서 §2.1, §7-3)
# ---------------------------------------------------------------------------

def staff_band(gray: np.ndarray) -> tuple:
    """오선(5선) 세로 밴드 (top, bottom) — 가로 잉크가 폭의 50% 이상인 행들의 min~max.

    오선이 흐리거나 기울어 밴드를 못 찾으면 이미지 높이의 30%/70%로 폴백(설계서 §2.4).
    """
    ink = gray < _INK_THRESHOLD
    h, w = gray.shape
    row_fill = ink.sum(axis=1) / w
    rows = np.where(row_fill > _STAFF_ROW_FILL)[0]
    if len(rows):
        return int(rows.min()), int(rows.max())
    return int(h * 0.30), int(h * 0.70)


def detect_content_bounds(gray: np.ndarray) -> tuple:
    """오선보 잉크의 bounding box (left, right, top, bottom)를 반환.

    right는 트레일링 공백을 제거한 오선보 우측끝(설계서 §7-3), top/bottom은 상/하 여백을
    제거한 세로 크롭 범위. left는 워터마크를 포함할 수 있으므로 실제 좌측 크롭은
    detect_watermark_crop_x로 덮어쓴다(설계서 §1.2 각 함수 책임 분리).
    """
    ink = gray < _INK_THRESHOLD
    cols = np.where(ink.any(axis=0))[0]
    rows = np.where(ink.any(axis=1))[0]
    if not len(cols) or not len(rows):
        h, w = gray.shape
        return 0, w - 1, 0, h - 1
    return int(cols.min()), int(cols.max()), int(rows.min()), int(rows.max())


# ---------------------------------------------------------------------------
# 바라인 탐지 / 균형 분할 (설계서 §2.2)
# ---------------------------------------------------------------------------

def _group_consecutive(indices) -> list:
    """정렬된 정수 배열을 간격 1(연속) 기준으로 (start, end) 구간 리스트로 그룹핑."""
    groups = []
    if len(indices) == 0:
        return groups
    s = prev = int(indices[0])
    for v in indices[1:]:
        v = int(v)
        if v - prev <= 1:
            prev = v
        else:
            groups.append((s, prev))
            s = prev = v
    groups.append((s, prev))
    return groups


def _is_fused_with_adjacent_note(fill_no_lines: np.ndarray, gs: int, ge: int, w: int, band_h: int) -> bool:
    """후보(gs..ge)가 인접 음표(노트헤드/스템/플래그)에 융합된 획인지 판정.

    fill_no_lines: 오선 라인 자체를 제외한, 밴드 내 컬럼별 잉크 비율(행간 공백 기준).
    바라인 바로 옆에 다른 음표가 있는 것 자체는 정상(모든 마디 경계 옆엔 음표가 있다).
    구분 신호는 "그 잉크가 스킵 구간을 지나서도 지속되는가"다 — 진짜 바라인은 옆 음표와의
    접촉이 안티에일리어싱 수준(offset 1~skip)에서 그치고 skip 이후엔 0으로 떨어지지만, 후보
    자체가 스템/플래그/부점머리의 일부라면 그 폭 전체(offset skip+1~skip+check_width)에 걸쳐
    잉크가 이어진다(설계서 §2.4, 실사용자 확인 823:col=2304 사례로 실측 검증).
    """
    skip = max(_ADJACENT_SKIP_MIN, round(band_h * _ADJACENT_SKIP_RATIO))
    check_width = max(_ADJACENT_SKIP_MIN, round(band_h * _ADJACENT_SKIP_RATIO))
    far_max = 0.0
    for off in range(skip + 1, skip + check_width + 1):
        if gs - off >= 0:
            far_max = max(far_max, fill_no_lines[gs - off])
        if ge + off < w:
            far_max = max(far_max, fill_no_lines[ge + off])
    return far_max > _ADJACENT_INK_THRESHOLD


def detect_barlines(gray: np.ndarray, band: tuple) -> list:
    """오선 밴드 안에서 col_fill > 0.85인 컬럼을 연속(간격≤3px) 그룹핑한 중심 리스트.

    바 라인 = 오선 전체 높이를 세로로 관통하는 얇은 검은 수직선이라 밴드 채움 비율이 높다.
    노트 스템은 밴드 일부만 채워 대개 0.85 미만이라 배제되지만, 온음표·8분음표 기둥+플래그처럼
    밴드 전체를 관통하는 획은 이 기준만으로 걸러지지 않는다 — 노트헤드/스템 인접 배제 필터
    (`_is_fused_with_adjacent_note`)로 2차 배제한다(설계서 §2.4).
    """
    top, bottom = band
    ink = gray < _INK_THRESHOLD
    band_ink = ink[top:bottom + 1, :]
    band_h = max(1, bottom - top + 1)
    col_fill = band_ink.sum(axis=0) / band_h
    cols = np.where(col_fill > _BARLINE_FILL)[0]
    if not len(cols):
        return []
    raw_groups = []
    s = prev = cols[0]
    for c in cols[1:]:
        if c - prev <= _BARLINE_GAP:
            prev = c
        else:
            raw_groups.append((int(s), int(prev)))
            s = prev = c
    raw_groups.append((int(s), int(prev)))

    # 오선 라인 자체를 제외한 "행간 공백" 잉크 비율 — 이 라인들은 후보 좌우 어디서든 항상
    # 가득 차 있어(수평으로 전체 폭 관통) 인접 판정 신호를 오염시키므로 반드시 제외해야 한다.
    w = gray.shape[1]
    row_fill_full = ink.sum(axis=1) / w
    staff_rows = np.where(row_fill_full > _STAFF_ROW_FILL)[0]
    line_groups = _group_consecutive(staff_rows)
    ink_no_lines = band_ink.copy()
    for ls, le in line_groups:
        rel_s = max(0, ls - top)
        rel_e = min(band_h - 1, le - top)
        if rel_s <= rel_e:
            ink_no_lines[rel_s:rel_e + 1, :] = False
    fill_no_lines = ink_no_lines.sum(axis=0) / band_h

    result = []
    for gs, ge in raw_groups:
        if not _is_fused_with_adjacent_note(fill_no_lines, gs, ge, w, band_h):
            result.append((gs + ge) // 2)
    return result


def choose_split_x(barlines: list, left: int, right: int) -> int:
    """중앙창 [0.30,0.70] 안의 바라인 중 좌우 폭 차이를 최소화하는 분할 col을 반환.

    폴백(설계서 §2.4): 중앙창에 바라인이 없으면 기하 중앙 (left+right)//2.
    """
    content_w = right - left
    lo = left + _CENTER_LO * content_w
    hi = left + _CENTER_HI * content_w
    candidates = [b for b in barlines if lo <= b <= hi]
    if not candidates:
        logger.warning("중앙창 바라인 후보 0개 — 기하 중앙으로 분할")
        return (left + right) // 2
    return min(candidates, key=lambda x: abs((x - left) - (right - x)))


def split_and_place(image: Image.Image, crop_box: tuple, split_x: int) -> tuple:
    """crop_box=(left, top, right, bottom)를 split_x에서 좌/우로 나눠 상/하 2장으로 반환.

    상단폭+하단폭 == right-left (겹침·누락 없는 무손실 분할, 설계서 §7-6). PIL crop의 right/
    bottom은 배타적이므로 호출자가 bottom에 +1을 넘긴다.
    """
    left, top, right, bottom = crop_box
    upper = image.crop((left, top, split_x, bottom))
    lower = image.crop((split_x, top, right, bottom))
    return upper, lower


# ---------------------------------------------------------------------------
# 슬라이드 조립 (설계서 §5.3, §8)
# ---------------------------------------------------------------------------

def _required_scale(img_w: int, img_h: int) -> float:
    """이미지를 밴드(폭·높이 둘 다)에 맞추는 데 필요한 축소 배율(종횡비 유지)."""
    return min(_BAND_W / img_w, _BAND_H / img_h)


def _place_with_scale(img_w: int, img_h: int, scale: float, band_top: int, h_align: str = "center") -> tuple:
    """주어진 배율로 이미지를 배치 → (left, top, w, h) EMU, 세로는 밴드 안 중앙정렬.

    가로 정렬은 `h_align`으로 지정한다 — 상단 줄은 "left", 하단 줄은 "right"(사용자 지시,
    §추가요구사항 2라운드: "첫 번째 단은 왼쪽 정렬, 두 번째 단은 오른쪽 정렬"). scale이 그
    이미지 자신의 `_required_scale`보다 작으면(다른 줄이 더 좁아 공통 배율이 낮게 잡힌 경우)
    밴드를 다 채우지 않고 남는 여백이 h_align 방향의 반대쪽에 생긴다 — 왜곡 없이 자연스럽게
    작아 보인다.
    """
    w_emu = int(round(img_w * scale))
    h_emu = int(round(img_h * scale))
    if h_align == "left":
        left = _BAND_LEFT
    elif h_align == "right":
        left = _BAND_LEFT + (_BAND_W - w_emu)
    else:
        left = _BAND_LEFT + (_BAND_W - w_emu) // 2
    top = band_top + (_BAND_H - h_emu) // 2
    return left, top, w_emu, h_emu


def _insert_run_before_end_para_rpr(p, new_r):
    """새 run을 <a:p>에 삽입한다. endParaRPr가 있으면 그 앞, 없으면 맨 끝.

    CT_TextParagraph 스키마 순서는 `pPr?, (run|br)*, endParaRPr?`다. run/br을 지우고
    endParaRPr는 남긴 단락에 새 run을 그냥 append하면 `pPr, endParaRPr, r` 순서가 되어
    스키마 위반(run이 endParaRPr 뒤에 옴)이 된다 — CLAUDE.md의 "pPr append 금지" 함정이
    `<a:p>` 자식 레벨(run vs endParaRPr)에서 재발한 사례(리뷰 지적, 02b_review_report.md).
    """
    eprp = p.find(qn("a:endParaRPr"))
    if eprp is not None:
        eprp.addprevious(new_r)
    else:
        p.append(new_r)


def _replace_title_paragraph(para, new_text: str):
    """단락의 run을 첫 run 서식으로 복제해 new_text로 교체 (서식 보존).

    missa_to_ppt._replace_para_text_clone과 동일한 안전 패턴을 이 모듈에 복제한다 —
    설계서 §1.4의 단방향 의존 규칙상 missa_to_ppt.py를 import할 수 없기 때문이다. run 인덱스를
    고정하거나 개수를 가정하지 않고, 첫 run의 rPr을 deepcopy해 텍스트만 바꾸므로 여러 run으로
    쪼개진 제목("화 답 송   시편 ", "67", "(66", ") ")에서도 색상/폰트가 보존된다(CLAUDE.md
    "run 개수 2개 가정 금지" 함정 회피). 새 run은 endParaRPr보다 앞에 삽입한다(위 함수 참고).
    """
    p = para._p
    runs = p.findall(qn("a:r"))
    template_r = runs[0] if runs else None
    for r in runs:
        p.remove(r)
    for br in p.findall(qn("a:br")):
        p.remove(br)
    if template_r is not None:
        new_r = copy.deepcopy(template_r)
        t_el = new_r.find(qn("a:t"))
        if t_el is not None:
            t_el.text = new_text
        _insert_run_before_end_para_rpr(p, new_r)
    else:
        from pptx.oxml import parse_xml as pptx_parse_xml
        A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
        new_r = pptx_parse_xml(
            f'<a:r xmlns:a="{A_NS}"><a:t>{new_text}</a:t></a:r>'
        )
        _insert_run_before_end_para_rpr(p, new_r)


def _update_title(slide, title: str):
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        t = shape.text_frame.text
        if "화 답 송" in t:
            _replace_title_paragraph(shape.text_frame.paragraphs[0], f"화 답 송   {title}")
            return


def _add_picture(slide, pil_img: Image.Image, scale: float, band_top: int, h_align: str = "center"):
    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    buf.seek(0)
    left, top, w, h = _place_with_scale(pil_img.width, pil_img.height, scale, band_top, h_align)
    slide.shapes.add_picture(buf, left, top, w, h)


def build_slide(template_path, title: str, up_img: Image.Image, low_img: Image.Image) -> Presentation:
    """고정 템플릿을 로드해 제목을 교체하고 상/하 악보 2장을 밴드에 배치한 Presentation 반환.

    상/하 두 이미지에 **동일한 배율**을 적용한다 — 각자 독립적으로 자기 밴드에 맞춰 스케일링하면
    마디선 제약 때문에 두 단의 폭이 크게 다른 악보에서 폭이 좁은 쪽이 상대적으로 확대되어 두
    줄의 음표 크기가 달라 보인다(사용자 지적, 02c). 공통 배율 = 두 이미지 각각의 필요 배율
    (`_required_scale`, 밴드 폭·높이 둘 다에 맞추는 데 필요한 축소율) 중 더 작은(더 제약이 큰)
    값. 이러면 음표 크기(배율)는 항상 같고, 폭이 좁은 쪽은 밴드를 다 채우지 않은 채 남은
    여백만큼만 비고, 두 줄 다 필요하면 함께 작아진다.

    가로 정렬은 상단 줄 왼쪽 정렬·하단 줄 오른쪽 정렬이다(사용자 지시, §추가요구사항 2라운드).

    템플릿 배경은 이미 정답이므로 _set_slide_bg_black류 배경 재설정을 하지 않는다(CLAUDE.md
    "슬라이드 복사 시 배경 재설정 금지"). 템플릿을 직접 로드하는 경로라 배경/레이아웃/마스터가
    그대로 유지된다.
    """
    prs = Presentation(str(template_path))
    slide = prs.slides[0]
    _update_title(slide, title)
    scale = min(
        _required_scale(up_img.width, up_img.height),
        _required_scale(low_img.width, low_img.height),
    )
    _add_picture(slide, up_img, scale, _TOP_BAND_TOP, h_align="left")
    _add_picture(slide, low_img, scale, _LOW_BAND_TOP, h_align="right")
    return prs


# ---------------------------------------------------------------------------
# 진입점 (설계서 §1.1)
# ---------------------------------------------------------------------------

def render_화답송_score_slide(image_path, title: str, template_path=ASSET_TEMPLATE) -> Presentation:
    """화답송 악보 원본 이미지 → 슬라이드 1장짜리 in-memory Presentation.

    부작용 없음(파일 저장은 호출자 몫). 출력 슬라이드 크기를 템플릿과 동일한
    12192000×6858000으로 유지해 이후 copy_slide_from_prs 스케일이 1:1이 되게 한다(설계서 §1.1).
    """
    gray = load_gray_array(image_path)
    image = load_flattened_image(image_path)

    crop_left = detect_watermark_crop_x(gray)
    _, crop_right, top, bottom = detect_content_bounds(gray)
    band = staff_band(gray)
    barlines = detect_barlines(gray, band)
    split_x = choose_split_x(barlines, crop_left, crop_right)

    up_img, low_img = split_and_place(image, (crop_left, top, crop_right, bottom + 1), split_x)
    return build_slide(template_path, title, up_img, low_img)
