# -*- coding: utf-8 -*-
"""프로그레션 테스트 — 아직 안정화되지 않은 신규 기능을 구현 전에 명세한다.

현재 진행 중: 화답송 악보 이미지(PNG/JPG) 입력 지원.
- 1차 마일스톤(missa_psalm_score_image.py — 이미지 → 슬라이드 변환): 완료·사용자 승인.
- 2차 마일스톤(missa_to_ppt.py 배선 — find_files/CLI 오버라이드/입력창/update_화답송):
  이번 라운드. 설계서 §4(통합 지점) 그대로 구현.

설계서: _workspace/01_architect_design.md §4(통합 지점), §7(1차 마일스톤 16개 테스트 대상 행동).
기대값은 3개 예제 세트(20260809/20260816/20260823) 실측 정답이다(추측 아님).

안정화되면 regression-qa가 test_missa_regression.py로 승격하고 이 파일에서 제거한다.
"""
import io
from pathlib import Path

import numpy as np
import pytest
from PIL import Image

import missa_psalm_score_image as ps

BASE = Path(__file__).resolve().parent
REF = BASE / "reference" / "화답송악보"
TEMPLATE = BASE / "assets" / "화답송_악보_template.pptx"

# 세트별 원본 이미지 경로와 실측 기대값 (설계서 §0.1, §2.3, §7)
CASES = {
    "20260809": {
        "img": REF / "20260809" / "20260809_043842947.jpg",
        "crop_left": 372,
        "crop_right": 3906,
        "split": 2146,
        "top_ratio": 0.502,
    },
    "20260816": {
        "img": REF / "20260816" / "20260816_043822155.png",
        "crop_left": 223,
        "crop_right": 1953,
        "split": 1201,
        "top_ratio": 0.565,
    },
    "20260823": {
        "img": REF / "20260823" / "20260823_043736869.png",
        "crop_left": 372,
        "crop_right": 3906,
        # split=2304(top_ratio 0.547)는 이전 라운드의 오검출 버그가 낳은 값이었다 — 실제로는
        # 8분음표 기둥+깃발(스템+플래그)이었고 진짜 마디선이 아니었다(사용자가 실제 생성물을
        # 육안 확대 검토해 발견, _workspace/02c 라운드). detect_barlines()의 노트헤드/스템 인접
        # 배제 필터 추가 후 실측 재검증한 참값은 split=1971(진짜 마디선, 확대 이미지로 확인:
        # 좌우가 깨끗한 공백인 순수 수직선). "값이 그럴듯하게 보인다"만으로 설계서 수치를
        # 맹신하지 않고, 매 라운드 실제 이미지 확대 검토로 재검증해야 한다는 교훈.
        "split": 1971,
        "top_ratio": 0.452,
    },
}


@pytest.fixture(scope="module", params=list(CASES))
def case(request):
    c = CASES[request.param]
    c = dict(c, name=request.param)
    return c


@pytest.fixture(scope="module")
def gray(case):
    return ps.load_gray_array(case["img"])


# ---------------------------------------------------------------------------
# 이미지 처리 (PowerPoint 불필요, 순수 함수 — 빠른 회귀 후보)
# ---------------------------------------------------------------------------

def test_01_watermark_crop_start(case, gray):
    """워터마크 크롭 시작 컬럼이 분리 공백 띠 우측에 온다 (±5px)."""
    x = ps.detect_watermark_crop_x(gray)
    assert abs(x - case["crop_left"]) <= 5, (case["name"], x)


def test_02_watermark_excluded_from_crop(case, gray):
    """워터마크가 크롭 결과에서 제외된다.

    (밀도 비교는 부적합 — 실측상 오선보 시작부 clef/조표가 워터마크 글자보다 잉크가
    더 촘촘하다.) 대신 두 사실로 강하게 단언한다: (1) 제거되는 좌측 영역에 실제
    워터마크 잉크가 존재하고(제거할 대상이 있음), (2) 크롭 경계가 분리 공백 띠 위에
    정확히 놓여 크롭 직전 컬럼이 완전 백색이다 → 워터마크 컬럼은 전부 크롭 좌측에 남아
    결과에서 빠진다.
    """
    crop_x = ps.detect_watermark_crop_x(gray)
    ink = gray < 128
    assert crop_x > 0, case["name"]
    assert ink[:, 10:crop_x].sum() > 0, ("제거 영역에 워터마크 잉크가 있어야 함", case["name"])
    assert ink[:, crop_x - 1].sum() == 0, ("크롭 경계가 공백 띠 위에 놓여야 함", case["name"])


def test_03_content_right_edge(case, gray):
    """오선보 유효 우측끝 검출 — 트레일링 공백 제거 (±5px)."""
    left, right, top, bottom = ps.detect_content_bounds(gray)
    assert abs(right - case["crop_right"]) <= 5, (case["name"], right)


def test_04_split_on_barline(case, gray):
    """분할점이 마디 경계(바라인) 위에 온다 — 중앙창 안 + 탐지 바라인 집합의 원소 (±20px)."""
    crop_left = ps.detect_watermark_crop_x(gray)
    left, right, top, bottom = ps.detect_content_bounds(gray)
    band = ps.staff_band(gray)
    barlines = ps.detect_barlines(gray, band)
    split = ps.choose_split_x(barlines, crop_left, right)
    assert abs(split - case["split"]) <= 20, (case["name"], split)
    # 중앙창 [0.30, 0.70] 안
    content_w = right - crop_left
    assert crop_left + 0.30 * content_w <= split <= crop_left + 0.70 * content_w
    # 폴백이 아닌 한 실제 탐지 바라인 중 하나여야 한다
    assert any(abs(split - b) <= 20 for b in barlines), (case["name"], barlines)


def test_05_balanced_two_lines(case, gray):
    """두 줄 폭 비율이 균형: |상단폭-하단폭|/content_w ≤ 0.15, 상단비율 실측 일치(±0.03)."""
    crop_left = ps.detect_watermark_crop_x(gray)
    left, right, top, bottom = ps.detect_content_bounds(gray)
    band = ps.staff_band(gray)
    split = ps.choose_split_x(ps.detect_barlines(gray, band), crop_left, right)
    content_w = right - crop_left
    up = split - crop_left
    low = right - split
    assert abs(up - low) / content_w <= 0.15, (case["name"], up, low)
    assert abs(up / content_w - case["top_ratio"]) <= 0.03, (case["name"], up / content_w)


def test_06_lossless_split(case, gray):
    """원본 픽셀 무손실 분할: 상단폭 + 하단폭 == crop_right - crop_left (겹침·누락 없음)."""
    crop_left = ps.detect_watermark_crop_x(gray)
    left, right, top, bottom = ps.detect_content_bounds(gray)
    band = ps.staff_band(gray)
    split = ps.choose_split_x(ps.detect_barlines(gray, band), crop_left, right)
    img = ps.load_flattened_image(CASES[case["name"]]["img"])
    up_img, low_img = ps.split_and_place(img, (crop_left, top, right, bottom + 1), split)
    assert up_img.width + low_img.width == right - crop_left, (case["name"],)
    assert up_img.height == low_img.height  # 같은 높이로 잘림


def test_07_rgba_flatten_safe():
    """RGBA 투명 입력이 검정 잉크로 오인되지 않는다 (알파 평탄화 검증).

    완전 투명 RGBA 이미지는 흰 배경 합성 후 그레이스케일에서 잉크(어두운 픽셀)가
    전혀 없어야 한다. 평탄화를 빠뜨리면 투명 픽셀이 0(검정)으로 변해 전부 잉크가 된다.
    """
    transparent = Image.new("RGBA", (100, 40), (0, 0, 0, 0))
    buf = io.BytesIO()
    transparent.save(buf, format="PNG")
    buf.seek(0)
    gray = ps.load_gray_array(buf)
    ink = gray < 128
    assert ink.mean() == 0.0, "투명 픽셀이 잉크로 오인됨 — 평탄화 누락"


def test_08_fallback_no_watermark_gap():
    """폴백: 워터마크 미검출 시 미크롭(crop_left==0) + 경고, 예외 없음.

    좌측 30%에 넓은 zero-run이 없는 합성 입력. 음표 손실보다 워터마크 잔존이 안전.
    """
    # 좌측 30% 안에 완전 백색 컬럼(zero-run)이 하나도 없도록 전 폭에 가로 잉크 바를 둔다
    arr = np.full((40, 300), 255, dtype=np.uint8)
    arr[10:30, :] = 0  # 모든 컬럼에 잉크 → 분리 공백 띠 없음
    x = ps.detect_watermark_crop_x(arr)
    assert x == 0


def test_09_fallback_no_barline_geometric_center():
    """폴백: 중앙창 바라인 0개 → 예외 없이 기하 중앙 근처 분할점 반환."""
    # 바라인(고밀도 세로선)이 전혀 없는 합성 이미지
    arr = np.full((40, 300), 255, dtype=np.uint8)
    arr[18:22, 20:280] = 0  # 얇은 가로 오선만
    band = ps.staff_band(arr)
    barlines = ps.detect_barlines(arr, band)
    split = ps.choose_split_x(barlines, 20, 280)
    assert 20 < split < 280
    assert abs(split - (20 + 280) // 2) <= 30  # 기하 중앙 근처


def test_09b_barline_rejects_notehead_fused_stem():
    """detect_barlines가 노트헤드/스템에 융합된 획을 마디선 후보에서 제외한다.

    실사용자가 20260823 미리보기에서 col=2304를 확대해 8분음표 기둥+깃발임을 확인한 사례의
    재발 방지 테스트(설계서 §2.4, 02c 라운드). 합성 이미지로 두 후보를 만든다:
    - col=100: 오선 전체를 관통하는 순수 수직선(좌우 공백) — 진짜 마디선.
    - col=150: 오선 전체를 관통하는 수직선 + 한쪽 행간에 노트헤드 모양 블록이 융합됨 — 가짜.
    둘 다 col_fill>0.85(밴드 전체 관통)라 1차 임계값만으로는 구분되지 않으므로, 노트헤드 인접
    배제 필터가 실제로 두 후보를 갈라내는지 확인한다.
    """
    w, h = 220, 60
    img = np.full((h, w), 255, dtype=np.uint8)
    for r in (10, 20, 30, 40, 50):
        img[r, :] = 0  # 오선 5줄, 폭 전체 관통
    img[10:51, 100:102] = 0  # 진짜 마디선: 순수 수직선, 좌우 공백
    img[10:51, 150:152] = 0  # 가짜 후보의 기둥(스템)
    img[24:37, 143:161] = 0  # 가짜 후보에 융합된 노트헤드(한쪽 행간에 폭넓게 번짐)

    band = ps.staff_band(img)
    barlines = ps.detect_barlines(img, band)
    assert any(abs(b - 100) <= 2 for b in barlines), ("진짜 마디선이 후보에서 빠짐", barlines)
    assert not any(abs(b - 150) <= 2 for b in barlines), (
        "노트헤드에 융합된 스템이 마디선으로 오검출됨", barlines
    )


# ---------------------------------------------------------------------------
# 슬라이드 생성 (템플릿 보존)
# ---------------------------------------------------------------------------

@pytest.fixture(scope="module")
def rendered(case):
    return ps.render_화답송_score_slide(CASES[case["name"]]["img"], "시편 00(00)", TEMPLATE)


def _iter_shapes(slide):
    return list(slide.shapes)


def test_10_exactly_one_slide(rendered):
    """출력 Presentation은 슬라이드 정확히 1장."""
    assert len(rendered.slides) == 1


def test_11_slide_size_preserved(rendered):
    """슬라이드 크기 12192000×6858000 유지 (copy_slide_from_prs 스케일 1:1 보장)."""
    assert rendered.slide_width == 12192000
    assert rendered.slide_height == 6858000


def test_12_template_text_preserved(rendered):
    """템플릿 라벨/저작권 텍스트박스가 그대로 보존된다."""
    texts = [
        sh.text_frame.text
        for sh in _iter_shapes(rendered.slides[0])
        if sh.has_text_frame
    ]
    joined = "\n".join(texts)
    assert "Responsorial" in joined
    assert "한국천주교중앙협의회" in joined
    assert "2022" in joined


def test_13_title_replaced(rendered):
    """제목박스 텍스트가 '화 답 송   {title}'로 교체, 다른 텍스트박스는 불변."""
    slide = rendered.slides[0]
    title_boxes = [
        sh for sh in _iter_shapes(slide)
        if sh.has_text_frame and "화 답 송" in sh.text_frame.text
    ]
    assert len(title_boxes) == 1
    assert title_boxes[0].text_frame.text.strip() == "화 답 송   시편 00(00)"
    # 저작권 박스는 제목 교체의 영향을 받지 않는다
    copyright_boxes = [
        sh for sh in _iter_shapes(slide)
        if sh.has_text_frame and "한국천주교중앙협의회" in sh.text_frame.text
    ]
    assert len(copyright_boxes) == 1


def test_13b_title_paragraph_child_order_schema_valid(rendered):
    """제목 교체 후 <a:p> 자식 순서가 CT_TextParagraph 스키마를 지킨다.

    스키마 시퀀스: pPr?, (run|br)*, endParaRPr?. run이 endParaRPr보다 뒤에 오면
    PowerPoint가 "손상된 파일"로 취급할 수 있는 스키마 위반이다. validate_pptx_structure()는
    XML 제어문자·끊어진 rId만 검사하고 요소 순서를 검사하지 않으므로(missa_to_ppt.py:6071),
    이 순서는 lxml로 직접 확인해야 한다(02b_review_report.md 리뷰 지적 — CLAUDE.md의
    "pPr append 금지" 함정이 <a:p> 자식 레벨(run vs endParaRPr)에서 재발했었음).
    """
    from pptx.oxml.ns import qn

    slide = rendered.slides[0]
    title_boxes = [
        sh for sh in _iter_shapes(slide)
        if sh.has_text_frame and "화 답 송" in sh.text_frame.text
    ]
    assert len(title_boxes) == 1
    p = title_boxes[0].text_frame.paragraphs[0]._p
    tags = [child.tag for child in p]
    eprp_tag = qn("a:endParaRPr")
    r_tag = qn("a:r")
    assert r_tag in tags, "제목 run이 존재해야 함"
    if eprp_tag in tags:
        r_indices = [i for i, t in enumerate(tags) if t == r_tag]
        eprp_index = tags.index(eprp_tag)
        assert max(r_indices) < eprp_index, (
            "run이 endParaRPr보다 앞에 와야 한다(스키마 위반)", tags
        )


def test_14_two_score_pictures_placed(rendered):
    """악보 이미지 PICTURE 2장 추가 (width>1M), 각각 상/하 밴드 안, 종횡비 유지."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    slide = rendered.slides[0]
    big_pics = [
        sh for sh in _iter_shapes(slide)
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE and sh.width > 1_000_000
    ]
    assert len(big_pics) == 2, [sh.width for sh in big_pics]
    tops = sorted(sh.top for sh in big_pics)
    # 상단 밴드 top≈1.0M, 하단 밴드 top≈3.65M (밴드 내 세로 중앙정렬이라 약간 아래로 이동)
    assert 1_000_000 <= tops[0] < 3_000_000, tops
    assert 3_000_000 <= tops[1] < 6_000_000, tops
    # 밴드 폭(11392000, left여백 400000) 안에 있고 왜곡 없이 배치
    for sh in big_pics:
        assert sh.left >= 400_000 - 1
        assert sh.left + sh.width <= 400_000 + 11_392_000 + 1
        assert sh.height <= 2_500_000 + 1


def test_14b_uniform_scale_across_both_lines():
    """상/하 두 단의 폭이 크게 다르면 공통(더 작은) 배율을 적용해 음표 크기를 통일한다.

    사용자 요구(02c 라운드): "일부 악보는 첫째 단과 둘째 단의 폭이 다를 수 있다 — 그 경우
    전체적인 비율을 조정해서... 첫째 단과 둘째 단의 높이가 더 작아지는 거지." 각자 독립적으로
    자기 밴드에 맞춰 스케일링하면(예전 방식) 폭이 좁은 쪽이 상대적으로 확대돼 두 줄의 음표
    크기가 달라 보인다. 원본 픽셀 높이가 같은 두 이미지(폭만 크게 다름)를 넣었을 때, 결과
    슬라이드의 두 PICTURE 높이(EMU)가 같아야 한다 — 배율이 공통이라는 강한 증거.
    """
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    wide_img = Image.new("RGB", (2000, 100), "white")
    narrow_img = Image.new("RGB", (300, 100), "white")
    prs = ps.build_slide(TEMPLATE, "시편 00(00)", wide_img, narrow_img)
    slide = prs.slides[0]
    big_pics = [
        sh for sh in slide.shapes
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE and sh.width > 1_000_000
    ]
    assert len(big_pics) == 2
    heights = sorted(sh.height for sh in big_pics)
    # 원본 픽셀 높이가 동일(100=100)하므로 배율이 같다면 결과 EMU 높이도 같아야 한다.
    assert abs(heights[0] - heights[1]) <= 2, (
        "두 단의 배율이 달라 높이가 다름 — 공통 배율 미적용", heights
    )
    widths = sorted(sh.width for sh in big_pics)
    # 폭 비율(2000:300 ≈ 6.67:1)이 결과에도 그대로 유지돼야 한다(공통 배율은 왜곡을 만들지 않음).
    assert abs(widths[1] / widths[0] - 2000 / 300) < 0.05


def test_15_residual_shapes_preserved(rendered):
    """잔재 4도형(12×7·41×17 EMU PICTURE) 그대로 — 템플릿 무손상."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    slide = rendered.slides[0]
    tiny_pics = [
        sh for sh in _iter_shapes(slide)
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE and sh.width < 1_000_000
    ]
    assert len(tiny_pics) == 4, [(sh.width, sh.height) for sh in tiny_pics]


def test_16_copy_slide_from_prs_compatible(rendered, tmp_path):
    """생성 Presentation을 타깃에 copy_slide_from_prs로 복사 시 손상 없이 rel까지 복사."""
    import missa_to_ppt as mtp
    import missa_ooxml_utils as ou
    from pptx import Presentation

    target = Presentation(str(TEMPLATE))
    ou.copy_slide_from_prs(target, len(target.slides), rendered, 0)
    out = tmp_path / "copied.pptx"
    target.save(str(out))
    problems = mtp.validate_pptx_structure(str(out))
    assert problems == [], problems


# ---------------------------------------------------------------------------
# 2차 마일스톤: missa_to_ppt.py 배선 (설계서 §4 통합 지점)
# ---------------------------------------------------------------------------

import missa_to_ppt as mtp  # noqa: E402  (모듈 상단 import 관례를 따르되, 1차 섹션과 분리해 배치)


def test_find_files_prefers_pptx_over_image(tmp_path):
    """화답송_pptx가 있으면 화답송_img 탐색 자체를 건너뛰고 pptx를 사용한다(기존 동작 보존)."""
    (tmp_path / "화답송 악보 시편 1.pptx").write_bytes(b"")
    (tmp_path / "photo.jpg").write_bytes(b"")
    files = mtp.find_files(str(tmp_path), {}, is_sunday=False)
    assert files["화답송_pptx"] is not None
    assert files["화답송_pptx"].name == "화답송 악보 시편 1.pptx"
    assert files["화답송_img"] is None


def test_find_files_named_image_when_no_pptx(tmp_path):
    """pptx가 없고 파일명에 '화답송'이 들어간 이미지가 있으면 그것을 화답송_img로 찾는다."""
    (tmp_path / "화답송 사진.jpg").write_bytes(b"")
    files = mtp.find_files(str(tmp_path), {}, is_sunday=False)
    assert files["화답송_pptx"] is None
    assert files["화답송_img"] is not None
    assert files["화답송_img"].name == "화답송 사진.jpg"


def test_find_files_unique_image_fallback_when_unnamed(tmp_path):
    """파일명에 '화답송'이 없어도(촬영 앱 타임스탬프 파일명) 폴더 내 유일 이미지면 폴백한다."""
    (tmp_path / "20260809_043842947.jpg").write_bytes(b"")
    files = mtp.find_files(str(tmp_path), {}, is_sunday=False)
    assert files["화답송_img"] is not None
    assert files["화답송_img"].name == "20260809_043842947.jpg"


def test_find_files_ambiguous_multiple_unnamed_images_no_fallback(tmp_path):
    """이름 매칭이 안 되고 이미지가 여러 개면 어떤 게 화답송 악보인지 모호하므로 폴백하지
    않는다 — 잘못된 이미지를 화답송으로 오인하는 것보다 미검출이 안전하다."""
    (tmp_path / "photo1.jpg").write_bytes(b"")
    (tmp_path / "photo2.png").write_bytes(b"")
    files = mtp.find_files(str(tmp_path), {}, is_sunday=False)
    assert files["화답송_img"] is None


def test_apply_hwadapsong_override_routes_image_extension():
    """--화답송 오버라이드가 이미지 확장자면 화답송_img로 가고, 반대쪽 키는 비워진다."""
    files = {"화답송_pptx": Path("auto_found.pptx"), "화답송_img": None}
    mtp.apply_화답송_override(files, "C:/mass/photo.jpg")
    assert files["화답송_img"] == Path("C:/mass/photo.jpg")
    assert files["화답송_pptx"] is None


def test_apply_hwadapsong_override_routes_pptx_extension():
    """--화답송 오버라이드가 pptx 확장자면 화답송_pptx로 가고, 반대쪽 키는 비워진다."""
    files = {"화답송_pptx": None, "화답송_img": Path("auto_found.jpg")}
    mtp.apply_화답송_override(files, "C:/mass/manual.pptx")
    assert files["화답송_pptx"] == Path("C:/mass/manual.pptx")
    assert files["화답송_img"] is None


def _화답송_json(title: str) -> dict:
    return {
        "화답송": {
            "title": title,
            "content": "◎ 후렴 테스트\n○ 절1\n○ 절2\n○ 절3",
        }
    }


def test_update_화답송_uses_image_when_no_pptx_end_to_end(tmp_path):
    """update_화답송이 화답송_img_path만 주어졌을 때 render_화답송_score_slide()를 거쳐
    만들어진 슬라이드가 실제 결과 PPT에 copy_slide_from_prs 경로로 정상 복사되는지 확인한다
    (실제 참조 예제 이미지로 end-to-end 검증, 설계서 §4.3)."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    prs = mtp.Presentation(
        str(BASE / "output" / "20260712" / "Template_20260628_연중 제13주일 (교황주일).pptx")
    )
    sections = mtp.find_sections(prs)
    json_data = _화답송_json("시편 99(98)")
    img_path = CASES["20260809"]["img"]

    mtp.update_화답송(prs, json_data, sections, None, 화답송_img_path=img_path, is_sunday=True)

    start = sections["화답송_start"]
    score_slide = prs.slides[start]  # i=0(짝수) => 악보 슬라이드
    big_pics = [
        sh for sh in score_slide.shapes
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE and sh.width > 1_000_000
    ]
    assert len(big_pics) == 2, "이미지 입력 경로로 악보 슬라이드에 상/하 악보 그림 2장이 있어야 함"

    texts = [sh.text_frame.text for sh in score_slide.shapes if sh.has_text_frame]
    assert any("시편 99(98)" in t for t in texts), texts

    out = tmp_path / "e2e_img.pptx"
    prs.save(str(out))
    assert mtp.validate_pptx_structure(str(out)) == []


def test_update_화답송_existing_pptx_path_unaffected(tmp_path):
    """기존 pptx 입력 경로가 이번 배선 변경으로 전혀 달라지지 않았는지 확인(회귀 방지)."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    prs = mtp.Presentation(
        str(BASE / "output" / "20260712" / "Template_20260628_연중 제13주일 (교황주일).pptx")
    )
    sections = mtp.find_sections(prs)
    json_data = _화답송_json("시편 67(66)")
    pptx_path = REF / "20260816" / "시편 67(66).pptx"

    mtp.update_화답송(prs, json_data, sections, pptx_path, 화답송_img_path=None, is_sunday=True)

    start = sections["화답송_start"]
    score_slide = prs.slides[start]
    big_pics = [
        sh for sh in score_slide.shapes
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE and sh.width > 1_000_000
    ]
    assert len(big_pics) == 2

    out = tmp_path / "e2e_pptx.pptx"
    prs.save(str(out))
    assert mtp.validate_pptx_structure(str(out)) == []


def test_update_화답송_prefers_pptx_over_image(monkeypatch):
    """화답송_pptx_path와 화답송_img_path가 둘 다 주어지면 pptx가 우선한다(설계서 §4.3 우선순위).

    render_화답송_score_slide가 절대 호출되지 않아야 함을 몽키패치로 직접 확인한다.
    """
    import missa_psalm_score_image

    def _must_not_be_called(*args, **kwargs):
        raise AssertionError("이미지 경로가 호출됨 — pptx가 우선이어야 하는데 위반됨")

    monkeypatch.setattr(
        missa_psalm_score_image, "render_화답송_score_slide", _must_not_be_called
    )

    prs = mtp.Presentation(
        str(BASE / "output" / "20260712" / "Template_20260628_연중 제13주일 (교황주일).pptx")
    )
    sections = mtp.find_sections(prs)
    json_data = _화답송_json("시편 1(1)")
    pptx_path = REF / "20260816" / "시편 67(66).pptx"
    img_path = CASES["20260809"]["img"]

    # 예외 없이 끝나야 함 = render_화답송_score_slide가 호출되지 않았다는 뜻
    mtp.update_화답송(
        prs, json_data, sections, pptx_path, 화답송_img_path=img_path, is_sunday=True
    )


# ---------------------------------------------------------------------------
# 공지사항 슬라이드 삽입 (설계서 _workspace/공지사항_슬라이드/01_architect_design.md)
# §6 테스트 대상 행동 1~9. 8번(GUI tkinter 확장자 검증)만 자동 테스트 제외.
# 기대 인덱스는 세 git 추적 픽스처의 실측 find_sections() 값이다(추측 아님):
#   20260712(주일) n=147, 2차봉헌_content_end=132(blank)
#   20260705(성수축복) n=144, 2차봉헌_content_end=129(blank)
#   20260624(평일) 2차봉헌_* 키 부재
# ---------------------------------------------------------------------------

from pptx import Presentation as _Presentation  # noqa: E402
from pptx.util import Emu as _Emu  # noqa: E402
from missa_content_updaters import insert_공지사항  # noqa: E402
from missa_ooxml_utils import all_slide_texts as _all_slide_texts  # noqa: E402

_FIXTURES = {
    "20260712": BASE / "output" / "20260712" / "20260712_연중 제15주일.pptx",
    "20260705": BASE / "output" / "20260705"
    / "20260705_한국 성직자들의 수호자 성 김대건 안드레아 사제 순교자 - 신심 미사.pptx",
    "20260624": BASE / "output" / "20260624" / "20260624_성 요한 세례자 탄생 대축일.pptx",
}


def _make_공지사항_pptx(path, n):
    """식별 가능한 텍스트를 가진 n장짜리 공지사항 PPT를 임시 생성한다.

    각 슬라이드에 '공지사항 슬라이드 {i}' 텍스트박스를 넣어 삽입 후 서식/순서 보존을
    강하게 단언할 수 있게 한다. n==0이면 슬라이드 없는 빈 PPT를 만든다.
    """
    prs = _Presentation()
    for i in range(n):
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout
        box = slide.shapes.add_textbox(_Emu(914400), _Emu(914400), _Emu(3657600), _Emu(914400))
        box.text_frame.text = f"공지사항 슬라이드 {i}"
    prs.save(str(path))
    return prs


def _src_texts(path):
    src = _Presentation(str(path))
    return [
        "\n".join(
            sh.text_frame.text for sh in s.shapes if sh.has_text_frame
        ).strip()
        for s in src.slides
    ]


def test_공지사항_01_insert_count_주일(tmp_path):
    """[삽입 총수] N장짜리 공지사항 → 반환 N+1, 전체 슬라이드 == 기존 + N + 1."""
    from missa_sections import find_sections

    n_src = 4
    src = tmp_path / "공지사항.pptx"
    _make_공지사항_pptx(src, n_src)

    prs = _Presentation(str(_FIXTURES["20260712"]))
    before = len(prs.slides)
    sec = find_sections(prs)
    added = insert_공지사항(prs, src, sec)

    assert added == n_src + 1
    assert len(prs.slides) == before + n_src + 1


def test_공지사항_02_one_divider_each_side(tmp_path):
    """[앞뒤 구분선 정확히 1장씩] 앞 구분선(기존 content_end)·뒤 구분선(신규) 모두 blank,
    그 뒤는 non-blank(마침 축복). 공지사항 경계에 2장 연속 blank가 없음."""
    from missa_sections import find_sections

    n_src = 3
    src = tmp_path / "공지사항.pptx"
    _make_공지사항_pptx(src, n_src)

    prs = _Presentation(str(_FIXTURES["20260712"]))
    sec = find_sections(prs)
    front = sec["2차봉헌_content_end"]  # 132 (실측)
    insert_공지사항(prs, src, sec)

    texts = _all_slide_texts(prs)
    is_blank = [not t.strip() for t in texts]

    # 앞 구분선: content_end 자리 그대로 blank
    assert is_blank[front], ("앞 구분선이 blank가 아님", front)
    # 공지사항 콘텐츠 n_src장: 모두 non-blank
    for i in range(n_src):
        assert not is_blank[front + 1 + i], ("공지사항 콘텐츠가 blank임", front + 1 + i)
    # 뒤 구분선: 콘텐츠 마지막 다음 자리 blank
    rear = front + 1 + n_src
    assert is_blank[rear], ("뒤 구분선이 blank가 아님", rear)
    # 2026-09-13 실측 버그의 end-to-end 재발 방지: 뒤 구분선도 앞 구분선과 동일하게
    # showMasterSp="0"이어야 한다(다르면 마스터 상속 요소가 노출돼 색이 달라 보임).
    front_show_master = prs.slides[front].element.get("showMasterSp")
    rear_show_master = prs.slides[rear].element.get("showMasterSp")
    assert rear_show_master == front_show_master, (
        "뒤 구분선의 showMasterSp가 앞 구분선과 다름", front_show_master, rear_show_master,
    )
    # 뒤 구분선 다음은 non-blank(마침 축복 기도)
    assert not is_blank[rear + 1], ("뒤 구분선 다음이 blank임 — 2장 연속 blank", rear + 1)
    # front-1(2차봉헌 마지막 콘텐츠)도 non-blank → front 앞뒤로 2장 연속 blank 없음
    assert not is_blank[front - 1], ("앞 구분선 앞이 blank임 — 2장 연속 blank", front - 1)


def test_공지사항_03_content_format_preserved(tmp_path):
    """[콘텐츠 서식 보존] 삽입된 공지사항 슬라이드 텍스트가 원본과 동일(순서 포함).
    뒤 구분선 슬라이드는 shapes 0개."""
    from missa_sections import find_sections

    n_src = 3
    src = tmp_path / "공지사항.pptx"
    _make_공지사항_pptx(src, n_src)
    src_texts = _src_texts(src)  # ['공지사항 슬라이드 0', '...1', '...2']

    prs = _Presentation(str(_FIXTURES["20260712"]))
    sec = find_sections(prs)
    front = sec["2차봉헌_content_end"]
    insert_공지사항(prs, src, sec)

    for i in range(n_src):
        slide = prs.slides[front + 1 + i]
        joined = "\n".join(
            sh.text_frame.text for sh in slide.shapes if sh.has_text_frame
        ).strip()
        assert joined == src_texts[i], (i, joined, src_texts[i])

    rear = front + 1 + n_src
    assert len(prs.slides[rear].shapes) == 0, "뒤 구분선은 원본 blank 복제라 shapes 0개여야 함"


def test_공지사항_04_no_2차봉헌_slot_skips_평일(tmp_path):
    """[2차봉헌 없는 평일 → 경고 후 skip] 예외 없이 반환 0, 슬라이드 수 불변."""
    from missa_sections import find_sections

    src = tmp_path / "공지사항.pptx"
    _make_공지사항_pptx(src, 2)

    prs = _Presentation(str(_FIXTURES["20260624"]))
    before = len(prs.slides)
    sec = find_sections(prs)
    assert sec.get("2차봉헌_content_end") is None  # 실측: 키 자체가 없음

    added = insert_공지사항(prs, src, sec)
    assert added == 0
    assert len(prs.slides) == before


def test_공지사항_05_insert_성수축복(tmp_path):
    """[성수축복(2차봉헌 있는 특수 주일)] content_end=129 기준 삽입 성립."""
    from missa_sections import find_sections

    n_src = 2
    src = tmp_path / "공지사항.pptx"
    _make_공지사항_pptx(src, n_src)

    prs = _Presentation(str(_FIXTURES["20260705"]))
    before = len(prs.slides)
    sec = find_sections(prs)
    front = sec["2차봉헌_content_end"]  # 129 (실측)
    added = insert_공지사항(prs, src, sec)

    assert added == n_src + 1
    assert len(prs.slides) == before + n_src + 1

    texts = _all_slide_texts(prs)
    is_blank = [not t.strip() for t in texts]
    assert is_blank[front]
    rear = front + 1 + n_src
    assert is_blank[rear]
    assert not is_blank[rear + 1]


def test_공지사항_06a_missing_path_returns_zero(tmp_path):
    """[로드 실패 방어] 존재하지 않는 경로 → 반환 0, 슬라이드 불변, 예외 없음."""
    from missa_sections import find_sections

    prs = _Presentation(str(_FIXTURES["20260712"]))
    before = len(prs.slides)
    sec = find_sections(prs)
    added = insert_공지사항(prs, tmp_path / "없는파일.pptx", sec)
    assert added == 0
    assert len(prs.slides) == before


def test_공지사항_06b_zero_slide_pptx_returns_zero(tmp_path):
    """[0장 방어] 슬라이드 0장짜리 공지사항 PPT → 반환 0, 슬라이드 불변."""
    from missa_sections import find_sections

    src = tmp_path / "empty.pptx"
    _make_공지사항_pptx(src, 0)

    prs = _Presentation(str(_FIXTURES["20260712"]))
    before = len(prs.slides)
    sec = find_sections(prs)
    added = insert_공지사항(prs, src, sec)
    assert added == 0
    assert len(prs.slides) == before


def test_공지사항_07_front_not_blank_skips(tmp_path):
    """[앞 구분선이 blank 아님 → skip] content_end가 non-blank를 가리키면 반환 0(방어 (c))."""
    from missa_sections import find_sections

    src = tmp_path / "공지사항.pptx"
    _make_공지사항_pptx(src, 2)

    prs = _Presentation(str(_FIXTURES["20260712"]))
    before = len(prs.slides)
    sec = find_sections(prs)
    # 132(blank) 대신 133(마침 축복, non-blank)을 가리키도록 조작
    sec = dict(sec, **{"2차봉헌_content_end": 133})
    added = insert_공지사항(prs, src, sec)
    assert added == 0
    assert len(prs.slides) == before


def test_공지사항_structure_valid_after_insert(tmp_path):
    """[구조 무손상] 삽입·저장 후 validate_pptx_structure()가 빈 리스트(끊어진 rId/제어문자 없음)."""
    from missa_sections import find_sections

    src = tmp_path / "공지사항.pptx"
    _make_공지사항_pptx(src, 3)

    prs = _Presentation(str(_FIXTURES["20260712"]))
    sec = find_sections(prs)
    insert_공지사항(prs, src, sec)

    out = tmp_path / "out.pptx"
    prs.save(str(out))
    assert mtp.validate_pptx_structure(str(out)) == []


def test_공지사항_08_cli_rejects_non_pptx(monkeypatch):
    """[CLI 확장자 검증] --공지사항에 .pptx 아닌 경로 → SystemExit(오류 종료)."""
    monkeypatch.setattr(
        "sys.argv", ["missa_to_ppt.py", "20260712", "--공지사항", "foo.docx", "--test"]
    )
    with pytest.raises(SystemExit):
        mtp.parse_args()


def test_공지사항_08b_cli_accepts_pptx_into_5tuple(monkeypatch):
    """[CLI 통과] .pptx면 parse_args가 5-튜플을 반환하고 5번째가 공지사항 경로다."""
    monkeypatch.setattr(
        "sys.argv", ["missa_to_ppt.py", "20260712", "--공지사항", "notice.pptx", "--test"]
    )
    result = mtp.parse_args()
    assert len(result) == 5
    assert result[4] == "notice.pptx"


def test_공지사항_08c_cli_default_none(monkeypatch):
    """[CLI 기본값] --공지사항 미지정이면 5번째가 None(선택 입력)."""
    monkeypatch.setattr("sys.argv", ["missa_to_ppt.py", "20260712", "--test"])
    result = mtp.parse_args()
    assert len(result) == 5
    assert result[4] is None


def test_공지사항_09_find_files_detects_by_name(tmp_path):
    """[자동 탐색] 파일명에 '공지사항'이 든 .pptx가 files['공지사항']에 잡힌다."""
    (tmp_path / "20260712_ref.pptx").write_bytes(b"")
    (tmp_path / "공지사항.pptx").write_bytes(b"")
    files = mtp.find_files(str(tmp_path), {}, is_sunday=False)
    assert files["공지사항"] is not None
    assert files["공지사항"].name == "공지사항.pptx"


def test_공지사항_09b_find_files_none_when_absent(tmp_path):
    """[미입력 무변경 가드 근거] 공지사항 파일이 없으면 files['공지사항']==None →
    main의 `if files.get('공지사항')` 가드가 False라 insert_공지사항이 호출되지 않는다."""
    (tmp_path / "20260712_ref.pptx").write_bytes(b"")
    files = mtp.find_files(str(tmp_path), {}, is_sunday=False)
    assert files.get("공지사항") is None
    assert not files.get("공지사항")  # 가드 falsy


# ---------------------------------------------------------------------------
# duplicate_slide()의 showMasterSp 속성 보존 (2026-09-13 실측 버그)
# insert_공지사항이 만든 뒤 구분 슬라이드(142, duplicate_slide 산출물)가 앞 구분 슬라이드
# (133, hand-authored)와 색이 달라 보였다. 두 슬라이드는 layout/배경이 100% 동일했으나
# 최상위 <p:sld>의 showMasterSp 속성이 앞쪽은 "0"(마스터 상속 요소 숨김), 뒤쪽은 부재
# (python-pptx add_slide 기본값 = 마스터 요소 노출)였던 것이 원인. duplicate_slide가
# spTree/이미지 rel/p:bg는 복사하면서 <p:sld> 자체의 showMasterSp 속성은 복사하지 않았다.
# ---------------------------------------------------------------------------

from missa_ooxml_utils import duplicate_slide as _duplicate_slide  # noqa: E402


def test_showmastersp_present_copied_to_duplicate():
    """원본 <p:sld showMasterSp="0">이면 복제본의 <p:sld>도 showMasterSp="0"이어야 한다.

    이 속성이 누락되면 layout/배경이 동일해도 마스터 배치 요소(장식·로고 등)가 복제본에서만
    노출돼 색이 달라 보인다. p:bg 복사만으로는 커버되지 않는 별도 속성이다.
    """
    prs = _Presentation()
    src = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    src.element.set("showMasterSp", "0")
    new_idx = _duplicate_slide(prs, 0)
    assert prs.slides[new_idx].element.get("showMasterSp") == "0"


def test_showmastersp_absent_stays_absent():
    """원본에 showMasterSp 속성이 없으면 복제본에도 추가하지 않는다(불필요한 속성 주입 금지).

    원본이 add_slide 기본 상태(속성 부재 = 마스터 요소 노출)면 복제본도 그 상태를 그대로
    유지해야 한다 — 원본에 없던 속성을 새로 만들어 넣으면 안 된다.
    """
    prs = _Presentation()
    src = prs.slides.add_slide(prs.slide_layouts[6])
    assert src.element.get("showMasterSp") is None  # 전제: add_slide 기본은 속성 부재
    new_idx = _duplicate_slide(prs, 0)
    assert prs.slides[new_idx].element.get("showMasterSp") is None


def test_showmastersp_real_fixture_handauthored_slide():
    """실측 재현: 20260712 idx132(hand-authored, showMasterSp="0")를 복제하면 복제본도 "0".

    2026-09-13 프로덕션 파일의 앞 구분 슬라이드(idx132)와 동일한 성격의 슬라이드다(git 추적
    픽스처 20260712 idx132도 showMasterSp="0"임을 실측). 이 슬라이드를 재사용해 뒤 구분
    슬라이드를 만들 때 속성이 보존돼야 두 슬라이드가 같은 색으로 렌더링된다.
    """
    prs = _Presentation(str(_FIXTURES["20260712"]))
    assert prs.slides[132].element.get("showMasterSp") == "0"  # 픽스처 실측 전제
    new_idx = _duplicate_slide(prs, 132)
    assert prs.slides[new_idx].element.get("showMasterSp") == "0"
