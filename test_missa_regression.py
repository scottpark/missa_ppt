"""
missa_to_ppt.py 회귀 테스트 스위트

목적
----
주일미사·평일미사 처리 로직의 회귀 테스트. 향후 토요일미사 지원이 추가되거나
다른 기능이 변경됐을 때, 기존 주일/평일 로직이 깨지지 않았는지 확인하는 기준선.

실행
----
    pip install pytest
    pytest test_missa_regression.py -v

테스트 데이터
------------
- 주일: 20260712 (2026-07-12, 연중 제15주일)
- 평일: 20260624 (2026-06-24, 성 요한 세례자 탄생 대축일, 수요일)

두 폴더 모두 참조 PPT·JSON·성가 파일이 준비되어 있어야 한다.

새 케이스(예: 토요일미사) 추가 시 MASS_CASES 리스트에 항목을 추가하면 된다.
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

import missa_to_ppt as m

REPO_ROOT = Path(__file__).resolve().parent

MASS_CASES = [
    {
        "id": "sunday",
        "date": "20260712",
        "is_sunday": True,
        "hymns": {"입당": "329", "봉헌": "221", "성체": "156", "2차봉헌": "221", "파견": "25"},
    },
    {
        "id": "weekday",
        "date": "20260624",
        "is_sunday": False,
        "hymns": {"입당": "329", "봉헌": "221", "성체": "156", "파견": "25"},
    },
]

READING_SECTIONS = [
    ("제1독서_start", "제1독서_end"),
    ("제2독서_start", "제2독서_end"),
    ("복음_start", "복음_end"),
]

HYMN_TYPES = ["입당", "봉헌", "성체", "2차봉헌", "파견"]


# ─────────────────────────────────────────────────────────────────
# 1. 단위 테스트 — 순수 함수
# ─────────────────────────────────────────────────────────────────

class TestIsSundayMass:
    """향후 토요일미사(전야미사) 판단 로직 추가 시 이 클래스에
    '토요일을 평일미사/주일 전야 중 무엇으로 볼 것인가' 케이스를 추가해서
    회귀를 잡는다."""

    def test_sunday(self):
        assert m.is_sunday_mass("20260712") is True  # 일요일

    def test_weekday_friday(self):
        assert m.is_sunday_mass("20260612") is False  # 금요일

    def test_weekday_wednesday(self):
        assert m.is_sunday_mass("20260624") is False  # 수요일

    def test_saturday_currently_treated_as_weekday(self):
        # 2026-07-11은 토요일. 현재 로직(주일 여부만 판단)에서는 '평일미사'로 처리된다.
        # 토요일미사 지원이 추가되면 이 테스트는 새 기대값으로 업데이트해야 한다.
        assert m.is_sunday_mass("20260711") is False


class TestWrapLineCount:
    def test_empty_text(self):
        assert m._wrap_line_count("") == 0

    def test_short_line_is_one_line(self):
        assert m._wrap_line_count("짧은 문장입니다.") == 1

    def test_long_text_wraps_to_multiple_lines(self):
        text = "가" * 60  # CHARS_PER_LINE(27) 초과
        assert m._wrap_line_count(text) >= 2


class TestParseIntoVerseUnits:
    def test_verse_numbers_are_detected(self):
        content = (
            "1 태초에 하느님께서 하늘과 땅을 창조하셨다. "
            "2 땅은 아직 모양을 갖추지 않고 비어 있었다."
        )
        units = m.parse_into_verse_units(content)
        assert len(units) >= 2
        assert units[0]["verse_num"] == "1"


# ─────────────────────────────────────────────────────────────────
# 2. 통합 테스트 — 실제 PPT 생성 + 구조 검증
# ─────────────────────────────────────────────────────────────────

def _find_output(date_str: str) -> Path:
    folder = REPO_ROOT / date_str
    candidates = [
        f for f in folder.glob(f"{date_str}_*.pptx")
        if not f.name.startswith("~$")
    ]
    assert candidates, f"{date_str} 출력 PPT를 찾을 수 없음"
    return max(candidates, key=lambda f: f.stat().st_mtime)


def _generate(case: dict) -> str:
    date_str = case["date"]
    folder = REPO_ROOT / date_str

    locked = list(folder.glob("~$*.pptx"))
    if locked:
        pytest.skip(f"{locked[0].name} 이(가) PowerPoint에서 열려 있어 생성을 건너뜀")

    args = [sys.executable, str(REPO_ROOT / "missa_to_ppt.py"), date_str]
    for k, v in case["hymns"].items():
        args += [f"--{k}", v]

    result = subprocess.run(
        args, cwd=REPO_ROOT, capture_output=True, text=True, encoding="utf-8",
    )
    assert result.returncode == 0, (
        f"{date_str} 생성 실패 (exit={result.returncode})\n"
        f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
    )
    assert "모든 검증 통과!" in result.stdout, (
        f"{date_str} 내부 검증(validate) 실패\nstdout:\n{result.stdout}"
    )
    return result.stdout


@pytest.fixture(scope="module", params=MASS_CASES, ids=lambda c: c["id"])
def generated_case(request):
    case = request.param
    log = _generate(case)
    path = _find_output(case["date"])
    prs = Presentation(str(path))
    sections = m.find_sections(prs)
    return {"case": case, "prs": prs, "sections": sections, "log": log, "path": path}


def test_output_opens_without_corruption(generated_case):
    # Presentation()이 예외 없이 로드되면 OOXML 구조 자체는 유효함(fixture에서 이미 검증됨).
    assert len(generated_case["prs"].slides) > 0


def test_liturgy_type_detected_correctly(generated_case):
    log = generated_case["log"]
    case = generated_case["case"]
    expected = "미사 유형: 주일미사" if case["is_sunday"] else "미사 유형: 평일미사"
    assert expected in log


def test_required_sections_present(generated_case):
    sections = generated_case["sections"]
    required = [
        "입당송", "제1독서_start", "제1독서_end",
        "화답송_start", "화답송_end",
        "복음환호송", "복음_start", "복음_end",
        "영성체송",
    ]
    missing = [k for k in required if k not in sections]
    assert not missing, f"누락된 섹션: {missing}"


def test_no_reading_slide_line_overflow(generated_case):
    """독서·복음 본문 슬라이드가 LINES_PER_SLIDE(9줄)를 넘으면 안 된다.
    2026-07-16 발견된 post-write 재조정 실패(연쇄 초과) 회귀 방지용."""
    prs, sections = generated_case["prs"], generated_case["sections"]
    problems = []
    for start_key, end_key in READING_SECTIONS:
        if start_key not in sections or end_key not in sections:
            continue
        s, e = sections[start_key], sections[end_key]
        for idx in range(s, e):
            lines = m._count_slide_lines_rendered(prs.slides[idx])
            if lines > m.LINES_PER_SLIDE:
                problems.append(f"{start_key} 슬라이드 {idx + 1}: {lines}줄 (>{m.LINES_PER_SLIDE})")
    assert not problems, "\n".join(problems)


def test_blackbg_covers_full_slide(generated_case):
    """BlackBg 도형이 존재하면 항상 슬라이드 전체 크기를 덮어야 한다.
    2026-07-16 발견된 _align_ending_slides_to_제2독서의 BlackBg 오분류 회귀 방지용."""
    prs = generated_case["prs"]
    problems = []
    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.name == "BlackBg":
                if not (
                    shape.left == 0 and shape.top == 0
                    and shape.width == prs.slide_width
                    and shape.height == prs.slide_height
                ):
                    problems.append(
                        f"슬라이드 {i + 1}: BlackBg 크기 이상 "
                        f"({shape.left},{shape.top},{shape.width},{shape.height})"
                    )
    assert not problems, "\n".join(problems)


def test_화답송_pattern_matches_mass_type(generated_case):
    """주일: 악보/텍스트 슬라이드 교대(짝수 인덱스=악보). 평일: 텍스트만, 악보 이미지 없음."""
    prs, sections, case = (
        generated_case["prs"], generated_case["sections"], generated_case["case"]
    )
    if "화답송_start" not in sections:
        pytest.skip("화답송 섹션 없음")
    s, e = sections["화답송_start"], sections["화답송_end"]
    problems = []
    for i in range(s, e):
        slide = prs.slides[i]
        has_picture = any(sh.shape_type == MSO_SHAPE_TYPE.PICTURE for sh in slide.shapes)
        rel_i = i - s
        if case["is_sunday"]:
            expected_picture = (rel_i % 2 == 0)
            if has_picture != expected_picture:
                problems.append(
                    f"화답송 슬라이드 {i + 1}: 악보 유무 불일치 "
                    f"(기대={expected_picture}, 실제={has_picture})"
                )
        else:
            if has_picture:
                problems.append(f"화답송 슬라이드 {i + 1}: 평일미사인데 악보 이미지 존재")
    assert not problems, "\n".join(problems)


def test_hymn_score_copy_matches_mass_type(generated_case):
    """주일: 성가 악보 슬라이드 복사됨. 평일: 번호만 갱신, 악보 슬라이드 없음."""
    sections, case = generated_case["sections"], generated_case["case"]
    problems = []
    for hymn in HYMN_TYPES:
        s_key, e_key = f"{hymn}_content_start", f"{hymn}_content_end"
        if s_key not in sections or e_key not in sections:
            continue
        n_content = sections[e_key] - sections[s_key]
        if case["is_sunday"]:
            if n_content == 0:
                problems.append(f"{hymn}: 주일미사인데 악보 슬라이드 없음")
        else:
            if n_content != 0:
                problems.append(f"{hymn}: 평일미사인데 악보 슬라이드 {n_content}장 존재")
    assert not problems, "\n".join(problems)
