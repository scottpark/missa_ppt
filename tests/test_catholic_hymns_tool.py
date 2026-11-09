# -*- coding: utf-8 -*-
"""tools/make_catholic_hymns.py — 가사 누락 방지 회귀 테스트(웹 악보 이미지 필요, 오프라인이면 skip)."""
import sys
from pathlib import Path

import numpy as np
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))
import make_catholic_hymns as m  # noqa: E402

TEST_SONGS = [434, 294, 50, 513, 203, 456, 40, 410, 400, 504, 506, 202]


def _img(n):
    try:
        return m.fetch_image(n)
    except Exception as e:  # 네트워크 불가
        pytest.skip(f"악보 이미지 받기 실패: {e}")


def test_203_last_system_keeps_lyrics():
    """203번 마지막 시스템은 오선↔가사 간격이 15px 이상이라, 예전엔 가사가 통째로 잘렸다."""
    gray = np.asarray(_img(203).convert("L"))
    bands = m.detect_systems(gray)
    assert bands[-1][1] > 680, bands          # 가사 줄(≈650~700)까지 포함
    assert m.lost_ink_rows(gray, bands) == []


@pytest.mark.parametrize("n", TEST_SONGS)
def test_no_lost_ink(n):
    gray = np.asarray(_img(n).convert("L"))
    bands = m.detect_systems(gray)
    assert bands
    assert m.lost_ink_rows(gray, bands) == []
    _crops, _sizes, _scales = m.score_groups(_img(n))
    assert not m.LOSS_WARN, m.LOSS_WARN


@pytest.mark.parametrize("n", TEST_SONGS)
def test_scale_steps(n):
    _crops, sizes, scales = m.score_groups(_img(n))
    assert all(s in (100, 92, 82) for s in scales)
    assert all(sc == 100 for sz, sc in zip(sizes, scales) if sz == 1)
