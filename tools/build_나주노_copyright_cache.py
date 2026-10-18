# -*- coding: utf-8 -*-
"""나주노 저작권 크롭 캐시 생성(개발 1회성, Tesseract 필요).

나주노 PDF는 스캔 래스터라 텍스트 검색이 불가하다. 개발자 PC에서 이 스크립트로 번호별
저작권 줄을 1회 OCR해 "Copyright ~ 트리거 문구 직전" 크롭의 정규화 bbox를 구해
assets/나주노_copyright_bbox_cache.json에 기록한다. 런타임/exe는 이 캐시만 읽고 Tesseract를
호출하지 않는다(요구사항 §배포 방식). 저작권이 없는 곡은 has_copyright:false로 기록한다.

주의: 생성된 크롭 PNG는 반드시 원본과 육안 대조해 bbox를 확정한다(OCR 정확도는 사람이 확인 —
요구사항). 이 스크립트는 좌표 초안을 만들 뿐이며, 최종 캐시 값은 사람이 검수한다.

사용:  python tools/build_나주노_copyright_cache.py 267 447        # 지정 번호 OCR
       python tools/build_나주노_copyright_cache.py --no-copyright 362 173 146
설정:  TESSDATA_PREFIX(kor/eng traineddata 폴더), Tesseract 실행 경로(_TESS_CMD).
"""
import argparse
import json
import os
import sys
from pathlib import Path

import numpy as np

BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE))

import missa_youth_hymn_pdf as hp  # noqa: E402

_TESS_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
_TESSDATA = str(BASE / "assets" / "tessdata")
CACHE = BASE / "assets" / "나주노_copyright_bbox_cache.json"
CROP_DIR = BASE / "_workspace" / "청년미사_1단계" / "crops"

_HUMAN_VERIFIED_FALSE = {362, 173, 146}  # 사람이 원본과 대조해 저작권 없음으로 확정한 곡
_INK = 128        # 잉크(어두운 픽셀) 임계 — hp._INK와 동일 관용
_PERIOD_MAX_W = 8  # 이 폭(px, 300dpi RENDER_DPI) 이하의 마지막 잉크 런은 꼬리 마침표로 본다


def _trailing_glyph_right_edge(gray, left, right_bound, y0, y1):
    """저작권자 이름 마지막 '글자'의 우측 x를 반환(꼬리 마침표 제외).

    트리거 유무와 무관하게 동작한다(항목 4-1 일반화): right_bound는 트리거가 있으면 trig_x,
    없으면(폴백) 마지막 단어 우측 등 임의의 탐색 상한이다. OCR 워드박스는 이 스캔의 한글을
    오독하고(447 '849+.', 267 '4173.') 마침표를 이름 단어에 붙여 잡으므로 우측 경계로 못 쓴다 —
    요구사항이 "마침표 자체도 제외"라고 못 박은 것을 OCR 박스로는 만족시킬 수 없다. 대신
    [left, right_bound) 밴드의 세로 잉크 열을 글자 런으로 끊어, 마지막 런이 마침표 폭
    (≤_PERIOD_MAX_W) 이하면 그 앞 런의 끝을 경계로 삼는다. 267의 내부 마침표(천태혁**.** 진경)는
    '진경' 뒤에 오므로 마지막 런이 아니라 그대로 보존된다."""
    band = gray[y0:y1, left:right_bound] < _INK
    col = band.sum(axis=0)
    runs = []
    i, n = 0, len(col)
    while i < n:
        if col[i] > 0:
            s = i
            while i < n and col[i] > 0:
                i += 1
            runs.append((left + s, left + i - 1))
        else:
            i += 1
    if not runs:
        return right_bound
    lx0, lx1 = runs[-1]
    if (lx1 - lx0 + 1) <= _PERIOD_MAX_W and len(runs) >= 2:
        return runs[-2][1]
    return lx1


def _all_naju_numbers():
    import re
    import fitz
    d = fitz.open(hp.SOURCES["나주노"]["pdf"])
    try:
        return sorted({int(m.group(1)) for _l, t, _p in d.get_toc()
                       for m in [re.match(r"^\s*(\d+)\b", t)] if m})
    finally:
        d.close()


def _is_shared_page(number: int) -> bool:
    return hp.naju_page_is_shared(hp.find_song_page("나주노", number))


def _has_bottom_text(number: int) -> bool:
    """페이지 하단 저작권 영역(hp._COPYRIGHT_ZONE_RATIO 아래 ~ 페이지 끝)에 글줄이 있나.

    세그먼트 분류에 의존하지 않고 **원시 행 잉크**로 본다: 저작권이 마지막 가사 줄과 가까워
    마지막 시스템 세그먼트에 병합된 곡(13·20·163·654)이 '텍스트 없음'으로 오판돼 has_copyright:
    false로 기록되던 결함(2026-10-18 독립 리뷰) 방지. 행당 잉크 15px 이상인 행이 3개 이상이면 글줄."""
    import fitz
    pi = hp.find_song_page("나주노", number)
    d = fitz.open(hp.SOURCES["나주노"]["pdf"])
    try:
        gray = np.asarray(hp.render_region(d[pi], None, hp.RENDER_DPI).convert("L"))
    finally:
        d.close()
    h = gray.shape[0]
    zone = gray[int(h * hp._COPYRIGHT_ZONE_RATIO - 40):] < _INK
    return int((zone.sum(axis=1) >= 15).sum()) >= 3


def _ocr_bbox(number: int, zone=None):
    """번호의 저작권 줄을 OCR해 (norm_bbox, trigger, crop_img) 반환. 실패 시 None.

    zone=(y_lo, y_hi): OCR을 이 픽셀 세로 구간으로 제한한다(공유 페이지 곡별 영역). 기본은 페이지 하단 15%."""
    import fitz
    import pytesseract
    from pytesseract import Output
    pytesseract.pytesseract.tesseract_cmd = _TESS_CMD
    os.environ.setdefault("TESSDATA_PREFIX", _TESSDATA)

    pi = hp.find_song_page("나주노", number)
    d = fitz.open(hp.SOURCES["나주노"]["pdf"])
    try:
        img = hp.render_region(d[pi], None, hp.RENDER_DPI)
    finally:
        d.close()
    W, H = img.size
    # 기본은 페이지 전체를 OCR한다(2026-10-18 독립 재리뷰): 짧은 곡(19·24·29·34번)은 저작권이 페이지
    # 중간에 있어 하단 15%만 보면 놓쳤다. 저작권 단어는 글자가 작고 흐려 'Copvright'처럼 오독되기도 해
    # 퍼지 매칭으로 찾고, 여러 개면 가장 아래(곡 바로 아래) 것을 쓴다.
    y0, y_end = (0, H) if zone is None else (int(zone[0]), int(zone[1]))
    data = pytesseract.image_to_data(img.crop((0, y0, W, y_end)), lang="eng",
                                     output_type=Output.DICT)
    words = [(data["text"][i].strip(), data["left"][i], y0 + data["top"][i],
              data["width"][i], data["height"][i])
             for i in range(len(data["text"])) if data["text"][i].strip()]
    import difflib
    cands = [w for w in words
             if difflib.SequenceMatcher(None, w[0].lower().strip(".,:;©"), "copyright").ratio() >= 0.75]
    if not cands:
        # 'Copyright'가 'Convrieht'처럼 심하게 오독된 곡(348번)은 같은 줄의 'Administered'(오독이 드물다)를
        # 닻으로 삼아 그 줄의 가장 왼쪽 단어를 저작권 시작으로 본다.
        anchors = [w for w in words
                   if difflib.SequenceMatcher(None, w[0].lower().strip(".,:;"), "administered").ratio() >= 0.8]
        if not anchors:
            return None
        anchor = max(anchors, key=lambda w: w[2])
        same_line = [w for w in words if abs(w[2] - anchor[2]) < 15]
        cands = [min(same_line, key=lambda w: w[1])]
    cw = max(cands, key=lambda w: w[2])
    line_y0 = min(w[2] for w in words if abs(w[2] - cw[2]) < 15)
    line_y1 = max(w[2] + w[4] for w in words if abs(w[2] - cw[2]) < 15)
    trig = None
    trig_x = None
    for t in hp.COPYRIGHT_TRIGGERS:
        head = t.split()[0]
        for w in words:
            if w[0].startswith(head) and w[1] > cw[1]:
                if trig_x is None or w[1] < trig_x:
                    trig_x, trig = w[1], t
    left = cw[1] - 3
    gray = np.asarray(img.convert("L"))
    # 항목 4-1: 트리거 유무와 무관하게 꼬리 마침표 트림을 적용. 폴백(트리거 없음)에서는
    # 탐색 상한을 마지막 단어 우측으로 둔다("…All rights reserved."처럼 트리거 없이 끝나는 줄).
    right_bound = trig_x if trig_x is not None else max(w[1] + w[3] for w in words)
    edge = _trailing_glyph_right_edge(gray, cw[1], right_bound, line_y0, line_y1)
    right = edge + 2  # 마지막 글자 끝 살짝 뒤(마침표 시작 전에서 끝남)
    box = (left, line_y0 - 4, right, line_y1 + 4)
    # 둘째 줄(All rights reserved…)이 첫 줄과 바짝 붙은 곡(692)은 +4 여유가 둘째 줄 윗끝을
    # 물어온다. 박스 안 행 잉크를 런으로 나눠 첫 줄 런 끝(+1)까지만 세로로 줄인다(2026-10-18).
    rows = (gray[box[1]:box[3], box[0]:box[2]] < _INK).sum(axis=1)
    ink_runs = hp._nonblank_runs(rows)
    if len(ink_runs) >= 2:
        box = (box[0], box[1] + ink_runs[0][0] - 1, box[2], box[1] + ink_runs[0][1] + 2)
    crop = img.crop(box)
    norm = [round(box[0] / W, 4), round(box[1] / H, 4),
            round(box[2] / W, 4), round(box[3] / H, 4)]
    return norm, trig, crop


# ---------------------------------------------------------------------------
# 공유 페이지(한 페이지에 곡 2~3개) — 곡별 세로 영역(y_range) 등록 (2026-10-18)
# ---------------------------------------------------------------------------
_TITLE_DIGIT_MIN_H = 95      # 곡 번호(굵은 큰 숫자)의 최소 글자 높이(px@300dpi). 본문 숫자(≈70px)와 구분
_TITLE_REGION_MARGIN = 40    # 영역 상단 = 다음 곡 번호 상단 − 이 값(제목·작곡가 줄 포함 여유)
PREVIEW_DIR = BASE / "_workspace" / "청년미사_1단계" / "shared_preview"


def _shared_pages():
    """TOC에서 곡이 둘 이상인 페이지 {page_index: [번호, ...]}(TOC 순서)."""
    import re
    import collections
    import fitz
    d = fitz.open(hp.SOURCES["나주노"]["pdf"])
    try:
        pages = collections.defaultdict(list)
        for _l, t, p in d.get_toc():
            m = re.match(r"^\s*(\d+)\b", t)
            if m:
                pages[p - 1].append(int(m.group(1)))
    finally:
        d.close()
    return {p: v for p, v in sorted(pages.items()) if len(v) > 1}


def _title_tops(img, expected: int):
    """페이지 이미지에서 곡 번호(큰 굵은 숫자)의 상단 y들(오름차순). OCR이 숫자를 오독해도(55→35)
    '큰 숫자 덩어리의 개수·위치'만 쓰므로 영향이 없다. 한 자리 번호(8·9)는 psm 11이 놓치는 경우가
    있어 psm 6으로 한 번 더 시도하며, 개수가 `expected`(페이지의 곡 수)와 맞는 결과를 채택한다."""
    import pytesseract
    from pytesseract import Output
    pytesseract.pytesseract.tesseract_cmd = _TESS_CMD
    os.environ.setdefault("TESSDATA_PREFIX", _TESSDATA)
    best = []
    for psm in (11, 6):
        data = pytesseract.image_to_data(
            img, lang="eng", config=f"--psm {psm} -c tessedit_char_whitelist=0123456789",
            output_type=Output.DICT)
        tops = sorted(data["top"][i] for i in range(len(data["text"]))
                      if data["text"][i].strip().isdigit() and data["height"][i] >= _TITLE_DIGIT_MIN_H)
        merged = []
        for y in tops:
            if not merged or y - merged[-1] > 300:   # 한 번호가 여러 조각으로 잡힌 경우 병합
                merged.append(y)
        if len(merged) == expected:
            return merged
        best = best or merged
    return best


def _register_shared(cache):
    """모든 공유 페이지 곡에 y_range와 저작권 항목을 등록하고 미리보기 PNG를 저장한다.
    큰 숫자 개수가 페이지의 곡 수와 다르면 그 페이지는 건너뛰고 보고한다(수동 확인)."""
    import fitz
    d = fitz.open(hp.SOURCES["나주노"]["pdf"])
    PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
    skipped = []
    try:
        for pi, nums in _shared_pages().items():
            img = hp.render_region(d[pi], None, hp.RENDER_DPI)
            W, H = img.size
            tops = _title_tops(img, len(nums))
            if len(tops) != len(nums):
                print(f"페이지 {pi + 1} {nums}: 큰 숫자 {len(tops)}개 ≠ 곡 {len(nums)}개 → 건너뜀(수동 확인)")
                skipped += nums
                continue
            edges = [0] + [max(0, y - _TITLE_REGION_MARGIN) for y in tops[1:]] + [H]
            for k, num in enumerate(nums):
                lo, hi = edges[k], edges[k + 1]
                gray = np.asarray(img.convert("L"))[lo:hi]
                bands = hp.detect_system_bands(gray)
                # 저작권 탐색 시작 = 마지막 '오선 세그먼트' 아래(밴드 아래가 아니다 — 저작권 줄이 마지막
                # 밴드에 가사 보충줄로 흡수돼 있을 수 있다)
                sys_bottoms = [b for _t, b, is_sys in hp._page_segments(gray)["segments"] if is_sys]
                last = lo + (max(sys_bottoms) if sys_bottoms else (hi - lo) // 2)
                res = _ocr_bbox(num, zone=(lo, hi))   # 곡 영역 전체 — 저작권이 마지막 시스템에 붙어 있어도 찾는다
                entry = {"y_range": [round(lo / H, 4), round(hi / H, 4)]}
                if res is None:
                    entry["has_copyright"] = False
                    print(num, f"-> y_range {entry['y_range']}, 저작권 없음(영역 안 Copyright 미검출)")
                else:
                    norm, trig, crop = res
                    entry.update({"has_copyright": True, "bbox": norm, "trigger": trig or ""})
                    crop.save(PREVIEW_DIR / f"COPY_{num}.png")
                    print(num, f"-> y_range {entry['y_range']}, bbox {norm}")
                cache[str(num)] = entry
    finally:
        d.close()
    if skipped:
        print("건너뛴 곡(수동 확인 필요):", skipped)


def _write_shared_previews():
    """저장된 캐시 기준으로 실제 생성에 쓰이는 곡 영역(naju_song_region_image)과 시스템 밴드(빨간 상자)를
    그린 미리보기를 저장한다 — 곡 하나만 들어갔는지 육안 확인용."""
    import fitz
    from PIL import ImageDraw
    d = fitz.open(hp.SOURCES["나주노"]["pdf"])
    try:
        for pi, nums in _shared_pages().items():
            for num in nums:
                img = hp.naju_song_region_image(d[pi], num)
                gray = np.asarray(img.convert("L"))
                dr = ImageDraw.Draw(img)
                for t, b in hp.detect_system_bands(gray):
                    dr.rectangle((10, t, img.width - 10, b), outline=(255, 0, 0), width=5)
                img.resize((img.width // 3, img.height // 3)).save(PREVIEW_DIR / f"SONG_{num}.png")
    finally:
        d.close()


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("numbers", nargs="*", type=int)
    ap.add_argument("--shared", action="store_true",
                    help="한 페이지에 곡이 둘 이상인 곡들의 곡별 y_range와 저작권을 등록(미리보기 PNG는 "
                         "_workspace/청년미사_1단계/shared_preview/)")
    ap.add_argument("--recheck-false", action="store_true",
                    help="캐시의 has_copyright:false 항목을 하단 원시 잉크로 재판정해, 글줄이 있으면 "
                         "항목을 지우고 OCR로 다시 만든다(과거 오기록 정정)")
    ap.add_argument("--all-missing", action="store_true",
                    help="TOC 전체 중 캐시에 없는 번호를 일괄 처리(기존 항목은 건드리지 않음). "
                         "하단에 텍스트 세그먼트가 없으면 has_copyright:false, 있는데 OCR이 못 "
                         "찾으면 미기록(목록 출력 → 수동 확인)")
    ap.add_argument("--no-copyright", action="store_true",
                    help="지정 번호를 has_copyright:false로 기록(OCR 안 함)")
    ns = ap.parse_args(argv)

    cache = json.loads(CACHE.read_text(encoding="utf-8")) if CACHE.exists() else {}
    CROP_DIR.mkdir(parents=True, exist_ok=True)
    if ns.shared:
        _register_shared(cache)
        CACHE.write_text(json.dumps(cache, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print("저장:", CACHE)
        _write_shared_previews()
        print("미리보기:", PREVIEW_DIR)
        if not (ns.numbers or ns.all_missing or ns.recheck_false):
            return
    numbers = list(ns.numbers)
    if ns.recheck_false:
        # 하단 잉크 유무가 아니라 OCR(페이지 전체, 퍼지 매칭)로 다시 판정하도록 false 항목을 비운다.
        for k in [k for k, v in cache.items() if k.isdigit() and not v.get("has_copyright")]:
            if cache[k].get("y_range"):
                continue   # 공유 페이지 곡 항목은 --shared가 관리한다(y_range를 지우면 안 된다)
            if int(k) in _HUMAN_VERIFIED_FALSE:
                continue   # 육안 확정(2026-09-10) 항목은 자동 재판정으로 건드리지 않는다
            print(k, "-> false 항목 비움(재 OCR 대상)")
            del cache[k]
    if ns.all_missing:
        shared = {n for n in _all_naju_numbers() if _is_shared_page(n)}
        if shared:
            print("공유 페이지 곡(자동 처리 제외, 수동 확인 필요):", sorted(shared))
        numbers += [n for n in _all_naju_numbers() if str(n) not in cache and n not in shared]
    unresolved = []
    for num in numbers:
        if ns.no_copyright:
            cache[str(num)] = {"has_copyright": False}
            print(num, "-> has_copyright: false")
            continue
        # (하단 잉크 유무로 false를 미리 확정하지 않는다 — 항상 페이지 전체 OCR로 판정한다)
        res = _ocr_bbox(num)
        if res is None:
            if ns.all_missing:
                # 하단 15%를 OCR했는데 'Copyright' 단어가 없다 → 저작권 없음(하단 잉크는 페이지 번호·
                # 마지막 시스템 가사 등). 표본 육안 확인 대상이므로 ocr_checked 표시를 남긴다.
                cache[str(num)] = {"has_copyright": False, "ocr_checked": True}
                print(num, "-> OCR에서 Copyright 없음 → has_copyright: false (ocr_checked)")
            else:
                print(num, "-> Copyright 미검출 (수동 확인 필요)")
                unresolved.append(num)
            continue
        norm, trig, crop = res
        p = CROP_DIR / f"COPY_{num}.png"
        crop.save(p)
        entry = {"has_copyright": True, "bbox": norm, "trigger": trig or ""}
        prev_text = cache.get(str(num), {}).get("text")
        if prev_text:  # 사람이 적어둔 확정 텍스트(육안 대조 기준)는 재생성 시 보존
            entry["text"] = prev_text
        cache[str(num)] = entry
        print(num, "-> bbox", norm, "trigger", trig, "| 육안 대조:", p)

    CACHE.write_text(json.dumps(cache, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print("저장:", CACHE)


if __name__ == "__main__":
    main()
