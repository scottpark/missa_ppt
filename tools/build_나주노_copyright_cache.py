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


def _ocr_bbox(number: int):
    """번호의 저작권 줄을 OCR해 (norm_bbox, trigger, crop_img) 반환. 실패 시 None."""
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
    y0 = int(H * 0.85)
    data = pytesseract.image_to_data(img.crop((0, y0, W, H)), lang="eng",
                                     output_type=Output.DICT)
    words = [(data["text"][i].strip(), data["left"][i], y0 + data["top"][i],
              data["width"][i], data["height"][i])
             for i in range(len(data["text"])) if data["text"][i].strip()]
    cw = next((w for w in words if w[0] == "Copyright"), None)
    if cw is None:
        return None
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
    crop = img.crop(box)
    norm = [round(box[0] / W, 4), round(box[1] / H, 4),
            round(box[2] / W, 4), round(box[3] / H, 4)]
    return norm, trig, crop


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("numbers", nargs="+", type=int)
    ap.add_argument("--no-copyright", action="store_true",
                    help="지정 번호를 has_copyright:false로 기록(OCR 안 함)")
    ns = ap.parse_args(argv)

    cache = json.loads(CACHE.read_text(encoding="utf-8")) if CACHE.exists() else {}
    CROP_DIR.mkdir(parents=True, exist_ok=True)
    for num in ns.numbers:
        if ns.no_copyright:
            cache[str(num)] = {"has_copyright": False}
            print(num, "-> has_copyright: false")
            continue
        res = _ocr_bbox(num)
        if res is None:
            print(num, "-> Copyright 미검출 (수동 확인 필요)")
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
