# -*- coding: utf-8 -*-
"""assets/Logo_brokenbaykcc.png(성당 로고) + "PPT" 텍스트를 결합해
assets/missa_to_ppt.ico(다중 해상도)를 생성한다.

디자인: 원형 배지형 — 로고 전체(알파 채널 기준 자동 크롭)를 캔버스에 배치하고,
우측 하단에 흰 테두리 + 적갈색(#871B24, 팝업 UI _UI['primary']와 동일 계열) 원형 배지,
그 안에 흰색 굵은 "PPT" 텍스트를 중앙 정렬한다(사용자 승인 시안, 2026-09-09).
"""
from __future__ import annotations

from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
SRC_LOGO = ROOT / 'assets' / 'Logo_brokenbaykcc.png'
OUT_ICO = ROOT / 'assets' / 'missa_to_ppt.ico'

MAROON = (135, 27, 36, 255)
WHITE = (255, 255, 255, 255)
BASE = 256  # 최종 png 마스터 크기. ICO 저장 시 이 마스터를 각 해상도로 리샘플한다.
ICO_SIZES = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]


def _load_logo(max_w: int, max_h: int) -> Image.Image:
    """가로세로 비율을 유지한 채 (max_w, max_h) 안에 맞춘다 (알파 bbox로 자동 크롭 후)."""
    im = Image.open(SRC_LOGO).convert('RGBA')
    bbox = im.getbbox()
    if bbox:
        im = im.crop(bbox)
    w, h = im.size
    scale = min(max_w / w, max_h / h)
    new_size = (max(1, round(w * scale)), max(1, round(h * scale)))
    return im.resize(new_size, Image.LANCZOS)


def _find_font(size: int) -> ImageFont.FreeTypeFont:
    for cand in (
        r'C:\Windows\Fonts\arialbd.ttf',
        r'C:\Windows\Fonts\segoeuib.ttf',
        r'C:\Windows\Fonts\calibrib.ttf',
    ):
        if Path(cand).exists():
            return ImageFont.truetype(cand, size)
    return ImageFont.load_default()


def _find_font_kr(size: int) -> ImageFont.FreeTypeFont:
    """§J9 — 2줄 배지의 윗줄('성인'/'청년'/'어린이')은 한글이라, `_find_font()`의
    Latin-only 후보(Arial/Segoe UI/Calibri Bold)로는 글리프가 없어 빈 사각형(tofu)으로
    나온다. Malgun Gothic Bold(맑은 고딕, 한글+라틴 모두 지원)로 별도 찾는다 — 'PPT'
    아랫줄도 같은 폰트로 그려 두 줄의 서체가 갈리지 않게 한다."""
    for cand in (
        r'C:\Windows\Fonts\malgunbd.ttf',
        r'C:\Windows\Fonts\malgun.ttf',
    ):
        if Path(cand).exists():
            return ImageFont.truetype(cand, size)
    return _find_font(size)


OUT_ICO_BY_TYPE = {
    '성인': ROOT / 'assets' / 'missa_to_ppt_성인.ico',
    '청년': ROOT / 'assets' / 'missa_to_ppt_청년.ico',
    '어린이': ROOT / 'assets' / 'missa_to_ppt_어린이.ico',
}


def build_master(lines: tuple[str, str] | None = None) -> Image.Image:
    """원형 배지형 마스터 이미지(256x256, RGBA)를 생성한다.

    §J9(2026-09-26) — `lines`가 None이면 기존 1줄 'PPT' 배지를 한 글자도 바꾸지 않고
    그린다(기본 `missa_to_ppt.ico`의 겉모습을 그대로 보존하려면 이 분기가 기존 코드와
    완전히 동일해야 한다). `lines=(윗줄, 아랫줄)`을 주면(예: `('성인', 'PPT')`) 2줄
    배지를 그린다 — 1줄일 때(badge_d*0.34)보다 작은 폰트(badge_d*0.22)를 써야 두 줄이
    겹치지 않는다.
    """
    canvas = Image.new('RGBA', (BASE, BASE), (0, 0, 0, 0))

    logo = _load_logo(int(BASE * 0.88), int(BASE * 0.80))
    lx = (BASE - logo.width) // 2
    ly = (BASE - logo.height) // 2 - int(BASE * 0.02)
    canvas.alpha_composite(logo, (lx, ly))

    badge_d = int(BASE * 0.52)
    bx = BASE - badge_d - int(BASE * 0.02)
    by = BASE - badge_d - int(BASE * 0.02)
    draw = ImageDraw.Draw(canvas)
    draw.ellipse([bx - 4, by - 4, bx + badge_d + 4, by + badge_d + 4], fill=WHITE)
    draw.ellipse([bx, by, bx + badge_d, by + badge_d], fill=MAROON)

    if lines is None:
        font = _find_font(int(badge_d * 0.34))
        text = 'PPT'
        bbox = draw.textbbox((0, 0), text, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        tx = bx + (badge_d - tw) // 2 - bbox[0]
        ty = by + (badge_d - th) // 2 - bbox[1]
        draw.text((tx, ty), text, font=font, fill=WHITE)
        return canvas

    cx = bx + badge_d // 2
    cy = by + badge_d // 2
    font = _find_font_kr(int(badge_d * 0.22))
    top_text, bottom_text = lines

    top_bbox = draw.textbbox((0, 0), top_text, font=font)
    bottom_bbox = draw.textbbox((0, 0), bottom_text, font=font)
    top_h = top_bbox[3] - top_bbox[1]
    bottom_h = bottom_bbox[3] - bottom_bbox[1]
    gap = int(badge_d * 0.06)
    total_h = top_h + gap + bottom_h
    top_y = cy - total_h // 2
    bottom_y = top_y + top_h + gap

    top_w = top_bbox[2] - top_bbox[0]
    top_x = cx - top_w // 2 - top_bbox[0]
    draw.text((top_x, top_y - top_bbox[1]), top_text, font=font, fill=WHITE)

    bottom_w = bottom_bbox[2] - bottom_bbox[0]
    bottom_x = cx - bottom_w // 2 - bottom_bbox[0]
    draw.text((bottom_x, bottom_y - bottom_bbox[1]), bottom_text, font=font, fill=WHITE)

    return canvas


def main() -> None:
    if not SRC_LOGO.exists():
        raise FileNotFoundError(f'로고 원본 없음: {SRC_LOGO}')

    # 기본/하위호환용 — 기존 1줄 'PPT' 배지, 겉모습 변경 없음.
    master = build_master()
    OUT_ICO.parent.mkdir(parents=True, exist_ok=True)
    master.save(OUT_ICO, format='ICO', sizes=ICO_SIZES)
    print(f'생성 완료: {OUT_ICO} ({len(ICO_SIZES)}개 해상도)')

    # §J9 — 미사유형별 2줄 배지(성인/청년/어린이 + 'PPT').
    for mass_type_kr, out_path in OUT_ICO_BY_TYPE.items():
        variant = build_master(lines=(mass_type_kr, 'PPT'))
        out_path.parent.mkdir(parents=True, exist_ok=True)
        variant.save(out_path, format='ICO', sizes=ICO_SIZES)
        print(f'생성 완료: {out_path} ({len(ICO_SIZES)}개 해상도)')


if __name__ == '__main__':
    main()
