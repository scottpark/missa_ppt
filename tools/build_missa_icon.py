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


def build_master() -> Image.Image:
    """원형 배지형 마스터 이미지(256x256, RGBA)를 생성한다."""
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

    font = _find_font(int(badge_d * 0.34))
    text = 'PPT'
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    tx = bx + (badge_d - tw) // 2 - bbox[0]
    ty = by + (badge_d - th) // 2 - bbox[1]
    draw.text((tx, ty), text, font=font, fill=WHITE)

    return canvas


def main() -> None:
    if not SRC_LOGO.exists():
        raise FileNotFoundError(f'로고 원본 없음: {SRC_LOGO}')
    master = build_master()
    OUT_ICO.parent.mkdir(parents=True, exist_ok=True)
    master.save(OUT_ICO, format='ICO', sizes=ICO_SIZES)
    print(f'생성 완료: {OUT_ICO} ({len(ICO_SIZES)}개 해상도)')


if __name__ == '__main__':
    main()
