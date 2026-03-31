from __future__ import annotations

import math
from pathlib import Path

from PIL import Image, ImageDraw


def _rounded_rect(draw: ImageDraw.ImageDraw, size: int, r: int, fill: str) -> None:
    draw.rounded_rectangle((0, 0, size, size), radius=r, fill=fill)


def _arc(draw: ImageDraw.ImageDraw, bbox, start_deg: float, end_deg: float, width: int, color: str) -> None:
    draw.arc(bbox, start=start_deg, end=end_deg, fill=color, width=width)


def render_icon(size: int) -> Image.Image:
    """
    生成一个类似插件图标风格的 macOS app 图标：
    - 深色圆角方底
    - 底部三道弧线（白色）
    """
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    bg = "#1e2429"
    white = "#ffffff"

    r = int(size * 0.19)
    _rounded_rect(d, size=size, r=r, fill=bg)

    stroke = max(2, int(size * 0.047))  # 128px≈6px
    cx = size / 2
    cy = size * 0.84

    def arc_radius(k: float) -> float:
        return size * k

    for k in (0.16, 0.26, 0.36):
        rad = arc_radius(k)
        bbox = (cx - rad, cy - rad, cx + rad, cy + rad)
        _arc(d, bbox, start_deg=210, end_deg=330, width=stroke, color=white)

    return img


def main() -> None:
    repo = Path(__file__).resolve().parents[1]
    build_dir = repo / "build"
    iconset_dir = build_dir / "icon.iconset"
    iconset_dir.mkdir(parents=True, exist_ok=True)

    # macOS iconset sizes
    sizes = [
        (16, False),
        (16, True),
        (32, False),
        (32, True),
        (128, False),
        (128, True),
        (256, False),
        (256, True),
        (512, False),
        (512, True),
    ]

    for base, retina in sizes:
        px = base * 2 if retina else base
        img = render_icon(px)
        name = f"icon_{base}x{base}{'@2x' if retina else ''}.png"
        img.save(iconset_dir / name, format="PNG")


if __name__ == "__main__":
    main()

