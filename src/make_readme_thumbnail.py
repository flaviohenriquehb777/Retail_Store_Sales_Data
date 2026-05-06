from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageEnhance


def main() -> None:
    src = Path("docs/assets/thumbnail.png")
    dst = Path("docs/assets/thumbnail_dark_rounded.png")

    im = Image.open(src).convert("RGBA")
    im = ImageEnhance.Brightness(im).enhance(0.78)
    im = ImageEnhance.Contrast(im).enhance(1.08)
    overlay = Image.new("RGBA", im.size, (10, 12, 16, 60))
    im = Image.alpha_composite(im, overlay)

    radius = int(min(im.size) * 0.035)
    mask = Image.new("L", im.size, 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle((0, 0, im.size[0], im.size[1]), radius=radius, fill=255)
    im.putalpha(mask)

    border = Image.new("RGBA", im.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(border)
    d.rounded_rectangle(
        (1, 1, im.size[0] - 2, im.size[1] - 2),
        radius=max(0, radius - 1),
        outline=(48, 54, 61, 255),
        width=2,
    )
    im = Image.alpha_composite(im, border)

    dst.parent.mkdir(parents=True, exist_ok=True)
    im.save(dst)
    print(dst.as_posix())


if __name__ == "__main__":
    main()

