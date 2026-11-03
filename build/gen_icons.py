"""Render emoji icons (16/32/80 px, transparent PNG) from build/icons.json. Needs Pillow + Segoe UI Emoji (Windows)."""
import json, sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

root = Path(__file__).resolve().parent
out = root.parent / "src" / "assets"
out.mkdir(parents=True, exist_ok=True)
font = ImageFont.truetype("seguiemj.ttf", 109)


def emoji(ch):
    img = Image.new("RGBA", (192, 192), (0, 0, 0, 0))
    ImageDraw.Draw(img).text((96, 96), ch, font=font, embedded_color=True, anchor="mm")
    return img.crop(img.getbbox() or (0, 0, 192, 192))


def fit(img, size, pad):
    box = size - 2 * pad
    k = min(box / img.width, box / img.height)
    img = img.resize((max(1, round(img.width * k)), max(1, round(img.height * k))), Image.LANCZOS)
    c = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    c.paste(img, ((size - img.width) // 2, (size - img.height) // 2), img)
    return c


def logo(size):
    """Gradient rounded tile with a bolt - used as the add-in icon."""
    s = 256
    g = Image.new("RGBA", (s, s))
    px = g.load()
    for y in range(s):
        for x in range(s):
            t = (x + y) / (2 * s)
            px[x, y] = (int(79 + (236 - 79) * t), int(70 + (72 - 70) * t), int(229 + (153 - 229) * t), 255)
    m = Image.new("L", (s, s), 0)
    ImageDraw.Draw(m).rounded_rectangle((0, 0, s - 1, s - 1), 56, fill=255)
    g.putalpha(m)
    g.alpha_composite(fit(emoji("⚡"), s, 52))
    return g.resize((size, size), Image.LANCZOS)


icons = json.loads((root / "icons.json").read_text(encoding="utf-8"))
for key, ch in icons.items():
    base = emoji(ch)
    for size in (16, 32, 80):
        fit(base, size, 1 if size == 16 else 3).save(out / f"{key}_{size}.png")
for size in (16, 32, 64, 80, 128):
    logo(size).save(out / f"logo_{size}.png")
print("icons:", len(icons) * 3 + 5)
