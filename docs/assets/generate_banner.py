#!/usr/bin/env python3
"""Bespoke README banners, one composition per project."""
import sys
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageFont

W, H = 2400, 640
CJK_B = "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"
CJK_R = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
FONT_DIR = Path("/usr/share/matplotlib/mpl-data/fonts/ttf")
WHITE = (248, 250, 252)


def f_cjk(size, bold=True):
    return ImageFont.truetype(CJK_B if bold else CJK_R, size, index=2)


def f_sans(size):
    return ImageFont.truetype(str(FONT_DIR / "DejaVuSans-Bold.ttf"), size)


def f_mono(size, bold=False):
    return ImageFont.truetype(str(FONT_DIR / ("DejaVuSansMono-Bold.ttf" if bold else "DejaVuSansMono.ttf")), size)


def use_cjk(text):
    return any(ord(ch) > 0x2E80 for ch in text)


CONFIGS = {
    "lab": dict(
        kicker="LAB MANAGEMENT  //  FACE + BARCODE",
        lines=[("实验室", 108), ("管理系统", 78)],
        subtitle="人脸识别签到  ·  扫码出入库  ·  Excel 台账  ·  单机 Tkinter 桌面端",
        metrics=[("2", "SERIAL PORTS"), ("9", "THEMES"), ("3", "TABLES"), ("2", "ROLES")],
        accent=(251, 191, 36), text_y=140, subtitle_y=412, metrics_y=486,
    ),
}


def base():
    t = np.linspace(0, 1, H)[:, None, None]
    img = np.repeat(np.array((5, 9, 17))[None, None, :] * (1 - t)
                    + np.array((11, 22, 38))[None, None, :] * t, W, axis=1)
    ys, xs = np.mgrid[0:H, 0:W].astype(float)
    img[((xs % 48) < 1) | ((ys % 48) < 1)] += 9.0
    return img


def vignette(canvas):
    ys, xs = np.mgrid[0:H, 0:W].astype(float)
    vx = (xs - W / 2) / (W / 2)
    vy = (ys - H / 2) / (H / 2)
    vig = 1.0 - 0.20 * np.clip(vx ** 2 + vy ** 2, 0, 1) ** 1.5
    arr = np.asarray(canvas.convert("RGB"), dtype=float) * vig[..., None]
    return Image.fromarray(np.clip(arr, 0, 255).astype(np.uint8), "RGB").convert("RGBA")


def halo(img, cx, cy, accent, radius=600, strength=0.20):
    ys, xs = np.mgrid[0:H, 0:W].astype(float)
    d = np.hypot(xs - cx, ys - cy) / radius
    img += (np.clip(1 - d, 0, 1) ** 2 * strength)[..., None] * np.array(accent) * 0.6


def draw_text(canvas, cfg):
    d = ImageDraw.Draw(canvas)
    accent = cfg["accent"]
    left, y = 140, cfg["text_y"]
    d.rectangle((96, cfg["text_y"] - 6, 104, cfg["metrics_y"] + 46), fill=accent + (190,))
    d.text((left, y - 14), cfg["kicker"], font=f_mono(26), fill=accent + (255,))
    y += 34
    for i, (text, size) in enumerate(cfg["lines"]):
        f = f_cjk(size) if use_cjk(text) else f_sans(size)
        d.text((left - 4, y), text, font=f, fill=(WHITE if i == 0 else accent) + (255,))
        y += int(size * 1.14) + 4
    d.text((left, cfg["subtitle_y"]), cfg["subtitle"], font=f_cjk(27, bold=False), fill=(148, 163, 184, 255))
    x = left
    for value, label in cfg["metrics"]:
        d.text((x, cfg["metrics_y"]), value, font=f_mono(30, True), fill=WHITE + (255,))
        wv = d.textlength(value, font=f_mono(30, True))
        d.text((x + wv + 10, cfg["metrics_y"] + 10), label, font=f_mono(19), fill=(100, 116, 139, 255))
        x += wv + 10 + d.textlength(label, font=f_mono(19)) + 42


# ---------------------------------------------------------------- motifs

def motif_barcode(d, accent):
    rng = np.random.default_rng(7)
    x, y0, y1 = 1610, 186, 470
    while x < 2250:
        w = int(rng.integers(3, 10))
        a = int(rng.integers(90, 210))
        d.rectangle([x, y0, x + w, y1], fill=accent + (a,))
        x += w + int(rng.integers(4, 12))
    beam_y = 320
    d.rectangle([1560, beam_y - 2, 2300, beam_y + 2], fill=(255, 214, 120, 235))
    for xx in (1560, 2300):
        d.polygon([(xx, beam_y - 9), (xx, beam_y + 9), (xx + (10 if xx < 1800 else -10), beam_y)],
                  fill=accent + (200,))
    b = 34
    for cx, cy, sx, sy in ((1550, 150, 1, 1), (2310, 150, -1, 1), (1550, 506, 1, -1), (2310, 506, -1, -1)):
        d.line([(cx, cy), (cx + sx * b, cy)], fill=accent + (150,), width=3)
        d.line([(cx, cy), (cx, cy + sy * b)], fill=accent + (150,), width=3)


def render(out):
    cfg = CONFIGS["lab"]
    accent = cfg["accent"]
    img = base()
    halo(img, 1900, 320, accent, 620, 0.20)

    ov = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    motif_barcode(d, accent)

    canvas = Image.fromarray(np.clip(img, 0, 255).astype(np.uint8), "RGB").convert("RGBA")
    canvas = Image.alpha_composite(canvas, ov.filter(ImageFilter.GaussianBlur(7)))
    canvas = Image.alpha_composite(canvas, ov.filter(ImageFilter.GaussianBlur(0.4)))
    canvas = vignette(canvas)
    draw_text(canvas, cfg)
    canvas.convert("RGB").save(out, optimize=True)
    print(f"wrote {out} ({Path(out).stat().st_size / 1024:.0f} KB)")


if __name__ == "__main__":
    render(sys.argv[1] if len(sys.argv) > 1 else "banner.png")
