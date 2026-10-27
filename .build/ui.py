import math
from PIL import Image, ImageDraw, ImageFont

S = 2
W, H = 540, 720

FONT_DISPLAY = "/usr/share/fonts/opentype/smiley-sans/SmileySans-Oblique.otf"
FONT_BOLD = "/usr/share/fonts/opentype/noto-cjk-otf/NotoSansCJKsc-Bold.otf"
FONT_REG = "/usr/share/fonts/opentype/noto-cjk-otf/NotoSansCJKsc-Regular.otf"
FONT_MONO = "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"

WHITE = (238, 244, 255)
MUTED = (170, 183, 210)
DIM = (122, 134, 160)
GREEN = (126, 255, 186)
GREEN_DEEP = (0, 208, 140)
AMBER = (255, 197, 92)
RED = (255, 128, 128)

_fcache = {}


def rgba(c):
    return c if len(c) == 4 else c + (255,)


def F(path, size):
    key = (path, int(size * S))
    if key not in _fcache:
        _fcache[key] = ImageFont.truetype(path, int(size * S))
    return _fcache[key]


class Canvas:
    def __init__(self, img):
        self.img = img
        self.draw = ImageDraw.Draw(img)

    @classmethod
    def blank(cls):
        top, bot = (10, 14, 28), (24, 30, 56)
        col = Image.new("RGBA", (1, H * S))
        for y in range(H * S):
            t = y / (H * S - 1)
            col.putpixel((0, y), tuple(int(top[i] + (bot[i] - top[i]) * t) for i in range(3)) + (255,))
        img = col.resize((W * S, H * S))

        gw, gh = 470 * S, 470 * S
        mask = Image.radial_gradient("L").resize((gw, gh))
        inv = mask.point(lambda v: int((255 - v) * 0.28))
        glow = Image.new("RGBA", (gw, gh), GREEN_DEEP + (255,))
        glow.putalpha(inv)
        img.alpha_composite(glow, (W * S - gw // 2, -gh // 2))

        layer = Image.new("RGBA", img.size, (0, 0, 0, 0))
        d = ImageDraw.Draw(layer)
        step = 30 * S
        for x in range(0, W * S, step):
            d.line([(x, 0), (x, H * S)], fill=(255, 255, 255, 7), width=1)
        for y in range(0, H * S, step):
            d.line([(0, y), (W * S, y)], fill=(255, 255, 255, 7), width=1)
        img.alpha_composite(layer)
        return cls(img)

    def _paint(self, fn, alpha):
        if alpha >= 255:
            fn(self.draw)
        else:
            layer = Image.new("RGBA", self.img.size, (0, 0, 0, 0))
            fn(ImageDraw.Draw(layer))
            self.img = Image.alpha_composite(self.img, layer)
            self.draw = ImageDraw.Draw(self.img)

    def rrect(self, x, y, w, h, r, fill=None, outline=None, width=1):
        fill = rgba(fill) if fill else None
        outline = rgba(outline) if outline else None
        a = max(fill[3] if fill else 0, outline[3] if outline else 0)

        def f(d):
            d.rounded_rectangle(
                [x * S, y * S, (x + w) * S - 1, (y + h) * S - 1],
                radius=r * S, fill=fill, outline=outline, width=max(1, int(round(width * S))),
            )
        self._paint(f, a)

    def line(self, x1, y1, x2, y2, fill, width=1):
        fill = rgba(fill)
        self._paint(lambda d: d.line([(x1 * S, y1 * S), (x2 * S, y2 * S)], fill=fill,
                                     width=max(1, int(round(width * S)))), fill[3])

    def circle(self, cx, cy, r, fill=None, outline=None, width=1):
        fill = rgba(fill) if fill else None
        outline = rgba(outline) if outline else None
        a = max(fill[3] if fill else 0, outline[3] if outline else 0)
        self._paint(lambda d: d.ellipse([(cx - r) * S, (cy - r) * S, (cx + r) * S, (cy + r) * S],
                                        fill=fill, outline=outline,
                                        width=max(1, int(round(width * S)))), a)

    def tw(self, s, font, tracking=0):
        if tracking == 0:
            return self.draw.textlength(s, font=font) / S
        return sum(self.draw.textlength(ch, font=font) for ch in s) / S + tracking * (len(s) - 1)

    def text(self, pos, s, font, fill, tracking=0, align="left"):
        fill = rgba(fill)
        x, y = pos
        if align == "right":
            x -= self.tw(s, font, tracking)
        if tracking == 0:
            self._paint(lambda d: d.text((x * S, y * S), s, font=font, fill=fill), fill[3])
            return
        cur = x

        def f(d):
            nonlocal cur
            for ch in s:
                d.text((cur * S, y * S), ch, font=font, fill=fill)
                cur += d.textlength(ch, font=font) / S + tracking
        self._paint(f, fill[3])

    def save(self, path):
        self.img.convert("RGB").resize((1080, 1440), Image.LANCZOS).save(path, quality=95)


def header(c, label):
    tw = c.tw(label, F(FONT_BOLD, 13))
    c.rrect(37, 30, 30 + tw + 20, 32, 16, fill=(255, 255, 255, 13), outline=(255, 255, 255, 28))
    c.circle(56, 46, 4.5, fill=GREEN)
    c.text((67, 38), label, F(FONT_BOLD, 13), (228, 238, 255))


def title(c, s, y, x=37, col=WHITE, size=52):
    c.text((x, y), s, F(FONT_DISPLAY, size), col)


def sub(c, s, y, x=37, col=MUTED, size=14):
    c.text((x, y), s, F(FONT_REG, size), col)


def footer(c, page, left):
    c.line(37, 666, 503, 666, (255, 255, 255, 22), 1)
    c.text((37, 678), left, F(FONT_REG, 12), DIM)
    num = "%02d / 06" % page
    c.text((503, 678), num, F(FONT_MONO, 11), DIM, align="right")


def icon(c, kind, x, y, s=34, col=GREEN):
    c.rrect(x, y, s, s, 10, fill=col + (24,), outline=col + (46,))
    cx, cy = x + s / 2, y + s / 2
    ink = col + (255,)
    lw = 2.2
    if kind == "yen":
        c.text((cx - 5, cy - 9), "¥", F(FONT_BOLD, 15), ink)
    elif kind == "check":
        c.line(cx - 7, cy, cx - 2, cy + 6, ink, lw)
        c.line(cx - 2, cy + 6, cx + 8, cy - 7, ink, lw)
    elif kind == "page":
        c.rrect(cx - 7, cy - 9, 14, 18, 3, outline=ink, width=lw)
        c.line(cx - 3, cy - 3, cx + 4, cy - 3, ink, 1.6)
        c.line(cx - 3, cy + 2, cx + 4, cy + 2, ink, 1.6)
    elif kind == "chart":
        c.line(cx - 8, cy + 8, cx + 8, cy + 8, ink, 1.6)
        for i, h in enumerate((7, 12, 9, 15)):
            c.line(cx - 6 + i * 4, cy + 8, cx - 6 + i * 4, cy + 8 - h, ink, 2.4)
    elif kind == "clock":
        c.circle(cx, cy, 9, outline=ink, width=lw)
        c.line(cx, cy, cx, cy - 5, ink, lw)
        c.line(cx, cy, cx + 5, cy + 2, ink, lw)
    elif kind == "box":
        c.rrect(cx - 8, cy - 6, 16, 13, 2, outline=ink, width=lw)
        c.line(cx - 8, cy - 1, cx + 8, cy - 1, ink, 1.6)
        c.line(cx, cy - 6, cx, cy + 7, ink, 1.6)
    elif kind == "list":
        for i in range(3):
            yy = cy - 6 + i * 6
            c.circle(cx - 6, yy, 1.3, fill=ink)
            c.line(cx - 1, yy, cx + 8, yy, ink, 1.8)
    elif kind == "grid":
        for i in range(2):
            for j in range(2):
                c.rrect(cx - 8 + i * 9, cy - 8 + j * 9, 7, 7, 2, outline=ink, width=1.6)


def phone():
    w, h = 176, 356
    img = Image.new("RGBA", (w * S, h * S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    def rr(x, y, ww, hh, r, fill=None, outline=None, width=1):
        d.rounded_rectangle([int(x * S), int(y * S), int((x + ww) * S), int((y + hh) * S)],
                            radius=int(r * S), fill=fill, outline=outline,
                            width=max(1, int(width * S)))

    rr(0, 0, w, h, 24, fill=(21, 27, 47, 255), outline=(255, 255, 255, 70), width=1.4)
    rr(8, 12, w - 16, h - 20, 18, fill=(12, 17, 32, 255), outline=(255, 255, 255, 20))
    rr(w / 2 - 22, 5, 44, 5, 2.5, fill=(40, 48, 70, 255))

    d.text((18 * S, 26 * S), "我的项目", font=F(FONT_BOLD, 13), fill=WHITE)
    d.rounded_rectangle([112 * S, 26 * S, 160 * S, 44 * S], radius=9 * S, fill=(0, 208, 140, 40))
    d.ellipse([119 * S, 33 * S, 124 * S, 38 * S], fill=GREEN)
    d.text((128 * S, 30 * S), "进行中", font=F(FONT_BOLD, 8), fill=GREEN)

    rr(16, 56, w - 32, 58, 12, fill=(0, 208, 140, 26), outline=(126, 255, 186, 40))
    d.text((26 * S, 66 * S), "本周完成", font=F(FONT_REG, 8), fill=(150, 170, 195))
    d.text((26 * S, 80 * S), "12", font=F(FONT_DISPLAY, 24), fill=GREEN)
    d.text((48 * S, 94 * S), "个任务", font=F(FONT_REG, 8), fill=(150, 170, 195))

    base_y = 160
    rr(16, 120, w - 32, 52, 10, fill=(255, 255, 255, 10))
    for i, hh in enumerate((14, 22, 17, 27, 19, 31, 24)):
        x = 22 + i * 18
        d.rounded_rectangle([x * S, (base_y - hh) * S, (x + 10) * S, base_y * S],
                            radius=4 * S, fill=(255, 255, 255, 40) if i != 5 else GREEN)

    rows = [("记账本", "已完成", GREEN), ("数据看板", "进行中", (255, 197, 92)), ("设置页", "待开始", (130, 145, 175))]
    for i, (name, st, col) in enumerate(rows):
        y = 176 + i * 34
        rr(16, y, w - 32, 30, 9, fill=(255, 255, 255, 14))
        d.ellipse([24 * S, (y + 11) * S, 32 * S, (y + 19) * S], fill=col + (255,))
        d.text((38 * S, (y + 5) * S), name, font=F(FONT_BOLD, 9), fill=WHITE)
        d.text((38 * S, (y + 17) * S), st, font=F(FONT_REG, 7), fill=(140, 155, 180))

    for i in range(3):
        cx = w / 2 - 10 + i * 10
        d.ellipse([(cx - 1.6) * S, (h - 22) * S, (cx + 1.6) * S, (h - 19) * S],
                  fill=GREEN if i == 0 else (255, 255, 255, 60))
    return img
