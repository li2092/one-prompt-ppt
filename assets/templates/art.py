"""生成模板用的底图。只用 Pillow，结果由种子决定，重跑得到同样的图。"""
import math
import os
import random

from PIL import Image, ImageChops, ImageDraw, ImageFilter

HERE = os.path.dirname(os.path.abspath(__file__))
ART = os.path.join(HERE, "art")
S = 2  # 先按两倍画，再缩小，线条更平滑


def _gradient(size, c1, c2, angle=0):
    w, h = size
    d = int(math.hypot(w, h)) + 4
    g = Image.linear_gradient("L").resize((d, d)).rotate(angle, resample=Image.BICUBIC)
    g = g.crop(((d - w) // 2, (d - h) // 2, (d - w) // 2 + w, (d - h) // 2 + h))
    return Image.composite(Image.new("RGB", size, c2), Image.new("RGB", size, c1), g)


def _glow(img, center, radius, color, alpha):
    r = int(radius)
    m = ImageChops.invert(Image.radial_gradient("L")).resize((2 * r, 2 * r))
    m = m.point(lambda v: int(max(0, v - 78) * 255 / 177 * alpha))   # 到内切圆边缘就降到零，不留方框的边
    layer = Image.new("RGB", img.size, color)
    mask = Image.new("L", img.size, 0)
    mask.paste(m, (int(center[0]) - r, int(center[1]) - r))
    return Image.composite(layer, img, mask)


def _bezier(p0, p1, p2, n=70):
    out = []
    for i in range(n + 1):
        t = i / n
        x = (1 - t) ** 2 * p0[0] + 2 * (1 - t) * t * p1[0] + t ** 2 * p2[0]
        y = (1 - t) ** 2 * p0[1] + 2 * (1 - t) * t * p1[1] + t ** 2 * p2[1]
        out.append((x, y))
    return out


def _arcs(size, seed, hubs, n, palette, width=(2, 4), spread=(0.18, 0.95), sector=(150, 400)):
    """从几个枢纽点向外画弧线。返回 (线条层, 节点层)，都是透明底。"""
    w, h = size
    rnd = random.Random(seed)
    lines = Image.new("RGBA", size, (0, 0, 0, 0))
    nodes = Image.new("RGBA", size, (0, 0, 0, 0))
    dl, dn = ImageDraw.Draw(lines), ImageDraw.Draw(nodes)
    for i in range(n):
        hub = hubs[i % len(hubs)]
        hx, hy = hub[0] * w, hub[1] * h
        ang = math.radians(rnd.uniform(*sector))
        dist = rnd.uniform(*spread) * w
        dx, dy = hx + math.cos(ang) * dist, hy + math.sin(ang) * dist * 0.62
        mx, my = (hx + dx) / 2, (hy + dy) / 2
        lift = dist * rnd.uniform(0.22, 0.36)
        ctrl = (mx, my - lift)
        col, a = rnd.choice(palette)
        dl.line(_bezier((hx, hy), ctrl, (dx, dy)), fill=col + (a,), width=rnd.randint(*width) * S // 2, joint="curve")
        r = rnd.uniform(3.5, 7) * S
        dn.ellipse((dx - r, dy - r, dx + r, dy + r), fill=col + (min(255, a + 70),))
    for hub in hubs:
        hx, hy = hub[0] * w, hub[1] * h
        for k, r in enumerate((9, 22, 40)):
            r *= S
            if k == 0:
                dn.ellipse((hx - r, hy - r, hx + r, hy + r), fill=(255, 255, 255, 255))
            else:
                dn.ellipse((hx - r, hy - r, hx + r, hy + r), outline=(255, 255, 255, 120 - 40 * k), width=S)
    return lines, nodes


def _dots(size, step, color, alpha):
    layer = Image.new("RGBA", size, (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    r = 1.6 * S
    for y in range(step // 2, size[1], step):
        for x in range(step // 2, size[0], step):
            d.ellipse((x - r, y - r, x + r, y + r), fill=color + (alpha,))
    return layer


def _fade(img, color, start, end, alpha, vertical=False):
    """从 start 到 end（占宽或高的比例）把 color 由浓到淡盖上去，给文字留出干净的区域。"""
    w, h = img.size
    mask = Image.new("L", (w, h), 0)
    d = ImageDraw.Draw(mask)
    span = w if not vertical else h
    a, b = int(start * span), int(end * span)
    step = 1 if b > a else -1
    for i, p in enumerate(range(a, b, step)):
        v = int(alpha * (1 - i / max(1, abs(b - a))))
        if vertical:
            d.line((0, p, w, p), fill=v)
        else:
            d.line((p, 0, p, h), fill=v)
    lo, hi = min(a, b), max(a, b)
    if step == 1 and lo > 0:
        d.rectangle((0, 0, lo, h) if not vertical else (0, 0, w, lo), fill=alpha)
    if step == -1 and hi < span:
        d.rectangle((hi, 0, w, h) if not vertical else (0, hi, w, h), fill=alpha)
    return Image.composite(Image.new("RGB", img.size, color), img, mask)


def _save(img, name, size, quality=90):
    os.makedirs(ART, exist_ok=True)
    path = os.path.join(ART, name)
    img = img.resize(size, Image.LANCZOS)
    if name.endswith(".png"):
        img.save(path, optimize=True)
    else:
        img.convert("RGB").save(path, quality=quality)
    return path


# ------------------------------------------------------------------ 航线
NAVY1, NAVY2 = (6, 14, 34), (10, 44, 110)
ROUTE = [((255, 255, 255), 95), ((255, 255, 255), 60), ((120, 200, 255), 120), ((34, 211, 238), 150), ((255, 107, 53), 220)]


def routes_dark(name, seed, hubs, n=48, size=(1920, 1080), fade=(0.0, 0.62, 215), sector=(150, 400)):
    big = (size[0] * S, size[1] * S)
    img = _gradient(big, NAVY1, NAVY2, 28)
    for hub in hubs:
        img = _glow(img, (hub[0] * big[0], hub[1] * big[1]), 0.5 * big[0], (20, 99, 255), 0.42)
    img = Image.alpha_composite(img.convert("RGBA"), _dots(big, 34 * S, (255, 255, 255), 16))
    lines, nodes = _arcs(big, seed, hubs, n, ROUTE, sector=sector)
    img = Image.alpha_composite(img, lines.filter(ImageFilter.GaussianBlur(9 * S)))
    img = Image.alpha_composite(img, lines)
    img = Image.alpha_composite(img, nodes.filter(ImageFilter.GaussianBlur(7 * S)))
    img = Image.alpha_composite(img, nodes).convert("RGB")
    if fade:
        img = _fade(img, NAVY1, fade[0], fade[1], fade[2])
    return _save(img, name, size)


def routes_light(name, seed, size=(1100, 460)):
    """正文页右上角用的淡色弧线，透明底，向左下逐渐消失。"""
    big = (size[0] * S, size[1] * S)
    pal = [((20, 99, 255), 70), ((20, 99, 255), 45), ((34, 170, 220), 60), ((255, 107, 53), 120)]
    lines, nodes = _arcs(big, seed, [(1.02, -0.1)], 16, pal, width=(2, 3), spread=(0.25, 1.0), sector=(95, 200))
    img = Image.alpha_composite(lines, nodes)
    mask = Image.linear_gradient("L").rotate(90, expand=True).resize(big)       # 左暗右亮
    mask = ImageChops.multiply(mask, ImageChops.invert(Image.linear_gradient("L").resize(big)))  # 上亮下暗
    img.putalpha(ImageChops.multiply(img.getchannel("A"), mask.point(lambda v: min(255, int(v * 1.5)))))
    return _save(img, name, size)


# ------------------------------------------------------------------ 极光
def mesh(name, seed, base, blobs, size=(1920, 1080), blur=150, grain=0.035, fade=None):
    big = (size[0] * S, size[1] * S)
    img = Image.new("RGB", big, base)
    for (x, y, r, color, alpha) in blobs:
        layer = Image.new("RGBA", big, color + (0,))     # 透明处也用同一个颜色，模糊后边缘不发灰
        d = ImageDraw.Draw(layer)
        cx, cy, rr = x * big[0], y * big[1], r * big[0]
        d.ellipse((cx - rr, cy - rr * 0.8, cx + rr, cy + rr * 0.8), fill=color + (alpha,))
        img = Image.alpha_composite(img.convert("RGBA"), layer.filter(ImageFilter.GaussianBlur(blur * S))).convert("RGB")
    if grain:
        noise = Image.effect_noise(big, 26).convert("RGB")
        img = Image.blend(img, noise, grain)
    if fade:
        img = _fade(img, fade[3], fade[0], fade[1], fade[2])
    return _save(img, name, size)


DEEP = (30, 24, 92)
DEEP_BLOBS = [(0.10, 0.15, 0.42, (79, 70, 229), 255), (0.62, 0.30, 0.36, (124, 58, 237), 235),
              (0.92, 0.88, 0.36, (255, 122, 89), 235), (0.98, 0.05, 0.22, (34, 211, 238), 190),
              (0.30, 0.98, 0.34, (37, 30, 120), 255)]
DEEP_BLOBS2 = [(0.85, 0.15, 0.40, (124, 58, 237), 255), (0.20, 0.25, 0.34, (79, 70, 229), 240),
               (0.10, 0.98, 0.34, (255, 122, 89), 225), (0.70, 0.95, 0.30, (34, 211, 238), 150)]
LIGHT = (246, 247, 251)
LIGHT_BLOBS = [(0.95, 0.02, 0.30, (150, 160, 255), 170), (0.70, -0.10, 0.24, (205, 175, 255), 150),
               (1.03, 0.45, 0.20, (255, 190, 165), 150), (0.00, 1.08, 0.26, (160, 225, 250), 120)]


def build_all():
    out = {}
    out["routes-cover"] = routes_dark("routes-cover.jpg", 7, [(0.74, 0.63)])
    out["routes-section"] = routes_dark("routes-section.jpg", 21, [(0.86, 0.86)], n=40, fade=(0.0, 0.5, 200), sector=(170, 300))
    out["routes-closing"] = routes_dark("routes-closing.jpg", 33, [(0.80, 0.30), (0.62, 0.82)], n=44)
    out["routes-side"] = routes_dark("routes-side.jpg", 5, [(0.5, 0.78)], n=30, size=(662, 1080), fade=None, sector=(190, 350))
    out["routes-corner"] = routes_light("routes-corner.png", 11)
    out["mesh-cover"] = mesh("mesh-cover.jpg", 1, DEEP, DEEP_BLOBS)
    out["mesh-section"] = mesh("mesh-section.jpg", 2, DEEP, DEEP_BLOBS2)
    out["mesh-light"] = mesh("mesh-light.jpg", 3, LIGHT, LIGHT_BLOBS, blur=170, grain=0.015)
    return out


if __name__ == "__main__":
    for k, v in build_all().items():
        print(k, v)
