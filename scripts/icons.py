"""Draw the ribbon menu icons as PNGs (standard library only).

Each icon is a 64x64 picture of what the button inserts. Shapes are
rasterised with 4x4 supersampling for smooth edges. Colours are fixed here
(the real shapes use the deck's theme colour); BLUE reads well in both light
and dark ribbons.
"""
import math
import struct
import zlib

SIZE = 64
SS = 4  # supersampling per axis

BLUE = (47, 111, 191)
GREY = (205, 210, 217)
WHITE = (255, 255, 255)
GREEN = (46, 158, 91)
AMBER = (240, 162, 2)
RED = (214, 69, 69)
SLATE = (163, 170, 181)

# ---------------------------------------------------------------- shapes
# Each shape is a function (x, y) -> bool on a 0..64 grid.


def circle(cx, cy, r):
    return lambda x, y: (x - cx) ** 2 + (y - cy) ** 2 <= r * r


def pie(cx, cy, r, pct):
    """Sector from 12 o'clock, clockwise, covering pct percent."""
    limit = 360 * pct / 100

    def inside(x, y):
        if (x - cx) ** 2 + (y - cy) ** 2 > r * r:
            return False
        ang = math.degrees(math.atan2(x - cx, -(y - cy))) % 360
        return ang <= limit
    return inside


def rect(x0, y0, x1, y1):
    return lambda x, y: x0 <= x <= x1 and y0 <= y <= y1


def stroke(x1, y1, x2, y2, w):
    """Line segment with round ends, width w."""
    dx, dy = x2 - x1, y2 - y1
    length2 = dx * dx + dy * dy

    def inside(x, y):
        t = max(0.0, min(1.0, ((x - x1) * dx + (y - y1) * dy) / length2)) if length2 else 0.0
        px, py = x1 + t * dx, y1 + t * dy
        return (x - px) ** 2 + (y - py) ** 2 <= (w / 2) ** 2
    return inside


def polyline(points, w):
    parts = [stroke(*points[i], *points[i + 1], w) for i in range(len(points) - 1)]
    return lambda x, y: any(p(x, y) for p in parts)


# ---------------------------------------------------------------- render


def render(layers):
    """layers: list of (shape_fn, rgb). Returns PNG bytes."""
    px = [[(0.0, 0.0, 0.0, 0.0)] * SIZE for _ in range(SIZE)]
    step = 1.0 / SS
    for fn, rgb in layers:
        for j in range(SIZE):
            row = px[j]
            for i in range(SIZE):
                hits = 0
                for sj in range(SS):
                    y = j + (sj + 0.5) * step
                    for si in range(SS):
                        if fn(i + (si + 0.5) * step, y):
                            hits += 1
                if not hits:
                    continue
                a = hits / (SS * SS)
                r0, g0, b0, a0 = row[i]
                out_a = a + a0 * (1 - a)
                mix = lambda c, c0: (c * a + c0 * a0 * (1 - a)) / out_a
                row[i] = (mix(rgb[0], r0), mix(rgb[1], g0), mix(rgb[2], b0), out_a)
    raw = bytearray()
    for row in px:
        raw.append(0)
        for r, g, b, a in row:
            raw += bytes((round(r), round(g), round(b), round(a * 255)))
    return _png(bytes(raw))


def _png(raw):
    def chunk(kind, data):
        c = kind + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)
    header = struct.pack(">IIBBBBB", SIZE, SIZE, 8, 6, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header)
            + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b""))

# ---------------------------------------------------------------- icons

C, R = 32, 28  # centre and radius for round icons


def harvey(pct):
    layers = [(circle(C, C, R), GREY)]
    if pct >= 100:
        layers.append((circle(C, C, R), BLUE))
    elif pct > 0:
        layers.append((pie(C, C, R, pct), BLUE))
    return render(layers)


def meter(level):
    layers = []
    for n in range(1, 4):
        h = 56 * (n + 1) / 4
        x0 = 6 + (n - 1) * 19
        layers.append((rect(x0, 60 - h, x0 + 14, 60), BLUE if n <= level else GREY))
    return render(layers)


TICK = [(17, 33), (27, 43), (46, 22)]


def status(kind):
    if kind == "complete":
        return render([(circle(C, C, R), GREEN), (polyline(TICK, 6.5), WHITE)])
    if kind == "inprogress":
        dots = []
        for (x1, y1), (x2, y2), n, first in ((TICK[0], TICK[1], 2, 0), (TICK[1], TICK[2], 3, 1)):
            for k in range(first, n + 1):
                f = k / n
                dots.append(circle(x1 + (x2 - x1) * f, y1 + (y2 - y1) * f, 3.6))
        return render([(circle(C, C, R), AMBER), (lambda x, y: any(d(x, y) for d in dots), WHITE)])
    if kind == "delayed":
        return render([(circle(C, C, R), RED),
                       (stroke(21, 21, 43, 43, 6.5), WHITE), (stroke(43, 21, 21, 43, 6.5), WHITE)])
    return render([(circle(C, C, R), SLATE), (stroke(20, 32, 44, 32, 6.5), WHITE)])


def tracker(kind):
    glyphs = {
        "number": [stroke(27, 22, 33, 18, 5), stroke(33, 18, 33, 46, 5)],                  # 1
        "letter": [stroke(22, 46, 32, 18, 5), stroke(32, 18, 42, 46, 5),
                   stroke(26, 37, 38, 37, 4.5)],                                             # A
        "custom": [circle(21, 32, 3.6), circle(32, 32, 3.6), circle(43, 32, 3.6)],           # ...
    }[kind]
    return render([(circle(C, C, R), BLUE), (lambda x, y: any(g(x, y) for g in glyphs), WHITE)])


def all_icons():
    """name -> PNG bytes. Names are used as image="..." in customUI14.xml."""
    icons = {}
    for pct in (0, 25, 50, 75, 100):
        icons[f"stHv{pct}"] = harvey(pct)
    for level, name in ((1, "Low"), (2, "Medium"), (3, "High")):
        icons[f"stMeter{name}"] = meter(level)
    for kind, name in (("complete", "Complete"), ("inprogress", "InProgress"),
                       ("delayed", "Delayed"), ("notstarted", "NotStarted")):
        icons[f"stStatus{name}"] = status(kind)
    for kind, name in (("number", "Number"), ("letter", "Letter"), ("custom", "Custom")):
        icons[f"stTracker{name}"] = tracker(kind)
    return icons


if __name__ == "__main__":
    import sys
    from pathlib import Path
    out = Path(sys.argv[1] if len(sys.argv) > 1 else "icons-preview")
    out.mkdir(exist_ok=True)
    for name, data in all_icons().items():
        (out / f"{name}.png").write_bytes(data)
    print(f"wrote {len(all_icons())} icons to {out}")
