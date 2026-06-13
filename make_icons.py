"""Generate Thusa PWA icons: navy tile, orange chat bubble, clock at 08:00."""
import math
from PIL import Image, ImageDraw

NAVY = (13, 27, 42, 255)
ORANGE = (244, 166, 35, 255)
WHITE = (255, 255, 255, 255)


def make(size, maskable=False):
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    if maskable:
        d.rectangle([0, 0, size, size], fill=NAVY)
        pad = size * 0.20
    else:
        d.rounded_rectangle([0, 0, size - 1, size - 1], radius=size * 0.22, fill=NAVY)
        pad = size * 0.15

    bx0, by0 = pad, pad * 1.05
    bx1, by1 = size - pad, size - pad * 1.45
    d.rounded_rectangle([bx0, by0, bx1, by1], radius=(by1 - by0) * 0.28, fill=ORANGE)
    tail_x = bx0 + (bx1 - bx0) * 0.22
    d.polygon([(tail_x, by1 - 2), (tail_x + size * 0.11, by1 - 2),
               (tail_x - size * 0.02, by1 + size * 0.10)], fill=ORANGE)

    cx, cy = (bx0 + bx1) / 2, (by0 + by1) / 2
    cr = (by1 - by0) * 0.34
    d.ellipse([cx - cr, cy - cr, cx + cr, cy + cr], fill=NAVY)
    lw = max(3, int(size * 0.028))
    # minute hand at 12, hour hand at 8 (08:00 — early-morning meetings)
    d.line([cx, cy, cx, cy - cr * 0.62], fill=WHITE, width=lw)
    a = math.radians(240)
    d.line([cx, cy, cx + math.sin(a) * cr * 0.45, cy - math.cos(a) * cr * 0.45],
           fill=WHITE, width=lw)
    return img


make(192).save('icons/icon-192.png')
make(512).save('icons/icon-512.png')
make(512, maskable=True).save('icons/icon-maskable-512.png')
print('icons done')
