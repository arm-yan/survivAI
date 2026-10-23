r"""Export the MarginKite kite mark (the same shape as KITE in
products/seller-profit-tracker/previews/build_previews.py) as square avatars.

Usage (from the repo root):
    .venv\Scripts\python.exe products\brand\build_logo.py

Writes marginkite-logo-1024.png (mark on the cream brand background) and
marginkite-logo-transparent-1024.png next to this file, plus 40 px circle-crop
check copies in _check/ (git-ignored).
"""

from pathlib import Path

from PIL import Image, ImageDraw

HERE = Path(__file__).resolve().parent
SIZE = 1024
SS = 4                      # supersampling for smooth edges
TEAL = (15, 139, 125)       # BRAND["teal"] #0F8B7D
CREAM = (246, 243, 236)     # BRAND["cream"] #F6F3EC
WHITE = (255, 255, 255)
# Mark geometry in the 64-unit viewBox used by the preview script.
KITE = [(32, 4), (56, 26), (32, 60), (8, 26)]
SPARS = [((32, 4), (32, 60)), ((8, 26), (56, 26))]
STROKE = 3
MARK_HEIGHT = 0.62          # kite height as a share of the canvas (stays inside a circle crop)


def draw(background):
    big = SIZE * SS
    img = Image.new("RGBA", (big, big), background)
    d = ImageDraw.Draw(img)
    scale = big * MARK_HEIGHT / 56          # kite spans y 4..60
    off = big / 2 - 32 * scale              # mark centre (32, 32) to canvas centre

    def pt(p):
        return (off + p[0] * scale, off + p[1] * scale)

    d.polygon([pt(p) for p in KITE], fill=TEAL)
    # Clip the white spars to the kite, as the SVG strokes appear inside it.
    spars = Image.new("L", (big, big), 0)
    sd = ImageDraw.Draw(spars)
    for a, b in SPARS:
        sd.line([pt(a), pt(b)], fill=255, width=round(STROKE * scale))
    mask = Image.new("L", (big, big), 0)
    ImageDraw.Draw(mask).polygon([pt(p) for p in KITE], fill=255)
    spars = Image.composite(spars, Image.new("L", (big, big), 0), mask)
    img.paste(Image.new("RGBA", (big, big), WHITE + (255,)), (0, 0), spars)
    return img.resize((SIZE, SIZE), Image.LANCZOS)


def circle_check(img, px=40):
    small = img.resize((px, px), Image.LANCZOS)
    m = Image.new("L", (px * 4, px * 4), 0)
    ImageDraw.Draw(m).ellipse((0, 0, px * 4 - 1, px * 4 - 1), fill=255)
    out = Image.new("RGBA", (px, px), (255, 255, 255, 0))
    out.paste(small, (0, 0), m.resize((px, px), Image.LANCZOS))
    return out


def main():
    check = HERE / "_check"
    check.mkdir(exist_ok=True)
    for name, bg in (("marginkite-logo-1024.png", CREAM + (255,)),
                     ("marginkite-logo-transparent-1024.png", (0, 0, 0, 0))):
        img = draw(bg)
        if bg[3] == 255:
            img = img.convert("RGB")
        img.save(HERE / name, optimize=True)
        circle_check(img.convert("RGBA")).save(check / name.replace("1024", "40-circle"))
        print(f"Built {HERE / name}")


if __name__ == "__main__":
    main()
