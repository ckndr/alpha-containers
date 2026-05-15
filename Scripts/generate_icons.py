"""
generate_icons.py
Alpha Containers — App Icon Generator
--------------------------------------
Run this ONCE to create proper PWA icons from your company logo.
Place this file in your AlphaContainers project folder and double-click,
or run: python generate_icons.py

Requires: pip install Pillow requests
"""

import os, sys, struct, zlib

# Scripts live in AlphaContainers/Scripts/ — icons go to root (where HTML is)
DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ── Try to use Pillow for high-quality icons ──────────────────
def try_pillow():
    try:
        from PIL import Image, ImageDraw, ImageFont
        import urllib.request
        return Image, ImageDraw, ImageFont, urllib.request
    except ImportError:
        return None

PIL = try_pillow()

if PIL is None:
    print("Installing Pillow...")
    os.system(f'"{sys.executable}" -m pip install Pillow --quiet')
    PIL = try_pillow()

if PIL is None:
    print("ERROR: Could not install Pillow. Run manually: pip install Pillow")
    input("Press Enter to exit...")
    sys.exit(1)

Image, ImageDraw, ImageFont, urllib_request = PIL

# ── Download company logo ─────────────────────────────────────
LOGO_URL = "http://www.alphacontainer.com.pk/images/abt_btn.gif"
LOGO_PATH = os.path.join(DIR, "_temp_logo.gif")

logo_img = None
print(f"Downloading logo from {LOGO_URL}...")
try:
    urllib_request.urlretrieve(LOGO_URL, LOGO_PATH)
    logo_img = Image.open(LOGO_PATH).convert("RGBA")
    print(f"  Logo downloaded: {logo_img.size[0]}x{logo_img.size[1]} px")
except Exception as e:
    print(f"  Could not download logo: {e}")
    print("  Will use text-based AC logo instead.")

def make_icon(size, logo=None):
    """Create a square icon at given size."""
    img = Image.new("RGBA", (size, size), (13, 31, 60, 255))  # navy background
    draw = ImageDraw.Draw(img)

    if logo:
        # Paste logo centred, scaled to 80% of icon
        target = int(size * 0.8)
        logo_resized = logo.copy()
        logo_resized.thumbnail((target, target), Image.LANCZOS)
        lw, lh = logo_resized.size
        x = (size - lw) // 2
        y = (size - lh) // 2
        img.paste(logo_resized, (x, y), logo_resized)
    else:
        # Draw "AC" text in gold on navy background
        gold = (232, 160, 32)
        # Rounded rectangle background
        pad = size // 8
        draw.rounded_rectangle([pad, pad, size-pad, size-pad],
                                radius=size//8, fill=(30, 47, 82))
        # Text
        font_size = size // 3
        try:
            # Try to load a system font
            for font_path in [
                "C:/Windows/Fonts/arialbd.ttf",
                "C:/Windows/Fonts/calibrib.ttf",
                "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            ]:
                if os.path.exists(font_path):
                    font = ImageFont.truetype(font_path, font_size)
                    break
            else:
                font = ImageFont.load_default()
        except Exception:
            font = ImageFont.load_default()

        text = "AC"
        bbox = draw.textbbox((0, 0), text, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        tx = (size - tw) // 2 - bbox[0]
        ty = (size - th) // 2 - bbox[1]
        draw.text((tx, ty), text, fill=gold, font=font)

    return img.convert("RGB")

# ── Generate icons ────────────────────────────────────────────
for size, filename in [(192, "icon-192.png"), (512, "icon-512.png")]:
    icon = make_icon(size, logo_img)
    path = os.path.join(DIR, filename)
    icon.save(path, "PNG", optimize=True)
    print(f"  Saved {filename} ({size}x{size})")

# Cleanup temp file
if os.path.exists(LOGO_PATH):
    os.remove(LOGO_PATH)

print()
print("=" * 50)
print("  Icons generated successfully!")
print("  Now run 'Update App HTML.bat' to push to GitHub.")
print("=" * 50)
print()
input("Press Enter to close...")
