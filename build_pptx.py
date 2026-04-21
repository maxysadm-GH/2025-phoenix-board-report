"""Generate PPTX by screenshotting the HTML slides with Playwright
and embedding them as full-bleed images. Pixel-perfect — the HTML is the design."""

from pptx import Presentation
from pptx.util import Inches, Emu
from playwright.sync_api import sync_playwright
from pathlib import Path
import sys

HERE = Path(__file__).parent
TMP = Path.home() / "tmp" / "pptx_shots"
TMP.mkdir(parents=True, exist_ok=True)

SLIDE_W_PX = 1280
# 16:9 widescreen
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def screenshot_slides(html_file, base_name):
    """Capture each .slide element as a high-res PNG."""
    url = f"file:///{html_file.as_posix()}?print=1"
    shots = []

    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page(
            viewport={"width": SLIDE_W_PX, "height": 2000},
            device_scale_factor=2,  # 2x for retina-quality
        )
        page.goto(url, wait_until="networkidle")
        page.wait_for_timeout(1500)  # Chart.js settle

        # Kill animations so elements are visible; hide hover tooltips
        page.evaluate("""() => {
            const s = document.createElement('style');
            s.textContent = `
                *, *::before, *::after {
                    animation: none !important;
                    transition: none !important;
                    opacity: 1 !important;
                    transform: none !important;
                }
                .tip::after, .tip::before {
                    display: none !important;
                }
            `;
            document.head.appendChild(s);
        }""")
        page.wait_for_timeout(300)

        # Get slide count and dimensions
        slide_data = page.evaluate("""() => {
            return [...document.querySelectorAll('.slide')].map((s, i) => ({
                idx: i,
                w: Math.round(s.getBoundingClientRect().width),
                h: Math.round(s.scrollHeight)
            }));
        }""")

        for sd in slide_data:
            i = sd['idx']
            # Isolate one slide
            page.evaluate("""(idx) => {
                document.querySelectorAll('.slide').forEach((s, j) => {
                    s.style.display = (j === idx) ? 'flex' : 'none';
                });
                const deck = document.querySelector('.deck');
                if (deck) { deck.style.gap = '0'; deck.style.padding = '0'; }
                document.body.style.background = '#fff';
            }""", i)
            page.wait_for_timeout(300)

            # Get the visible slide's bounding box
            box = page.evaluate("""(idx) => {
                const s = document.querySelectorAll('.slide')[idx];
                const r = s.getBoundingClientRect();
                return { x: r.x, y: r.y, width: r.width, height: r.height };
            }""", i)

            out_png = TMP / f"{base_name}_slide{i+1}.png"
            page.screenshot(
                path=str(out_png),
                clip={"x": box['x'], "y": box['y'],
                      "width": box['width'], "height": box['height']},
            )
            shots.append((out_png, box['width'], box['height']))
            print(f"  {out_png.name}: {box['width']}x{box['height']}")

        browser.close()
    return shots


def build_pptx(shots, out_path):
    """Create a PPTX with each screenshot as a full-width slide image."""
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    slide_w_emu = SLIDE_W
    slide_h_emu = SLIDE_H

    for png_path, w_px, h_px in shots:
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

        # Scale image to fill slide width, center vertically
        aspect = h_px / w_px
        img_w = slide_w_emu
        img_h = int(img_w * aspect)

        # If image is taller than slide, scale to fit height instead
        if img_h > int(slide_h_emu):
            img_h = slide_h_emu
            img_w = int(img_h / aspect)

        # Center
        left = (int(slide_w_emu) - int(img_w)) // 2
        top = (int(slide_h_emu) - int(img_h)) // 2

        slide.shapes.add_picture(
            str(png_path),
            Emu(left), Emu(top),
            Emu(int(img_w)), Emu(int(img_h)),
        )

    prs.save(str(out_path))
    print(f"  -> {out_path}")


def run(base):
    html_file = HERE / f"{base}.html"
    if not html_file.exists():
        print(f"  SKIP: {html_file} not found")
        return

    print(f"\n=== {base} ===")
    print("Screenshotting slides...")
    shots = screenshot_slides(html_file, base)

    print("Building PPTX...")
    out = HERE / f"{base}.pptx"
    build_pptx(shots, out)


if __name__ == '__main__':
    targets = sys.argv[1:] if len(sys.argv) > 1 else ['editorial', 'brand']
    for t in targets:
        run(t)
    print("\nDone.")
