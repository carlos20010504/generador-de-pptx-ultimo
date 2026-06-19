"""Convert a PPTX to per-slide PNGs using Microsoft Office COM (Windows-only).
Usage: python scripts/pptx_to_images.py <input.pptx> <output_dir> [width=1280]
"""
import os
import sys
from pathlib import Path
import comtypes.client


def export_slides(pptx_path: str, out_dir: str, width: int = 1280) -> int:
    pptx_abs = os.path.abspath(pptx_path)
    out_abs = os.path.abspath(out_dir)
    os.makedirs(out_abs, exist_ok=True)

    ppt = comtypes.client.CreateObject("PowerPoint.Application")
    ppt.Visible = 1
    deck = ppt.Presentations.Open(pptx_abs, ReadOnly=True, WithWindow=False)
    n = 0
    try:
        for i, slide in enumerate(deck.Slides, start=1):
            out_png = os.path.join(out_abs, f"slide-{i:02d}.png")
            slide.Export(out_png, "PNG", width, int(width * 9 / 16))
            n += 1
    finally:
        deck.Close()
        ppt.Quit()
    return n


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python scripts/pptx_to_images.py <input.pptx> <output_dir> [width]")
        sys.exit(1)
    width = int(sys.argv[3]) if len(sys.argv) > 3 else 1280
    n = export_slides(sys.argv[1], sys.argv[2], width)
    print(f"Exported {n} slides to {sys.argv[2]}")
