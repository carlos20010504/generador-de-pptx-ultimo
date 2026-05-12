"""PPTX → PNG slide preview generation.

Strategy: best-effort. If we can render true previews via PowerPoint COM
(Windows-only, requires Microsoft Office installed), do so. Otherwise we
return a soft failure so the UI can still let the user download — the
preview is a nice-to-have, not a blocker.

Public entry-points:
- `generate_previews(pptx_path, out_dir, width=1280)` → returns dict
- `cmd_preview_pptx(args)` → CLI handler used by `python -m socya_pipeline preview-pptx`
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import List

# Width in pixels for the rendered slide. 1280 keeps the PNGs small enough
# to ship over the wire while still readable.
DEFAULT_WIDTH = 1280


def _render_with_powerpoint(pptx_path: str, out_dir: str, width: int
                              ) -> List[str]:
    """Use Microsoft PowerPoint via COM (Windows + Office only) to export
    each slide as a PNG. Raises on failure. Returns list of PNG paths."""
    import comtypes.client  # type: ignore[import-untyped]

    pptx_abs = os.path.abspath(pptx_path)
    out_abs = os.path.abspath(out_dir)
    os.makedirs(out_abs, exist_ok=True)

    ppt = None
    deck = None
    paths: List[str] = []
    try:
        ppt = comtypes.client.CreateObject("PowerPoint.Application")
        # Visible=False is unsupported on some PPT builds; we set Visible=1
        # then immediately export and close — never shows a window thanks to
        # WithWindow=False.
        ppt.Visible = 1
        deck = ppt.Presentations.Open(
            pptx_abs, ReadOnly=True, Untitled=False, WithWindow=False,
        )
        for i, slide in enumerate(deck.Slides, start=1):
            out_png = os.path.join(out_abs, f"slide-{i:02d}.png")
            slide.Export(out_png, "PNG", width, int(width * 9 / 16))
            paths.append(out_png)
    finally:
        try:
            if deck is not None:
                deck.Close()
        except Exception:
            pass
        try:
            if ppt is not None:
                ppt.Quit()
        except Exception:
            pass
    return paths


def generate_previews(pptx_path: str, out_dir: str,
                       width: int = DEFAULT_WIDTH) -> dict:
    """Try to render PNG previews of every slide in `pptx_path`.

    Returns a stable dict shape so callers can react without try/except:
        {
          "ok": bool,
          "renderer": "powerpoint" | None,
          "slides": [{"index": int, "filename": str}, ...],
          "error": str | None,
        }
    """
    if not Path(pptx_path).is_file():
        return {"ok": False, "renderer": None, "slides": [],
                "error": "PPTX no encontrado."}

    if sys.platform != "win32":
        return {"ok": False, "renderer": None, "slides": [],
                "error": "Preview no disponible en este sistema (requiere Windows + PowerPoint)."}
    try:
        paths = _render_with_powerpoint(pptx_path, out_dir, width)
        return {
            "ok": True,
            "renderer": "powerpoint",
            "slides": [
                {"index": i, "filename": os.path.basename(p)}
                for i, p in enumerate(paths)
            ],
            "error": None,
        }
    except ImportError as e:
        return {"ok": False, "renderer": None, "slides": [],
                "error": f"Falta dependencia para preview: {e}"}
    except Exception as e:
        # PowerPoint can fail for many reasons (not installed, license,
        # locked file, etc.). Always return a soft error so the caller
        # can still serve the download.
        return {"ok": False, "renderer": None, "slides": [],
                "error": f"PowerPoint no pudo renderizar el preview: {str(e)[:200]}"}


def cmd_preview_pptx(args) -> None:
    """CLI entry: emit preview status JSON to stdout."""
    width = getattr(args, "width", None) or DEFAULT_WIDTH
    result = generate_previews(args.input, args.output_dir, width=width)
    sys.stdout.write(json.dumps(result, ensure_ascii=True))
