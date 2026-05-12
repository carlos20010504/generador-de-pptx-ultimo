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


# PowerPoint's SaveAs format constant for PDF.
_PP_SAVE_AS_PDF = 32


def export_to_pdf(pptx_path: str, pdf_path: str) -> dict:
    """Export a PPTX to PDF via PowerPoint COM (Windows-only). Returns
    {"ok": bool, "path": str|None, "error": str|None}. Same fail-soft
    contract as `generate_previews` so callers don't need try/except."""
    if not Path(pptx_path).is_file():
        return {"ok": False, "path": None,
                "error": "PPTX no encontrado."}
    if sys.platform != "win32":
        return {"ok": False, "path": None,
                "error": "Export a PDF no disponible en este sistema (requiere Windows + PowerPoint)."}

    pptx_abs = os.path.abspath(pptx_path)
    pdf_abs = os.path.abspath(pdf_path)
    os.makedirs(os.path.dirname(pdf_abs) or ".", exist_ok=True)

    try:
        import comtypes.client  # type: ignore[import-untyped]
    except ImportError as e:
        return {"ok": False, "path": None,
                "error": f"Falta dependencia para PDF: {e}"}

    ppt = None
    deck = None
    try:
        ppt = comtypes.client.CreateObject("PowerPoint.Application")
        ppt.Visible = 1
        deck = ppt.Presentations.Open(
            pptx_abs, ReadOnly=True, Untitled=False, WithWindow=False,
        )
        # Fallback first to ExportAsFixedFormat (richer, supports notes etc.);
        # if not available on this PowerPoint build, use SaveAs(PDF).
        try:
            deck.ExportAsFixedFormat(pdf_abs, 2)  # 2 == ppFixedFormatTypePDF
        except Exception:
            deck.SaveAs(pdf_abs, _PP_SAVE_AS_PDF)
    except Exception as e:
        return {"ok": False, "path": None,
                "error": f"PowerPoint no pudo exportar a PDF: {str(e)[:200]}"}
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

    if not Path(pdf_abs).is_file():
        return {"ok": False, "path": None,
                "error": "PDF no se generó (PowerPoint terminó sin escribir el archivo)."}
    return {"ok": True, "path": pdf_abs, "error": None}


def cmd_export_pdf(args) -> None:
    """CLI entry: emit PDF export status JSON to stdout."""
    result = export_to_pdf(args.input, args.output)
    sys.stdout.write(json.dumps(result, ensure_ascii=True))
