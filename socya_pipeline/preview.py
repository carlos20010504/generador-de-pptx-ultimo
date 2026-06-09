"""PPTX → PNG slide preview generation (+ PDF export).

Strategy: best-effort, platform-dispatched.
- Windows: PowerPoint COM (pixel-perfect, requires Microsoft Office).
- Linux/macOS: headless LibreOffice (`soffice`) + poppler (`pdf2image`) —
  this is what runs in the Docker/Render deployment.
Either way, failure returns a soft error so the UI can still let the user
download — the preview is a nice-to-have, not a blocker.

Public entry-points:
- `generate_previews(pptx_path, out_dir, width=1280)` → returns dict
- `export_to_pdf(pptx_path, pdf_path)` → returns dict
- `cmd_preview_pptx(args)` / `cmd_export_pdf(args)` → CLI handlers used by
  `python -m socya_pipeline {preview-pptx,export-pdf}`
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path
from typing import List, Optional

# Width in pixels for the rendered slide. 1280 keeps the PNGs small enough
# to ship over the wire while still readable.
DEFAULT_WIDTH = 1280

# LibreOffice can take a while on the first conversion (it cold-starts a
# headless soffice). 180s is generous but bounded so a wedged soffice can't
# hang the request forever.
_SOFFICE_TIMEOUT_S = 180


def _find_soffice() -> Optional[str]:
    """Locate the LibreOffice binary. Returns its path or None if absent.

    On the Debian-based Docker image `soffice` is on PATH (libreoffice-impress);
    we also check `libreoffice` and the canonical absolute paths as a fallback
    so a thin PATH in some container shells still resolves it."""
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for candidate in ("/usr/bin/soffice", "/usr/bin/libreoffice"):
        if Path(candidate).is_file():
            return candidate
    return None


def _soffice_convert_to_pdf(soffice: str, pptx_abs: str, out_dir: str) -> str:
    """Convert a PPTX to PDF with headless LibreOffice into `out_dir`.

    Returns the produced PDF path. Raises RuntimeError on failure.

    A per-call `UserInstallation` profile isolates each conversion so
    concurrent requests don't collide on LibreOffice's single-instance lock
    (without it the 2nd concurrent soffice silently no-ops)."""
    profile = Path(tempfile.gettempdir()) / f"lo_profile_{os.getpid()}_{uuid.uuid4().hex}"
    cmd = [
        soffice,
        f"-env:UserInstallation=file://{profile.as_posix()}",
        "--headless", "--norestore", "--nologo", "--nofirststartwizard",
        "--convert-to", "pdf:impress_pdf_Export",
        "--outdir", out_dir,
        pptx_abs,
    ]
    try:
        proc = subprocess.run(
            cmd, capture_output=True, text=True, timeout=_SOFFICE_TIMEOUT_S,
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError("LibreOffice tardó demasiado en convertir el PPTX.")
    finally:
        shutil.rmtree(profile, ignore_errors=True)

    produced = Path(out_dir) / (Path(pptx_abs).stem + ".pdf")
    if not produced.is_file():
        detail = (proc.stderr or proc.stdout or "").strip()[:200]
        raise RuntimeError(
            f"LibreOffice no generó el PDF (code={proc.returncode}). {detail}"
        )
    return str(produced)


def _render_with_libreoffice(pptx_path: str, out_dir: str, width: int
                              ) -> List[str]:
    """Use headless LibreOffice + poppler to export each slide as a PNG on
    Linux/macOS. Raises on failure. Returns list of PNG paths.

    Strategy: PPTX -> PDF (one soffice call) -> rasterize each PDF page to PNG.
    We go via PDF because `soffice --convert-to png` only exports the first
    slide, whereas poppler renders every page."""
    soffice = _find_soffice()
    if soffice is None:
        raise RuntimeError("LibreOffice (soffice) no está instalado en el sistema.")
    # Lazy import: pdf2image needs poppler at runtime (present in the Docker
    # image, absent on the Windows dev box where the PowerPoint path is used).
    from pdf2image import convert_from_path  # type: ignore[import-untyped]

    pptx_abs = os.path.abspath(pptx_path)
    out_abs = os.path.abspath(out_dir)
    os.makedirs(out_abs, exist_ok=True)

    paths: List[str] = []
    with tempfile.TemporaryDirectory(prefix="socya_pdf_") as pdf_tmp:
        pdf_path = _soffice_convert_to_pdf(soffice, pptx_abs, pdf_tmp)
        # size=(width, None) sets the pixel width and keeps the aspect ratio.
        images = convert_from_path(pdf_path, size=(width, None))
        for i, img in enumerate(images, start=1):
            out_png = os.path.join(out_abs, f"slide-{i:02d}.png")
            img.save(out_png, "PNG")
            paths.append(out_png)
    return paths


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

    # Windows local uses PowerPoint COM (pixel-perfect); Linux/Docker uses
    # headless LibreOffice + poppler. Both raise on failure and are wrapped
    # below in the same fail-soft contract so the caller can still serve the
    # download even if previews can't be produced.
    if sys.platform == "win32":
        renderer, render = "powerpoint", _render_with_powerpoint
    else:
        renderer, render = "libreoffice", _render_with_libreoffice
    try:
        paths = render(pptx_path, out_dir, width)
        # Verifica que cada PNG efectivamente exista en disco. PowerPoint COM
        # a veces reporta éxito antes de que Windows termine de flush'ear el
        # archivo — sin este check el route /api/pptx-preview servía 404 y
        # el frontend mostraba spinner infinito.
        verified = []
        for i, p in enumerate(paths):
            try:
                if Path(p).is_file() and Path(p).stat().st_size > 100:
                    verified.append({"index": i, "filename": os.path.basename(p)})
            except OSError:
                continue
        return {
            "ok": len(verified) > 0,
            "renderer": renderer,
            "slides": verified,
            "error": (None if verified
                       else f"{renderer} exportó pero no se encontraron PNGs válidos en disco."),
        }
    except ImportError as e:
        return {"ok": False, "renderer": None, "slides": [],
                "error": f"Falta dependencia para preview: {e}"}
    except Exception as e:
        # The renderer can fail for many reasons (not installed, license,
        # locked file, soffice timeout, etc.). Always return a soft error so
        # the caller can still serve the download.
        return {"ok": False, "renderer": None, "slides": [],
                "error": f"{renderer} no pudo renderizar el preview: {str(e)[:200]}"}


def cmd_preview_pptx(args) -> None:
    """CLI entry: emit preview status JSON to stdout."""
    width = getattr(args, "width", None) or DEFAULT_WIDTH
    result = generate_previews(args.input, args.output_dir, width=width)
    sys.stdout.write(json.dumps(result, ensure_ascii=True))


# PowerPoint's SaveAs format constant for PDF.
_PP_SAVE_AS_PDF = 32


def export_to_pdf(pptx_path: str, pdf_path: str) -> dict:
    """Export a PPTX to PDF. Returns {"ok": bool, "path": str|None,
    "error": str|None}. Same fail-soft contract as `generate_previews` so
    callers don't need try/except.

    Dispatches by platform: Windows uses PowerPoint COM; Linux/macOS uses
    headless LibreOffice."""
    if not Path(pptx_path).is_file():
        return {"ok": False, "path": None,
                "error": "PPTX no encontrado."}

    pptx_abs = os.path.abspath(pptx_path)
    pdf_abs = os.path.abspath(pdf_path)
    os.makedirs(os.path.dirname(pdf_abs) or ".", exist_ok=True)

    if sys.platform == "win32":
        return _export_pdf_powerpoint(pptx_abs, pdf_abs)
    return _export_pdf_libreoffice(pptx_abs, pdf_abs)


def _export_pdf_powerpoint(pptx_abs: str, pdf_abs: str) -> dict:
    """PPTX -> PDF via PowerPoint COM (Windows + Office only)."""
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


def _export_pdf_libreoffice(pptx_abs: str, pdf_abs: str) -> dict:
    """PPTX -> PDF via headless LibreOffice (Linux/macOS). soffice always
    names its output `<input-stem>.pdf` inside --outdir, so we convert into a
    temp dir and move the result to the caller's requested path."""
    soffice = _find_soffice()
    if soffice is None:
        return {"ok": False, "path": None,
                "error": "Export a PDF no disponible (LibreOffice no instalado)."}
    try:
        with tempfile.TemporaryDirectory(prefix="socya_pdf_") as pdf_tmp:
            produced = _soffice_convert_to_pdf(soffice, pptx_abs, pdf_tmp)
            # Path destino: reemplaza si ya existía (shutil.move falla al mover
            # sobre un archivo en algunos FS, así que lo quitamos antes).
            if Path(pdf_abs).exists():
                Path(pdf_abs).unlink()
            shutil.move(produced, pdf_abs)
    except Exception as e:
        return {"ok": False, "path": None,
                "error": f"LibreOffice no pudo exportar a PDF: {str(e)[:200]}"}

    if not Path(pdf_abs).is_file():
        return {"ok": False, "path": None,
                "error": "PDF no se generó (LibreOffice terminó sin escribir el archivo)."}
    return {"ok": True, "path": pdf_abs, "error": None}


def cmd_export_pdf(args) -> None:
    """CLI entry: emit PDF export status JSON to stdout."""
    result = export_to_pdf(args.input, args.output)
    sys.stdout.write(json.dumps(result, ensure_ascii=True))
