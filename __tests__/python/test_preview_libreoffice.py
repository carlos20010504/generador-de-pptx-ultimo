"""Tests for the Linux/LibreOffice preview + PDF path in socya_pipeline.preview.

LibreOffice and poppler only exist in the Docker image, not on the Windows dev
box, so every external call is mocked: `subprocess.run` (soffice) and
`pdf2image.convert_from_path` (poppler). These run cross-platform and lock in
the renderer's contract and the concurrency-safety design (isolated soffice
profile + Impress PDF filter + one PNG per slide via PDF)."""
import sys
import types
from pathlib import Path

import pytest

from socya_pipeline import preview


def _install_fake_soffice(monkeypatch, *, returncode=0):
    """Force the Linux branch and capture the soffice command.

    The fake `subprocess.run` creates the PDF soffice would have produced
    (`<input-stem>.pdf` inside `--outdir`) so downstream code sees a real file.
    Returns a list that will hold the captured argv."""
    captured = []
    monkeypatch.setattr(preview.sys, "platform", "linux")
    monkeypatch.setattr(preview, "_find_soffice", lambda: "/usr/bin/soffice")

    def fake_run(cmd, *args, **kwargs):
        captured.append(cmd)
        out_dir = cmd[cmd.index("--outdir") + 1]
        pptx_abs = cmd[-1]
        produced = Path(out_dir) / (Path(pptx_abs).stem + ".pdf")
        produced.write_bytes(b"%PDF-1.4 fake pdf payload for tests")
        return types.SimpleNamespace(returncode=returncode, stdout="", stderr="")

    monkeypatch.setattr(preview.subprocess, "run", fake_run)
    return captured


def _install_fake_pdf2image(monkeypatch, n_pages):
    """Inject a fake `pdf2image` module whose images write real PNG bytes."""
    class _FakeImg:
        def save(self, path, fmt):
            Path(path).write_bytes(b"\x89PNG\r\n\x1a\n" + b"0" * 200)

    fake_mod = types.ModuleType("pdf2image")
    fake_mod.convert_from_path = lambda pdf_path, size=None, **kw: [
        _FakeImg() for _ in range(n_pages)
    ]
    monkeypatch.setitem(sys.modules, "pdf2image", fake_mod)


def test_generate_previews_libreoffice_happy_path(monkeypatch, tmp_path):
    captured = _install_fake_soffice(monkeypatch)
    _install_fake_pdf2image(monkeypatch, n_pages=3)

    pptx = tmp_path / "deck.pptx"
    pptx.write_bytes(b"fake pptx")
    out_dir = tmp_path / "previews"

    result = preview.generate_previews(str(pptx), str(out_dir), width=900)

    assert result["ok"] is True
    assert result["renderer"] == "libreoffice"
    assert [s["filename"] for s in result["slides"]] == [
        "slide-01.png", "slide-02.png", "slide-03.png",
    ]
    for s in result["slides"]:
        assert (out_dir / s["filename"]).stat().st_size > 100

    # Concurrency-safety + correctness design must hold:
    cmd = captured[0]
    assert any(a.startswith("-env:UserInstallation=file://") for a in cmd), \
        "soffice must use an isolated per-call profile"
    assert "--headless" in cmd
    assert cmd[cmd.index("--convert-to") + 1] == "pdf:impress_pdf_Export"


def test_generate_previews_soffice_absent_fails_soft(monkeypatch, tmp_path):
    monkeypatch.setattr(preview.sys, "platform", "linux")
    monkeypatch.setattr(preview, "_find_soffice", lambda: None)

    pptx = tmp_path / "deck.pptx"
    pptx.write_bytes(b"fake pptx")

    result = preview.generate_previews(str(pptx), str(tmp_path / "out"))

    assert result["ok"] is False
    assert result["slides"] == []
    assert "libreoffice" in result["error"].lower()


def test_generate_previews_missing_pptx():
    result = preview.generate_previews("/no/such/deck.pptx", "/tmp/out")
    assert result["ok"] is False
    assert result["error"] == "PPTX no encontrado."


def test_export_pdf_libreoffice_happy_path(monkeypatch, tmp_path):
    _install_fake_soffice(monkeypatch)

    pptx = tmp_path / "deck.pptx"
    pptx.write_bytes(b"fake pptx")
    dest = tmp_path / "nested" / "out.pdf"

    result = preview.export_to_pdf(str(pptx), str(dest))

    assert result["ok"] is True
    assert Path(result["path"]).is_file()
    assert dest.is_file()


def test_export_pdf_soffice_absent_fails_soft(monkeypatch, tmp_path):
    monkeypatch.setattr(preview.sys, "platform", "linux")
    monkeypatch.setattr(preview, "_find_soffice", lambda: None)

    pptx = tmp_path / "deck.pptx"
    pptx.write_bytes(b"fake pptx")

    result = preview.export_to_pdf(str(pptx), str(tmp_path / "out.pdf"))

    assert result["ok"] is False
    assert result["path"] is None
    assert "libreoffice" in result["error"].lower()


def test_export_pdf_missing_pptx():
    result = preview.export_to_pdf("/no/such/deck.pptx", "/tmp/out.pdf")
    assert result["ok"] is False
    assert result["error"] == "PPTX no encontrado."
