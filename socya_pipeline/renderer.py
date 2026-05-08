"""Pure execute renderer. No fallbacks. No template fillers."""
import io
from pathlib import Path
from typing import List, Optional
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

PRIMARY = RGBColor(0x0F, 0x17, 0x2A)
ACCENT = RGBColor(0x25, 0x63, 0xEB)
TEXT_DARK = RGBColor(0x33, 0x33, 0x33)

def render_pptx(slides: List[dict], presentation_meta: dict,
                  template_path: Path, output_path: Path) -> None:
    prs = Presentation(str(template_path))
    # Strip any pre-existing slides from the template so we control the deck
    while prs.slides and len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]

    blank_layout = _pick_blank_layout(prs)

    for slide_def in slides:
        stype = slide_def.get("type")
        if stype == "title":
            _add_title(prs, blank_layout, slide_def["data"])
        elif stype == "kpi_row":
            _add_kpi_row(prs, blank_layout, slide_def)
        elif stype == "chart":
            _add_chart(prs, blank_layout, slide_def)
        elif stype == "table":
            _add_table(prs, blank_layout, slide_def)
        elif stype == "text_bullets":
            _add_bullets(prs, blank_layout, slide_def)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))


def _pick_blank_layout(prs):
    for layout in prs.slide_layouts:
        if "Blank" in layout.name or "blanco" in layout.name.lower():
            return layout
    return prs.slide_layouts[0]


def _add_title(prs, layout, data):
    s = prs.slides.add_slide(layout)
    _put_text(s, data.get("title", ""), Inches(0.7), Inches(2.2), Inches(8.5),
                Inches(1.5), font_size=44, bold=True, color=PRIMARY)
    _put_text(s, data.get("subtitle", ""), Inches(0.7), Inches(3.8), Inches(8.5),
                Inches(1.0), font_size=20, color=TEXT_DARK)


def _add_kpi_row(prs, layout, slide_def):
    s = prs.slides.add_slide(layout)
    _put_text(s, slide_def.get("title", ""), Inches(0.5), Inches(0.5), Inches(9),
                Inches(0.7), font_size=24, bold=True, color=PRIMARY)
    kpis = slide_def["data"].get("kpis", [])[:4]
    if not kpis:
        return
    width = Inches(8.5 / max(1, len(kpis)))
    for i, k in enumerate(kpis):
        x = Inches(0.5) + width * i
        _put_text(s, k["value"], x, Inches(2.2), width, Inches(1.0),
                    font_size=36, bold=True, color=ACCENT, align="center")
        _put_text(s, k["label"], x, Inches(3.4), width, Inches(0.6),
                    font_size=14, color=TEXT_DARK, align="center")


def _add_chart(prs, layout, slide_def):
    s = prs.slides.add_slide(layout)
    _put_text(s, slide_def.get("title", ""), Inches(0.5), Inches(0.4), Inches(9),
                Inches(0.7), font_size=22, bold=True, color=PRIMARY)
    data = slide_def["data"]
    img = _build_chart_png(data)
    s.shapes.add_picture(img, Inches(0.7), Inches(1.4), Inches(8.5), Inches(5.0))
    if slide_def.get("narrative"):
        _put_text(s, slide_def["narrative"], Inches(0.7), Inches(6.6),
                    Inches(8.5), Inches(0.6), font_size=12, color=TEXT_DARK)


def _build_chart_png(data: dict) -> io.BytesIO:
    chart_type = data.get("chart_type", "bar")
    labels = data.get("labels", [])
    values = data.get("values", [])
    fig, ax = plt.subplots(figsize=(8, 4.5), dpi=120)
    if chart_type == "pie":
        ax.pie(values, labels=labels, autopct="%1.1f%%",
               colors=plt.cm.Blues([0.4 + 0.5 * i / max(1, len(labels))
                                       for i in range(len(labels))]))
    elif chart_type == "line":
        ax.plot(labels, values, marker="o", color="#2563EB", linewidth=2)
        ax.tick_params(axis="x", rotation=45)
    else:  # bar
        ax.bar(labels, values, color="#2563EB")
        ax.tick_params(axis="x", rotation=30)
    ax.set_title(data.get("name", ""), fontsize=12)
    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def _add_table(prs, layout, slide_def):
    s = prs.slides.add_slide(layout)
    _put_text(s, slide_def.get("title", ""), Inches(0.5), Inches(0.4), Inches(9),
                Inches(0.7), font_size=22, bold=True, color=PRIMARY)
    data = slide_def["data"]
    headers = data.get("headers", [])
    rows = data.get("rows", [])
    if not headers or not rows:
        return
    n_cols = len(headers)
    n_rows = len(rows) + 1
    table_shape = s.shapes.add_table(n_rows, n_cols, Inches(0.5), Inches(1.4),
                                       Inches(9), Inches(0.4 * n_rows + 0.4))
    tbl = table_shape.table
    for i, h in enumerate(headers):
        cell = tbl.cell(0, i)
        cell.text = str(h)
        for r in cell.text_frame.paragraphs:
            for run in r.runs:
                run.font.bold = True
                run.font.size = Pt(11)
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        cell.fill.solid(); cell.fill.fore_color.rgb = PRIMARY
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, val in enumerate(row[:n_cols]):
            tbl.cell(r_idx, c_idx).text = "" if val is None else str(val)


def _add_bullets(prs, layout, slide_def):
    s = prs.slides.add_slide(layout)
    _put_text(s, slide_def.get("title", ""), Inches(0.5), Inches(0.4), Inches(9),
                Inches(0.7), font_size=22, bold=True, color=PRIMARY)
    bullets = slide_def["data"].get("bullets", [])
    body = s.shapes.add_textbox(Inches(0.7), Inches(1.5), Inches(8.5), Inches(5))
    tf = body.text_frame
    tf.word_wrap = True
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"• {b}"
        for run in p.runs:
            run.font.size = Pt(16)
            run.font.color.rgb = TEXT_DARK


def _put_text(slide, text, x, y, w, h, *, font_size=14, bold=False,
                color=None, align="left"):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = str(text or "")
    if align == "center":
        from pptx.enum.text import PP_ALIGN
        p.alignment = PP_ALIGN.CENTER
    for run in p.runs:
        run.font.size = Pt(font_size)
        run.font.bold = bold
        if color is not None:
            run.font.color.rgb = color
