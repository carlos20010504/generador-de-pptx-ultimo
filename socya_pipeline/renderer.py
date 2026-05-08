"""Pure execute renderer with editorial-grade visual design.

Layout system inspired by SIG_Revision_Direccion_2025 (the 'good' reference deck):
- 16:9 widescreen (13.33"x7.5")
- Forest green + gold palette
- Georgia for titles and big KPIs, Calibri for body text
- Eyebrow + accent bar + title + subtitle pattern on each slide
- Footer band with pagination on each slide
- KPI cards with white fill + colored top header band
- No fallbacks. No template fillers.
"""
import io
from pathlib import Path
from typing import List, Optional
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ──────────────────── Design system ────────────────────

# Palette (matches reference deck: forest green + gold)
FOREST_DARK = RGBColor(0x1B, 0x3B, 0x2F)   # primary, titles, footer
FOREST_MID  = RGBColor(0x2C, 0x5F, 0x2D)   # KPI accent A
FOREST_LITE = RGBColor(0x4A, 0x7C, 0x3A)   # KPI accent B
GOLD        = RGBColor(0xC9, 0xA2, 0x27)   # eyebrow, KPI accent C, accent line
TERRACOTTA  = RGBColor(0xB8, 0x50, 0x42)   # warning/alert accent
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
TEXT_GRAY   = RGBColor(0x6B, 0x6B, 0x6B)
TEXT_DARK   = RGBColor(0x33, 0x33, 0x33)
PAGE_BG     = RGBColor(0xFA, 0xFA, 0xF5)   # off-white page background
ZEBRA_BG    = RGBColor(0xF2, 0xF0, 0xE6)   # subtle warm gray for zebra rows

KPI_COLORS = [FOREST_MID, FOREST_LITE, GOLD, TERRACOTTA]
CHART_PALETTE = ["#2C5F2D", "#4A7C3A", "#C9A227", "#B85042", "#1B3B2F", "#84B59F"]

HEADING_FONT = "Georgia"
BODY_FONT = "Calibri"

# Layout constants (16:9 = 13.33 x 7.50 inches)
SLIDE_W = 13.33
SLIDE_H = 7.50
MARGIN = 0.5
CONTENT_W = SLIDE_W - 2 * MARGIN  # 12.33

EYEBROW_Y = 0.45
ACCENT_Y = 0.78
TITLE_Y = 0.90
SUBTITLE_Y = 1.85
CONTENT_TOP = 2.45
CONTENT_BOTTOM = 6.30
PAGINATION_Y = 7.05


def render_pptx(slides: List[dict], presentation_meta: dict,
                template_path: Path, output_path: Path) -> None:
    """Render slides to a PPTX. presentation_meta drives the title slide + footer."""
    prs = Presentation(str(template_path))
    # Force 16:9 widescreen
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    # Strip pre-existing slides from the template — we control the deck
    while len(prs.slides._sldIdLst) > 0:
        rId = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]

    blank_layout = _pick_blank_layout(prs)
    deck_title = (presentation_meta.get("title") or "").strip()
    deck_subtitle = (presentation_meta.get("subtitle") or "").strip()
    footer_tagline = _build_footer_tagline(deck_title)

    total = len(slides)
    for i, slide_def in enumerate(slides):
        stype = slide_def.get("type")
        slide = prs.slides.add_slide(blank_layout)
        _add_page_background(slide)

        if stype == "title":
            _add_title_slide(slide, slide_def, deck_title, deck_subtitle,
                             i + 1, total)
            continue

        eyebrow = slide_def.get("eyebrow") or _eyebrow_from_type(stype, i)
        _add_eyebrow(slide, eyebrow)
        _add_title_block(slide, slide_def.get("title", ""),
                         _slide_subtitle(slide_def))

        if stype == "kpi_row":
            _add_kpi_row(slide, slide_def)
        elif stype == "chart":
            _add_chart_slide(slide, slide_def)
        elif stype == "table":
            _add_table_slide(slide, slide_def)
        elif stype == "text_bullets":
            _add_bullets_slide(slide, slide_def)

        _add_footer(slide, footer_tagline, i + 1, total)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))


# ──────────────────── Slide builders ────────────────────

def _add_title_slide(slide, slide_def: dict, deck_title: str, deck_subtitle: str,
                     page: int, total: int) -> None:
    """Cover slide: dark forest background, gold accent, big white title."""
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0,
                                  Inches(SLIDE_W), Inches(SLIDE_H))
    bg.line.fill.background()
    bg.fill.solid(); bg.fill.fore_color.rgb = FOREST_DARK

    eyebrow = slide_def.get("eyebrow") or "PRESENTACIÓN EJECUTIVA"
    _put_text(slide, eyebrow.upper(), Inches(MARGIN), Inches(0.7),
              Inches(CONTENT_W), Inches(0.4),
              font=BODY_FONT, font_size=12, bold=True, color=GOLD,
              letter_spacing_em=0.15)

    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   Inches(MARGIN), Inches(1.2),
                                   Inches(0.6), Inches(0.10))
    bar.line.fill.background()
    bar.fill.solid(); bar.fill.fore_color.rgb = GOLD

    title_text = slide_def.get("data", {}).get("title") or deck_title or "Presentación"
    _put_text(slide, title_text, Inches(MARGIN), Inches(2.2),
              Inches(CONTENT_W), Inches(2.5),
              font=HEADING_FONT, font_size=54, bold=True, color=WHITE,
              line_spacing=1.05)

    subtitle_text = slide_def.get("data", {}).get("subtitle") or deck_subtitle
    if subtitle_text:
        _put_text(slide, subtitle_text, Inches(MARGIN), Inches(5.0),
                  Inches(CONTENT_W * 0.85), Inches(1.0),
                  font=BODY_FONT, font_size=18, color=RGBColor(0xCF, 0xD7, 0xCB),
                  line_spacing=1.3)

    _put_text(slide, f"{page:02d} / {total:02d}",
              Inches(SLIDE_W - 1.8), Inches(SLIDE_H - 0.7),
              Inches(1.3), Inches(0.4),
              font=BODY_FONT, font_size=11, bold=True, color=GOLD,
              align="right")


def _add_kpi_row(slide, slide_def: dict) -> None:
    """3 or 4 big KPI cards across the slide."""
    kpis = slide_def.get("data", {}).get("kpis", [])[:4]
    if not kpis:
        return

    n = len(kpis)
    gap = 0.25
    total_gap = gap * (n - 1)
    card_w = (CONTENT_W - total_gap) / n
    card_h = 3.6
    card_y = CONTENT_TOP

    use_big_numbers = n <= 3
    number_size = 96 if use_big_numbers else 56

    for i, k in enumerate(kpis):
        x = MARGIN + i * (card_w + gap)
        accent = KPI_COLORS[i % len(KPI_COLORS)]
        _add_kpi_card(slide, x, card_y, card_w, card_h,
                      header=k.get("label", ""),
                      value=k.get("value", ""),
                      description=k.get("description", ""),
                      accent_color=accent,
                      number_size=number_size)


def _add_kpi_card(slide, x: float, y: float, w: float, h: float,
                  header: str, value: str, description: str,
                  accent_color: RGBColor, number_size: int) -> None:
    """KPI card: white body + colored top band + huge number + description."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                    Inches(x), Inches(y), Inches(w), Inches(h))
    card.adjustments[0] = 0.04
    card.fill.solid(); card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = RGBColor(0xE5, 0xE5, 0xDF)
    card.line.width = Pt(0.5)

    header_h = 0.5
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                    Inches(x), Inches(y),
                                    Inches(w), Inches(header_h))
    band.line.fill.background()
    band.fill.solid(); band.fill.fore_color.rgb = accent_color

    if header:
        _put_text(slide, header.upper(),
                  Inches(x + 0.2), Inches(y),
                  Inches(w - 0.4), Inches(header_h),
                  font=BODY_FONT, font_size=11, bold=True, color=WHITE,
                  align="left", anchor="middle", letter_spacing_em=0.08)

    _put_text(slide, str(value),
              Inches(x + 0.2), Inches(y + header_h + 0.3),
              Inches(w - 0.4), Inches(h - header_h - 1.4),
              font=HEADING_FONT, font_size=number_size, bold=True,
              color=accent_color, align="center", anchor="middle",
              line_spacing=1.0)

    if description:
        _put_text(slide, description,
                  Inches(x + 0.3), Inches(y + h - 1.0),
                  Inches(w - 0.6), Inches(0.9),
                  font=BODY_FONT, font_size=11, color=TEXT_GRAY,
                  align="center", anchor="top", line_spacing=1.25)


def _add_chart_slide(slide, slide_def: dict) -> None:
    """Chart on the left, narrative card on the right."""
    data = slide_def.get("data", {})
    narrative = (slide_def.get("narrative") or "").strip()

    chart_w = 7.7
    chart_h = CONTENT_BOTTOM - CONTENT_TOP - 0.1
    chart_x = MARGIN
    chart_y = CONTENT_TOP

    img = _build_chart_png(data, width_in=chart_w, height_in=chart_h)
    slide.shapes.add_picture(img, Inches(chart_x), Inches(chart_y),
                              Inches(chart_w), Inches(chart_h))

    if narrative:
        card_x = MARGIN + chart_w + 0.3
        card_w = CONTENT_W - chart_w - 0.3
        card_y = CONTENT_TOP
        card_h = chart_h
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                        Inches(card_x), Inches(card_y),
                                        Inches(card_w), Inches(card_h))
        card.adjustments[0] = 0.05
        card.fill.solid(); card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = RGBColor(0xE5, 0xE5, 0xDF)
        card.line.width = Pt(0.5)
        header_h = 0.55
        band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                        Inches(card_x), Inches(card_y),
                                        Inches(card_w), Inches(header_h))
        band.line.fill.background()
        band.fill.solid(); band.fill.fore_color.rgb = FOREST_MID
        _put_text(slide, "LECTURA EJECUTIVA",
                  Inches(card_x + 0.25), Inches(card_y),
                  Inches(card_w - 0.5), Inches(header_h),
                  font=BODY_FONT, font_size=11, bold=True, color=WHITE,
                  anchor="middle", letter_spacing_em=0.08)
        _put_text(slide, narrative,
                  Inches(card_x + 0.3), Inches(card_y + header_h + 0.2),
                  Inches(card_w - 0.6), Inches(card_h - header_h - 0.4),
                  font=BODY_FONT, font_size=14, color=TEXT_DARK,
                  line_spacing=1.4, anchor="top")


def _add_table_slide(slide, slide_def: dict) -> None:
    """Table with colored header band and zebra-striped rows."""
    data = slide_def.get("data", {})
    headers = data.get("headers", [])
    rows = data.get("rows", [])
    if not headers or not rows:
        return

    max_cols = 6
    max_rows = 12
    headers = [_clean_header(h) for h in headers[:max_cols]]
    rows = [list(r)[:max_cols] for r in rows[:max_rows]]
    n_cols = len(headers)
    n_rows = len(rows) + 1

    table_x = MARGIN
    table_y = CONTENT_TOP
    table_w = CONTENT_W
    available_h = CONTENT_BOTTOM - CONTENT_TOP - 0.2
    row_h = min(0.4, available_h / n_rows)
    table_h = row_h * n_rows

    table_shape = slide.shapes.add_table(n_rows, n_cols,
                                           Inches(table_x), Inches(table_y),
                                           Inches(table_w), Inches(table_h))
    tbl = table_shape.table

    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        cell.text = ""
        cell.fill.solid(); cell.fill.fore_color.rgb = FOREST_DARK
        tf = cell.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.08)
        tf.margin_top = tf.margin_bottom = Inches(0.04)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = str(h).upper()
        run.font.name = BODY_FONT
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = WHITE

    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.text = ""
            cell.fill.solid()
            cell.fill.fore_color.rgb = ZEBRA_BG if i % 2 == 0 else WHITE
            tf = cell.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_right = Inches(0.08)
            tf.margin_top = tf.margin_bottom = Inches(0.04)
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = "" if val is None else str(val)
            run.font.name = BODY_FONT
            run.font.size = Pt(10)
            run.font.color.rgb = TEXT_DARK


def _add_bullets_slide(slide, slide_def: dict) -> None:
    """Bullets rendered as cards in a single column with gold left accent."""
    bullets = slide_def.get("data", {}).get("bullets", [])[:6]
    if not bullets:
        return

    card_w = CONTENT_W
    available_h = CONTENT_BOTTOM - CONTENT_TOP - 0.3
    n = len(bullets)
    gap = 0.18
    card_h = (available_h - gap * (n - 1)) / n
    card_h = min(card_h, 1.0)
    card_h = max(card_h, 0.5)

    for i, bullet in enumerate(bullets):
        y = CONTENT_TOP + i * (card_h + gap)
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                        Inches(MARGIN), Inches(y),
                                        Inches(card_w), Inches(card_h))
        card.adjustments[0] = 0.05
        card.fill.solid(); card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = RGBColor(0xE5, 0xE5, 0xDF)
        card.line.width = Pt(0.5)
        accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                          Inches(MARGIN), Inches(y),
                                          Inches(0.12), Inches(card_h))
        accent.line.fill.background()
        accent.fill.solid(); accent.fill.fore_color.rgb = GOLD
        _put_text(slide, f"{i + 1:02d}",
                  Inches(MARGIN + 0.3), Inches(y),
                  Inches(0.7), Inches(card_h),
                  font=HEADING_FONT, font_size=22, bold=True, color=FOREST_DARK,
                  align="left", anchor="middle")
        _put_text(slide, bullet,
                  Inches(MARGIN + 1.1), Inches(y),
                  Inches(card_w - 1.4), Inches(card_h),
                  font=BODY_FONT, font_size=14, color=TEXT_DARK,
                  align="left", anchor="middle", line_spacing=1.3)


# ──────────────────── Common chrome ────────────────────

def _add_page_background(slide) -> None:
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0,
                                  Inches(SLIDE_W), Inches(SLIDE_H))
    bg.line.fill.background()
    bg.fill.solid(); bg.fill.fore_color.rgb = PAGE_BG


def _add_eyebrow(slide, text: str) -> None:
    if text:
        _put_text(slide, text.upper(),
                  Inches(MARGIN), Inches(EYEBROW_Y),
                  Inches(CONTENT_W), Inches(0.3),
                  font=BODY_FONT, font_size=10, bold=True, color=GOLD,
                  letter_spacing_em=0.12)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   Inches(MARGIN), Inches(ACCENT_Y),
                                   Inches(0.4), Inches(0.08))
    bar.line.fill.background()
    bar.fill.solid(); bar.fill.fore_color.rgb = GOLD


def _add_title_block(slide, title: str, subtitle: str) -> None:
    if title:
        _put_text(slide, title,
                  Inches(MARGIN), Inches(TITLE_Y),
                  Inches(CONTENT_W), Inches(0.95),
                  font=HEADING_FONT, font_size=28, bold=True, color=FOREST_DARK,
                  line_spacing=1.1)
    if subtitle:
        _put_text(slide, subtitle,
                  Inches(MARGIN), Inches(SUBTITLE_Y),
                  Inches(CONTENT_W), Inches(0.55),
                  font=BODY_FONT, font_size=12, color=TEXT_GRAY,
                  line_spacing=1.3)


def _add_footer(slide, tagline: str, page: int, total: int) -> None:
    sep = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   Inches(MARGIN), Inches(PAGINATION_Y - 0.05),
                                   Inches(0.4), Inches(0.04))
    sep.line.fill.background()
    sep.fill.solid(); sep.fill.fore_color.rgb = GOLD

    _put_text(slide, tagline,
              Inches(MARGIN + 0.6), Inches(PAGINATION_Y),
              Inches(CONTENT_W - 1.6), Inches(0.3),
              font=BODY_FONT, font_size=9, color=TEXT_GRAY,
              letter_spacing_em=0.05)

    _put_text(slide, f"{page} / {total}",
              Inches(SLIDE_W - 1.3), Inches(PAGINATION_Y),
              Inches(0.8), Inches(0.3),
              font=BODY_FONT, font_size=9, bold=True, color=FOREST_DARK,
              align="right")


# ──────────────────── Charts (matplotlib) ────────────────────

def _build_chart_png(data: dict, width_in: float = 7.5,
                       height_in: float = 4.0) -> io.BytesIO:
    chart_type = data.get("chart_type", "bar")
    labels = [str(l) for l in data.get("labels", [])]
    values = list(data.get("values", []))
    name = data.get("name", "")

    fig, ax = plt.subplots(figsize=(width_in, height_in), dpi=150,
                             facecolor="#FAFAF5")
    ax.set_facecolor("#FAFAF5")

    n = len(labels)
    colors = (CHART_PALETTE * ((n // len(CHART_PALETTE)) + 1))[:n]

    if chart_type == "pie":
        wedges, texts, autotexts = ax.pie(
            values, labels=labels, autopct="%1.1f%%",
            colors=colors, startangle=90, counterclock=False,
            wedgeprops={"linewidth": 2, "edgecolor": "#FAFAF5"},
            textprops={"fontsize": 10, "fontfamily": "Calibri", "color": "#333333"},
        )
        for at in autotexts:
            at.set_color("white")
            at.set_fontweight("bold")
    elif chart_type == "line":
        ax.plot(labels, values, marker="o", color=CHART_PALETTE[0],
                 linewidth=2.5, markersize=7, markerfacecolor=CHART_PALETTE[2],
                 markeredgecolor=CHART_PALETTE[0], markeredgewidth=2)
        ax.set_title(name, fontsize=12, fontfamily="Georgia",
                     color="#1B3B2F", pad=12, loc="left", fontweight="bold")
        ax.tick_params(axis="x", rotation=35, labelsize=9, colors="#6B6B6B")
        ax.tick_params(axis="y", labelsize=9, colors="#6B6B6B")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color("#D5D5CD")
        ax.spines["bottom"].set_color("#D5D5CD")
        ax.grid(axis="y", linestyle="--", linewidth=0.5, color="#D5D5CD", alpha=0.7)
    else:  # bar
        bars = ax.bar(labels, values, color=colors, edgecolor="#FAFAF5",
                       linewidth=1.5, width=0.65)
        ax.set_title(name, fontsize=12, fontfamily="Georgia",
                     color="#1B3B2F", pad=12, loc="left", fontweight="bold")
        ax.tick_params(axis="x", rotation=25, labelsize=9, colors="#6B6B6B")
        ax.tick_params(axis="y", labelsize=9, colors="#6B6B6B")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color("#D5D5CD")
        ax.spines["bottom"].set_color("#D5D5CD")
        ax.grid(axis="y", linestyle="--", linewidth=0.5, color="#D5D5CD", alpha=0.7)
        for b, v in zip(bars, values):
            ax.text(b.get_x() + b.get_width() / 2, b.get_height(),
                     _format_value_short(v),
                     ha="center", va="bottom", fontsize=9,
                     fontfamily="Calibri", color="#333333", fontweight="bold")

    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight",
                 facecolor=fig.get_facecolor())
    plt.close(fig)
    buf.seek(0)
    return buf


def _format_value_short(v) -> str:
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    if abs(f) >= 1_000_000:
        return f"{f / 1_000_000:.1f}M"
    if abs(f) >= 1_000:
        return f"{f / 1_000:.1f}K"
    if f.is_integer():
        return str(int(f))
    return f"{f:.1f}"


# ──────────────────── Helpers ────────────────────

def _pick_blank_layout(prs):
    for layout in prs.slide_layouts:
        if "Blank" in layout.name or "blanco" in layout.name.lower():
            return layout
    return prs.slide_layouts[-1]


def _eyebrow_from_type(stype: str, idx: int) -> str:
    return {
        "kpi_row": "INDICADORES CLAVE",
        "chart": "ANÁLISIS VISUAL",
        "table": "DETALLE OPERATIVO",
        "text_bullets": "HALLAZGOS Y CONCLUSIONES",
    }.get(stype, f"SECCIÓN {idx + 1:02d}")


def _slide_subtitle(slide_def: dict) -> str:
    return (slide_def.get("subtitle")
            or slide_def.get("data", {}).get("subtitle")
            or "")


def _build_footer_tagline(deck_title: str) -> str:
    parts = []
    if deck_title:
        parts.append(deck_title)
    parts.append("PRESENTACIÓN GENERADA AUTOMÁTICAMENTE")
    return "  ·  ".join(parts).upper()


def _clean_header(value) -> str:
    s = str(value or "").strip()
    if s.startswith("Unnamed:") or not s or s.lower() in ("nan", "none", "null"):
        return ""
    return s


def _put_text(slide, text, x, y, w, h, *, font: str = BODY_FONT,
              font_size: float = 12, bold: bool = False,
              color: Optional[RGBColor] = None,
              align: str = "left", anchor: str = "top",
              line_spacing: float = 1.2,
              letter_spacing_em: float = 0.0) -> None:
    """Add a text box. align: left/center/right. anchor: top/middle/bottom."""
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.04)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    if anchor == "middle":
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    elif anchor == "bottom":
        tf.vertical_anchor = MSO_ANCHOR.BOTTOM
    else:
        tf.vertical_anchor = MSO_ANCHOR.TOP

    p = tf.paragraphs[0]
    p.alignment = {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
    }.get(align, PP_ALIGN.LEFT)
    p.line_spacing = line_spacing

    run = p.add_run()
    run.text = str(text or "")
    run.font.name = font
    run.font.size = Pt(font_size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    if letter_spacing_em:
        try:
            rPr = run._r.get_or_add_rPr()
            rPr.set("spc", str(int(letter_spacing_em * font_size * 100)))
        except Exception:
            pass
