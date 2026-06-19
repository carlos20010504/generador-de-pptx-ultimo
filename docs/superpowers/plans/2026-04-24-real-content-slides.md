# Real Content Slides Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans

**Goal:** Eliminate all generic boilerplate text from the PPTX generator and replace it with real content extracted directly from Excel cells.

**Architecture:** All text-generating functions in `generate_template_presentation.py` that produce meta-commentary are replaced with extractors that read actual cell values. A new `extract_real_bullets()` function serves as the universal content extractor used across all slide types. Narrative tables get Option C: bullets slide + table slide.

**Tech Stack:** Python, python-pptx, existing generate_template_presentation.py

---

### Task 1: Add `extract_real_bullets()` — universal content extractor

**Files:**
- Modify: `generate_template_presentation.py` (after `build_table_signal_profile`)

- [ ] Add the function after `build_table_signal_profile`:

```python
def extract_real_bullets(headers, rows, max_bullets=6):
    """Extract real cell content as smart formatted bullets. No meta-commentary."""
    if not headers or not rows:
        return []

    profile = build_table_signal_profile(headers, rows)
    basis = derive_table_basis(headers, rows)
    label_index = basis.get("label_index")
    numeric_index = basis.get("numeric_index")

    # Score each column for informativeness
    col_scores = []
    for idx, header in enumerate(headers):
        vals = [clean_text(r[idx]) for r in rows if idx < len(r)]
        non_empty = [v for v in vals if not is_noise_text(v) and len(v) >= 6]
        if not non_empty:
            col_scores.append(-1)
            continue
        avg_len = sum(len(v) for v in non_empty) / len(non_empty)
        unique_ratio = len(set(v.casefold() for v in non_empty)) / len(non_empty)
        score = len(non_empty) * 0.4 + min(avg_len, 60) * 0.3 + unique_ratio * 0.3
        if is_identifier_header(header):
            score -= 20
        if is_person_like_header(header):
            score -= 10
        col_scores.append(score)

    # Best text column (highest score among non-numeric columns)
    best_text_idx = None
    best_text_score = -1
    short_label_idx = None  # short categorical col to use as prefix
    for idx, score in enumerate(col_scores):
        if score < 0:
            continue
        vals = [clean_text(r[idx]) for r in rows if idx < len(r)]
        non_empty = [v for v in vals if not is_noise_text(v)]
        if not non_empty:
            continue
        avg_len = sum(len(v) for v in non_empty) / len(non_empty)
        numeric_ratio = sum(1 for v in non_empty if try_number(v) is not None) / len(non_empty)
        if numeric_ratio > 0.5:
            continue
        if avg_len <= 20 and len(set(v.casefold() for v in non_empty)) <= 8:
            if short_label_idx is None:
                short_label_idx = idx
        if score > best_text_score:
            best_text_score = score
            best_text_idx = idx

    if best_text_idx is None:
        return []

    bullets = []
    seen = set()
    for row in rows:
        if len(bullets) >= max_bullets:
            break
        if best_text_idx >= len(row):
            continue
        main_val = clean_text(row[best_text_idx], 120)
        if not main_val or is_noise_text(main_val) or len(main_val) < 6:
            continue
        key = main_val.casefold()
        if key in seen:
            continue
        seen.add(key)

        # Smart formatting: [short_label] long_description
        if short_label_idx is not None and short_label_idx != best_text_idx and short_label_idx < len(row):
            prefix = clean_text(row[short_label_idx], 25)
            if prefix and not is_noise_text(prefix) and prefix.casefold() != main_val.casefold():
                bullet = f"[{prefix}] {main_val}"
            else:
                bullet = main_val
        elif numeric_index is not None and numeric_index < len(row):
            num_val = try_number(row[numeric_index])
            if num_val is not None:
                metric_kind = infer_metric_kind(headers[numeric_index], [num_val])
                bullet = f"{main_val}: {format_metric(num_val, metric_kind)}"
            else:
                bullet = main_val
        else:
            bullet = main_val

        # Truncate to 85 chars
        bullet = clean_text(bullet)
        if len(bullet) > 85:
            bullet = bullet[:82] + "..."
        if len(bullet) >= 8:
            bullets.append(bullet)

    return bullets
```

- [ ] Commit: `git add generate_template_presentation.py && git commit -m "feat: add extract_real_bullets universal content extractor"`

---

### Task 2: Replace `build_table_summary_lines` with real content

**Files:**
- Modify: `generate_template_presentation.py:1713-1777`

- [ ] Replace the entire `build_table_summary_lines` function:

```python
def build_table_summary_lines(table_name, headers, rows):
    """Extract real content lines from a table — no meta-commentary."""
    bullets = extract_real_bullets(headers, rows, max_bullets=5)
    if bullets:
        return bullets

    # Fallback: numeric summary if no good text columns
    basis = derive_table_basis(headers, rows)
    numeric_index = basis.get("numeric_index")
    label_index = basis.get("label_index")
    lines = []
    if numeric_index is not None and label_index is not None:
        numeric_header = clean_text(headers[numeric_index], 40) or "valor"
        metric_kind = infer_metric_kind(numeric_header)
        ranked = []
        for row in rows:
            if label_index >= len(row) or numeric_index >= len(row):
                continue
            label = clean_text(row[label_index], 42)
            val = try_number(row[numeric_index])
            if label and not is_noise_text(label) and val is not None:
                ranked.append((label, val))
        ranked.sort(key=lambda x: x[1], reverse=True)
        for label, val in ranked[:5]:
            lines.append(f"{label}: {format_metric(val, metric_kind)}")
    return lines[:5]
```

- [ ] Commit: `git add generate_template_presentation.py && git commit -m "feat: replace build_table_summary_lines with real content extractor"`

---

### Task 3: Replace `build_native_sheet_highlights` with real content

**Files:**
- Modify: `generate_template_presentation.py:1223-1256`

- [ ] Replace the entire function:

```python
def build_native_sheet_highlights(sheet_type, table_name, headers, rows):
    """Extract real cell content from the sheet — no hardcoded sheet-type phrases."""
    bullets = extract_real_bullets(headers, rows, max_bullets=5)
    if bullets:
        return bullets

    # Fallback: show first non-empty cells from the most informative columns
    lines = []
    seen = set()
    for row in (rows or [])[:8]:
        for idx, header in enumerate(headers):
            if idx >= len(row):
                continue
            val = clean_text(row[idx], 90)
            if not val or is_noise_text(val) or len(val) < 8:
                continue
            key = val.casefold()
            if key in seen:
                continue
            seen.add(key)
            header_label = clean_text(header, 28)
            if header_label and not is_generic_header(header_label):
                lines.append(f"{header_label}: {val[:80]}")
            else:
                lines.append(val[:85])
            if len(lines) >= 4:
                return lines
    return lines
```

- [ ] Commit: `git add generate_template_presentation.py && git commit -m "feat: replace build_native_sheet_highlights with real content extractor"`

---

### Task 4: Replace `build_narrative_focus_lines` — remove filler prefix

**Files:**
- Modify: `generate_template_presentation.py:1780-1814`

- [ ] Replace the entire function:

```python
def build_narrative_focus_lines(table_name, headers, rows):
    """Real cell content from narrative tables — no 'La hoja X es mayoritariamente narrativa' prefix."""
    profile = build_table_signal_profile(headers, rows)
    if profile["table_kind"] != "narrative":
        return []
    return extract_real_bullets(headers, rows, max_bullets=6)
```

- [ ] Commit: `git add generate_template_presentation.py && git commit -m "feat: replace build_narrative_focus_lines with real content"`

---

### Task 5: Redesign `render_table_summary_slide` — show real bullets, no "Resumen ejecutivo"

**Files:**
- Modify: `generate_template_presentation.py:3230-3340` (render_table_summary_slide)

- [ ] Read the full function to get the render boundaries, then replace the body content section. The "Resumen ejecutivo de la tabla" header and generic lines need replacing with actual bullet points from the table data. Replace from the `text_frame.clear()` call onward:

```python
def render_table_summary_slide(slide, block, excel_path, content_index, total_blocks):
    add_content_title(slide, block["title"], block.get("subtitle"))

    left = SLIDE4["content_x"]
    top = SLIDE4["content_y"]
    width = SLIDE4["content_w"]
    trace_caption = build_traceability_caption(block.get("traceability"))

    basis_chip = slide.shapes.add_shape(1, left, top, Inches(4.3), Inches(0.34))
    basis_chip.fill.solid()
    basis_chip.fill.fore_color.rgb = RGBColor(245, 247, 250)
    basis_chip.line.color.rgb = COLOR_LINE
    basis_text = slide.shapes.add_textbox(left + Inches(0.12), top + Inches(0.05), Inches(4.02), Inches(0.18))
    basis_text.text_frame.text = clean_text(block.get("source_basis") or "Contenido real del Excel")
    style_text_frame(basis_text.text_frame, font_size=8.8, color=COLOR_BLUE, bold=True)

    if trace_caption:
        trace_box = slide.shapes.add_textbox(left + Inches(4.5), top + Inches(0.05), width - Inches(4.55), Inches(0.18))
        trace_box.text_frame.text = clean_text(trace_caption, 96)
        style_text_frame(trace_box.text_frame, font_size=7.6, color=COLOR_SOFT, bold=False, align=PP_ALIGN.RIGHT)

    accent = slide.shapes.add_shape(1, left, top + Inches(0.52), Inches(0.18), Inches(4.3))
    accent.fill.solid()
    accent.fill.fore_color.rgb = COLOR_ORANGE
    accent.line.color.rgb = COLOR_ORANGE

    lines = safe_list(block.get("lines"))
    content_box = slide.shapes.add_textbox(
        left + Inches(0.35), top + Inches(0.58), width - Inches(0.45), Inches(4.85)
    )
    tf = content_box.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(6)
    tf.margin_right = Pt(4)
    tf.margin_top = Pt(6)
    tf.margin_bottom = Pt(4)
    tf.clear()

    for line_index, line in enumerate(lines[:8]):
        p = tf.paragraphs[0] if line_index == 0 else tf.add_paragraph()
        p.text = clean_text(line, 110)
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(8)
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(13.5)
            run.font.color.rgb = COLOR_TEXT

    add_content_footer(slide, excel_path, content_index, total_blocks)
```

- [ ] Commit: `git add generate_template_presentation.py && git commit -m "feat: redesign table_summary slide to show real bullets instead of Resumen ejecutivo"`

---

### Task 6: Add Option C for narrative tables — bullets slide + table slide

**Files:**
- Modify: `generate_template_presentation.py` in `build_table_blocks`

- [ ] In `build_table_blocks`, when a narrative table has real bullets, generate a proper table slide (not just text). Find the section after `kpi_rendered = False` and add before the `if not detail_candidate: continue` block:

```python
        # Option C: for narrative tables with real content — show bullets + force table render
        if profile["table_kind"] == "narrative" and not detail_candidate:
            real_bullets = extract_real_bullets(headers, rows, max_bullets=6)
            if real_bullets and len(rows) <= 20 and len(headers) <= 6:
                # Force a table slide even for narrative content (with truncation)
                for subset_headers, subset_rows, label in split_table_columns(headers, rows):
                    rows_per_page = min(8, resolve_rows_per_table_page(subset_headers, subset_rows))
                    total_pages = max(1, math.ceil(len(subset_rows) / rows_per_page))
                    for page_index in range(min(total_pages, 2)):
                        chunk = subset_rows[page_index * rows_per_page:(page_index + 1) * rows_per_page]
                        if not chunk:
                            continue
                        page_label = f"Pagina {page_index + 1}/{total_pages}" if total_pages > 1 else ""
                        suffix = f"{label} · {page_label}".strip(" ·") if label or page_label else ""
                        table_basis = derive_table_basis(subset_headers, subset_rows)
                        blocks.append({
                            "type": "table",
                            "title": table_name,
                            "subtitle": suffix,
                            "headers": subset_headers,
                            "rows": chunk,
                            "source_basis": build_source_basis_text(table_basis["dimension_label"], table_basis["metric_label"], table_basis["aggregation"]),
                            "traceability": build_traceability(source_sheet, table_basis["aggregation"], truth["score"], row_count=len(chunk), col_count=len(subset_headers), numeric_coverage=truth["numeric_coverage"]),
                        })
                        if len(blocks) >= max_table_blocks:
                            return blocks
```

- [ ] Commit: `git add generate_template_presentation.py && git commit -m "feat: add Option C for narrative tables — bullets + table slide"`

---

### Task 7: Redesign Slide 3 synthesis — replace generic lines with real top values

**Files:**
- Modify: `generate_template_presentation.py:2374-2413` (`build_slide_3_sections`)

- [ ] Replace `build_slide_3_sections`:

```python
def build_slide_3_sections(data, excel_path):
    """Build slide 3 with real data insights — no generic meta-commentary."""

    # BREVE RESUMEN: real top values from charts and KPIs
    summary_lines = []
    for chart in safe_list(data.get("graficas_automaticas"))[:3]:
        insight = compact_line(chart.get("insight_auto"), 80)
        if insight and len(insight) >= 12:
            summary_lines.append(insight)
    for kpi in safe_list(data.get("kpis_automaticos"))[:3]:
        label = clean_text(kpi.get("label"), 36)
        val = kpi.get("value")
        if label and val is not None:
            kind = infer_metric_kind(label, [val] if isinstance(val, (int, float)) else [])
            formatted = format_metric(val, kind) if isinstance(val, (int, float)) else clean_text(str(val), 40)
            summary_lines.append(f"{label}: {formatted}")
    # Pull from derived charts too
    for block in build_derived_chart_blocks(data)[:2]:
        insight = compact_line(block.get("insight"), 80)
        if insight and len(insight) >= 12 and insight not in summary_lines:
            summary_lines.append(insight)

    # OBJETIVOS: derived from real column names of main table
    objectives_lines = []
    if data.get("muestra_tabla"):
        headers, rows = extract_table_payload(data.get("muestra_tabla"))
        real_headers = [clean_text(h, 42) for h in headers if h and not is_generic_header(h)]
        if real_headers:
            objectives_lines.append(f"Analizar: {' | '.join(real_headers[:4])}")
        bullets = extract_real_bullets(headers, rows, max_bullets=2)
        objectives_lines.extend(bullets)

    # ELEMENTOS PRIORITARIOS: real top rows/values from first table with data
    priority_lines = []
    for table_name, table, source_sheet in collect_table_candidates(data)[:3]:
        h, r = extract_table_payload(table)
        if not h or not r:
            continue
        bullets = extract_real_bullets(h, r, max_bullets=3)
        priority_lines.extend(bullets)
        if len(priority_lines) >= 4:
            break

    summary_lines = expand_lines_for_slide(unique_texts(summary_lines, limit=SLIDE3_MAX_BULLETS, min_len=10), max_len=76, max_items=4)
    objectives_lines = expand_lines_for_slide(unique_texts(objectives_lines, limit=SLIDE3_MAX_BULLETS, min_len=10), max_len=76, max_items=4)
    priority_lines = expand_lines_for_slide(unique_texts(priority_lines, limit=SLIDE3_MAX_BULLETS, min_len=10), max_len=76, max_items=4)

    return [
        ("BREVE RESUMEN", summary_lines or ["Ver datos en las diapositivas siguientes."]),
        ("COLUMNAS ANALIZADAS", objectives_lines or ["Datos extraidos directamente del Excel."]),
        ("DATOS DESTACADOS", priority_lines or ["Consultar tablas y graficas en el cuerpo de la presentacion."]),
    ]
```

- [ ] Commit: `git add generate_template_presentation.py && git commit -m "feat: redesign slide 3 synthesis with real top values from data"`

---

### Task 8: Redesign Slide 2 cards — replace meta-counts with real data

**Files:**
- Modify: `generate_template_presentation.py:2299-2370` (`build_slide_2_cards`)

- [ ] Replace `build_slide_2_cards`:

```python
def build_slide_2_cards(data, excel_path):
    """Build slide 2 with real content — sheet names, real columns, real KPIs."""
    metadata = data.get("metadatos") or {}
    sheet_names = safe_list(metadata.get("hojas_encontradas"))
    sheet_count = len(sheet_names)
    mode_label = "Boardroom" if IS_BOARDROOM_MODE else "Ejecutivo"

    # Card 1: real sheet names (not generic topic)
    sheets_preview = " | ".join(clean_text(s, 22) for s in sheet_names[:4] if s)
    if len(sheet_names) > 4:
        sheets_preview += f" (+{len(sheet_names) - 4})"
    card_1_body = f"Hojas: {sheets_preview}" if sheets_preview else f"{sheet_count} hojas detectadas"

    # Card 2: real column names from main table
    card_2_body = ""
    if data.get("muestra_tabla"):
        headers, rows = extract_table_payload(data.get("muestra_tabla"))
        real_cols = [clean_text(h, 28) for h in headers if h and not is_generic_header(h)]
        if real_cols:
            card_2_body = "Columnas: " + " | ".join(real_cols[:4])
    if not card_2_body:
        primary = clean_text((data.get("resumen_generico") or {}).get("hoja_principal"), 38)
        card_2_body = f"Hoja principal: {primary}" if primary else f"{sheet_count} hojas analizadas"

    # Card 3: top real KPI value
    card_3_body = f"{mode_label} con datos reales."
    for kpi in safe_list(data.get("kpis_automaticos"))[:2]:
        label = clean_text(kpi.get("label"), 32)
        val = kpi.get("value")
        if label and val is not None and isinstance(val, (int, float)):
            kind = infer_metric_kind(label, [val])
            card_3_body = f"{label}: {format_metric(val, kind)}"
            break
    # Fallback to chart insight
    if card_3_body == f"{mode_label} con datos reales.":
        for chart in safe_list(data.get("graficas_automaticas"))[:1]:
            insight = compact_line(chart.get("insight_auto"), 44)
            if insight:
                card_3_body = insight
                break

    # Card 4: first real data finding
    card_4_body = ""
    for table_name, table, _ in collect_table_candidates(data)[:2]:
        h, r = extract_table_payload(table)
        bullets = extract_real_bullets(h, r, max_bullets=1)
        if bullets:
            card_4_body = bullets[0][:80]
            break
    if not card_4_body:
        card_4_body = f"Base: {clean_text((data.get('metadatos') or {}).get('archivo'), 44)}"

    return [
        {
            "badge": "01",
            "title": "Hojas detectadas",
            "body": card_1_body,
            "x": Inches(1.2), "y": Inches(1.85), "w": Inches(3.55), "h": Inches(1.2),
        },
        {
            "badge": "02",
            "title": "Datos analizados",
            "body": card_2_body,
            "x": Inches(6.15), "y": Inches(1.85), "w": Inches(3.55), "h": Inches(1.2),
        },
        {
            "badge": "03",
            "title": "Indicador clave",
            "body": card_3_body,
            "x": Inches(1.2), "y": Inches(3.55), "w": Inches(3.55), "h": Inches(1.2),
        },
        {
            "badge": "04",
            "title": "Dato destacado",
            "body": card_4_body,
            "x": Inches(6.15), "y": Inches(3.55), "w": Inches(3.55), "h": Inches(1.2),
        },
    ]
```

- [ ] Commit: `git add generate_template_presentation.py && git commit -m "feat: redesign slide 2 cards with real sheet names, columns and KPIs"`

---

### Task 9: Fix `build_technical_highlights` — remove generic count lines

**Files:**
- Modify: `generate_template_presentation.py:1280-1356` (`build_technical_highlights`)

- [ ] Replace lines that produce generic counts with real data. Replace the lines block assembly:

```python
def build_technical_highlights(data, excel_path, limit=4):
    """Real technical highlights — sheet names and real column names, not row counts."""
    metadata = data.get("metadatos") or {}
    lines = []

    # Sheet names are real data
    sheet_names = safe_list(metadata.get("hojas_encontradas"))
    if sheet_names:
        names_str = " | ".join(clean_text(s, 24) for s in sheet_names[:5] if s)
        if names_str:
            lines.append(f"Hojas: {names_str}")

    # Real column names from main table
    if data.get("muestra_tabla"):
        headers, rows = extract_table_payload(data.get("muestra_tabla"))
        real_cols = [clean_text(h, 32) for h in headers if h and not is_generic_header(h)]
        if real_cols:
            lines.append("Columnas: " + " | ".join(real_cols[:5]))

    # Real KPI if available
    for kpi in safe_list(data.get("kpis_automaticos"))[:2]:
        label = clean_text(kpi.get("label"), 36)
        val = kpi.get("value")
        if label and val is not None and isinstance(val, (int, float)):
            kind = infer_metric_kind(label, [val])
            lines.append(f"{label}: {format_metric_full(val, kind)}")

    return unique_texts([compact_line(line, 84) for line in lines if line], limit=limit, min_len=10)
```

- [ ] Commit: `git add generate_template_presentation.py && git commit -m "feat: replace build_technical_highlights generic counts with real column and KPI data"`

---

### Task 10: Test full pipeline with FONDOS FIJOS and a generic Excel

- [ ] Run with audit Excel:
```bash
cd "C:\Users\cpinzon\Desktop\generador-pptx-socya-main ultimo"
python -X utf8 generate_template_presentation.py "FONDOS FIJOS  (1).xlsx" test_final.pptx
python -m markitdown test_final.pptx | grep -v "^$" | head -120
```
Expected: No lines containing "La tabla concentra", "dataset contiene", "principalmente documental", "evitando sobrecargar"

- [ ] Create and run with a generic sales Excel:
```python
import pandas as pd
df = pd.DataFrame({
    'Producto': ['A','B','C','D','E'],
    'Ventas': [500000,350000,280000,190000,120000],
    'Region': ['Norte','Sur','Centro','Norte','Sur']
})
df.to_excel('test_generic.xlsx', index=False)
```
```bash
python -X utf8 generate_template_presentation.py test_generic.xlsx test_generic.pptx
python -m markitdown test_generic.pptx | grep -v "^$" | head -80
```
Expected: Real product names and values visible in slides, no generic meta-text

- [ ] Clean up test files:
```bash
del test_final.pptx test_final.audit.json test_generic.xlsx test_generic.pptx test_generic.audit.json
```

- [ ] Final commit:
```bash
git add -A
git commit -m "feat: complete real-content slides — all generic boilerplate replaced with real Excel data"
```
