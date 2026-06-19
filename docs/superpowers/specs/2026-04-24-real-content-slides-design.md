# Real Content Slides — Design Spec
Date: 2026-04-24

## Problem
The PPTX generator produces slides filled with generic meta-commentary ("La tabla concentra 6 filas utiles...", "El dataset contiene 42 registros...") instead of real data from the Excel. This happens for any Excel type, not just audits.

## Goal
Every slide must contain ONLY real information extracted from Excel cells. No meta-commentary, no structural descriptions, no boilerplate.

## What Gets Removed
- `build_table_summary_lines` generic text about table structure
- `build_native_sheet_highlights` hardcoded sheet-type phrases
- `build_intro_briefing` meta-counts and hardcoded topic labels
- Slide 2/3 generic objectives and planning text
- Any text describing the data instead of showing it

## What Replaces It

### 1. `extract_real_bullets(headers, rows, max=6)` — New core function
Extracts real cell content as formatted bullets using these rules:
- Find the most informative column (highest unique variety, 10-80 char values)
- If table has short col + long col: format as `[Short] Long description`
- If table has category + metric: format as `Category: formatted_value`
- Skip any cell with noise values (—, N/A, <8 chars)
- Truncate bullets at 85 chars

### 2. `render_table_summary_slide` redesign
Replace generic "resumen ejecutivo" text with:
- Header: source basis (already real)
- Body: `extract_real_bullets()` output — actual cell content
- No meta-commentary lines at all

### 3. Slide 2 (description) — real content only
- Card 1: sheet names from the Excel (real)
- Card 2: main column names + top KPI value (real)
- Card 3: dominant category + % (real, e.g. "Centro: 36% de Ventas")
- Card 4: first real insight from the data

### 4. Slide 3 (synthesis) — real insights
- BREVE RESUMEN: top 2 real values/findings from the data
- OBJETIVOS: derived from actual column names (not hardcoded audit phrases)
- ELEMENTOS PRIORITARIOS: real top rows/values, not structural counts

### 5. Narrative tables → Option C (2 slides)
- Slide A: `extract_real_bullets()` — real cell content as bullets
- Slide B: PPTX table with actual rows (same as operational tables)

### 6. `build_native_sheet_highlights` — replaced
Instead of "La hoja debe escalar solo hallazgos...", extract actual first meaningful cells from the sheet as context lines.

## Universal Rule
If there is no real content to show for a slide → don't generate it. Better fewer slides than slides with filler.

## Files Changed
- `generate_template_presentation.py`: all text generation functions
- `organizer.py`: topic inference (remove hardcoded labels)
