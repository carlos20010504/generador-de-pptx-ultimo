import pptxgen from 'pptxgenjs';

export interface SocyaSlideJSON {
  type: "title" | "kpi_row" | "chart" | "table" | "text_bullets" | "closing";
  title: string;
  subtitle?: string;
  content: any; // specific per type
  detail_link?: string;
}

// THEME / PALETTE — Same Socya brand colors
const BRAND = {
  primary: '6BBF1F',
  dark: '4A8A0F',
  amber: 'F5A623',
  red: 'E74C3C',
  blue: '3498DB',
  purple: '9B59B6',
  white: 'FFFFFF',
  text: '333333',
  muted: '777777',
  bgLight: 'F9f9f9',
  border: 'E2E8F0',
  shadow: 'D1D5DB'
};

const KPI_COLORS = [BRAND.primary, BRAND.amber, BRAND.red, BRAND.blue, BRAND.purple];

// Badge styles mapping
const BADGE_STYLES: Record<string, { bg: string; text: string }> = {
  'CONTABILIZADO': { bg: 'D1FAE5', text: '065F46' },
  'LEGALIZADO': { bg: 'DBEAFE', text: '1E40AF' },
  'RECHAZADO': { bg: 'FEE2E2', text: '991B1B' },
  'SOLICITADO': { bg: 'E0F2FE', text: '0C4A6E' },
  'PENDIENTE': { bg: 'FEF3C7', text: '92400E' },
  'COMPLETADO': { bg: 'D1FAE5', text: '065F46' },
  'EN EJECUCION': { bg: 'DBEAFE', text: '1E40AF' },
  'EN PROCESO': { bg: 'FEF3C7', text: '92400E' },
  'ALTO': { bg: 'FEE2E2', text: '991B1B' },
  'MEDIO': { bg: 'FEF3C7', text: '92400E' },
  'BAJO': { bg: 'D1FAE5', text: '065F46' },
  'SI': { bg: 'D1FAE5', text: '065F46' },
  'SÍ': { bg: 'D1FAE5', text: '065F46' },
  'NO': { bg: 'FEE2E2', text: '991B1B' },
};

function getBadgeStyle(val: string) {
  const key = String(val).toUpperCase().trim();
  if (BADGE_STYLES[key]) return BADGE_STYLES[key];
  
  // Fuzzy match
  for (const [k, style] of Object.entries(BADGE_STYLES)) {
    if (key.includes(k) || k.includes(key)) return style;
  }
  return null;
}

// DIMENSIONS
const W = 13.33;
const H = 7.5;

function addCornerDecorations(slide: any) {
  slide.addShape('ellipse', { x: -0.5, y: -0.5, w: 1.5, h: 1.5, line: { color: 'FFFFFF', width: 2 }, fill: { transparency: 100 } });
  slide.addShape('ellipse', { x: W - 1, y: H - 1, w: 1.5, h: 1.5, line: { color: 'FFFFFF', width: 2 }, fill: { transparency: 100 } });

  for (let i = 0; i < 5; i++) {
    for (let j = 0; j < 5; j++) {
      if (i + j < 6) {
        slide.addShape('ellipse', { x: W - 1.5 + (i * 0.2), y: 0.2 + (j * 0.2), w: 0.05, h: 0.05, fill: { color: 'FFFFFF' } });
      }
    }
  }
}

function addProfessionalMotifs(slide: any) {
  // Diagonal stripe in background (very light)
  slide.addShape('rect', { 
    x: -2, y: H - 2, w: W + 4, h: 1, 
    fill: { color: 'FFFFFF', transparency: 90 }, 
    rotate: -15 
  });
  
  // Subtle white lines
  slide.addShape('line', { x: 0.6, y: H - 0.4, w: 4, h: 0, line: { color: BRAND.primary, width: 1, transparency: 50 } });
}

function addSideBanner(slide: any, slideIndex: number) {
  slide.addShape('rect', { x: 0, y: 0, w: 0.4, h: H, fill: { color: BRAND.primary } });
  slide.addText("SOCYA - PRESENTACIÓN INTELIGENTE", { 
    x: -3.5, y: H / 2, w: H, h: 0.4, 
    rotate: -90, 
    fontSize: 10, color: 'FFFFFF', bold: true, align: 'center', fontFace: 'Calibri'
  });
  
  // Slide number at bottom left
  slide.addText(String(slideIndex), {
    x: 0, y: H - 0.4, w: 0.4, h: 0.4,
    fontSize: 10, color: 'FFFFFF', bold: true, align: 'center', fontFace: 'Calibri'
  });
}

function addLogoAndHeader(slide: any, title: string, subtitle?: string) {
  // Logo placeholder - Leaf-like shape + Text
  slide.addShape('ellipse', { x: W - 1.85, y: 0.25, w: 0.7, h: 0.7, fill: { color: BRAND.white }, line: { color: BRAND.primary, width: 2 } });
  slide.addShape('ellipse', { x: W - 1.78, y: 0.32, w: 0.55, h: 0.55, fill: { color: BRAND.primary } });
  
  slide.addText('SOCYA', { 
    x: W - 1.3, y: 0.35, w: 1.2, h: 0.5, 
    color: BRAND.dark, bold: true, fontSize: 18, align: 'left', fontFace: 'Calibri' 
  });

  // Header title
  slide.addText(title || "Sección", { 
    x: 0.6, y: 0.3, w: W - 3, h: 0.6, 
    fontSize: 32, bold: true, color: BRAND.dark, fontFace: 'Calibri' 
  });
  // Green divider line - thicker and more modern
  slide.addShape('rect', { x: 0.6, y: 0.95, w: 5, h: 0.08, fill: { color: BRAND.primary } });
  
  // Subtitle if provided
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.6, y: 1.1, w: W - 3, h: 0.35,
      fontSize: 14, color: BRAND.muted, italic: true, fontFace: 'Calibri'
    });
  }
}

function addFooter(slide: any) {
  const now = new Date().toLocaleDateString('es-CO');
  slide.addText(`© ${new Date().getFullYear()} Fundación Socya | Reporte de Auditoría | Generado: ${now}`, {
    x: 0.6, y: H - 0.4, w: W - 1, h: 0.3,
    fontSize: 8, color: BRAND.muted, align: 'right', fontFace: 'Calibri'
  });
}

function addDetailLink(slide: any, link?: string) {
  if (!link) return;
  slide.addShape('roundRect', { x: W / 2 - 1.5, y: H - 0.7, w: 3, h: 0.4, fill: { color: BRAND.bgLight }, line: { color: BRAND.primary }, rectRadius: 0.1 });
  slide.addText("Ver datos completos →", { 
    x: W / 2 - 1.5, y: H - 0.7, w: 3, h: 0.4, 
    fontSize: 12, bold: true, color: BRAND.primary, align: 'center',
    hyperlink: { url: link, tooltip: "Ver Excel completo" }
  });
}

// RENDERERS
export function renderSocyaSlide(prs: pptxgen, slideData: SocyaSlideJSON, slideIndex: number) {
  const slide = prs.addSlide();

  // title y closing — same design
  if (slideData.type === 'title' || slideData.type === 'closing') {
    slide.background = { color: BRAND.primary };
    addCornerDecorations(slide);
    addProfessionalMotifs(slide);

    if (slideData.type === 'title') {
      // White circle logo
      slide.addShape('ellipse', { x: W / 2 - 1.2, y: 1.2, w: 2.4, h: 2.4, fill: { color: 'FFFFFF' } });
      slide.addText('SOCYA', { x: W / 2 - 1.2, y: 1.2, w: 2.4, h: 2.4, fontSize: 28, bold: true, color: BRAND.primary, align: 'center', valign: 'middle', fontFace: 'Calibri' });
      
      slide.addText(slideData.title, { x: 1, y: 4.2, w: W - 2, h: 1.2, fontSize: 44, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Calibri' });
      if (slideData.subtitle) {
        slide.addText(slideData.subtitle, { x: 1, y: 5.5, w: W - 2, h: 0.6, fontSize: 22, color: 'FFFFFF', align: 'center', fontFace: 'Calibri', italic: true });
      }
      
      // Decorative bottom line
      slide.addShape('rect', { x: W / 2 - 2, y: 6.5, w: 4, h: 0.05, fill: { color: 'FFFFFF' } });
    } else {
      slide.addText(slideData.title || "¡Gracias!", { x: 1, y: 2.5, w: W - 2, h: 1.5, fontSize: 48, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Calibri' });
      if (slideData.subtitle) {
        slide.addText(slideData.subtitle, { x: 1, y: 4.0, w: W - 2, h: 0.8, fontSize: 24, color: 'FFFFFF', align: 'center', fontFace: 'Calibri' });
      }
      
      slide.addShape('rect', { x: W / 2 - 1.5, y: 5.5, w: 3, h: 0.05, fill: { color: 'FFFFFF' } });
    }
    return;
  }

  // Common inner pages — same design
  slide.background = { color: 'FFFFFF' };
  addSideBanner(slide, slideIndex);
  addLogoAndHeader(slide, slideData.title, slideData.subtitle);
  addFooter(slide);

  const contentStartY = slideData.subtitle ? 1.5 : 1.3;

  if (slideData.type === 'kpi_row') {
    const kpis = slideData.content || [];
    const count = Math.min(kpis.length, 5);
    const cardW = 2.2;
    const spacing = (W - 1.2 - (cardW * count)) / (count + 1);
    
    kpis.slice(0, 5).forEach((kpi: any, idx: number) => {
      const cx = 0.8 + spacing + idx * (cardW + spacing);
      const cy = 2.2;
      const cardColor = KPI_COLORS[idx % KPI_COLORS.length];
      
      // Card Shadow
      slide.addShape('roundRect', { x: cx + 0.05, y: cy + 0.05, w: cardW, h: 2.8, fill: { color: BRAND.shadow, transparency: 80 }, rectRadius: 0.1 });
      
      // Main Card
      slide.addShape('roundRect', { x: cx, y: cy, w: cardW, h: 2.8, fill: { color: cardColor }, rectRadius: 0.1 });
      
      // Top accent stripe in white
      slide.addShape('rect', { x: cx + 0.2, y: cy + 0.2, w: cardW - 0.4, h: 0.02, fill: { color: BRAND.white, transparency: 50 } });

      slide.addText(String(kpi.value), { 
        x: cx, y: cy + 0.4, w: cardW, h: 1.0, 
        fontSize: 42, bold: true, color: 'FFFFFF', align: 'center', fontFace: 'Calibri' 
      });
      
      slide.addText(String(kpi.label).toUpperCase(), { 
        x: cx + 0.1, y: cy + 1.5, w: cardW - 0.2, h: 0.6, 
        fontSize: 12, bold: true, color: 'FFFFFF', align: 'center', wrap: true, fontFace: 'Calibri'
      });
      
      // Subtitle under the label
      if (kpi.subtitle) {
        slide.addShape('rect', { x: cx + 0.5, y: cy + 2.1, w: cardW - 1, h: 0.01, fill: { color: BRAND.white, transparency: 70 } });
        slide.addText(String(kpi.subtitle), { 
          x: cx + 0.1, y: cy + 2.2, w: cardW - 0.2, h: 0.4, 
          fontSize: 10, color: 'FFFFFF', align: 'center', italic: true, wrap: true 
        });
      }
    });
  }
  else if (slideData.type === 'chart') {
    const ch = slideData.content;
    if (ch && ch.labels && ch.values) {
      const chartType = ch.barDir === 'bar' ? prs.ChartType.bar : prs.ChartType.bar;
      slide.addChart(chartType, [{ name: ch.name || "Dato", labels: ch.labels, values: ch.values }], {
        x: 1, y: contentStartY + 0.3, w: W - 2, h: H - contentStartY - 1,
        chartColors: KPI_COLORS,
        showValue: true, barDir: ch.barDir === 'bar' ? 'bar' : 'col',
        valAxisLabelColor: BRAND.text, catAxisLabelColor: BRAND.text,
        catAxisLabelFontSize: 10,
        valAxisLabelFontSize: 9,
      });
    }
  }
  else if (slideData.type === 'table') {
    const tbl = slideData.content;
    if (tbl && tbl.headers && tbl.rows) {
      // Strip purely 'Col_X' generic columns
      const validIndices = tbl.headers.map((h: string, idx: number) => h.startsWith('Col_') ? -1 : idx).filter((i: number) => i !== -1);
      
      let displayHeaders = validIndices.map((i: number) => tbl.headers[i]);
      // Show up to 12 rows per slide (increased from 8)
      const maxRows = 12;
      let displayRows = tbl.rows.slice(0, maxRows).map((row: any[]) => validIndices.map((i: number) => row[i]));

      // Calculate dynamic column widths based on number of columns
      const numCols = displayHeaders.length;
      const tableW = W - 1.6;
      
      // Determine available height for table
      const hasDetailLink = !!slideData.detail_link;
      const tableStartY = contentStartY + 0.3;
      const tableEndY = hasDetailLink ? H - 1 : H - 0.4;
      const availableH = tableEndY - tableStartY;
      
      // Dynamic row height based on row count
      const totalRows = displayRows.length + 1; // +1 for header
      const rowH = Math.min(0.4, Math.max(0.28, availableH / totalRows));
      
      // Dynamic font sizes based on column count
      const headerFontSize = numCols > 5 ? 10 : 12;
      const cellFontSize = numCols > 5 ? 8 : 9;
      const cellTruncate = numCols > 5 ? 40 : 55;

      // Table container shadow (subtle)
      slide.addShape('rect', { x: 0.82, y: tableStartY + 0.02, w: tableW, h: availableH, fill: { color: BRAND.shadow, transparency: 80 } });

      const formattedRows = [
        displayHeaders.map((h: string) => ({ 
          text: h, 
          options: { bold: true, fill: BRAND.primary, color: 'FFFFFF', fontSize: headerFontSize, align: 'center' as const, wrap: true } 
        })),
        ...displayRows.map((row: any[], rIdx: number) => 
          row.map((cell: any, cIdx: number) => {
            const cellText = String(cell ?? '—').substring(0, cellTruncate);
            const badge = getBadgeStyle(cellText);
            
            // If it's a badge-like column (usually the last one or named Estado)
            const isStatusCol = displayHeaders[cIdx].toLowerCase().includes('estado') || 
                               displayHeaders[cIdx].toLowerCase().includes('confirm');
            
            if (isStatusCol && badge) {
              return { 
                text: cellText, 
                options: { 
                  fill: badge.bg, 
                  color: badge.text, 
                  fontSize: cellFontSize, 
                  bold: true,
                  align: 'center' as const,
                  wrap: true 
                } 
              };
            }

            return { 
              text: cellText, 
              options: { 
                fill: rIdx % 2 === 0 ? 'F5F5F5' : 'FFFFFF', 
                color: BRAND.text, 
                fontSize: cellFontSize, 
                wrap: true 
              } 
            };
          })
        )
      ];
      
      slide.addTable(formattedRows, { 
        x: 0.8, y: tableStartY, w: tableW, 
        rowH: rowH, 
        border: { pt: 1, color: BRAND.border },
        autoPage: false
      });
    }
  }
  else if (slideData.type === 'text_bullets') {
    const items = slideData.content || [];
    
    items.slice(0, 6).forEach((item: string, idx: number) => {
      const col = idx % 2;
      const row = Math.floor(idx / 2);
      const cx = 1.2 + col * 5.8;
      const cy = contentStartY + 0.4 + row * 1.6;
      
      // Card aspect for bullet point
      slide.addShape('roundRect', { x: cx, y: cy, w: 5.2, h: 1.3, fill: { color: BRAND.bgLight }, line: { color: BRAND.border }, rectRadius: 0.05 });
      
      // Icon Circle
      slide.addShape('ellipse', { x: cx + 0.2, y: cy + 0.2, w: 0.4, h: 0.4, fill: { color: BRAND.primary } });
      slide.addText('✓', { x: cx + 0.2, y: cy + 0.2, w: 0.4, h: 0.4, color: 'FFFFFF', bold: true, align: 'center', fontSize: 14, fontFace: 'Calibri' });
      
      // Text
      slide.addText(item, { 
        x: cx + 0.7, y: cy + 0.15, w: 4.3, h: 1.0, 
        fontSize: 13, color: BRAND.text, wrap: true, valign: 'middle', fontFace: 'Calibri' 
      });
    });
  }

  // Add detail link if exists — same design
  if (slideData.detail_link) {
    addDetailLink(slide, slideData.detail_link);
  }
}

export function buildSocyaPresentation(prs: pptxgen, slides: SocyaSlideJSON[]) {
  slides.forEach((s, idx) => renderSocyaSlide(prs, s, idx + 1));
}
