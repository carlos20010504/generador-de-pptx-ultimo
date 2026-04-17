// @ts-nocheck

// ─── Slide Dimensions ─────────────────────────────────────────────────────────
const W   = 13.33;  // Width (widescreen)
const H   = 7.5;    // Height
const MX  = 0.5;    // Margin X
const MW  = W - MX * 2;  // Content width
const CH  = 1.35;   // Content Y start (below header)
const FY  = H - 0.42;   // Footer Y
const MAX_Y = H - 0.5;  // Max content Y
const TABLE_W = 12.3;   // Max table width

// ─── Helpers ──────────────────────────────────────────────────────────────────

function hex(c) { return c ? String(c).replace('#', '') : 'FFFFFF'; }

function safeStr(v) {
  if (v === null || v === undefined) return '—';
  if (typeof v === 'object' && 'value' in v) return String(v.icon ? `${v.icon} ${v.value}` : v.value);
  return String(v);
}

// ─── Theme-specific header drawing ────────────────────────────────────────────

function drawHeaderDark(slide, title, subtitle, theme) {
  if (theme.sidebarW > 0) {
    slide.addShape('rect', {
      x: 0, y: 0, w: theme.sidebarW, h: H,
      fill: { color: hex(theme.primary) },
    });
  }
  slide.addShape('rect', {
    x: 0, y: 0, w: W, h: 1.15,
    fill: { color: hex(theme.footerBg) },
  });
  slide.addText(String(title || ''), {
    x: MX + theme.sidebarW, y: 0.1,
    w: MW - theme.sidebarW - 0.3, h: 0.75,
    fontSize: 24, bold: true, color: hex(theme.text),
    fontFace: 'Calibri', valign: 'middle',
  });
  if (subtitle) {
    slide.addText(String(subtitle), {
      x: MX + theme.sidebarW, y: 0.85,
      w: MW - 0.3, h: 0.26,
      fontSize: 10, color: hex(theme.textMuted),
      fontFace: 'Calibri', italic: true,
    });
  }
}

function drawHeaderLight(slide, title, subtitle, theme) {
  slide.addShape('rect', {
    x: 0, y: 0, w: W, h: 0.08,
    fill: { color: hex(theme.primary) },
  });
  slide.addText(String(title || ''), {
    x: MX, y: 0.15, w: MW, h: 0.78,
    fontSize: 26, bold: true, color: hex(theme.primary),
    fontFace: 'Calibri', valign: 'middle',
  });
  if (subtitle) {
    slide.addText(String(subtitle), {
      x: MX, y: 0.92, w: MW, h: 0.3,
      fontSize: 11, color: hex(theme.textMuted),
      fontFace: 'Calibri', italic: true,
    });
  }
}

function drawHeader(slide, title, subtitle, theme) {
  if (theme.isDark) drawHeaderDark(slide, title, subtitle, theme);
  else drawHeaderLight(slide, title, subtitle, theme);
}

function drawFooter(slide, theme, meta) {
  slide.addShape('rect', {
    x: 0, y: FY, w: W, h: H - FY,
    fill: { color: hex(theme.footerBg) },
  });
  slide.addText(String(meta || ''), {
    x: MX, y: FY + 0.04, w: 7, h: 0.3,
    fontSize: 8, color: hex(theme.footerText), fontFace: 'Calibri',
  });
  slide.slideNumber = {
    x: W - 1.2, y: FY + 0.04, w: 1.0, h: 0.3,
    color: hex(theme.footerText), fontSize: 8, align: 'right',
  };
}

function makeBg(slide, theme) {
  slide.background = { color: hex(theme.slideBg) };
}

// ─── Transitions & Decorations (unique per generation) ──────────────────────

const SLIDE_TRANSITIONS = [
  { type: 'fade', speed: 'slow' },
  { type: 'fade', speed: 'med' },
  { type: 'push', speed: 'med' },
  { type: 'cover', speed: 'fast' },
  { type: 'wipe', speed: 'med' },
  { type: 'push', speed: 'fast' },
  { type: 'fade', speed: 'fast' },
];

function applySlideEnhancements(slide, theme) {
  // Random transition
  slide.transition = SLIDE_TRANSITIONS[Math.floor(Math.random() * SLIDE_TRANSITIONS.length)];

  // 25% chance of subtle decorative accent (varies per slide)
  const r = Math.random();
  if (r < 0.12) {
    slide.addShape('ellipse', {
      x: W - 1.5, y: -0.3, w: 1.2, h: 1.2,
      fill: { color: hex(theme.primary), transparency: 92 },
    });
  } else if (r < 0.25) {
    slide.addShape('rect', {
      x: W - 2.5, y: MAX_Y + 0.15, w: 2, h: 0.04,
      fill: { color: hex(theme.secondary), transparency: 70 },
    });
  }
}

// ─── KPI Cards ────────────────────────────────────────────────────────────────

function renderKPIs(slide, kpis, y, theme) {
  if (!kpis || !kpis.length) return y;
  const count = Math.min(kpis.length, 5);
  const kw = (MW - (count - 1) * 0.18) / count;

  kpis.slice(0, count).forEach((kpi, i) => {
    const kx = MX + i * (kw + 0.18);
    slide.addShape('roundRect', {
      x: kx, y, w: kw, h: 0.88,
      fill: { color: hex(theme.kpiCardBg) },
      line: { color: hex(theme.divider), width: 0.5 },
      rectRadius: 0.07,
    });
    // Accent bar top
    slide.addShape('rect', {
      x: kx, y, w: kw, h: 0.06,
      fill: { color: hex(theme.primary) },
    });
    slide.addText(String(kpi.value ?? '—'), {
      x: kx + 0.06, y: y + 0.1, w: kw - 0.12, h: 0.44,
      fontSize: 18, bold: true, color: hex(theme.primary),
      align: 'center', fontFace: 'Calibri',
    });
    slide.addText(String(kpi.label ?? ''), {
      x: kx + 0.06, y: y + 0.55, w: kw - 0.12, h: 0.28,
      fontSize: 8, color: hex(theme.textMuted),
      align: 'center', fontFace: 'Calibri',
    });
  });
  return y + 1.02;
}

// ─── Table Renderer ───────────────────────────────────────────────────────────

function renderTable(slide, columns, colW, rows, y, theme, overflow) {
  if (!rows?.length || !columns?.length) return y;

  const numCols = columns.length;
  const safeColW = colW && colW.length === numCols
    ? colW
    : Array(numCols).fill(TABLE_W / numCols);

  const tRows = [];

  // Header row
  tRows.push(columns.map(c => ({
    text: String(c.header || c.key || ''),
    options: {
      bold: true,
      color: hex(theme.tableHeaderText),
      fill: { color: hex(theme.tableHeaderBg) },
      fontSize: 9.5,
      align: 'center',
      fontFace: 'Calibri',
      valign: 'middle',
    },
  })));

  // Data rows
  rows.forEach((r, i) => {
    const bgColor = i % 2 === 0 ? hex(theme.tableRow1) : hex(theme.tableRow2);
    tRows.push(columns.map(c => {
      const val = r[c.key];
      let txt = safeStr(val);
      let fill = bgColor;
      let txtColor = hex(theme.tableText);
      let align = 'left';

      if (typeof val === 'object' && val !== null && 'bg' in val) {
        fill = hex(val.bg) || bgColor;
        txtColor = hex(val.text) || hex(theme.tableText);
        align = 'center';
        txt = val.icon ? `${val.icon} ${val.value}` : String(val.value ?? '');
      }
      if (['currency', 'numeric', 'percent'].includes(c.type)) align = 'right';
      if (c.type === 'id') align = 'center';

      return {
        text: txt,
        options: {
          fill: { color: fill },
          color: txtColor,
          fontSize: 8.5,
          align,
          fontFace: 'Calibri',
          valign: 'middle',
        },
      };
    }));
  });

  const rowH = 0.27;
  const tblH = rowH * tRows.length;
  const maxAvailH = MAX_Y - y - 0.3;
  const maxRows = Math.min(tRows.length, Math.floor(maxAvailH / rowH));
  const finalRows = tRows.slice(0, Math.max(2, maxRows));
  const finalY = Math.min(y, MAX_Y - tblH - 0.1);

  try {
    slide.addTable(finalRows, {
      x: MX,
      y: Math.max(CH, finalY),
      w: TABLE_W,
      colW: safeColW,
      rowH,
      border: { pt: 0.5, color: hex(theme.divider) },
    });
    if (overflow > 0) {
      slide.addText(`+${overflow.toLocaleString('es-CO')} registros adicionales en el archivo original`, {
        x: MX,
        y: Math.max(CH, finalY) + finalRows.length * rowH + 0.06,
        w: TABLE_W,
        h: 0.22,
        fontSize: 8,
        italic: true,
        color: hex(theme.textMuted),
        align: 'right',
      });
    }
    return Math.max(CH, finalY) + finalRows.length * rowH + 0.18;
  } catch (e) {
    console.error('[Renderer] renderTable error:', e);
    return y;
  }
}

// ─── Chart Renderer ───────────────────────────────────────────────────────────

function renderChart(prs, slide, ch, x, y, w, h, theme) {
  if (!ch || !ch.data || !ch.data.length) return;
  const clean = ch.data.filter(d => !isNaN(Number(d.value)) && Number(d.value) !== 0);
  if (!clean.length) return;

  const labels = clean.map(d => String(d.label ?? ''));
  const values = clean.map(d => Number(d.value) || 0);
  const colors = clean.map((d, i) =>
    hex(d.color || theme.chartColors[i % theme.chartColors.length])
  );

  let cy = y;
  let ch2 = h;

  if (ch.title) {
    slide.addText(String(ch.title), {
      x, y: cy, w, h: 0.3,
      fontSize: 10, bold: true,
      color: hex(theme.primary),
      align: 'center',
      fontFace: 'Calibri',
    });
    cy += 0.33;
    ch2 -= 0.33;
  }

  const isDoughnut = ch.type === 'donut' || ch.type === 'doughnut';
  const isLine = ch.type === 'line' || ch.type === 'area';
  const isHBar = ch.type === 'bar_horizontal';

  const ctype = isDoughnut
    ? prs.ChartType.doughnut
    : isLine
      ? prs.ChartType.line
      : prs.ChartType.bar;

  try {
    slide.addChart(ctype, [{ name: 'Data', labels, values }], {
      x, y: cy, w, h: ch2,
      chartColors: colors,
      showValue: true,
      dataLabelFontSize: 8,
      dataLabelColor: theme.isDark ? 'FFFFFF' : '111827',
      showLegend: isDoughnut || clean.length <= 8,
      legendPos: isDoughnut ? 'r' : 'b',
      legendFontSize: 8,
      legendFontColor: hex(theme.textMuted),
      barDir: isHBar ? 'bar' : 'col',
      holeSize: isDoughnut ? 60 : 0,
      valAxisLabelColor: hex(theme.textMuted),
      catAxisLabelColor: hex(theme.textMuted),
      valAxisLineShow: false,
      catAxisLineShow: !isLine,
      gridLineDashType: isDoughnut ? undefined : 'sysDot',
      dataLabelFormatCode: values.some(v => v > 10000) ? '#,##0' : undefined,
    });
  } catch (e) {
    console.error('[Renderer] renderChart error:', ch.type, e.message);
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE RENDERERS
// ══════════════════════════════════════════════════════════════════════════════

// ─── COVER — 6 unique designs ─────────────────────────────────────────────────

function renderCover(prs, s, theme) {
  const slide = prs.addSlide();
  slide.transition = SLIDE_TRANSITIONS[Math.floor(Math.random() * SLIDE_TRANSITIONS.length)];
  const tid = theme.id;

  if (tid === 'dark_professional') {
    slide.addShape('rect', { x: 0, y: 0, w: 4.2, h: H, fill: { color: '000000' } });
    slide.addShape('rect', { x: 4.2, y: 0, w: W - 4.2, h: H, fill: { color: hex(theme.coverBg) } });
    slide.addShape('rect', { x: 4.1, y: 0, w: 0.18, h: H, fill: { color: hex(theme.primary) } });
    slide.addShape('ellipse', { x: -1, y: -1, w: 3.5, h: 3.5, fill: { color: hex(theme.primary), transparency: 88 } });
    slide.addShape('ellipse', { x: 0.5, y: H - 2, w: 2, h: 2, fill: { color: hex(theme.secondary), transparency: 85 } });
    slide.addText(String(s.title || ''), { x: 4.6, y: 1.8, w: 8.3, h: 2, fontSize: 40, bold: true, color: hex(theme.text), fontFace: 'Calibri', wrap: true });
    if (s.subtitle) slide.addText(String(s.subtitle), { x: 4.6, y: 3.95, w: 8.3, h: 0.6, fontSize: 15, color: hex(theme.textMuted), fontFace: 'Calibri', italic: true });
    const orgLine = [s.organization, s.date].filter(Boolean).join('   ·   ');
    if (orgLine) slide.addText(orgLine, { x: 4.6, y: H - 1.3, w: 8.3, h: 0.4, fontSize: 12, color: hex(theme.textMuted) });
    slide.addText(String(s.organization || 'Socya'), { x: 0.3, y: H / 2 - 0.8, w: 3.5, h: 1.6, fontSize: 20, bold: true, color: hex(theme.primary), align: 'center' });

  } else if (tid === 'minimal_white') {
    slide.background = { color: 'FFFFFF' };
    slide.addShape('rect', { x: 0, y: 0, w: W, h: 0.12, fill: { color: hex(theme.primary) } });
    slide.addShape('rect', { x: 0, y: H - 0.12, w: W, h: 0.12, fill: { color: hex(theme.secondary) } });
    slide.addText(String(s.title || ''), { x: 1.5, y: 1.8, w: W - 3, h: 2.2, fontSize: 44, bold: true, color: hex(theme.primary), fontFace: 'Calibri', align: 'center', wrap: true });
    if (s.subtitle) slide.addText(String(s.subtitle), { x: 2, y: 4.1, w: W - 4, h: 0.7, fontSize: 16, color: hex(theme.textMuted), fontFace: 'Calibri', align: 'center', italic: true });
    slide.addShape('rect', { x: W / 2 - 1.5, y: 4.9, w: 3, h: 0.04, fill: { color: hex(theme.secondary) } });
    const ol = [s.organization, s.date].filter(Boolean).join('   |   ');
    slide.addText(ol, { x: 2, y: 5.5, w: W - 4, h: 0.4, fontSize: 12, color: hex(theme.textMuted), align: 'center' });

  } else if (tid === 'gradient_modern') {
    slide.addShape('rect', { x: 0, y: 0, w: W, h: H, fill: { color: hex(theme.coverBg) } });
    slide.addShape('roundRect', { x: 1.8, y: 1.3, w: 9.7, h: 4.8, fill: { color: 'FFFFFF' }, line: { color: 'C7D2FE', width: 1 }, rectRadius: 0.2 });
    slide.addShape('roundRect', { x: 1.8, y: 1.3, w: 9.7, h: 0.6, fill: { color: hex(theme.primary) }, rectRadius: 0.1 });
    slide.addText(String(s.title || ''), { x: 2.2, y: 2.2, w: 8.9, h: 2, fontSize: 38, bold: true, color: hex(theme.primary), fontFace: 'Calibri', wrap: true });
    if (s.subtitle) slide.addText(String(s.subtitle), { x: 2.2, y: 4.2, w: 8.9, h: 0.55, fontSize: 14, color: hex(theme.textMuted), italic: true });
    slide.addText([s.organization, s.date].filter(Boolean).join('  ·  '), { x: 2.2, y: 5.0, w: 8.9, h: 0.4, fontSize: 11, color: hex(theme.textMuted) });

  } else if (tid === 'corporate_bold') {
    slide.addShape('rect', { x: 0, y: 0, w: W, h: H, fill: { color: hex(theme.coverBg) } });
    slide.addShape('rect', { x: 0, y: 0, w: 5.5, h: H, fill: { color: '000D33' } });
    slide.addShape('rect', { x: 5.3, y: 0, w: 0.22, h: H, fill: { color: hex(theme.primary) } });
    slide.addText(String(s.organization || 'Socya'), { x: 0.4, y: H / 2 - 0.7, w: 4.7, h: 1.4, fontSize: 30, bold: true, color: hex(theme.primary), align: 'center', fontFace: 'Calibri' });
    slide.addText(String(s.title || ''), { x: 5.8, y: 1.5, w: 7.2, h: 2.5, fontSize: 38, bold: true, color: 'FFFFFF', fontFace: 'Calibri', wrap: true });
    if (s.subtitle) slide.addText(String(s.subtitle), { x: 5.8, y: 4.1, w: 7.2, h: 0.6, fontSize: 14, color: hex(theme.secondary), italic: true });
    slide.addText(String(s.date || ''), { x: 5.8, y: 5.5, w: 7.2, h: 0.4, fontSize: 12, color: hex(theme.textMuted) });

  } else if (tid === 'green_data') {
    slide.addShape('rect', { x: 0, y: 0, w: W, h: H, fill: { color: hex(theme.coverBg) } });
    for (let gx = 0; gx < W; gx += 1.33)
      slide.addShape('rect', { x: gx, y: 0, w: 0.01, h: H, fill: { color: hex(theme.divider), transparency: 50 } });
    for (let gy = 0; gy < H; gy += 0.75)
      slide.addShape('rect', { x: 0, y: gy, w: W, h: 0.01, fill: { color: hex(theme.divider), transparency: 50 } });
    slide.addShape('roundRect', { x: 0.8, y: 1.5, w: 11.7, h: 3.5, fill: { color: hex(theme.cardBg) }, rectRadius: 0.1 });
    slide.addShape('rect', { x: 0.8, y: 1.5, w: 0.22, h: 3.5, fill: { color: hex(theme.primary) } });
    slide.addText(String(s.title || ''), { x: 1.3, y: 2.0, w: 10.9, h: 2, fontSize: 38, bold: true, color: hex(theme.primary), fontFace: 'Consolas', wrap: true });
    if (s.subtitle) slide.addText(String(s.subtitle), { x: 1.3, y: 4.1, w: 10.9, h: 0.5, fontSize: 13, color: hex(theme.textMuted), italic: true });
    slide.addText([s.organization, s.date].filter(Boolean).join('  //  '), { x: 1.3, y: 5.3, w: 10.9, h: 0.4, fontSize: 11, color: hex(theme.secondary) });

  } else { // sunset_analytics
    slide.addShape('rect', { x: 0, y: 0, w: W, h: H, fill: { color: hex(theme.coverBg) } });
    slide.addShape('ellipse', { x: W - 4, y: -1.5, w: 6, h: 6, fill: { color: hex(theme.primary), transparency: 75 } });
    slide.addShape('ellipse', { x: W - 2.5, y: H - 3, w: 4, h: 4, fill: { color: hex(theme.secondary), transparency: 80 } });
    slide.addShape('ellipse', { x: -2, y: H - 3, w: 4, h: 4, fill: { color: hex(theme.accent), transparency: 85 } });
    slide.addText(String(s.title || ''), { x: 1, y: 1.8, w: 11.3, h: 2.2, fontSize: 44, bold: true, color: hex(theme.text), fontFace: 'Calibri', wrap: true, align: 'center' });
    if (s.subtitle) slide.addText(String(s.subtitle), { x: 2, y: 4.1, w: 9.3, h: 0.6, fontSize: 15, color: hex(theme.textMuted), align: 'center', italic: true });
    slide.addShape('rect', { x: W / 2 - 2, y: 4.9, w: 4, h: 0.06, fill: { color: hex(theme.primary) } });
    slide.addText([s.organization, s.date].filter(Boolean).join('  ·  '), { x: 2, y: 5.3, w: 9.3, h: 0.4, fontSize: 12, color: hex(theme.textMuted), align: 'center' });
  }
}

// ─── SECTION DIVIDER ──────────────────────────────────────────────────────────

function renderSectionDivider(prs, s, theme) {
  const slide = prs.addSlide();
  slide.transition = SLIDE_TRANSITIONS[Math.floor(Math.random() * SLIDE_TRANSITIONS.length)];

  if (theme.isDark) {
    slide.addShape('rect', { x: 0, y: 0, w: W, h: H, fill: { color: hex(theme.slideBg) } });
    // Decorative accent strip
    slide.addShape('rect', { x: 0, y: H * 0.35, w: W, h: 0.08, fill: { color: hex(theme.primary), transparency: 40 } });
    slide.addShape('rect', { x: 0, y: H * 0.35 + 0.14, w: W * 0.6, h: 0.04, fill: { color: hex(theme.secondary), transparency: 60 } });
  } else {
    slide.addShape('rect', { x: 0, y: 0, w: W, h: H, fill: { color: hex(theme.primary) } });
  }

  // Section number (if provided)
  if (s.sectionNumber) {
    slide.addText(String(s.sectionNumber).padStart(2, '0'), {
      x: MX, y: H * 0.2, w: 2, h: 1,
      fontSize: 60, bold: true,
      color: theme.isDark ? hex(theme.primary) : 'FFFFFF',
      fontFace: 'Calibri',
      transparency: theme.isDark ? 0 : 30,
    });
  }

  // Title
  const titleY = s.sectionNumber ? H * 0.38 : H * 0.3;
  slide.addText(String(s.title || ''), {
    x: MX, y: titleY, w: MW, h: 1.5,
    fontSize: 42, bold: true,
    color: theme.isDark ? hex(theme.text) : 'FFFFFF',
    fontFace: 'Calibri', wrap: true,
  });

  // Subtitle
  if (s.subtitle) {
    slide.addText(String(s.subtitle), {
      x: MX, y: titleY + 1.6, w: MW * 0.7, h: 0.7,
      fontSize: 16,
      color: theme.isDark ? hex(theme.textMuted) : 'FFFFFFCC',
      fontFace: 'Calibri', italic: true,
    });
  }
}

// ─── STAT CALLOUT ─────────────────────────────────────────────────────────────

function renderStatCallout(prs, s, theme, meta) {
  const slide = prs.addSlide();
  makeBg(slide, theme);
  applySlideEnhancements(slide, theme);
  drawHeader(slide, s.title, s.subtitle, theme);
  drawFooter(slide, theme, meta);

  const stats = (s.stats || []).slice(0, 4);
  const count = stats.length;
  if (!count) return;

  const cardW = (MW - (count - 1) * 0.3) / count;
  const cardH = MAX_Y - CH - 0.3;
  const cardY = CH + 0.1;

  stats.forEach((stat, i) => {
    const cx = MX + i * (cardW + 0.3);

    // Card background
    slide.addShape('roundRect', {
      x: cx, y: cardY, w: cardW, h: cardH,
      fill: { color: hex(theme.cardBg) },
      line: { color: hex(theme.divider), width: 0.7 },
      rectRadius: 0.1,
    });

    // Top accent bar
    slide.addShape('rect', {
      x: cx, y: cardY, w: cardW, h: 0.08,
      fill: { color: hex(theme.chartColors[i % theme.chartColors.length]) },
    });

    // Icon/emoji
    if (stat.icon) {
      slide.addText(String(stat.icon), {
        x: cx, y: cardY + 0.3, w: cardW, h: 0.5,
        fontSize: 28, align: 'center',
      });
    }

    // Big number
    slide.addText(String(stat.value ?? '—'), {
      x: cx + 0.15, y: cardY + (stat.icon ? 0.9 : 0.5), w: cardW - 0.3, h: 1.2,
      fontSize: 36, bold: true,
      color: hex(theme.primary),
      align: 'center', fontFace: 'Calibri',
    });

    // Label
    slide.addText(String(stat.label ?? ''), {
      x: cx + 0.15, y: cardY + (stat.icon ? 2.2 : 1.8), w: cardW - 0.3, h: 0.5,
      fontSize: 12, bold: true,
      color: hex(theme.text),
      align: 'center', fontFace: 'Calibri',
    });

    // Description
    if (stat.description) {
      slide.addText(String(stat.description), {
        x: cx + 0.2, y: cardY + (stat.icon ? 2.8 : 2.4), w: cardW - 0.4, h: 0.8,
        fontSize: 9,
        color: hex(theme.textMuted),
        align: 'center', fontFace: 'Calibri',
      });
    }
  });
}

// ─── INSIGHTS / FINDINGS ──────────────────────────────────────────────────────

function renderInsights(prs, s, theme, meta) {
  const slide = prs.addSlide();
  makeBg(slide, theme);
  applySlideEnhancements(slide, theme);
  drawHeader(slide, s.title || 'Hallazgos Clave', s.subtitle, theme);
  drawFooter(slide, theme, meta);

  const findings = (s.findings || []).slice(0, 6);
  const twoCol = findings.length > 3;
  const colW = twoCol ? (MW - 0.3) / 2 : MW;
  const itemH = twoCol ? 1.4 : 1.5;

  findings.forEach((f, i) => {
    const col = twoCol ? i % 2 : 0;
    const row = twoCol ? Math.floor(i / 2) : i;
    const fx = MX + col * (colW + 0.3);
    const fy = CH + row * (itemH + 0.15);

    // Card background
    slide.addShape('roundRect', {
      x: fx, y: fy, w: colW, h: itemH,
      fill: { color: hex(theme.cardBg) },
      line: { color: hex(theme.divider), width: 0.5 },
      rectRadius: 0.08,
    });

    // Left accent bar
    const accentColor = f.severity === 'high' ? 'EF4444'
      : f.severity === 'medium' ? 'F59E0B'
      : f.severity === 'low' ? '10B981'
      : hex(theme.primary);

    slide.addShape('rect', {
      x: fx, y: fy, w: 0.12, h: itemH,
      fill: { color: accentColor },
    });

    // Icon
    const icon = f.severity === 'high' ? '⚠' : f.severity === 'medium' ? '◷' : '✓';
    slide.addText(icon, {
      x: fx + 0.22, y: fy + 0.12, w: 0.35, h: 0.35,
      fontSize: 16, align: 'center',
    });

    // Finding title
    slide.addText(String(f.title || ''), {
      x: fx + 0.62, y: fy + 0.12, w: colW - 0.82, h: 0.35,
      fontSize: 11, bold: true,
      color: hex(theme.text), fontFace: 'Calibri',
      valign: 'middle',
    });

    // Finding description
    if (f.description) {
      slide.addText(String(f.description), {
        x: fx + 0.22, y: fy + 0.55, w: colW - 0.42, h: itemH - 0.65,
        fontSize: 9,
        color: hex(theme.textMuted), fontFace: 'Calibri',
      });
    }
  });
}

// ─── DUAL CHART ───────────────────────────────────────────────────────────────

function renderDualChart(prs, s, theme, meta) {
  const slide = prs.addSlide();
  makeBg(slide, theme);
  applySlideEnhancements(slide, theme);
  drawHeader(slide, s.title, s.subtitle, theme);
  drawFooter(slide, theme, meta);

  let y = CH;
  if (s.kpis && s.kpis.length > 0) y = renderKPIs(slide, s.kpis, y, theme) + 0.08;

  const chartH = MAX_Y - y - 0.1;
  const chartW = (MW - 0.4) / 2;

  // Left chart
  if (s.chartLeft && s.chartLeft.data && s.chartLeft.data.length > 0) {
    renderChart(prs, slide, s.chartLeft, MX, y, chartW, chartH, theme);
  }

  // Right chart
  if (s.chartRight && s.chartRight.data && s.chartRight.data.length > 0) {
    renderChart(prs, slide, s.chartRight, MX + chartW + 0.4, y, chartW, chartH, theme);
  }
}

// ─── INDEX ────────────────────────────────────────────────────────────────────

function renderIndex(prs, s, theme, meta) {
  const slide = prs.addSlide();
  makeBg(slide, theme);
  applySlideEnhancements(slide, theme);
  drawHeader(slide, s.title || 'Contenido', s.subtitle, theme);
  drawFooter(slide, theme, meta);

  const items = (s.items || []).slice(0, 16);
  const perCol = 8;
  const colW = (MW - 0.6) / 2;

  items.forEach((item, i) => {
    const col = Math.floor(i / perCol);
    const row = i % perCol;
    const x = MX + col * (colW + 0.6);
    const y = CH + row * 0.62;

    slide.addShape('roundRect', {
      x, y, w: colW, h: 0.5,
      fill: { color: hex(theme.cardBg) },
      line: { color: hex(theme.divider), width: 0.5 },
      rectRadius: 0.05,
    });
    slide.addShape('ellipse', {
      x: x + 0.1, y: y + 0.06, w: 0.36, h: 0.36,
      fill: { color: hex(theme.primary) },
    });
    slide.addText(String(item.num ?? i + 1), {
      x: x + 0.1, y: y + 0.06, w: 0.36, h: 0.36,
      fontSize: 11, bold: true,
      color: theme.isDark ? '000000' : 'FFFFFF',
      align: 'center', valign: 'middle',
    });
    slide.addText(String(item.label || '').substring(0, 50), {
      x: x + 0.55, y, w: colW - 0.62, h: 0.5,
      fontSize: 9.5, bold: true, color: hex(theme.text),
      fontFace: 'Calibri', valign: 'middle',
    });
  });
}

// ─── EXECUTIVE SUMMARY ────────────────────────────────────────────────────────

function renderExecutiveSummary(prs, s, theme, meta) {
  const slide = prs.addSlide();
  makeBg(slide, theme);
  applySlideEnhancements(slide, theme);
  drawHeader(slide, s.title, s.subtitle, theme);
  drawFooter(slide, theme, meta);

  let y = CH;
  y = renderKPIs(slide, s.kpis || [], y, theme);
  y += 0.15;

  const leftW = MW * 0.55;
  const rightW = MW - leftW - 0.35;
  const colY = y;
  const panelH = MAX_Y - colY - 0.2;

  // Status distribution panel (left)
  slide.addShape('roundRect', {
    x: MX, y: colY, w: leftW, h: panelH,
    fill: { color: hex(theme.cardBg) },
    line: { color: hex(theme.divider), width: 0.7 },
    rectRadius: 0.07,
  });
  slide.addText('DISTRIBUCIÓN POR ESTADO', {
    x: MX + 0.18, y: colY + 0.12, w: leftW - 0.36, h: 0.28,
    fontSize: 10, bold: true, color: hex(theme.primary), fontFace: 'Calibri',
  });

  (s.statusDistribution || []).slice(0, 6).forEach((d, i) => {
    const dy = colY + 0.5 + i * 0.46;
    const barW = leftW - 1.6;
    slide.addShape('rect', {
      x: MX + 0.18, y: dy, w: barW, h: 0.26,
      fill: { color: hex(theme.divider) },
    });
    const pct = Math.max(0.08, 1 - i * 0.14);
    slide.addShape('rect', {
      x: MX + 0.18, y: dy, w: barW * pct, h: 0.26,
      fill: { color: hex(d.color || theme.primary) },
    });
    slide.addText(String(d.label), {
      x: MX + 0.22, y: dy, w: barW - 0.1, h: 0.26,
      fontSize: 8.5, color: theme.isDark ? 'FFFFFF' : '111827', valign: 'middle',
    });
    slide.addText(String(d.count), {
      x: MX + leftW - 1.3, y: dy, w: 0.9, h: 0.26,
      fontSize: 9, bold: true, color: hex(theme.primary), align: 'right',
    });
  });

  // Sheet coverage panel (right)
  slide.addShape('roundRect', {
    x: MX + leftW + 0.35, y: colY, w: rightW, h: panelH,
    fill: { color: hex(theme.cardBg) },
    line: { color: hex(theme.divider), width: 0.7 },
    rectRadius: 0.07,
  });
  slide.addText('COBERTURA DE ANÁLISIS', {
    x: MX + leftW + 0.52, y: colY + 0.12, w: rightW - 0.3, h: 0.28,
    fontSize: 10, bold: true, color: hex(theme.primary),
  });

  (s.sheetSummary || []).slice(0, 10).forEach((sh, i) => {
    const sy = colY + 0.5 + i * 0.32;
    const dot = sh.layout === 'chart' ? '📊' : sh.layout === 'mixed' ? '📋' : '🗂';
    slide.addText(`${dot} ${String(sh.name).substring(0, 28)}`, {
      x: MX + leftW + 0.52, y: sy, w: rightW - 1.1, h: 0.28,
      fontSize: 9, color: hex(theme.text),
    });
    slide.addText(String(sh.rows), {
      x: MX + leftW + rightW - 0.55, y: sy, w: 0.45, h: 0.28,
      fontSize: 9, bold: true, color: hex(theme.primary), align: 'right',
    });
  });
}

// ─── TABLE SLIDE ──────────────────────────────────────────────────────────────

function renderTableSlide(prs, s, theme, meta) {
  const slide = prs.addSlide();
  makeBg(slide, theme);
  applySlideEnhancements(slide, theme);
  drawHeader(slide, s.title, s.subtitle, theme);
  drawFooter(slide, theme, meta);

  let y = CH;
  if (s.kpis && s.kpis.length > 0) y = renderKPIs(slide, s.kpis, y, theme) + 0.08;
  renderTable(slide, s.columns, s.colW, s.rows, y, theme, s.overflow ?? 0);
}

// ─── CHART SLIDE ──────────────────────────────────────────────────────────────

function renderChartSlide(prs, s, theme, meta) {
  const slide = prs.addSlide();
  makeBg(slide, theme);
  applySlideEnhancements(slide, theme);
  drawHeader(slide, s.title, s.subtitle, theme);
  drawFooter(slide, theme, meta);

  let y = CH;
  if (s.kpis && s.kpis.length > 0) y = renderKPIs(slide, s.kpis, y, theme) + 0.08;

  if (s.chart && s.chart.data && s.chart.data.length > 0) {
    renderChart(prs, slide, s.chart, MX, y, MW, MAX_Y - y - 0.1, theme);
  }
}

// ─── MIXED SLIDE (table left + chart right) ───────────────────────────────────

function renderMixedSlide(prs, s, theme, meta) {
  const slide = prs.addSlide();
  makeBg(slide, theme);
  applySlideEnhancements(slide, theme);
  drawHeader(slide, s.title, s.subtitle, theme);
  drawFooter(slide, theme, meta);

  let y = CH;
  if (s.kpis && s.kpis.length > 0) y = renderKPIs(slide, s.kpis, y, theme) + 0.08;

  const hasTable = s.rows && s.rows.length > 0 && s.columns && s.columns.length > 0;
  const hasChart = s.chart && s.chart.data && s.chart.data.length > 0;

  if (hasTable && hasChart) {
    const tblW = 5.8;
    const chX  = MX + tblW + 0.35;
    const chW  = W - chX - MX;
    const chH  = MAX_Y - y - 0.1;

    // Table on left — only first 4 cols to fit
    const leftCols = s.columns.slice(0, 4);
    const leftColW = leftCols.map(() => parseFloat((tblW / leftCols.length).toFixed(2)));
    renderTable(slide, leftCols, leftColW, s.rows, y, theme, s.overflow ?? 0);

    // Chart on right
    renderChart(prs, slide, s.chart, chX, y, chW, chH, theme);

  } else if (hasTable) {
    renderTable(slide, s.columns, s.colW, s.rows, y, theme, s.overflow ?? 0);
  } else if (hasChart) {
    renderChart(prs, slide, s.chart, MX, y, MW, MAX_Y - y - 0.1, theme);
  }
}

// ─── TEXT BULLETS ─────────────────────────────────────────────────────────────

function renderTextBullets(prs, s, theme, meta) {
  const slide = prs.addSlide();
  makeBg(slide, theme);
  applySlideEnhancements(slide, theme);
  drawHeader(slide, s.title, s.subtitle, theme);
  drawFooter(slide, theme, meta);

  const sections = s.sections || [];
  const twoCol = sections.length > 2;
  const colW = twoCol ? (MW - 0.3) / 2 : MW;
  const sectionH = (MAX_Y - CH) / (twoCol ? Math.ceil(sections.length / 2) : sections.length);

  sections.slice(0, 4).forEach((sec, i) => {
    const col = twoCol ? Math.floor(i / 2) : 0;
    const row = twoCol ? i % 2 : i;
    const sx = MX + col * (colW + 0.3);
    const sy = CH + row * sectionH;

    slide.addShape('roundRect', {
      x: sx, y: sy, w: colW, h: 0.38,
      fill: { color: hex(theme.primary) },
      rectRadius: 0.06,
    });
    slide.addText(`${sec.icon || ''} ${sec.label || ''}`, {
      x: sx + 0.12, y: sy, w: colW - 0.24, h: 0.38,
      bold: true,
      color: theme.isDark ? '000000' : 'FFFFFF',
      fontSize: 11,
      valign: 'middle',
    });
    let cy = sy + 0.45;
    (sec.bullets || []).slice(0, 5).forEach(b => {
      slide.addText(`▸  ${String(b)}`, {
        x: sx + 0.12, y: cy, w: colW - 0.24, h: 0.3,
        fontSize: 10, color: hex(theme.text),
      });
      cy += 0.32;
    });
  });
}

// ─── OVERFLOW DETAIL ─────────────────────────────────────────────────────────

function renderOverflowDetail(prs, s, theme, meta) {
  const slide = prs.addSlide();
  makeBg(slide, theme);
  applySlideEnhancements(slide, theme);
  drawHeader(slide, s.title, s.subtitle || 'Datos adicionales disponibles en el archivo original', theme);
  drawFooter(slide, theme, meta);

  slide.addText(String(s.message || ''), {
    x: MX, y: CH + 0.1, w: MW, h: 0.55,
    fontSize: 13, color: hex(theme.textMuted),
    italic: true, align: 'center',
  });

  // Navigation hint
  slide.addShape('roundRect', {
    x: MW / 2 - 1.5, y: CH + 0.75, w: 4, h: 0.45,
    fill: { color: hex(theme.primary), transparency: 15 },
    line: { color: hex(theme.primary), width: 0.8 },
    rectRadius: 0.08,
  });
  slide.addText('◀ Ir al Inicio (Diapositiva 1)', {
    x: MW / 2 - 1.5, y: CH + 0.75, w: 4, h: 0.45,
    fontSize: 10, bold: true, color: hex(theme.primary),
    align: 'center', valign: 'middle',
    hyperlink: { slide: '1', tooltip: 'Volver a la portada' },
  });

  const sections = (s.sections || []).slice(0, 8);
  const totalRows = sections.reduce((sum, sec) => sum + (sec.rows ?? 0), 0);

  // Total records badge
  if (totalRows > 0) {
    slide.addShape('roundRect', {
      x: MX, y: CH + 1.35, w: MW, h: 0.4,
      fill: { color: hex(theme.cardBg) },
      line: { color: hex(theme.divider), width: 0.5 },
      rectRadius: 0.06,
    });
    slide.addText(`📊 Total: ${totalRows.toLocaleString('es-CO')} registros en ${sections.length} secciones no incluidas`, {
      x: MX + 0.2, y: CH + 1.35, w: MW - 0.4, h: 0.4,
      fontSize: 10, color: hex(theme.textMuted), align: 'center', valign: 'middle',
    });
  }
  const cardStartY = totalRows > 0 ? CH + 1.95 : CH + 1.35;
  sections.forEach((sec, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const sx = MX + col * (MW / 2 + 0.15);
    const sy = cardStartY + row * 0.82;

    slide.addShape('roundRect', {
      x: sx, y: sy, w: MW / 2 - 0.15, h: 0.65,
      fill: { color: hex(theme.cardBg) },
      line: { color: hex(theme.primary), width: 0.7 },
      rectRadius: 0.08,
    });
    // Accent stripe
    slide.addShape('rect', {
      x: sx, y: sy, w: 0.08, h: 0.65,
      fill: { color: hex(theme.chartColors[i % theme.chartColors.length]) },
    });
    slide.addText(`📄 ${String(sec.name).substring(0, 36)}`, {
      x: sx + 0.18, y: sy + 0.04, w: MW / 2 - 0.5, h: 0.3,
      fontSize: 11, bold: true, color: hex(theme.primary),
      hyperlink: { slide: '1', tooltip: `Datos de: ${sec.name}` },
    });
    slide.addText(`${(sec.rows ?? 0).toLocaleString('es-CO')} registros disponibles en Excel`, {
      x: sx + 0.18, y: sy + 0.37, w: MW / 2 - 0.5, h: 0.22,
      fontSize: 9, color: hex(theme.textMuted),
    });
  });
}

// ─── CLOSING ─────────────────────────────────────────────────────────────────

function renderClosing(prs, s, theme) {
  const slide = prs.addSlide();
  slide.transition = { type: 'fade', speed: 'slow' };

  if (theme.isDark) {
    slide.addShape('rect', { x: 0, y: 0, w: W, h: H, fill: { color: hex(theme.coverBg) } });
    slide.addShape('ellipse', { x: -1, y: -1, w: 4, h: 4, fill: { color: hex(theme.primary), transparency: 85 } });
    slide.addShape('ellipse', { x: W - 2.5, y: H - 2.5, w: 4, h: 4, fill: { color: hex(theme.secondary), transparency: 88 } });
  } else {
    slide.addShape('rect', { x: 0, y: 0, w: W, h: H, fill: { color: hex(theme.primary) } });
  }

  slide.addText(String(s.title || '¡Gracias!'), {
    x: 1, y: H / 2 - 1.2, w: W - 2, h: 1.3,
    fontSize: 48, bold: true, color: 'FFFFFF',
    fontFace: 'Calibri', align: 'center',
    shadow: { type: 'outer', color: '000000', opacity: 0.4, blur: 8, offset: 4, angle: 45 },
  });
  if (s.message) {
    slide.addText(String(s.message), {
      x: 1.5, y: H / 2 + 0.25, w: W - 3, h: 0.55,
      fontSize: 14, color: 'FFFFFF',
      align: 'center', italic: true, fontFace: 'Calibri',
    });
  }
  if (s.organization) {
    slide.addText(String(s.organization), {
      x: 1.5, y: H - 1.0, w: W - 3, h: 0.4,
      fontSize: 12, color: 'FFFFFF',
      align: 'center', fontFace: 'Calibri',
    });
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// MAIN BUILDER
// ══════════════════════════════════════════════════════════════════════════════

export function buildEntirePresentation(prs, ast) {
  const theme = ast.theme || {};
  const meta = `${ast.meta?.title || 'Reporte'} · ${new Date().toLocaleDateString('es-CO')}`;

  for (const s of (ast.slides || [])) {
    if (!s || !s.type) continue;
    try {
      switch (s.type) {
        case 'cover':             renderCover(prs, s, theme); break;
        case 'section_divider':   renderSectionDivider(prs, s, theme); break;
        case 'index':             renderIndex(prs, s, theme, meta); break;
        case 'executive_summary': renderExecutiveSummary(prs, s, theme, meta); break;
        case 'stat_callout':      renderStatCallout(prs, s, theme, meta); break;
        case 'insights':
        case 'findings':          renderInsights(prs, s, theme, meta); break;
        case 'dual_chart':        renderDualChart(prs, s, theme, meta); break;
        case 'table':
        case 'kpi_and_table':
        case 'action_table':
        case 'coso_model':
        case 'matrix_table':      renderTableSlide(prs, s, theme, meta); break;
        case 'chart':
        case 'chart_only':        renderChartSlide(prs, s, theme, meta); break;
        case 'mixed':             renderMixedSlide(prs, s, theme, meta); break;
        case 'text_bullets':
        case 'risk_card':         renderTextBullets(prs, s, theme, meta); break;
        case 'overflow_detail':   renderOverflowDetail(prs, s, theme, meta); break;
        case 'closing':
        case 'closing_data':      renderClosing(prs, s, theme); break;
        default:
          console.warn(`[Renderer] Tipo desconocido: ${s.type}, renderizando como tabla`);
          if (s.rows || s.columns) renderTableSlide(prs, s, theme, meta);
      }
    } catch (e) {
      console.error(`[Renderer] Error en slide "${s.title}" (${s.type}):`, e?.message || e);
    }
  }
}
