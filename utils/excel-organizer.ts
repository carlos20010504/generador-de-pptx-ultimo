// @ts-nocheck
import * as XLSX from 'xlsx';

export type OrganizerMode = 'charts' | 'tables' | 'mixed' | 'auto';

// ─── Column Type Detection ────────────────────────────────────────────────────

function toNum(val: any): number | null {
  if (val === null || val === undefined || val === '') return null;
  const n = parseFloat(String(val).replace(/[$€£¥,\s]/g, '').replace(/COP|USD/gi, ''));
  return isNaN(n) ? null : n;
}

function normKey(s: any): string {
  return String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-zA-Z0-9]/g, '').toUpperCase().trim();
}

function isUnnamed(v: any): boolean {
  if (v === null || v === undefined || String(v).trim() === '') return true;
  const s = String(v).toLowerCase().trim();
  return s.startsWith('unnamed') || /^col[_ -]?\d+$/i.test(s);
}

function detectColType(header: string, values: any[]): string {
  const nk = normKey(header);
  const sample = values.filter(v => v !== null && v !== undefined && String(v).trim() !== '').slice(0, 30);
  if (!sample.length) return 'empty';

  if (/ID$|^ID|COD|CODIGO|FOLIO|CONS/.test(nk)) return 'id';
  if (/FECHA|DATE|PERIODO|YEAR|MONTH|MES/.test(nk)) return 'date';
  if (/PORC|PCT|PERCENT|TASA/.test(nk)) return 'percent';
  if (/VALOR|TOTAL|COSTO|PRECIO|MONTO|GASTO|SALAR|COMIS|PAGO/.test(nk)) return 'currency';

  const nums = sample.map(v => toNum(v)).filter(n => n !== null) as number[];
  const numRatio = nums.length / sample.length;

  if (numRatio >= 0.75) {
    const avg = nums.reduce((a, b) => a + b, 0) / nums.length;
    if (nums.every(n => n >= 0 && n <= 1) && nums.length >= 5) return 'percent';
    if (avg > 10000) return 'currency';
    return 'numeric';
  }

  const unique = new Set(sample.map(v => normKey(String(v))));
  if (unique.size <= 2 && (unique.has('SI') || unique.has('NO') || unique.has('TRUE'))) return 'boolean';
  if (unique.size <= 10 && sample.length >= 4) return 'categorical';
  return 'text';
}

// ─── Header Detection ─────────────────────────────────────────────────────────

function findHeaderRow(rows: any[][]): number {
  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    const nonempty = (rows[i] ?? []).filter((c: any) => c !== null && c !== undefined && String(c).trim() !== '');
    if (nonempty.length >= 2) return i;
  }
  return 0;
}

function columnHasData(rows: any[][], colIdx: number, minFill = 0.10): boolean {
  if (!rows.length) return false;
  const filled = rows.filter(r => {
    const v = r[colIdx];
    return v !== null && v !== undefined && String(v).trim() !== '';
  }).length;
  return filled / rows.length >= minFill;
}

function cleanColumns(headers: any[], bodyRows: any[][]): { headers: string[]; rows: any[][] } {
  const valid: number[] = [];
  headers.forEach((h, idx) => {
    if (!isUnnamed(h) && columnHasData(bodyRows, idx, 0.08)) valid.push(idx);
  });
  if (!valid.length) return { headers: headers.map(h => String(h ?? '')), rows: bodyRows };
  return {
    headers: valid.map(i => String(headers[i] ?? '').trim()),
    rows: bodyRows.map(row => valid.map(i => row[i] ?? null)),
  };
}

// ─── Slide Type Decision ──────────────────────────────────────────────────────

function decideType(headers: string[], bodyRows: any[][], mode: OrganizerMode): string {
  const colCount = headers.length;
  if (colCount <= 1) return 'TEXT';

  // Detect column types
  const types = headers.map((h, idx) => detectColType(h, bodyRows.slice(0, 30).map(r => r[idx])));
  const numericCount = types.filter(t => ['currency', 'numeric', 'percent'].includes(t)).length;
  const catCount = types.filter(t => ['categorical', 'boolean'].includes(t)).length;
  const rowCount = bodyRows.length;

  // Force modes
  if (mode === 'tables') return 'TABLE';
  if (mode === 'charts') {
    if (numericCount >= 1 && colCount <= 5) return 'CHART';
    return colCount <= 1 ? 'TEXT' : 'TABLE';
  }

  // Auto / mixed logic
  if (colCount === 2 && numericCount === 1) return 'CHART';
  if (numericCount >= 1 && catCount >= 1 && colCount <= 6 && rowCount <= 30) return 'MIXTO';
  if (numericCount >= 1 && colCount >= 4) return 'TABLE';
  if (catCount >= 1 && numericCount >= 1) return 'MIXTO';
  return 'TABLE';
}

// ─── Title Inference ─────────────────────────────────────────────────────────

function inferTitle(cleanRows: any[][]): { title: string; subtitle: string } {
  let title = '', subtitle = '';
  for (let i = 0; i < Math.min(cleanRows.length, 8); i++) {
    const row = cleanRows[i] ?? [];
    const cell = (row[0] != null ? String(row[0]) : '').trim();
    const nonEmpty = row.filter((c: any) => c !== null && c !== undefined && String(c).trim() !== '');
    if (cell.toUpperCase().startsWith('TITLE:')) { title = cell.substring(6).trim(); continue; }
    if (cell.toUpperCase().startsWith('SUBTITLE:')) { subtitle = cell.substring(9).trim(); continue; }
    if (nonEmpty.length === 1 && cell.length > 3) {
      if (!title) title = cell;
      else if (!subtitle) { subtitle = cell; break; }
    }
    if (nonEmpty.length >= 2) break;
  }
  return { title, subtitle };
}

// ─── MAIN EXPORT ─────────────────────────────────────────────────────────────

export async function autoOrganizeExcel(file: File, mode: OrganizerMode = 'auto'): Promise<Blob> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = async (e) => {
      try {
        const data = e.target?.result;
        if (!data) throw new Error('No se pudo leer el archivo.');

        const wb = XLSX.read(data, { type: 'array', cellDates: false });
        const newWb = XLSX.utils.book_new();

        for (const sheetName of wb.SheetNames) {
          const ws = wb.Sheets[sheetName];
          const rawRows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as any[][];

          // Filter completely empty rows
          const cleanRows = rawRows.filter(r => r && r.some((c: any) => c !== null && c !== undefined && String(c).trim() !== ''));
          if (!cleanRows.length) continue;

          const headerIdx = findHeaderRow(cleanRows);
          const { title: inferredTitle, subtitle: inferredSubtitle } = inferTitle(cleanRows.slice(0, headerIdx));

          const rawHeaders = cleanRows[headerIdx] ?? [];
          const bodyRows = cleanRows.slice(headerIdx + 1);

          const { headers: cleanHeaders, rows: filteredRows } = cleanColumns(rawHeaders, bodyRows);

          // Decide visualization type
          const vizType = decideType(cleanHeaders, filteredRows, mode);

          // Build label
          const finalTitle = inferredTitle || sheetName;
          const finalSubtitle = inferredSubtitle || `Organizado automáticamente — ${vizType}`;

          // Build output rows
          const outRows: any[][] = [];
          outRows.push([`TITLE: ${finalTitle}`]);
          outRows.push([`SUBTITLE: ${finalSubtitle}`]);
          outRows.push([`TYPE: ${vizType}`]);
          outRows.push([]); // blank row
          outRows.push(cleanHeaders);
          filteredRows.forEach(row => outRows.push(row));

          const newWs = XLSX.utils.aoa_to_sheet(outRows);
          XLSX.utils.book_append_sheet(newWb, newWs, sheetName);

          console.log(`[Organizer] "${sheetName}" → ${vizType} (${filteredRows.length} filas, ${cleanHeaders.length} cols)`);
        }

        if (newWb.SheetNames.length === 0) throw new Error('El archivo no contiene hojas válidas.');

        const out = XLSX.write(newWb, { bookType: 'xlsx', type: 'array' });
        resolve(new Blob([out], { type: 'application/octet-stream' }));
      } catch (err) {
        console.error('[Organizer] Error:', err);
        reject(err);
      }
    };

    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}
