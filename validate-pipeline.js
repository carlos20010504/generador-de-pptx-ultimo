/**
 * validate-pipeline.js
 * Tests the Excel parsing & decision logic using pure Node.js (no TypeScript compilation needed)
 * Mirrors the logic in excel-parser.ts to verify it works with Comisiones V1.xlsx
 */
const XLSX = require('xlsx');
const path = require('path');

// ── Helpers ──────────────────────────────────────────────────────────────────
function toNum(val) {
  if (val === null || val === undefined || val === '') return null;
  const n = parseFloat(String(val).replace(/[$€£¥,\s]/g,'').replace(/COP|USD/gi,''));
  return isNaN(n) ? null : n;
}
function normKey(s) {
  return String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-zA-Z0-9]/g,'').toUpperCase().trim();
}
function truncate(str, max=45) {
  const s = String(str ?? '').trim();
  return s.length > max ? s.slice(0,max-1)+'…' : (s||'—');
}

// ── Column Type Detection ─────────────────────────────────────────────────────
function detectColumnType(colName, values) {
  const sample = values.filter(v => v!==null && v!==undefined && String(v).trim()!=='').slice(0,30);
  if (!sample.length) return 'empty';
  const nk = normKey(colName);
  if (/ID$|^ID|COD|CODIGO|FOLIO|NUMERO|CONS/.test(nk)) return 'id';
  if (/FECHA|DATE|DIA|MES|AÑO|YEAR|MONTH|PERIODO/.test(nk)) return 'date';
  if (/PORC|PCT|PERCENT|TASA|RATE/.test(nk)) return 'percent';
  if (/VALOR|TOTAL|COSTO|PRECIO|MONTO|SALAR|GASTO|INGRES|COMIS|EGRESO|PAGO/.test(nk)) return 'currency';
  const nums = sample.map(v => toNum(v)).filter(n => n!==null);
  const numRatio = nums.length / sample.length;
  if (numRatio >= 0.75) {
    const avg = nums.reduce((a,b)=>a+b,0)/nums.length;
    if (nums.every(n=>n>=0&&n<=1) && nums.length>=5) return 'percent';
    if (avg > 10000) return 'currency';
    return 'numeric';
  }
  const unique = new Set(sample.map(v=>normKey(String(v))));
  if (unique.size<=2 && (unique.has('SI')||unique.has('NO')||unique.has('TRUE'))) return 'boolean';
  if (unique.size<=10 && sample.length>=4) return 'categorical';
  return 'text';
}

// ── Layout Decision ───────────────────────────────────────────────────────────
function decideLayout(colDefs, rowCount) {
  const numericCount  = colDefs.filter(c=>['currency','numeric','percent'].includes(c.type)).length;
  const catCount      = colDefs.filter(c=>['categorical','boolean'].includes(c.type)).length;
  const totalCols     = colDefs.length;
  if (rowCount<=8 && numericCount>=2 && totalCols<=5) return 'kpi_cards';
  if (totalCols===2 && numericCount===1) return 'chart';
  if (rowCount<=25 && numericCount>=1 && catCount>=1 && totalCols>=3 && totalCols<=6) return 'mixed';
  if (numericCount>=1 && catCount>=1 && totalCols<=4 && rowCount>10) return 'mixed';
  return 'table';
}

// ── Chart Type ────────────────────────────────────────────────────────────────
function selectChartType(labelType, distinctCount, hasTime, values) {
  if (hasTime || labelType==='date') return 'line';
  const allPos = values.every(v=>v>=0);
  if (distinctCount<=7 && allPos) return 'donut';
  if (distinctCount>8) return 'bar_horizontal';
  return 'bar';
}

// ── Header Detection ──────────────────────────────────────────────────────────
function findHeaderRow(rows) {
  for (let i=0; i<Math.min(rows.length,15); i++) {
    const ne = (rows[i]||[]).filter(c=>c!==null&&c!==undefined&&String(c).trim()!=='');
    if (ne.length >= 2) return i;
  }
  return 0;
}

// ── Main ──────────────────────────────────────────────────────────────────────
function validateFile(filePath) {
  console.log('\n' + '═'.repeat(60));
  console.log('📊 Validando:', path.basename(filePath));
  console.log('═'.repeat(60));

  const wb = XLSX.readFile(filePath);
  const errors = [];
  const warnings = [];
  let totalSlides = 2; // cover + exec summary

  wb.SheetNames.forEach(sn => {
    const ws = wb.Sheets[sn];
    const rawRows = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
    const cleanRows = rawRows.filter(r => r && r.some(c => c!==null && String(c).trim()!==''));
    if (!cleanRows.length) {
      warnings.push(`Hoja "${sn}": vacía, ignorada`);
      return;
    }

    const hi = findHeaderRow(cleanRows);
    const headers = (cleanRows[hi]||[]).map(h=>String(h??'').trim()).filter(h=>h);
    const bodyRows = cleanRows.slice(hi+1).filter(r=>r&&r.some(c=>c!==null&&String(c).trim()!==''));

    if (!headers.length) {
      errors.push(`Hoja "${sn}": no se detectaron encabezados`);
      return;
    }
    if (!bodyRows.length) {
      warnings.push(`Hoja "${sn}": sin filas de datos`);
      return;
    }

    // Detect column types
    const colDefs = headers.map((h,idx) => {
      const vals = bodyRows.slice(0,30).map(r=>r[idx]??null);
      const type = detectColumnType(h, vals);
      return { key:h, header:h, idx, type };
    }).filter(c=>c.type!=='empty');

    // Decide layout
    const layout = decideLayout(colDefs, bodyRows.length);

    // For chart/mixed: validate we have the right columns
    const numCols = colDefs.filter(c=>['currency','numeric','percent'].includes(c.type));
    const catCols = colDefs.filter(c=>['categorical','boolean','text'].includes(c.type));

    // Validate table columns fit
    const maxCols = 8;
    const useableCols = colDefs.slice(0, maxCols).length;

    // Validate no more than 8 cols in table
    if (colDefs.length > maxCols) {
      warnings.push(`Hoja "${sn}": ${colDefs.length} cols → se mostrarán las primeras ${maxCols}`);
    }

    // Compute slide count for this sheet
    const rowsPerSlide = 18;
    const sheetSlides = layout === 'table' ? Math.ceil(bodyRows.length/rowsPerSlide) : 1;
    // If table also generates a chart companion
    const hasChartCompanion = layout === 'table' && numCols.length && catCols.length && bodyRows.length > 2;
    totalSlides += sheetSlides + (hasChartCompanion ? 1 : 0);

    // Validate chart data would be non-empty
    if ((layout==='chart'||layout==='mixed') && (!numCols.length || !catCols.length)) {
      warnings.push(`Hoja "${sn}": layout=${layout} pero faltan columnas numéricas o categóricas → usará tabla`);
    }

    console.log(`✅ ${sn}:`);
    console.log(`   Filas: ${bodyRows.length}, Cols usables: ${useableCols}/${colDefs.length}`);
    console.log(`   Tipos: ${colDefs.map(c=>c.type).join(', ')}`);
    console.log(`   Layout → ${layout} | Slides generadas: ${sheetSlides}`);
  });

  totalSlides += 2; // index + closing
  console.log('\n📋 RESUMEN:');
  console.log(`   Total slides estimadas: ${totalSlides} ${totalSlides>25?'⚠️ (se consolidarán a 25)':'✅'}`);

  if (warnings.length) {
    console.log('\n⚠️  Advertencias:');
    warnings.forEach(w => console.log('   -', w));
  }
  if (errors.length) {
    console.log('\n❌ ERRORES:');
    errors.forEach(e => console.log('   -', e));
  } else {
    console.log('\n✅ Sin errores críticos — el pipeline debería funcionar correctamente');
  }
  return errors.length === 0;
}

// Run validation
const files = [
  'test_data.xlsx',
  'Comisiones V1.xlsx',
];

let allPassed = true;
files.forEach(f => {
  try {
    const ok = validateFile(f);
    if (!ok) allPassed = false;
  } catch(e) {
    console.error(`❌ Error leyendo ${f}:`, e.message);
    allPassed = false;
  }
});

console.log('\n' + '═'.repeat(60));
console.log(allPassed ? '🎉 VALIDACIÓN COMPLETA — Todo OK' : '💥 VALIDACIÓN FALLIDA — Ver errores');
console.log('═'.repeat(60));
process.exit(allPassed ? 0 : 1);
