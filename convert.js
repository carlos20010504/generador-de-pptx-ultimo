const fs = require('fs');

const inPath = 'Exceltopptxstructure.txt';
const outPath = 'utils/excel-parser.ts';

let txt = fs.readFileSync(inPath, 'utf8');
if (txt.includes('\u0000')) {
  txt = fs.readFileSync(inPath, 'utf16le');
}

// 1. replace requires
txt = txt.replace(/const\s+XLSX\s*=\s*require\('xlsx'\);/g, "import * as XLSX from 'xlsx';");
txt = txt.replace(/const\s+path\s*=\s*require\('path'\);/g, "");
txt = txt.replace(/const\s+fs\s*=\s*require\('fs'\);/g, "");

// 2. fix buildPresentation
// Find: function buildPresentation(filePath, options = {}) {
// Let's replace the whole method signature and the file reading part.
const funcStart = "function buildPresentation(filePath, options = {}) {";
const funcEnd = "const wb       = XLSX.readFile(filePath, { cellDates: false, cellNF: true });";

// The replacement signature will be:
const newSig = `export function parsePresentationFromWorkbook(wb, options = {}) {`;
const replaceRegex = /function buildPresentation\s*\([\s\S]*?XLSX\.readFile\([^;]+;\s*/;

if (replaceRegex.test(txt)) {
  txt = txt.replace(replaceRegex, `
export function parsePresentationFromWorkbook(wb: any, options: any = {}) {
  const {
    validate      = true,
    datasetUrl    = 'https://mi-servidor.com/dataset/Comisiones_V1.xlsx',
    organization  = 'Auditoría Interna',
  } = options;

  log.info(\`Procesando workbook...\`);
  `);
} else {
  console.error("Regex did not match buildPresentation!");
}

// Strip out the CLI "if (require.main === module)" stuff to avoid TS errors
txt = txt.replace(/if\s*\(require\.main\s*===\s*module\)\s*\{[\s\S]*\}[\s\S]*$/g, "");

// Change the exports
txt = txt.replace(/module\.exports\s*=\s*\{[\s\S]*?^};/m, "");

fs.writeFileSync(outPath, txt, 'utf8');
console.log('Conversion successful. Wrote utils/excel-parser.ts');
