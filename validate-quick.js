// Quick validation script — tests the Excel→PPTX pipeline
const XLSX = require('xlsx');
const path = require('path');

// Read test Excel
const testFile = path.join(__dirname, 'test_data.xlsx');
const wb = XLSX.readFile(testFile);

console.log('\n=== Excel Structure ===');
console.log('Sheets:', wb.SheetNames);
wb.SheetNames.forEach(sn => {
  const ws = wb.Sheets[sn];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
  console.log(`  "${sn}": ${rows.length} rows, ${rows.length ? Object.keys(rows[0]).length : 0} cols`);
  if (rows.length > 0) {
    console.log('    Headers:', Object.keys(rows[0]).slice(0, 6).join(', '));
    console.log('    Sample:', JSON.stringify(rows[0]).substring(0, 120));
  }
});

console.log('\n✅ Excel read successfully. Pipeline ready for testing.');
console.log('To test: upload test_data.xlsx via the web UI at localhost:3000');
