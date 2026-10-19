/**
 * Writes fixtures/equipment/out/Equipment.xlsx from workbook.json.
 * Every sheet gets exactly one formal Excel Table, as the workbook conventions require.
 */
import { mkdir } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import ExcelJS from 'exceljs';
import { equipmentWorkbook } from './index.js';

const outDir = fileURLToPath(new URL('./out/', import.meta.url));
const outFile = `${outDir}${equipmentWorkbook.file}`;

const book = new ExcelJS.Workbook();
book.creator = 'McSheety fixture generator';

for (const sheet of equipmentWorkbook.sheets) {
  const ws = book.addWorksheet(sheet.sheet);
  ws.addTable({
    name: sheet.table,
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    style: { theme: 'TableStyleMedium2', showRowStripes: true },
    columns: sheet.columns.map((name) => ({ name, filterButton: true })),
    rows: sheet.rows.map((row) => [...row]),
  });
  sheet.columns.forEach((name, i) => {
    const widest = Math.max(name.length, ...sheet.rows.map((r) => String(r[i] ?? '').length));
    ws.getColumn(i + 1).width = Math.min(Math.max(widest + 2, 10), 40);
  });
}

await mkdir(outDir, { recursive: true });
await book.xlsx.writeFile(outFile);
console.log(`Wrote ${outFile}`);
