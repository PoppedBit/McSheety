/** Reads formal Excel Tables from a local .xlsx into the same shapes Graph returns. */
import type { CellValue, RawTable } from '@mcsheety/core';
import { columnIndex } from '@mcsheety/core';
import ExcelJS from 'exceljs';

export interface LocalTable extends RawTable {
  readonly sheet: string;
  readonly name: string;
}

const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);

/** Converts an exceljs cell value to what Graph reports in `values` (dates as serials, blanks as ""). */
function toGraphValue(value: ExcelJS.CellValue): { value: CellValue; formula: string | undefined } {
  if (value === null || value === undefined) return { value: '', formula: undefined };
  if (value instanceof Date)
    return { value: (value.getTime() - EXCEL_EPOCH_MS) / 86_400_000, formula: undefined };
  if (typeof value === 'object') {
    if ('formula' in value || 'sharedFormula' in value) {
      const f = value as ExcelJS.CellFormulaValue;
      const inner = toGraphValue(f.result ?? '').value;
      return { value: inner, formula: `=${f.formula ?? ''}` };
    }
    if ('richText' in value)
      return { value: value.richText.map((r) => r.text).join(''), formula: undefined };
    if ('text' in value) return { value: String(value.text), formula: undefined };
    if ('error' in value) return { value: String(value.error), formula: undefined };
    return { value: '', formula: undefined };
  }
  return { value, formula: undefined };
}

function parseRef(ref: string): { c1: number; r1: number; c2: number; r2: number } {
  const m = /^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/i.exec(ref);
  if (!m) throw new Error(`Unsupported table range "${ref}"`);
  return {
    c1: columnIndex(m[1] ?? 'A'),
    r1: Number(m[2]),
    c2: columnIndex(m[3] ?? 'A'),
    r2: Number(m[4]),
  };
}

export async function readWorkbookTables(file: string): Promise<LocalTable[]> {
  const book = new ExcelJS.Workbook();
  await book.xlsx.readFile(file);
  const tables: LocalTable[] = [];
  for (const ws of book.worksheets) {
    for (const entry of ws.getTables()) {
      const table = (Array.isArray(entry) ? entry[0] : entry) as unknown as {
        table: { name: string; ref?: string; tableRef?: string; totalsRow?: boolean };
      };
      const ref = table.table.tableRef ?? table.table.ref;
      if (!ref) continue;
      const { c1, r1, c2, r2 } = parseRef(ref);
      const lastBodyRow = table.table.totalsRow ? r2 - 1 : r2;
      const header: string[] = [];
      for (let c = c1; c <= c2; c++)
        header.push(String(toGraphValue(ws.getRow(r1).getCell(c + 1).value).value));
      const rows: CellValue[][] = [];
      const formulas: CellValue[][] = [];
      for (let r = r1 + 1; r <= lastBodyRow; r++) {
        const values: CellValue[] = [];
        const fs: CellValue[] = [];
        for (let c = c1; c <= c2; c++) {
          const { value, formula } = toGraphValue(ws.getRow(r).getCell(c + 1).value);
          values.push(value);
          fs.push(formula ?? value);
        }
        rows.push(values);
        formulas.push(fs);
      }
      tables.push({ sheet: ws.name, name: table.table.name, columns: header, rows, formulas });
    }
  }
  return tables;
}
