import { mkdtemp, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import type { FixtureCell } from '../../../fixtures/equipment/index.js';
import { equipmentWorkbook } from '../../../fixtures/equipment/index.js';
import { validateTables } from './validate.js';
import { readWorkbookTables } from './xlsx.js';

async function writeFixture(
  edit?: (sheet: string, rows: FixtureCell[][]) => FixtureCell[][],
): Promise<string> {
  const book = new ExcelJS.Workbook();
  for (const s of equipmentWorkbook.sheets) {
    const ws = book.addWorksheet(s.sheet);
    const rows = edit
      ? edit(
          s.sheet,
          s.rows.map((r) => [...r]),
        )
      : s.rows.map((r) => [...r]);
    ws.addTable({
      name: s.table,
      ref: 'A1',
      headerRow: true,
      columns: s.columns.map((name) => ({ name })),
      rows,
    });
  }
  const dir = await mkdtemp(join(tmpdir(), 'mcsheety-cli-'));
  const file = join(dir, 'Equipment.xlsx');
  await writeFile(file, Buffer.from(await book.xlsx.writeBuffer()));
  return file;
}

describe('mcsheety validate', () => {
  it('accepts the Equipment fixture', async () => {
    const report = validateTables(await readWorkbookTables(await writeFixture()));
    expect(report.issues).toEqual([]);
    expect(report).toMatchObject({ ok: true, tables: ['Equipment', 'Locations'] });
  });

  it('reports duplicate ids, bad cells and broken policies', async () => {
    const file = await writeFixture((sheet, rows) => {
      if (sheet === 'Equipment') {
        rows[1] = ['eq-001', ...rows[1]!.slice(1)];
        rows[2]![2] = 'Retired';
      }
      if (sheet === '_rowPolicies') rows[0]![4] = 'startsWith';
      return rows;
    });
    const report = validateTables(await readWorkbookTables(file));
    expect(report.ok).toBe(false);
    const messages = report.issues.map((i) => i.message).join('\n');
    expect(messages).toContain('Unsupported operator "startsWith"');
    expect(report.tables).not.toContain('Equipment');

    const dataOnly = await writeFixture((sheet, rows) => {
      if (sheet === 'Equipment') {
        rows[1] = ['eq-001', ...rows[1]!.slice(1)];
        rows[2]![2] = 'Retired';
      }
      return rows;
    });
    const dataReport = validateTables(await readWorkbookTables(dataOnly));
    const dataMessages = dataReport.issues.map((i) => i.message);
    expect(dataMessages).toContain('Duplicate id "eq-001"; these rows are hidden until fixed.');
    expect(dataMessages).toContain('eq-003: status must be one of: Active, Inactive.');
  });
});
