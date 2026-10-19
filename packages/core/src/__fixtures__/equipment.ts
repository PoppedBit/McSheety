/** Test helpers over fixtures/equipment/workbook.json. Not part of the published package. */
import type { FixtureSheet } from '../../../../fixtures/equipment/index.js';
import { equipmentWorkbook } from '../../../../fixtures/equipment/index.js';
import type { WorkbookSource } from '../metadata.js';
import { parseWorkbookModel } from '../metadata.js';
import type { CellValue, Identity, RawTable, TableSchema, WorkbookModel } from '../types.js';

export type SheetEdit = (sheet: FixtureSheet) => FixtureSheet;

/** Returns sheets with optional per-sheet edits applied (for negative tests). */
export function fixtureSheets(edits: Record<string, SheetEdit> = {}): Map<string, FixtureSheet> {
  return new Map(equipmentWorkbook.sheets.map((s) => [s.table, edits[s.sheet]?.(s) ?? s] as const));
}

export const toRaw = (sheet: FixtureSheet): RawTable => ({
  columns: sheet.columns,
  rows: sheet.rows,
});

export function sourceFrom(sheets: Map<string, FixtureSheet>): WorkbookSource {
  const get = (name: string): RawTable | undefined => {
    const s = sheets.get(name);
    return s ? toRaw(s) : undefined;
  };
  const dataTables = new Map(
    [...sheets.values()]
      .filter((s) => !s.table.startsWith('mc_'))
      .map((s) => [s.table, { columns: s.columns }] as const),
  );
  return {
    config: get('mc_config'),
    schema: get('mc_schema'),
    permissions: get('mc_permissions'),
    columns: get('mc_columns'),
    rowPolicies: get('mc_rowPolicies'),
    dataTables,
  };
}

export function fixtureModel(edits: Record<string, SheetEdit> = {}): WorkbookModel {
  return parseWorkbookModel(sourceFrom(fixtureSheets(edits))).model;
}

export function table(model: WorkbookModel, name = 'Equipment'): TableSchema {
  const t = model.tables.get(name);
  if (!t) throw new Error(`table ${name} missing from model`);
  return t;
}

export function rawData(name = 'Equipment', edits: Record<string, SheetEdit> = {}): RawTable {
  const sheet = [...fixtureSheets(edits).values()].find((s) => s.sheet === name);
  if (!sheet) throw new Error(`sheet ${name} missing`);
  return toRaw(sheet);
}

/** Replaces matching rows' cells. `match` gets the row; `patch` maps column -> new value. */
export function editRows(
  match: (row: readonly CellValue[], columns: readonly string[]) => boolean,
  patch: Record<string, CellValue>,
): SheetEdit {
  return (sheet) => ({
    ...sheet,
    rows: sheet.rows.map((row) =>
      match(row, sheet.columns)
        ? sheet.columns.map((c, i) => (c in patch ? (patch[c] ?? null) : (row[i] ?? null)))
        : row,
    ),
  });
}

export const cell = (
  row: readonly CellValue[],
  columns: readonly string[],
  name: string,
): CellValue => row[columns.indexOf(name)] ?? null;

const base = { tenantId: 'tenant-1', attributes: {} };
export const users = {
  viewer: { ...base, userId: 'u-viewer', roles: ['viewer'] },
  editorOps: {
    ...base,
    userId: 'u-ed-ops',
    roles: ['editor'],
    attributes: { departmentId: 'dept-ops' },
  },
  editorNoDept: { ...base, userId: 'u-ed-none', roles: ['editor'] },
  manager: {
    ...base,
    userId: 'u-mgr',
    roles: ['manager'],
    attributes: { departmentId: 'dept-eng' },
  },
  admin: { ...base, userId: 'u-admin', roles: ['admin'] },
  nobody: { ...base, userId: 'u-none', roles: [] },
  spoofed: { ...base, userId: 'u-spoof', roles: ['superuser', 'Admin'] },
} satisfies Record<string, Identity>;
