import workbook from './workbook.json' with { type: 'json' };

export type FixtureCell = string | number | boolean | null;

export interface FixtureSheet {
  readonly sheet: string;
  readonly table: string;
  readonly columns: readonly string[];
  readonly rows: readonly (readonly FixtureCell[])[];
}

export interface FixtureWorkbook {
  readonly file: string;
  readonly sheets: readonly FixtureSheet[];
}

export const equipmentWorkbook: FixtureWorkbook = workbook;

export function fixtureSheet(name: string): FixtureSheet {
  const sheet = equipmentWorkbook.sheets.find((s) => s.sheet === name);
  if (!sheet) throw new Error(`Fixture sheet not found: ${name}`);
  return sheet;
}
