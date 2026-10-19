import { McSheetyError } from './errors.js';
import type { CellValue, RawTable, RecordValues, TableSchema } from './types.js';
import { decodeCell } from './values.js';

export interface FieldIssue {
  readonly field: string;
  readonly message: string;
}

export interface DecodedRecord {
  readonly id: string;
  /**
   * Zero-based body row index observed in this read. Only valid for the current operation:
   * sorting or editing in Excel changes it, so it is never a record identity.
   */
  readonly index: number;
  readonly values: RecordValues;
  /** Cells that could not be decoded (their value is null in `values`). */
  readonly issues: readonly FieldIssue[];
  /** Columns whose cell holds a formula; writes to these are rejected. */
  readonly formulaColumns: ReadonlySet<string>;
}

export interface TableProblems {
  /** Body rows with a blank or invalid id; unreachable through the API. */
  readonly rowsWithoutId: readonly number[];
  /** IDs shared by more than one row; excluded from reads and mutations until fixed. */
  readonly duplicateIds: readonly string[];
}

export interface DecodedTable {
  readonly records: readonly DecodedRecord[];
  readonly problems: TableProblems;
  readonly duplicateIds: ReadonlySet<string>;
}

const isFormula = (v: CellValue | undefined): boolean => typeof v === 'string' && v.startsWith('=');

export function decodeTable(table: TableSchema, raw: RawTable): DecodedTable {
  const position = new Map(raw.columns.map((name, i) => [name, i]));
  const decoded: DecodedRecord[] = [];
  const rowsWithoutId: number[] = [];
  const seen = new Map<string, number>();

  raw.rows.forEach((cells, index) => {
    const values: Record<string, CellValue> = {};
    const issues: FieldIssue[] = [];
    const formulaColumns = new Set<string>();
    for (const column of table.columns) {
      const at = position.get(column.name);
      const cell = at === undefined ? null : (cells[at] ?? null);
      if (at !== undefined && isFormula(raw.formulas?.[index]?.[at]))
        formulaColumns.add(column.name);
      const result = decodeCell(column, cell);
      if (result.ok) values[column.name] = result.value;
      else {
        values[column.name] = null;
        issues.push({ field: column.name, message: `${column.name} ${result.reason}.` });
      }
    }
    const id = values['id'];
    if (typeof id !== 'string' || id === '') {
      rowsWithoutId.push(index + 1);
      return;
    }
    seen.set(id, (seen.get(id) ?? 0) + 1);
    decoded.push({ id, index, values, issues, formulaColumns });
  });

  const duplicateIds = new Set([...seen].filter(([, n]) => n > 1).map(([id]) => id));
  return {
    records: decoded.filter((r) => !duplicateIds.has(r.id)),
    problems: { rowsWithoutId, duplicateIds: [...duplicateIds] },
    duplicateIds,
  };
}

/**
 * Locates a record by immutable id. Throws NOT_FOUND when absent and CONFLICT when the id
 * is duplicated in the sheet (manual edits), since the target row is then ambiguous.
 */
export function findRecord(table: DecodedTable, id: string): DecodedRecord {
  if (table.duplicateIds.has(id)) {
    throw new McSheetyError('CONFLICT', 'The record id is duplicated in the workbook.');
  }
  const record = table.records.find((r) => r.id === id);
  if (!record) throw new McSheetyError('NOT_FOUND');
  return record;
}

/**
 * Deterministic serialization of a record view, used as revision input. Callers hash the
 * caller's projection, so revisions do not reveal changes to columns the caller cannot read.
 */
export function canonicalRecord(values: RecordValues): string {
  const entries = Object.keys(values)
    .sort()
    .map((key) => [key, values[key] ?? null]);
  return JSON.stringify(entries);
}
