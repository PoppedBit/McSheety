/**
 * Storage contract implemented by @mcsheety/adapter-graph (and the in-memory test adapter).
 * Adapters move cells; they make no authorization decisions.
 */
import type { DataTableInfo } from './metadata.js';
import type { CellValue, FieldType, RawTable } from './types.js';

export interface WorkbookContext {
  readonly correlationId: string;
  /** Delegated Graph token for the current user (acquired via OBO). Server-side only; never log. */
  readonly getGraphToken: () => Promise<string>;
  /** Set inside `withSession`. */
  readonly sessionId?: string;
}

/** Where a table sat in its worksheet when it was read. */
export interface TableLocation {
  readonly sheet: string;
  /** 0-based column index of the table's first column (A = 0). */
  readonly startColumn: number;
  /** 1-based worksheet row of the table header. Body row i is at startRow + 1 + i. */
  readonly startRow: number;
}

/** A consistent read of one table (header, body values, formulas) plus its location. */
export interface TableSnapshot extends RawTable {
  readonly table: string;
  readonly location: TableLocation;
}

/** How each cell should be encoded when written, by header position. */
export type ColumnTypes = readonly (FieldType | undefined)[];

export interface WorkbookAdapter {
  /** Excel table name -> header names, for every table in the workbook. */
  listTables(ctx: WorkbookContext): Promise<ReadonlyMap<string, DataTableInfo>>;

  /** Reads a table's header, body values and formulas in one consistent snapshot. */
  readTable(ctx: WorkbookContext, table: string): Promise<TableSnapshot>;

  /**
   * Appends one row (values in header order). Not retried on ambiguous failures: callers
   * check for the new id before retrying, so a timeout can never create duplicates.
   */
  appendRow(
    ctx: WorkbookContext,
    table: string,
    row: readonly CellValue[],
    types: ColumnTypes,
  ): Promise<void>;

  /**
   * Writes only the changed cells of one body row. The adapter first re-reads the row and
   * throws CONFLICT unless it still equals `snapshot.rows[rowIndex]`; it returns the row as
   * observed after the write.
   */
  updateCells(
    ctx: WorkbookContext,
    snapshot: TableSnapshot,
    rowIndex: number,
    changes: ReadonlyMap<number, CellValue>,
    types: ColumnTypes,
  ): Promise<readonly CellValue[]>;

  /** Deletes one body row after the same re-read check as `updateCells`. */
  deleteRow(ctx: WorkbookContext, snapshot: TableSnapshot, rowIndex: number): Promise<void>;

  /**
   * Runs `fn` inside a persistent workbook session (changes are saved) and closes it after.
   * Graph operations inside run sequentially.
   */
  withSession<T>(ctx: WorkbookContext, fn: (ctx: WorkbookContext) => Promise<T>): Promise<T>;
}

/** 0-based column index -> Excel column letters (0 -> A, 26 -> AA). */
export function columnLetters(index: number): string {
  let n = index + 1;
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Excel column letters -> 0-based index (A -> 0). */
export function columnIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/** Cells compare equal when Graph would report them the same way (blank "" equals null). */
export function sameRow(a: readonly CellValue[], b: readonly CellValue[]): boolean {
  const width = Math.max(a.length, b.length);
  for (let i = 0; i < width; i++) {
    const x = a[i] ?? '';
    const y = b[i] ?? '';
    if (x !== y) return false;
  }
  return true;
}
