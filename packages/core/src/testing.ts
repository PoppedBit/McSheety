/**
 * InMemoryWorkbookAdapter: a deterministic stand-in for Graph in tests and local demos.
 * It is NOT a shipping storage adapter; it has no persistence, auth or concurrency control.
 */
import type { ColumnTypes, TableSnapshot, WorkbookAdapter, WorkbookContext } from './adapter.js';
import { sameRow } from './adapter.js';
import { McSheetyError } from './errors.js';
import type { DataTableInfo } from './metadata.js';
import type { CellValue } from './types.js';

export interface InMemoryTable {
  readonly sheet: string;
  readonly table: string;
  columns: string[];
  rows: CellValue[][];
  formulas?: CellValue[][];
}

export interface InMemoryHooks {
  /** Called before every mutation; throw to simulate Graph failures, or edit tables to race. */
  beforeWrite?: (op: 'append' | 'update' | 'delete', table: string) => void;
}

export class InMemoryWorkbookAdapter implements WorkbookAdapter {
  readonly tables = new Map<string, InMemoryTable>();
  readonly calls: string[] = [];
  hooks: InMemoryHooks = {};
  private sessionCounter = 0;

  constructor(tables: readonly InMemoryTable[]) {
    for (const t of tables) {
      this.tables.set(t.table, {
        ...t,
        columns: [...t.columns],
        rows: t.rows.map((r) => [...r]),
        ...(t.formulas ? { formulas: t.formulas.map((r) => [...r]) } : {}),
      });
    }
  }

  private get(table: string): InMemoryTable {
    const t = this.tables.get(table);
    if (!t) throw new McSheetyError('WORKBOOK_UNAVAILABLE');
    return t;
  }

  /** Like Graph, every operation needs the caller's delegated token. */
  private async authorize(ctx: WorkbookContext): Promise<void> {
    await ctx.getGraphToken();
  }

  async listTables(ctx: WorkbookContext): Promise<ReadonlyMap<string, DataTableInfo>> {
    await this.authorize(ctx);
    this.calls.push('listTables');
    return new Map([...this.tables].map(([name, t]) => [name, { columns: [...t.columns] }]));
  }

  async readTable(ctx: WorkbookContext, table: string): Promise<TableSnapshot> {
    await this.authorize(ctx);
    this.calls.push(`readTable:${table}`);
    const t = this.get(table);
    return {
      table,
      columns: [...t.columns],
      rows: t.rows.map((r) => [...r]),
      formulas: (t.formulas ?? t.rows).map((r) => [...r]),
      location: { sheet: t.sheet, startColumn: 0, startRow: 1 },
    };
  }

  async appendRow(
    ctx: WorkbookContext,
    table: string,
    row: readonly CellValue[],
    _types?: ColumnTypes,
  ): Promise<void> {
    await this.authorize(ctx);
    this.calls.push(`appendRow:${table}`);
    this.hooks.beforeWrite?.('append', table);
    const t = this.get(table);
    t.rows.push(t.columns.map((_, i) => row[i] ?? ''));
    t.formulas?.push(t.columns.map((_, i) => row[i] ?? ''));
  }

  private verify(snapshot: TableSnapshot, rowIndex: number): InMemoryTable {
    const t = this.get(snapshot.table);
    const expected = snapshot.rows[rowIndex];
    const current = t.rows[rowIndex];
    if (
      !expected ||
      !current ||
      !sameRow(expected, current) ||
      !sameRow(t.columns, snapshot.columns)
    ) {
      throw new McSheetyError('CONFLICT');
    }
    return t;
  }

  async updateCells(
    ctx: WorkbookContext,
    snapshot: TableSnapshot,
    rowIndex: number,
    changes: ReadonlyMap<number, CellValue>,
  ): Promise<readonly CellValue[]> {
    await this.authorize(ctx);
    this.calls.push(`updateCells:${snapshot.table}`);
    this.hooks.beforeWrite?.('update', snapshot.table);
    const t = this.verify(snapshot, rowIndex);
    const row = t.rows[rowIndex] as CellValue[];
    for (const [col, value] of changes) {
      row[col] = value ?? '';
      const f = t.formulas?.[rowIndex];
      if (f) f[col] = value ?? '';
    }
    return [...row];
  }

  async deleteRow(ctx: WorkbookContext, snapshot: TableSnapshot, rowIndex: number): Promise<void> {
    await this.authorize(ctx);
    this.calls.push(`deleteRow:${snapshot.table}`);
    this.hooks.beforeWrite?.('delete', snapshot.table);
    const t = this.verify(snapshot, rowIndex);
    t.rows.splice(rowIndex, 1);
    t.formulas?.splice(rowIndex, 1);
  }

  async withSession<T>(ctx: WorkbookContext, fn: (ctx: WorkbookContext) => Promise<T>): Promise<T> {
    const sessionId = `mem-session-${++this.sessionCounter}`;
    this.calls.push(`openSession:${sessionId}`);
    try {
      return await fn({ ...ctx, sessionId });
    } finally {
      this.calls.push(`closeSession:${sessionId}`);
    }
  }
}
