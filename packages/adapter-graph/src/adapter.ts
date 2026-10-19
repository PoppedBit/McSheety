/**
 * WorkbookAdapter over Microsoft Graph v1.0 Excel APIs for one admin-configured workbook
 * (drive ID + item ID). Request paths never include caller-supplied workbook identifiers.
 *
 * Endpoint choices (verify in the MC-002 spike, see docs/decisions/0002):
 * - read:   tables/{t}/headerRowRange + tables/{t}/dataBodyRange (values, formulas, address)
 * - append: POST tables/{t}/rows/add
 * - update: PATCH worksheets/{ws}/range(address=...) for each run of changed cells only, so
 *           formulas and other columns in the row are never overwritten
 * - delete: POST worksheets/{ws}/range(address=<row>)/delete { shift: "Up" }
 * Every update/delete first re-reads the target row and requires it to equal the snapshot.
 */
import type {
  CellValue,
  ColumnTypes,
  DataTableInfo,
  TableLocation,
  TableSnapshot,
  WorkbookAdapter,
  WorkbookContext,
} from '@mcsheety/core';
import { McSheetyError, columnIndex, columnLetters, sameRow } from '@mcsheety/core';
import type { TextEncoding } from './encoding.js';
import { encodeCell } from './encoding.js';
import type { GraphHttpOptions } from './http.js';
import { GraphHttp } from './http.js';

export interface GraphWorkbookAdapterOptions extends GraphHttpOptions {
  readonly driveId: string;
  readonly itemId: string;
  /** Product safeguard (implementation.md §1); reads of larger tables fail. */
  readonly maxRows?: number;
  readonly textEncoding?: TextEncoding;
}

interface RangeResponse {
  readonly address: string;
  readonly values: CellValue[][];
  readonly formulas?: CellValue[][];
}

interface Page<T> {
  readonly value: T[];
  readonly '@odata.nextLink'?: string;
}

/** Splits "Sheet!A1:G9" or "'My Sheet'!B3:H3" into sheet name, first column and first row. */
export function parseAddress(address: string): {
  sheet: string;
  startColumn: number;
  startRow: number;
} {
  const bang = address.lastIndexOf('!');
  const rawSheet = address.slice(0, bang);
  const sheet = rawSheet.startsWith("'") ? rawSheet.slice(1, -1).replaceAll("''", "'") : rawSheet;
  const m = /^\$?([A-Z]+)\$?(\d+)/i.exec(address.slice(bang + 1));
  if (bang < 0 || !m?.[1] || !m[2])
    throw new McSheetyError('GRAPH_ERROR', undefined, { cause: { address } });
  return { sheet, startColumn: columnIndex(m[1]), startRow: Number(m[2]) };
}

const seg = encodeURIComponent;

export class GraphWorkbookAdapter implements WorkbookAdapter {
  private readonly http: GraphHttp;
  private readonly root: string;
  private readonly maxRows: number;
  private readonly textEncoding: TextEncoding;

  constructor(options: GraphWorkbookAdapterOptions) {
    this.http = new GraphHttp(options);
    this.root = `/drives/${seg(options.driveId)}/items/${seg(options.itemId)}/workbook`;
    this.maxRows = options.maxRows ?? 2000;
    this.textEncoding = options.textEncoding ?? 'apostrophe';
  }

  private get<T>(ctx: WorkbookContext, path: string): Promise<T> {
    return this.http.request<T>(ctx, { method: 'GET', path, idempotent: true });
  }

  private rowAddress(
    location: TableLocation,
    width: number,
    rowIndex: number,
    from = 0,
    to = width - 1,
  ): string {
    const row = location.startRow + 1 + rowIndex;
    const a = columnLetters(location.startColumn + from);
    const b = columnLetters(location.startColumn + to);
    return `${a}${row}:${b}${row}`;
  }

  private rangePath(location: TableLocation, address: string): string {
    return `${this.root}/worksheets/${seg(location.sheet)}/range(address='${address}')`;
  }

  async listTables(ctx: WorkbookContext): Promise<ReadonlyMap<string, DataTableInfo>> {
    const names: string[] = [];
    let next: string | undefined = `${this.root}/tables?$select=name`;
    while (next) {
      const page: Page<{ name: string }> = await this.get(ctx, next);
      names.push(...page.value.map((t) => t.name));
      next = page['@odata.nextLink'];
    }
    const result = new Map<string, DataTableInfo>();
    for (const name of names) {
      const header = await this.get<RangeResponse>(
        ctx,
        `${this.root}/tables/${seg(name)}/headerRowRange?$select=values`,
      );
      result.set(name, { columns: (header.values[0] ?? []).map((v) => String(v)) });
    }
    return result;
  }

  async readTable(ctx: WorkbookContext, table: string): Promise<TableSnapshot> {
    const base = `${this.root}/tables/${seg(table)}`;
    const header = await this.get<RangeResponse>(
      ctx,
      `${base}/headerRowRange?$select=address,values`,
    );
    const body = await this.get<RangeResponse>(
      ctx,
      `${base}/dataBodyRange?$select=address,values,formulas`,
    );
    if (body.values.length > this.maxRows) {
      throw new McSheetyError(
        'WORKBOOK_UNAVAILABLE',
        'The table exceeds the configured row limit.',
        {
          retryable: false,
        },
      );
    }
    return {
      table,
      columns: (header.values[0] ?? []).map((v) => String(v)),
      rows: body.values,
      formulas: body.formulas ?? body.values,
      location: parseAddress(header.address),
    };
  }

  async appendRow(
    ctx: WorkbookContext,
    table: string,
    row: readonly CellValue[],
    types: ColumnTypes,
  ): Promise<void> {
    await this.http.request(ctx, {
      method: 'POST',
      path: `${this.root}/tables/${seg(table)}/rows/add`,
      body: { values: [row.map((v, i) => encodeCell(v, types[i], this.textEncoding))] },
      idempotent: false,
    });
  }

  private async readRow(
    ctx: WorkbookContext,
    snapshot: TableSnapshot,
    rowIndex: number,
  ): Promise<CellValue[]> {
    const address = this.rowAddress(snapshot.location, snapshot.columns.length, rowIndex);
    const range = await this.get<RangeResponse>(
      ctx,
      `${this.rangePath(snapshot.location, address)}?$select=values`,
    );
    return range.values[0] ?? [];
  }

  private async verifyRow(
    ctx: WorkbookContext,
    snapshot: TableSnapshot,
    rowIndex: number,
  ): Promise<void> {
    const expected = snapshot.rows[rowIndex];
    const current = await this.readRow(ctx, snapshot, rowIndex);
    if (!expected || !sameRow(expected, current)) {
      throw new McSheetyError('CONFLICT', 'The record changed in the workbook since it was read.');
    }
  }

  async updateCells(
    ctx: WorkbookContext,
    snapshot: TableSnapshot,
    rowIndex: number,
    changes: ReadonlyMap<number, CellValue>,
    types: ColumnTypes,
  ): Promise<readonly CellValue[]> {
    await this.verifyRow(ctx, snapshot, rowIndex);
    const columns = [...changes.keys()].sort((a, b) => a - b);
    // Group adjacent changed columns into one PATCH each.
    const runs: number[][] = [];
    for (const c of columns) {
      const last = runs.at(-1);
      if (last && last.at(-1) === c - 1) last.push(c);
      else runs.push([c]);
    }
    for (const run of runs) {
      const first = run[0] as number;
      const last = run.at(-1) as number;
      const address = this.rowAddress(
        snapshot.location,
        snapshot.columns.length,
        rowIndex,
        first,
        last,
      );
      await this.http.request(ctx, {
        method: 'PATCH',
        path: this.rangePath(snapshot.location, address),
        body: {
          values: [run.map((c) => encodeCell(changes.get(c) ?? null, types[c], this.textEncoding))],
        },
        idempotent: true,
      });
    }
    return this.readRow(ctx, snapshot, rowIndex);
  }

  async deleteRow(ctx: WorkbookContext, snapshot: TableSnapshot, rowIndex: number): Promise<void> {
    await this.verifyRow(ctx, snapshot, rowIndex);
    const address = this.rowAddress(snapshot.location, snapshot.columns.length, rowIndex);
    await this.http.request(ctx, {
      method: 'POST',
      path: `${this.rangePath(snapshot.location, address)}/delete`,
      body: { shift: 'Up' },
      idempotent: false,
    });
  }

  async withSession<T>(ctx: WorkbookContext, fn: (ctx: WorkbookContext) => Promise<T>): Promise<T> {
    const session = await this.http.request<{ id: string }>(ctx, {
      method: 'POST',
      path: `${this.root}/createSession`,
      body: { persistChanges: true },
      idempotent: true,
    });
    const inSession: WorkbookContext = { ...ctx, sessionId: session.id };
    try {
      return await fn(inSession);
    } finally {
      try {
        await this.http.request(inSession, {
          method: 'POST',
          path: `${this.root}/closeSession`,
          idempotent: true,
        });
      } catch {
        // Sessions also expire on their own; a failed close must not mask the result.
      }
    }
  }
}
