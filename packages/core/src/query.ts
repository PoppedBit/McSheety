/**
 * List queries (implementation.md §6 "Querying behavior", §10). Filtering, sorting and paging
 * run on the caller's projected records, after row authorization. A column hidden on a row
 * behaves as "no value" there, so filters and sorts cannot be used as oracles.
 */
import { McSheetyError } from './errors.js';
import { readableColumns } from './policy.js';
import type { CellValue, Identity, RecordValues, TableSchema } from './types.js';
import { compareValues, parseInput } from './values.js';

export type Condition =
  | { readonly op: 'eq'; readonly value: CellValue }
  | { readonly op: 'ne'; readonly value: CellValue }
  | { readonly op: 'in'; readonly values: readonly CellValue[] };

export interface SortKey {
  readonly field: string;
  readonly direction: 'asc' | 'desc';
}

export interface ListQuery {
  readonly where: ReadonlyMap<string, Condition>;
  readonly sort: readonly SortKey[];
  readonly limit: number;
  readonly offset: number;
}

export interface ListQueryParams {
  /** JSON object, e.g. {"status":"Active","locationId":{"in":["loc-1","loc-2"]}}. */
  readonly where?: string | undefined;
  /** Comma-separated fields; prefix with "-" for descending, e.g. "name,-createdAt". */
  readonly sort?: string | undefined;
  readonly limit?: string | undefined;
  readonly cursor?: string | undefined;
}

export interface QueryLimits {
  readonly defaultLimit: number;
  readonly maxLimit: number;
  readonly maxFilters: number;
  readonly maxInValues: number;
  readonly maxSortKeys: number;
}

export const DEFAULT_QUERY_LIMITS: QueryLimits = {
  defaultLimit: 50,
  maxLimit: 100,
  maxFilters: 10,
  maxInValues: 50,
  maxSortKeys: 3,
};

const invalid = (message: string): McSheetyError => new McSheetyError('INVALID_REQUEST', message);

/** Small stable string hash (FNV-1a, 32-bit); binds a cursor to the query it came from. */
function fingerprint(text: string): string {
  let h = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) {
    h ^= text.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return (h >>> 0).toString(36);
}

function queryFingerprint(where: ReadonlyMap<string, Condition>, sort: readonly SortKey[]): string {
  return fingerprint(JSON.stringify([[...where].sort(([a], [b]) => (a < b ? -1 : 1)), sort]));
}

const toBase64Url = (s: string): string =>
  btoa(s).replaceAll('+', '-').replaceAll('/', '_').replace(/=+$/, '');
const fromBase64Url = (s: string): string => atob(s.replaceAll('-', '+').replaceAll('_', '/'));

/** Opaque cursor: an offset into the caller's authorized, filtered, sorted result. */
export function encodeCursor(offset: number, query: Pick<ListQuery, 'where' | 'sort'>): string {
  return toBase64Url(JSON.stringify({ o: offset, q: queryFingerprint(query.where, query.sort) }));
}

function decodeCursor(cursor: string, expected: string): number {
  try {
    const parsed = JSON.parse(fromBase64Url(cursor)) as { o?: unknown; q?: unknown };
    if (
      typeof parsed.o === 'number' &&
      Number.isInteger(parsed.o) &&
      parsed.o >= 0 &&
      parsed.q === expected
    ) {
      return parsed.o;
    }
  } catch {
    // fall through
  }
  throw invalid('The cursor is invalid or does not match this query.');
}

export function parseListQuery(
  table: TableSchema,
  identity: Identity,
  params: ListQueryParams,
  limits: QueryLimits = DEFAULT_QUERY_LIMITS,
): ListQuery {
  const readable = new Map(readableColumns(table, identity).map((c) => [c.name, c]));
  const column = (field: string) => {
    const c = readable.get(field);
    if (!c) throw invalid(`Unknown field ${field}.`);
    return c;
  };

  const where = new Map<string, Condition>();
  if (params.where !== undefined && params.where !== '') {
    let parsed: unknown;
    try {
      parsed = JSON.parse(params.where);
    } catch {
      throw invalid('where must be a JSON object.');
    }
    if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
      throw invalid('where must be a JSON object.');
    }
    const entries = Object.entries(parsed as Record<string, unknown>);
    if (entries.length > limits.maxFilters)
      throw invalid(`At most ${limits.maxFilters} filters are allowed.`);
    for (const [field, raw] of entries) {
      const col = column(field);
      const value = (v: unknown): CellValue => {
        const r = parseInput(col, v, { forFilter: true });
        if (!r.ok) throw invalid(`Filter on ${field}: value ${r.reason}.`);
        return r.value;
      };
      if (raw === null || typeof raw !== 'object') {
        where.set(field, { op: 'eq', value: value(raw) });
        continue;
      }
      const keys = Object.keys(raw);
      const op = keys[0];
      if (keys.length !== 1 || (op !== 'eq' && op !== 'ne' && op !== 'in')) {
        throw invalid(`Filter on ${field} must use exactly one of eq, ne, in.`);
      }
      const operand = (raw as Record<string, unknown>)[op];
      if (op === 'in') {
        if (
          !Array.isArray(operand) ||
          operand.length === 0 ||
          operand.length > limits.maxInValues
        ) {
          throw invalid(`Filter on ${field}: in needs 1-${limits.maxInValues} values.`);
        }
        where.set(field, { op, values: operand.map(value) });
      } else {
        where.set(field, { op, value: value(operand) });
      }
    }
  }

  const sort: SortKey[] = [];
  if (params.sort !== undefined && params.sort !== '') {
    const parts = params.sort.split(',').map((p) => p.trim());
    if (parts.length > limits.maxSortKeys)
      throw invalid(`At most ${limits.maxSortKeys} sort fields are allowed.`);
    for (const part of parts) {
      const desc = part.startsWith('-');
      const field = desc ? part.slice(1) : part;
      column(field);
      sort.push({ field, direction: desc ? 'desc' : 'asc' });
    }
  }

  let limit = limits.defaultLimit;
  if (params.limit !== undefined && params.limit !== '') {
    limit = Number(params.limit);
    if (!Number.isInteger(limit) || limit < 1 || limit > limits.maxLimit) {
      throw invalid(`limit must be an integer from 1 to ${limits.maxLimit}.`);
    }
  }

  const offset =
    params.cursor === undefined || params.cursor === ''
      ? 0
      : decodeCursor(params.cursor, queryFingerprint(where, sort));

  return { where, sort, limit, offset };
}

function matches(record: RecordValues, field: string, condition: Condition): boolean {
  // A column absent from the caller's projection of this row never matches.
  if (!(field in record)) return false;
  const actual = record[field] ?? null;
  switch (condition.op) {
    case 'eq':
      return actual === condition.value;
    case 'ne':
      return actual !== condition.value;
    case 'in':
      return condition.values.includes(actual);
  }
}

export interface QueryPage {
  readonly items: readonly RecordValues[];
  readonly nextCursor: string | null;
}

/** Filters, sorts (ties broken by id) and pages already-projected records. */
export function applyListQuery(records: readonly RecordValues[], query: ListQuery): QueryPage {
  const filtered = records.filter((r) =>
    [...query.where].every(([field, condition]) => matches(r, field, condition)),
  );
  const keys: SortKey[] = [...query.sort, { field: 'id', direction: 'asc' }];
  const sorted = [...filtered].sort((a, b) => {
    for (const key of keys) {
      const av = key.field in a ? (a[key.field] ?? null) : null;
      const bv = key.field in b ? (b[key.field] ?? null) : null;
      if (av === null || bv === null) {
        if (av !== bv) return av === null ? 1 : -1; // nulls last in both directions
        continue;
      }
      const c = compareValues(av, bv);
      if (c !== 0) return key.direction === 'asc' ? c : -c;
    }
    return 0;
  });
  const items = sorted.slice(query.offset, query.offset + query.limit);
  const nextOffset = query.offset + items.length;
  return {
    items,
    nextCursor: nextOffset < sorted.length ? encodeCursor(nextOffset, query) : null,
  };
}
