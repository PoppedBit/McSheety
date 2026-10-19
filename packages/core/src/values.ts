import type { CellValue, ColumnSchema } from './types.js';

export type Parsed<T> =
  { readonly ok: true; readonly value: T } | { readonly ok: false; readonly reason: string };

const ok = <T>(value: T): Parsed<T> => ({ ok: true, value });
const fail = (reason: string): Parsed<never> => ({ ok: false, reason });

/** Excel's per-cell text limit. */
export const EXCEL_MAX_TEXT_LENGTH = 32767;

const DATE_RE = /^(\d{4})-(\d{2})-(\d{2})$/;
const TIME_RE = /^(\d{2}):(\d{2})(?::(\d{2}))?$/;
const DATETIME_RE =
  /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2})(\.\d{1,3})?)?(Z|[+-]\d{2}:\d{2})$/;
const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const NUMERIC_TEXT_RE = /^-?(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?$/;
/**
 * Leading characters that make Excel treat typed input as a formula (or that are commonly
 * abused for CSV/formula injection). Rejected at the start of user-supplied text.
 */
const FORMULA_LEAD_RE = /^[=+\-@\t\r\n]/;
/** Day 0 of the Excel 1900 date system as used for serials >= 61 (1900-03-01 onward). */
const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;

const pad = (n: number, width = 2): string => String(n).padStart(width, '0');

function isValidYmd(y: number, m: number, d: number): boolean {
  if (m < 1 || m > 12 || d < 1) return false;
  const daysInMonth = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return d <= daysInMonth;
}

export function parseDate(text: string): Parsed<string> {
  const m = DATE_RE.exec(text);
  if (!m || !isValidYmd(Number(m[1]), Number(m[2]), Number(m[3]))) {
    return fail('must be a date in YYYY-MM-DD format');
  }
  return ok(text);
}

export function parseTime(text: string): Parsed<string> {
  const m = TIME_RE.exec(text);
  if (!m) return fail('must be a time in HH:mm:ss format');
  const [h, min, s] = [Number(m[1]), Number(m[2]), Number(m[3] ?? '0')];
  if (h > 23 || min > 59 || s > 59) return fail('must be a time in HH:mm:ss format');
  return ok(`${pad(h)}:${pad(min)}:${pad(s)}`);
}

/** Accepts ISO 8601 with an explicit `Z` or offset and normalizes to UTC (`...Z`). */
export function parseDateTime(text: string): Parsed<string> {
  const m = DATETIME_RE.exec(text);
  const invalid = 'must be an ISO 8601 date-time with a timezone, e.g. 2026-10-16T14:05:00Z';
  if (!m) return fail(invalid);
  const [h, min, s] = [Number(m[4]), Number(m[5]), Number(m[6] ?? '0')];
  if (!isValidYmd(Number(m[1]), Number(m[2]), Number(m[3])) || h > 23 || min > 59 || s > 59) {
    return fail(invalid);
  }
  const date = new Date(text);
  if (Number.isNaN(date.getTime())) return fail(invalid);
  return ok(formatDateTime(date));
}

/** UTC ISO string without a `.000` millisecond suffix. */
export function formatDateTime(date: Date): string {
  return date.toISOString().replace('.000Z', 'Z');
}

function excelSerialToDate(serial: number): Parsed<string> {
  if (!Number.isInteger(serial) || serial < 61) {
    return fail('is an Excel date serial that cannot be converted safely');
  }
  const d = new Date(EXCEL_EPOCH_MS + serial * MS_PER_DAY);
  return ok(`${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`);
}

function excelFractionToTime(fraction: number): Parsed<string> {
  if (fraction < 0 || fraction >= 1) return fail('is an Excel time serial out of range');
  const total = Math.round(fraction * 86_400) % 86_400;
  return ok(
    `${pad(Math.floor(total / 3600))}:${pad(Math.floor(total / 60) % 60)}:${pad(total % 60)}`,
  );
}

function checkConstraints(
  column: ColumnSchema,
  value: Exclude<CellValue, null>,
): Parsed<CellValue> {
  if (column.type === 'enum' && !column.enumValues?.includes(String(value))) {
    return fail(`must be one of: ${(column.enumValues ?? []).join(', ')}`);
  }
  if (typeof value === 'number') {
    if (column.min !== undefined && value < column.min)
      return fail(`must be at least ${column.min}`);
    if (column.max !== undefined && value > column.max)
      return fail(`must be at most ${column.max}`);
  }
  if (typeof value === 'string') {
    const limit = Math.min(column.maxLength ?? EXCEL_MAX_TEXT_LENGTH, EXCEL_MAX_TEXT_LENGTH);
    if (value.length > limit) return fail(`must be at most ${limit} characters`);
  }
  return ok(value);
}

/**
 * Read path: converts a value stored in Excel into the column's normalized form.
 * Blank cells ("" or null) are null. Conversions are explicit; anything ambiguous fails.
 */
export function decodeCell(column: ColumnSchema, cell: CellValue): Parsed<CellValue> {
  if (cell === null || cell === '') return ok(null);
  let decoded: Parsed<Exclude<CellValue, null>>;
  switch (column.type) {
    case 'string':
      decoded =
        typeof cell === 'string'
          ? ok(cell)
          : typeof cell === 'number'
            ? ok(String(cell))
            : fail('must be text');
      break;
    case 'enum':
    case 'uuid':
      decoded = typeof cell === 'string' ? ok(cell) : fail('must be text');
      if (decoded.ok && column.type === 'uuid' && !UUID_RE.test(cell as string)) {
        decoded = fail('must be a UUID');
      }
      break;
    case 'number':
      if (typeof cell === 'number' && Number.isFinite(cell)) decoded = ok(cell);
      else if (typeof cell === 'string' && NUMERIC_TEXT_RE.test(cell.trim())) {
        decoded = ok(Number(cell.trim()));
      } else decoded = fail('must be a number');
      break;
    case 'boolean':
      if (typeof cell === 'boolean') decoded = ok(cell);
      else if (typeof cell === 'string' && /^(true|false)$/i.test(cell.trim())) {
        decoded = ok(cell.trim().toLowerCase() === 'true');
      } else decoded = fail('must be TRUE or FALSE');
      break;
    case 'date':
      decoded =
        typeof cell === 'number'
          ? excelSerialToDate(cell)
          : typeof cell === 'string'
            ? parseDate(cell)
            : fail('must be a date');
      break;
    case 'time':
      decoded =
        typeof cell === 'number'
          ? excelFractionToTime(cell)
          : typeof cell === 'string'
            ? parseTime(cell)
            : fail('must be a time');
      break;
    case 'datetime':
      // A numeric serial has no timezone; converting it would silently guess one.
      decoded =
        typeof cell === 'string'
          ? parseDateTime(cell)
          : fail('must be stored as ISO 8601 text with a timezone');
      break;
  }
  return decoded.ok ? checkConstraints(column, decoded.value) : decoded;
}

/**
 * Write path: strictly validates a JSON value supplied by an API client. No implicit
 * coercion between JSON types (e.g. "12" is not a number).
 */
export function parseInput(
  column: ColumnSchema,
  value: unknown,
  options: { readonly forFilter?: boolean } = {},
): Parsed<CellValue> {
  if (value === null || value === undefined) return ok(null);
  let parsed: Parsed<Exclude<CellValue, null>>;
  switch (column.type) {
    case 'string':
    case 'enum':
    case 'uuid': {
      if (typeof value !== 'string') return fail('must be a string');
      if (value === '') return ok(null);
      // Filter values are never written to the workbook, so formula-like text is harmless there.
      if (!options.forFilter && FORMULA_LEAD_RE.test(value)) {
        return fail('must not start with =, +, -, @ or a control character');
      }
      if (column.type === 'uuid') {
        parsed = UUID_RE.test(value) ? ok(value.toLowerCase()) : fail('must be a UUID');
      } else parsed = ok(value);
      break;
    }
    case 'number':
      if (typeof value !== 'number' || !Number.isFinite(value)) return fail('must be a number');
      parsed = ok(value);
      break;
    case 'boolean':
      if (typeof value !== 'boolean') return fail('must be true or false');
      parsed = ok(value);
      break;
    case 'date':
      parsed = typeof value === 'string' ? parseDate(value) : fail('must be a string');
      break;
    case 'time':
      parsed = typeof value === 'string' ? parseTime(value) : fail('must be a string');
      break;
    case 'datetime':
      parsed = typeof value === 'string' ? parseDateTime(value) : fail('must be a string');
      break;
  }
  return parsed.ok ? checkConstraints(column, parsed.value) : parsed;
}

/** Total order for sorting: nulls last, then numbers, booleans (false < true), strings. */
export function compareValues(a: CellValue, b: CellValue): number {
  if (a === b) return 0;
  if (a === null) return 1;
  if (b === null) return -1;
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'boolean' && typeof b === 'boolean') return a ? 1 : -1;
  const sa = String(a);
  const sb = String(b);
  return sa < sb ? -1 : sa > sb ? 1 : 0;
}
