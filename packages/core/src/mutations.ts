/**
 * Write preparation (implementation.md §7 "Order of checks for writes"). These functions
 * decide whether a mutation may happen and what the resulting row is; they never touch Graph.
 */
import type { ErrorDetail } from './errors.js';
import { McSheetyError } from './errors.js';
import {
  qualifyingRoles,
  readableColumns,
  rolesGranting,
  writableColumnsForRole,
} from './policy.js';
import type { DecodedRecord, DecodedTable } from './records.js';
import type {
  CellValue,
  ColumnSchema,
  Identity,
  RecordValues,
  TableAction,
  TableSchema,
} from './types.js';
import { formatDateTime, parseInput } from './values.js';

export interface MutationDeps {
  readonly now: () => Date;
  readonly newId: () => string;
}

export interface PreparedUpdate {
  /** Only the columns whose value changes. */
  readonly changes: RecordValues;
  /** The full record after the update. */
  readonly next: RecordValues;
}

function asObject(input: unknown): Record<string, unknown> {
  if (typeof input !== 'object' || input === null || Array.isArray(input)) {
    throw new McSheetyError('INVALID_REQUEST', 'The request body must be a JSON object.');
  }
  return input as Record<string, unknown>;
}

/**
 * Fields the caller may name at all. Columns the caller can neither read nor write are
 * reported exactly like nonexistent ones, so requests cannot probe for hidden columns.
 */
function resolveFields(
  table: TableSchema,
  identity: Identity,
  body: Record<string, unknown>,
): ColumnSchema[] {
  const known = new Set(readableColumns(table, identity).map((c) => c.name));
  for (const role of identity.roles) {
    for (const c of writableColumnsForRole(table, role)) known.add(c.name);
  }
  const unknown = Object.keys(body).filter((k) => !known.has(k));
  if (unknown.length > 0) {
    throw new McSheetyError('INVALID_REQUEST', 'The request contains unknown fields.', {
      details: unknown.map((field) => ({
        field,
        code: 'unknown_field',
        message: `Unknown field ${field}.`,
      })),
    });
  }
  return Object.keys(body).map((k) => table.columnsByName.get(k) as ColumnSchema);
}

function parseFields(
  fields: readonly ColumnSchema[],
  body: Record<string, unknown>,
  details: ErrorDetail[],
): Record<string, CellValue> {
  const parsed: Record<string, CellValue> = {};
  for (const column of fields) {
    const result = parseInput(column, body[column.name]);
    if (result.ok) parsed[column.name] = result.value;
    else {
      parsed[column.name] = null;
      details.push({
        field: column.name,
        code: 'invalid',
        message: `${column.name} ${result.reason}.`,
      });
    }
  }
  return parsed;
}

/** True if a single role grants the action, may write every field and passes every row check. */
function someRoleAuthorizes(
  table: TableSchema,
  identity: Identity,
  action: TableAction,
  fields: readonly ColumnSchema[],
  rows: readonly RecordValues[],
): boolean {
  const candidates = rows.reduce<string[]>(
    (roles, row) => roles.filter((r) => qualifyingRoles(table, identity, action, row).includes(r)),
    rolesGranting(table, identity, action),
  );
  return candidates.some((role) => {
    const writable = new Set(writableColumnsForRole(table, role).map((c) => c.name));
    return fields.every((f) => writable.has(f.name));
  });
}

function requireAction(table: TableSchema, identity: Identity, action: TableAction): void {
  if (rolesGranting(table, identity, action).length === 0) throw new McSheetyError('FORBIDDEN');
}

function checkRequired(
  columns: readonly ColumnSchema[],
  values: RecordValues,
  details: ErrorDetail[],
): void {
  const alreadyInvalid = new Set(details.map((d) => d.field));
  for (const column of columns) {
    if (alreadyInvalid.has(column.name)) continue;
    if (column.required && (values[column.name] ?? null) === null) {
      const message =
        column.default?.kind === 'claim'
          ? `${column.name} is required and your account has no trusted ${column.default.claim} attribute.`
          : `${column.name} is required.`;
      details.push({ field: column.name, code: 'required', message });
    }
  }
}

/**
 * Uniqueness within the documented consistency envelope: checked against the rows read for
 * this request. Concurrent writers can still race (implementation.md §9).
 */
function checkUnique(
  columns: readonly ColumnSchema[],
  values: RecordValues,
  table: DecodedTable,
  selfId?: string,
): void {
  for (const column of columns) {
    const value = values[column.name] ?? null;
    if (!column.unique || value === null) continue;
    const clash =
      (column.name === 'id' && table.duplicateIds.has(String(value))) ||
      table.records.some((r) => r.id !== selfId && r.values[column.name] === value);
    if (clash) {
      throw new McSheetyError('CONFLICT', `A record with the same ${column.name} already exists.`);
    }
  }
}

function throwIfInvalid(details: readonly ErrorDetail[]): void {
  if (details.length > 0) throw new McSheetyError('VALIDATION_FAILED', undefined, { details });
}

export function prepareCreate(
  table: TableSchema,
  identity: Identity,
  input: unknown,
  existing: DecodedTable,
  deps: MutationDeps,
): RecordValues {
  requireAction(table, identity, 'create');
  const body = asObject(input);
  const fields = resolveFields(table, identity, body);
  const details: ErrorDetail[] = [];
  const submitted = parseFields(fields, body, details);

  const candidate: Record<string, CellValue> = {};
  for (const column of table.columns) {
    if (column.name in submitted) {
      candidate[column.name] = submitted[column.name] ?? null;
      continue;
    }
    const def = column.default;
    candidate[column.name] =
      def === undefined
        ? null
        : def.kind === 'generated'
          ? deps.newId()
          : def.kind === 'now'
            ? formatDateTime(deps.now())
            : def.kind === 'claim'
              ? (identity.attributes[def.claim] ?? null)
              : def.value;
  }

  if (!someRoleAuthorizes(table, identity, 'create', fields, [candidate])) {
    throw new McSheetyError(
      'FORBIDDEN',
      'You cannot create this record or set one of these fields.',
    );
  }
  checkRequired(table.columns, candidate, details);
  throwIfInvalid(details);
  checkUnique(table.columns, candidate, existing);
  return candidate;
}

export function prepareUpdate(
  table: TableSchema,
  identity: Identity,
  existing: DecodedRecord,
  input: unknown,
  all: DecodedTable,
): PreparedUpdate {
  requireAction(table, identity, 'update');
  const body = asObject(input);
  const fields = resolveFields(table, identity, body);
  if (fields.length === 0) {
    throw new McSheetyError('INVALID_REQUEST', 'The update must change at least one field.');
  }
  const details: ErrorDetail[] = [];
  const patch = parseFields(fields, body, details);
  const next: RecordValues = { ...existing.values, ...patch };

  if (!someRoleAuthorizes(table, identity, 'update', fields, [existing.values, next])) {
    throw new McSheetyError('FORBIDDEN', 'You cannot make this change to this record.');
  }
  for (const column of fields) {
    if (existing.formulaColumns.has(column.name)) {
      details.push({
        field: column.name,
        code: 'formula_cell',
        message: `${column.name} is calculated by a formula in the workbook and cannot be changed.`,
      });
    }
  }
  checkRequired(fields, next, details);
  throwIfInvalid(details);
  checkUnique(fields, next, all, existing.id);

  const changes: Record<string, CellValue> = {};
  for (const column of fields) {
    const value = next[column.name] ?? null;
    if (value !== (existing.values[column.name] ?? null)) changes[column.name] = value;
  }
  return { changes, next };
}

export function authorizeDelete(
  table: TableSchema,
  identity: Identity,
  existing: DecodedRecord,
): void {
  requireAction(table, identity, 'delete');
  if (!someRoleAuthorizes(table, identity, 'delete', [], [existing.values])) {
    throw new McSheetyError('FORBIDDEN', 'You cannot delete this record.');
  }
}
