import { describe, expect, it } from 'vitest';
import type { SheetEdit } from './__fixtures__/equipment.js';
import { cell, editRows, fixtureSheets, sourceFrom } from './__fixtures__/equipment.js';
import { parseWorkbookModel } from './metadata.js';

const parse = (edits: Record<string, SheetEdit> = {}) =>
  parseWorkbookModel(sourceFrom(fixtureSheets(edits)));
const errors = (r: ReturnType<typeof parse>) => r.issues.filter((i) => i.severity === 'error');

const rowWhere =
  (table: string, column: string) => (row: readonly unknown[], cols: readonly string[]) =>
    cell(row as never, cols, 'table') === table && cell(row as never, cols, 'column') === column;

describe('parseWorkbookModel with the Equipment fixture', () => {
  const { model, issues } = parse();

  it('parses cleanly', () => {
    expect(issues).toEqual([]);
    expect([...model.tables.keys()]).toEqual(['Equipment', 'Locations']);
    expect(model.formatVersion).toBe(1);
    expect(model.applicationId).toBe('mcsheety-equipment-fixture');
    expect([...model.roles].sort()).toEqual(['admin', 'editor', 'manager', 'viewer']);
  });

  it('builds column schemas in Excel header order with policies and defaults', () => {
    const eq = model.tables.get('Equipment');
    expect(eq?.excelTable).toBe('tbl_Equipment');
    expect(eq?.columns.map((c) => c.name)).toEqual([
      'id',
      'name',
      'status',
      'locationId',
      'departmentId',
      'purchasePrice',
      'createdAt',
    ]);
    const status = eq?.columnsByName.get('status');
    expect(status?.enumValues).toEqual(['Active', 'Inactive']);
    expect(status?.default).toEqual({ kind: 'literal', value: 'Active' });
    expect(eq?.columnsByName.get('departmentId')?.default).toEqual({
      kind: 'claim',
      claim: 'departmentId',
    });
    expect(eq?.columnsByName.get('purchasePrice')?.min).toBe(0);
    expect(eq?.columnsByName.get('purchasePrice')?.readRoles.roles).toEqual(
      new Set(['manager', 'admin']),
    );
    expect(eq?.columnsByName.get('name')?.readRoles.any).toBe(true);
    expect(eq?.rowPolicies).toHaveLength(4);
  });
});

describe('parseWorkbookModel fails closed', () => {
  it('disables only the affected table for an invalid row policy operator', () => {
    const result = parse({
      _rowPolicies: editRows((row, cols) => cell(row, cols, 'action') === 'read', {
        operator: 'contains',
      }),
    });
    expect(result.model.tables.has('Equipment')).toBe(false);
    expect(result.model.tables.has('Locations')).toBe(true);
    expect(errors(result).map((e) => e.message)).toContain(
      'Unsupported operator "contains" (supported: equals, allowAll).',
    );
  });

  it('returns no tables when a metadata table header is malformed', () => {
    const result = parse({
      _columns: (s) => ({ ...s, columns: ['table', 'column', 'readers', 'writeRoles'] }),
    });
    expect(result.model.tables.size).toBe(0);
  });

  it('returns no tables for an unsupported formatVersion', () => {
    const result = parse({
      _mcsheety: editRows((row) => row[0] === 'formatVersion', { value: '2' }),
    });
    expect(result.model.tables.size).toBe(0);
  });

  it('returns no tables when a required metadata table is missing', () => {
    const sheets = fixtureSheets();
    sheets.delete('mc_permissions');
    expect(parseWorkbookModel(sourceFrom(sheets)).model.tables.size).toBe(0);
  });

  it('denies a column with no column policy and warns', () => {
    const result = parse({
      _columns: (s) => ({
        ...s,
        rows: s.rows.filter((r) => !rowWhere('Equipment', 'status')(r, s.columns)),
      }),
    });
    const status = result.model.tables.get('Equipment')?.columnsByName.get('status');
    expect(status?.readRoles).toEqual({ any: false, roles: new Set() });
    expect(status?.writeRoles).toEqual({ any: false, roles: new Set() });
    expect(result.issues).toContainEqual(
      expect.objectContaining({ severity: 'warning', table: 'Equipment' }),
    );
  });

  it('disables a table whose schema column is missing from the Excel headers', () => {
    const result = parse({
      Equipment: (s) => ({ ...s, columns: s.columns.map((c) => (c === 'status' ? 'state' : c)) }),
    });
    expect(result.model.tables.has('Equipment')).toBe(false);
  });

  it('rejects non-boolean permission cells', () => {
    const result = parse({
      _permissions: editRows((row) => row[0] === 'Equipment' && row[1] === 'viewer', {
        read: 'yes',
      }),
    });
    expect(result.model.tables.has('Equipment')).toBe(false);
  });

  it('requires id to be server-owned', () => {
    const result = parse({
      _columns: editRows(rowWhere('Equipment', 'id'), { writeRoles: 'admin' }),
    });
    expect(result.model.tables.has('Equipment')).toBe(false);
  });

  it('rejects "*" mixed with role names', () => {
    const result = parse({
      _columns: editRows(rowWhere('Equipment', 'name'), { readRoles: '*,admin' }),
    });
    expect(result.model.tables.has('Equipment')).toBe(false);
  });

  it('rejects unknown types and invalid defaults', () => {
    expect(
      parse({
        _schema: editRows(rowWhere('Equipment', 'name'), { type: 'text' }),
      }).model.tables.has('Equipment'),
    ).toBe(false);
    expect(
      parse({
        _schema: editRows(rowWhere('Equipment', 'status'), { default: 'Retired' }),
      }).model.tables.has('Equipment'),
    ).toBe(false);
  });

  it('rejects row policies that reference unknown columns', () => {
    const result = parse({
      _rowPolicies: editRows((row, cols) => cell(row, cols, 'action') === 'read', {
        column: 'dept',
      }),
    });
    expect(result.model.tables.has('Equipment')).toBe(false);
  });

  it('accepts textual TRUE/FALSE in metadata booleans', () => {
    const result = parse({
      _permissions: editRows((row) => row[0] === 'Equipment' && row[1] === 'viewer', {
        read: 'TRUE',
      }),
    });
    expect(result.model.tables.get('Equipment')?.permissions.get('viewer')?.has('read')).toBe(true);
  });
});
