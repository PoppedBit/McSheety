import { describe, expect, it } from 'vitest';
import { cell, editRows, fixtureModel, rawData, table, users } from './__fixtures__/equipment.js';
import { McSheetyError } from './errors.js';
import type { MutationDeps } from './mutations.js';
import { authorizeDelete, prepareCreate, prepareUpdate } from './mutations.js';
import { decodeTable, findRecord } from './records.js';
import type { Identity, RawTable } from './types.js';

const deps: MutationDeps = {
  now: () => new Date('2026-10-17T09:30:00.000Z'),
  newId: () => '6f1c2f9e-6d0c-4a59-9b8e-1f2a3b4c5d6e',
};

const model = fixtureModel();
const equipment = table(model);
const data = decodeTable(equipment, rawData());

function codeOf(fn: () => unknown): string | undefined {
  try {
    fn();
  } catch (e) {
    if (e instanceof McSheetyError) return e.code;
    throw e;
  }
  return undefined;
}

describe('prepareCreate', () => {
  it('fills server-owned fields and trusted defaults for an editor', () => {
    const row = prepareCreate(
      equipment,
      users.editorOps,
      { name: 'Pump Z', locationId: 'loc-2' },
      data,
      deps,
    );
    expect(row).toEqual({
      id: deps.newId(),
      name: 'Pump Z',
      status: 'Active',
      locationId: 'loc-2',
      departmentId: 'dept-ops',
      purchasePrice: null,
      createdAt: '2026-10-17T09:30:00Z',
    });
  });

  it('forbids an editor from setting departmentId (not writable)', () => {
    const body = { name: 'X', locationId: 'loc-1', departmentId: 'dept-eng' };
    expect(codeOf(() => prepareCreate(equipment, users.editorOps, body, data, deps))).toBe(
      'FORBIDDEN',
    );
  });

  it('treats hidden columns like unknown fields so they cannot be probed', () => {
    const body = { name: 'X', locationId: 'loc-1', purchasePrice: 1 };
    expect(codeOf(() => prepareCreate(equipment, users.editorOps, body, data, deps))).toBe(
      'INVALID_REQUEST',
    );
    expect(codeOf(() => prepareCreate(equipment, users.editorOps, { nope: 1 }, data, deps))).toBe(
      'INVALID_REQUEST',
    );
  });

  it('forbids create without the table action or without a trusted attribute', () => {
    const body = { name: 'X', locationId: 'loc-1' };
    expect(codeOf(() => prepareCreate(equipment, users.viewer, body, data, deps))).toBe(
      'FORBIDDEN',
    );
    expect(codeOf(() => prepareCreate(equipment, users.editorNoDept, body, data, deps))).toBe(
      'FORBIDDEN',
    );
    expect(codeOf(() => prepareCreate(equipment, users.spoofed, body, data, deps))).toBe(
      'FORBIDDEN',
    );
  });

  it('lets a manager create in any department and validates required fields', () => {
    const row = prepareCreate(
      equipment,
      users.manager,
      { name: 'Hoist', locationId: 'loc-1', departmentId: 'dept-ops', status: 'Inactive' },
      data,
      deps,
    );
    expect(row['departmentId']).toBe('dept-ops');
    expect(
      codeOf(() => prepareCreate(equipment, users.manager, { locationId: 'loc-1' }, data, deps)),
    ).toBe('VALIDATION_FAILED');
  });

  it('rejects formula injection and bad enum values with field details', () => {
    try {
      prepareCreate(
        equipment,
        users.editorOps,
        { name: '=HYPERLINK("x")', locationId: 'loc-1', status: 'Gone' },
        data,
        deps,
      );
      expect.unreachable();
    } catch (e) {
      expect(e).toBeInstanceOf(McSheetyError);
      expect((e as McSheetyError).code).toBe('VALIDATION_FAILED');
      expect((e as McSheetyError).details?.map((d) => d.field)).toEqual(['name', 'status']);
    }
  });

  it('enforces uniqueness against existing rows', () => {
    const locations = table(model, 'Locations');
    const locData = decodeTable(locations, rawData('Locations'));
    expect(
      codeOf(() => prepareCreate(locations, users.manager, { name: 'Head Office' }, locData, deps)),
    ).toBe('CONFLICT');
  });

  it('rejects non-object bodies', () => {
    expect(codeOf(() => prepareCreate(equipment, users.manager, [1], data, deps))).toBe(
      'INVALID_REQUEST',
    );
  });
});

describe('prepareUpdate', () => {
  const eq001 = findRecord(data, 'eq-001');
  const eq002 = findRecord(data, 'eq-002');

  it('returns only changed columns for an authorized editor', () => {
    const { changes, next } = prepareUpdate(
      equipment,
      users.editorOps,
      eq001,
      { status: 'Inactive', name: 'Pump A' },
      data,
    );
    expect(changes).toEqual({ status: 'Inactive' });
    expect(next['status']).toBe('Inactive');
  });

  it('forbids editing a row outside the editor’s department', () => {
    expect(
      codeOf(() => prepareUpdate(equipment, users.editorOps, eq002, { status: 'Active' }, data)),
    ).toBe('FORBIDDEN');
  });

  it('never allows id changes and requires at least one field', () => {
    expect(codeOf(() => prepareUpdate(equipment, users.admin, eq001, { id: 'eq-999' }, data))).toBe(
      'FORBIDDEN',
    );
    expect(codeOf(() => prepareUpdate(equipment, users.admin, eq001, {}, data))).toBe(
      'INVALID_REQUEST',
    );
  });

  it('checks both the existing and the resulting row', () => {
    // Let editors write departmentId so that only the row rule can stop the move.
    const m = fixtureModel({
      _columns: editRows((r, c) => cell(r, c, 'column') === 'departmentId', {
        writeRoles: 'editor,manager,admin',
      }),
    });
    const t = table(m);
    const d = decodeTable(t, rawData());
    const own = findRecord(d, 'eq-001');
    expect(
      codeOf(() => prepareUpdate(t, users.editorOps, own, { departmentId: 'dept-eng' }, d)),
    ).toBe('FORBIDDEN');
    expect(
      prepareUpdate(t, users.editorOps, own, { departmentId: 'dept-ops', name: 'P' }, d).changes,
    ).toEqual({
      name: 'P',
    });
  });

  it('does not combine one role’s rows with another role’s fields', () => {
    // viewer: may update any row, but only `status`. editor: may write `name`, but only own department.
    const m = fixtureModel({
      _permissions: editRows((r) => r[0] === 'Equipment' && r[1] === 'viewer', { update: true }),
      _columns: editRows((r, c) => cell(r, c, 'column') === 'status', {
        writeRoles: 'viewer,editor,manager,admin',
      }),
    });
    const t = table(m);
    const d = decodeTable(t, rawData());
    const both: Identity = { ...users.editorOps, roles: ['viewer', 'editor'] };
    const foreign = findRecord(d, 'eq-002');
    expect(codeOf(() => prepareUpdate(t, both, foreign, { name: 'Renamed' }, d))).toBe('FORBIDDEN');
    expect(prepareUpdate(t, both, foreign, { status: 'Active' }, d).changes).toEqual({
      status: 'Active',
    });
  });

  it('rejects writes to formula-bearing cells', () => {
    const raw: RawTable = {
      ...rawData(),
      formulas: rawData().rows.map((row, i) =>
        row.map((v, j) => (i === 0 && j === 1 ? '=UPPER("pump a")' : v)),
      ),
    };
    const d = decodeTable(equipment, raw);
    expect(
      codeOf(() =>
        prepareUpdate(equipment, users.manager, findRecord(d, 'eq-001'), { name: 'B' }, d),
      ),
    ).toBe('VALIDATION_FAILED');
  });

  it('rejects clearing a required field', () => {
    expect(codeOf(() => prepareUpdate(equipment, users.manager, eq001, { name: null }, data))).toBe(
      'VALIDATION_FAILED',
    );
  });
});

describe('authorizeDelete', () => {
  it('allows admin, forbids manager (no delete permission)', () => {
    const rec = findRecord(data, 'eq-001');
    expect(() => authorizeDelete(equipment, users.admin, rec)).not.toThrow();
    expect(codeOf(() => authorizeDelete(equipment, users.manager, rec))).toBe('FORBIDDEN');
  });
});
