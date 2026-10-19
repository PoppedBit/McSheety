import { describe, expect, it } from 'vitest';
import { fixtureModel, rawData, table, users } from './__fixtures__/equipment.js';
import { capabilities, projectRecord, readableColumns } from './policy.js';
import { decodeTable } from './records.js';
import type { Identity } from './types.js';

const model = fixtureModel();
const equipment = table(model);
const records = decodeTable(equipment, rawData()).records;

const visible = (who: Identity) =>
  records.map((r) => projectRecord(equipment, who, r.values)).filter((r) => r !== undefined);

describe('read authorization', () => {
  it('viewer sees every row but not purchasePrice', () => {
    const rows = visible(users.viewer);
    expect(rows).toHaveLength(8);
    for (const r of rows) expect(r).not.toHaveProperty('purchasePrice');
    expect(Object.keys(rows[0] ?? {})).toEqual([
      'id',
      'name',
      'status',
      'locationId',
      'departmentId',
      'createdAt',
    ]);
  });

  it('editor sees only rows in their trusted department', () => {
    const rows = visible(users.editorOps);
    expect(rows.map((r) => r['id'])).toEqual(['eq-001', 'eq-003', 'eq-005']);
  });

  it('editor without a trusted department attribute sees nothing (fail closed)', () => {
    expect(visible(users.editorNoDept)).toEqual([]);
  });

  it('manager sees purchasePrice on all rows', () => {
    const rows = visible(users.manager);
    expect(rows).toHaveLength(8);
    expect(rows[0]).toHaveProperty('purchasePrice', 123.5);
  });

  it('unknown or wrongly-cased roles grant nothing', () => {
    expect(visible(users.spoofed)).toEqual([]);
    expect(visible(users.nobody)).toEqual([]);
    expect(readableColumns(equipment, users.spoofed)).toEqual([]);
  });

  it('a second unrestricted role widens row access (any role may grant)', () => {
    const both: Identity = { ...users.editorOps, roles: ['editor', 'viewer'] };
    expect(visible(both)).toHaveLength(8);
  });
});

describe('capabilities', () => {
  it('describes only the caller’s own effective rights', () => {
    const [eq, loc] = capabilities(model, users.editorOps);
    expect(eq).toEqual({
      table: 'Equipment',
      actions: ['read', 'create', 'update'],
      readableColumns: ['id', 'name', 'status', 'locationId', 'departmentId', 'createdAt'],
      creatableColumns: ['name', 'status', 'locationId'],
      updatableColumns: ['name', 'status', 'locationId'],
      rowScope: { read: 'restricted', create: 'restricted', update: 'restricted' },
    });
    expect(loc?.actions).toEqual(['read']);
  });

  it('lists no tables for callers without roles', () => {
    expect(capabilities(model, users.nobody)).toEqual([]);
  });

  it('marks admin delete as unrestricted', () => {
    expect(capabilities(model, users.admin)[0]?.rowScope.delete).toBe('all');
  });
});
