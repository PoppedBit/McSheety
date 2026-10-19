import type { CellValue, ColumnTypes, WorkbookContext } from '@mcsheety/core';
import { McSheetyError } from '@mcsheety/core';
import { InMemoryWorkbookAdapter } from '@mcsheety/core/testing';
import { beforeEach, describe, expect, it } from 'vitest';
import { memoryWorkbook, scopeFor, users } from './__fixtures__/memory.js';
import type { AuditEvent } from './service.js';
import { McSheetyService, computeRevision } from './service.js';

const NEW_ID = '0b9f3c2e-1d4a-4e5f-8a6b-7c8d9e0f1a2b';
let adapter: InMemoryWorkbookAdapter;
let audits: AuditEvent[];
let service: McSheetyService;

function makeService(a: InMemoryWorkbookAdapter) {
  return new McSheetyService({
    adapter: a,
    now: () => new Date('2026-10-17T10:00:00Z'),
    newId: () => NEW_ID,
    audit: (e) => audits.push(e),
  });
}

beforeEach(() => {
  adapter = memoryWorkbook();
  audits = [];
  service = makeService(adapter);
});

const mutations = () => adapter.calls.filter((c) => /^(appendRow|updateCells|deleteRow)/.test(c));
const equipmentRow = (id: string) => {
  const t = adapter.tables.get('tbl_Equipment');
  const row = t?.rows.find((r) => r[0] === id);
  return t && row && Object.fromEntries(t.columns.map((c, i) => [c, row[i]]));
};

describe('reads', () => {
  it('lists only readable tables', async () => {
    expect((await service.listTables(scopeFor(users.viewer))).data).toEqual([
      { name: 'Equipment' },
      { name: 'Locations' },
    ]);
    expect((await service.listTables(scopeFor(users.nobody))).data).toEqual([]);
  });

  it('returns only readable schema columns with edit flags', async () => {
    const { columns } = await service.schema(scopeFor(users.editorOps), 'Equipment');
    expect(columns.map((c) => c.name)).not.toContain('purchasePrice');
    const byName = Object.fromEntries(columns.map((c) => [c.name, c]));
    expect(byName['name']).toMatchObject({ creatable: true, updatable: true });
    expect(byName['departmentId']).toMatchObject({ creatable: false, default: 'server' });
    expect(byName['status']).toMatchObject({
      enumValues: ['Active', 'Inactive'],
      default: { value: 'Active' },
    });
  });

  it('lists the editor’s rows with redacted columns and revisions', async () => {
    const result = await service.listRows(scopeFor(users.editorOps), 'Equipment', {});
    expect(result.data.map((r) => r.id)).toEqual(['eq-001', 'eq-003', 'eq-005']);
    for (const r of result.data) {
      expect(r.fields).not.toHaveProperty('purchasePrice');
      expect(r.revision).toMatch(/^[\w-]{27}$/);
    }
  });

  it('answers NOT_FOUND for invisible, missing and inaccessible resources alike', async () => {
    const code = (p: Promise<unknown>) =>
      p.then(
        () => 'ok',
        (e: McSheetyError) => e.code,
      );
    expect(await code(service.getRow(scopeFor(users.editorOps), 'Equipment', 'eq-002'))).toBe(
      'NOT_FOUND',
    );
    expect(await code(service.getRow(scopeFor(users.editorOps), 'Equipment', 'eq-999'))).toBe(
      'NOT_FOUND',
    );
    expect(await code(service.getRow(scopeFor(users.viewer), 'Payroll', 'x'))).toBe('NOT_FOUND');
    expect(await code(service.listRows(scopeFor(users.nobody), 'Equipment', {}))).toBe('NOT_FOUND');
  });

  it('surfaces cell problems only for columns the caller can read', async () => {
    const t = adapter.tables.get('tbl_Equipment');
    (t?.rows[0] as CellValue[])[5] = 'lots'; // purchasePrice manually broken in Excel
    expect(
      (await service.getRow(scopeFor(users.viewer), 'Equipment', 'eq-001')).issues,
    ).toBeUndefined();
    expect((await service.getRow(scopeFor(users.manager), 'Equipment', 'eq-001')).issues).toEqual([
      { field: 'purchasePrice', message: 'purchasePrice must be a number.' },
    ]);
  });

  it('revisions do not reveal changes to hidden columns', async () => {
    const before = await service.getRow(scopeFor(users.viewer), 'Equipment', 'eq-001');
    const mgrBefore = await service.getRow(scopeFor(users.manager), 'Equipment', 'eq-001');
    (adapter.tables.get('tbl_Equipment')?.rows[0] as CellValue[])[5] = 999;
    expect((await service.getRow(scopeFor(users.viewer), 'Equipment', 'eq-001')).revision).toBe(
      before.revision,
    );
    expect(
      (await service.getRow(scopeFor(users.manager), 'Equipment', 'eq-001')).revision,
    ).not.toBe(mgrBefore.revision);
  });

  it('hides tables whose metadata is broken (fail closed)', async () => {
    const rp = adapter.tables.get('mc_rowPolicies');
    (rp?.rows[0] as CellValue[])[4] = 'regex';
    await expect(service.listRows(scopeFor(users.viewer), 'Equipment', {})).rejects.toMatchObject({
      code: 'NOT_FOUND',
    });
  });
});

describe('createRow', () => {
  it('creates a row with server-owned values inside a session and audits field names', async () => {
    const record = await service.createRow(scopeFor(users.editorOps), 'Equipment', {
      name: 'Pump Z',
      locationId: 'loc-2',
    });
    expect(record.fields).toMatchObject({
      id: NEW_ID,
      departmentId: 'dept-ops',
      status: 'Active',
      createdAt: '2026-10-17T10:00:00Z',
    });
    expect(equipmentRow(NEW_ID)).toMatchObject({ name: 'Pump Z', purchasePrice: '' });
    expect(adapter.calls[0]).toMatch(/^openSession/);
    expect(adapter.calls.at(-1)).toMatch(/^closeSession/);
    expect(audits).toEqual([
      expect.objectContaining({
        action: 'create',
        recordId: NEW_ID,
        outcome: 'success',
        changedFields: ['name', 'locationId'],
      }),
    ]);
  });

  it('denied creates never reach a workbook mutation', async () => {
    await expect(
      service.createRow(scopeFor(users.viewer), 'Equipment', { name: 'X', locationId: 'loc-1' }),
    ).rejects.toMatchObject({ code: 'FORBIDDEN' });
    expect(mutations()).toEqual([]);
    expect(audits[0]).toMatchObject({ outcome: 'denied', errorCode: 'FORBIDDEN' });
  });

  it('treats an append that landed despite an error as success (no blind retry)', async () => {
    class FlakyAfterWrite extends InMemoryWorkbookAdapter {
      override async appendRow(
        ctx: WorkbookContext,
        table: string,
        row: readonly CellValue[],
        types: ColumnTypes,
      ) {
        await super.appendRow(ctx, table, row, types);
        throw new McSheetyError('WORKBOOK_UNAVAILABLE');
      }
    }
    const flaky = new FlakyAfterWrite([...adapter.tables.values()]);
    const record = await makeService(flaky).createRow(scopeFor(users.editorOps), 'Equipment', {
      name: 'P',
      locationId: 'loc-1',
    });
    expect(record.id).toBe(NEW_ID);
    expect(flaky.calls.filter((c) => c.startsWith('appendRow'))).toHaveLength(1);
    expect(flaky.tables.get('tbl_Equipment')?.rows.filter((r) => r[0] === NEW_ID)).toHaveLength(1);
  });

  it('propagates an append failure that did not land', async () => {
    adapter.hooks.beforeWrite = () => {
      throw new McSheetyError('WORKBOOK_UNAVAILABLE');
    };
    await expect(
      service.createRow(scopeFor(users.editorOps), 'Equipment', { name: 'P', locationId: 'loc-1' }),
    ).rejects.toMatchObject({ code: 'WORKBOOK_UNAVAILABLE' });
  });
});

describe('updateRow', () => {
  it('requires If-Match before touching the workbook', async () => {
    await expect(
      service.updateRow(
        scopeFor(users.editorOps),
        'Equipment',
        'eq-001',
        { status: 'Inactive' },
        undefined,
      ),
    ).rejects.toMatchObject({ code: 'PRECONDITION_REQUIRED' });
    expect(adapter.calls).toEqual([]);
  });

  it('updates with a current revision and returns a new one', async () => {
    const scope = scopeFor(users.editorOps);
    const before = await service.getRow(scope, 'Equipment', 'eq-001');
    const after = await service.updateRow(
      scope,
      'Equipment',
      'eq-001',
      { status: 'Inactive' },
      `"${before.revision}"`,
    );
    expect(after.fields['status']).toBe('Inactive');
    expect(after.revision).not.toBe(before.revision);
    expect(equipmentRow('eq-001')?.['status']).toBe('Inactive');
    expect(audits.at(-1)).toMatchObject({
      action: 'update',
      changedFields: ['status'],
      outcome: 'success',
    });
  });

  it('rejects stale revisions', async () => {
    await expect(
      service.updateRow(
        scopeFor(users.editorOps),
        'Equipment',
        'eq-001',
        { status: 'Inactive' },
        '"stale"',
      ),
    ).rejects.toMatchObject({ code: 'CONFLICT' });
    expect(mutations()).toEqual([]);
  });

  it('detects a concurrent edit between the read and the write', async () => {
    const scope = scopeFor(users.editorOps);
    const { revision } = await service.getRow(scope, 'Equipment', 'eq-001');
    adapter.hooks.beforeWrite = () => {
      (adapter.tables.get('tbl_Equipment')?.rows[0] as CellValue[])[1] = 'Edited in Excel';
    };
    await expect(
      service.updateRow(scope, 'Equipment', 'eq-001', { status: 'Inactive' }, revision),
    ).rejects.toMatchObject({
      code: 'CONFLICT',
    });
    expect(equipmentRow('eq-001')?.['status']).toBe('Active');
  });

  it('finds records by id even after rows are re-sorted in Excel', async () => {
    const scope = scopeFor(users.editorOps);
    const { revision } = await service.getRow(scope, 'Equipment', 'eq-005');
    adapter.tables.get('tbl_Equipment')?.rows.reverse();
    await service.updateRow(scope, 'Equipment', 'eq-005', { name: 'Forklift E2' }, revision);
    expect(equipmentRow('eq-005')?.['name']).toBe('Forklift E2');
    expect(equipmentRow('eq-004')?.['name']).toBe('Generator D');
  });

  it('cannot update rows outside the caller’s scope (NOT_FOUND, not FORBIDDEN)', async () => {
    const rev = computeRevision({});
    await expect(
      service.updateRow(
        scopeFor(users.editorOps),
        'Equipment',
        'eq-002',
        { status: 'Active' },
        rev,
      ),
    ).rejects.toMatchObject({ code: 'NOT_FOUND' });
  });
});

describe('deleteRow', () => {
  it('lets admin delete and forbids manager without reaching Graph mutations', async () => {
    const admin = scopeFor(users.admin);
    const { revision } = await service.getRow(admin, 'Equipment', 'eq-002');

    const manager = scopeFor(users.manager);
    const mgrRev = (await service.getRow(manager, 'Equipment', 'eq-002')).revision;
    await expect(service.deleteRow(manager, 'Equipment', 'eq-002', mgrRev)).rejects.toMatchObject({
      code: 'FORBIDDEN',
    });
    expect(mutations()).toEqual([]);

    await service.deleteRow(admin, 'Equipment', 'eq-002', revision);
    expect(equipmentRow('eq-002')).toBeUndefined();
    expect(audits.at(-1)).toMatchObject({
      action: 'delete',
      recordId: 'eq-002',
      outcome: 'success',
    });
  });
});
