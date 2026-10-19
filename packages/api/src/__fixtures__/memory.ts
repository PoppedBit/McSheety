/** Test helpers: an in-memory workbook seeded from fixtures/equipment. */
import type { Identity, WorkbookContext } from '@mcsheety/core';
import { InMemoryWorkbookAdapter } from '@mcsheety/core/testing';
import { equipmentWorkbook } from '../../../../fixtures/equipment/index.js';
import type { RequestScope } from '../service.js';

export function memoryWorkbook(): InMemoryWorkbookAdapter {
  return new InMemoryWorkbookAdapter(
    equipmentWorkbook.sheets.map((s) => ({
      sheet: s.sheet,
      table: s.table,
      columns: [...s.columns],
      rows: s.rows.map((r) => r.map((v) => v ?? '')),
    })),
  );
}

const base = { tenantId: 'tenant-1', attributes: {} };
export const users = {
  viewer: { ...base, userId: 'u-viewer', roles: ['viewer'] },
  editorOps: {
    ...base,
    userId: 'u-ed-ops',
    roles: ['editor'],
    attributes: { departmentId: 'dept-ops' },
  },
  manager: { ...base, userId: 'u-mgr', roles: ['manager'] },
  admin: { ...base, userId: 'u-admin', roles: ['admin'] },
  nobody: { ...base, userId: 'u-none', roles: [] },
} satisfies Record<string, Identity>;

export function scopeFor(identity: Identity): RequestScope {
  const ctx: WorkbookContext = {
    correlationId: `corr-${identity.userId}`,
    getGraphToken: () => Promise.resolve('graph'),
  };
  return { identity, ctx };
}
