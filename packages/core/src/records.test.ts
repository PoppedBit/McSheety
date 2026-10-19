import { describe, expect, it } from 'vitest';
import { editRows, fixtureModel, rawData, table } from './__fixtures__/equipment.js';
import type { McSheetyError } from './errors.js';
import { canonicalRecord, decodeTable, findRecord } from './records.js';

const equipment = table(fixtureModel());

describe('decodeTable', () => {
  it('decodes fixture rows without issues', () => {
    const decoded = decodeTable(equipment, rawData());
    expect(decoded.records).toHaveLength(8);
    expect(decoded.records.every((r) => r.issues.length === 0)).toBe(true);
    expect(decoded.records[5]?.values['purchasePrice']).toBeNull();
  });

  it('excludes rows with duplicate or blank ids and reports them', () => {
    const raw = rawData('Equipment', {
      Equipment: editRows((row) => row[0] === 'eq-002' || row[0] === 'eq-003', { id: 'eq-dup' }),
    });
    const withBlank = {
      ...raw,
      rows: [...raw.rows, ['', 'Orphan', 'Active', 'loc-1', 'dept-ops', 1, '2026-10-16T00:00:00Z']],
    };
    const decoded = decodeTable(equipment, withBlank);
    expect(decoded.records.map((r) => r.id)).not.toContain('eq-dup');
    expect(decoded.problems).toEqual({ rowsWithoutId: [9], duplicateIds: ['eq-dup'] });
    expect(() => findRecord(decoded, 'eq-dup')).toThrow(
      expect.objectContaining({ code: 'CONFLICT' }) as McSheetyError,
    );
    expect(() => findRecord(decoded, 'eq-404')).toThrow(
      expect.objectContaining({ code: 'NOT_FOUND' }) as McSheetyError,
    );
  });

  it('surfaces invalid cells as field issues with null values', () => {
    const raw = rawData('Equipment', {
      Equipment: editRows((row) => row[0] === 'eq-001', {
        status: 'Broken',
        purchasePrice: 'lots',
      }),
    });
    const rec = findRecord(decodeTable(equipment, raw), 'eq-001');
    expect(rec.values['status']).toBeNull();
    expect(rec.issues.map((i) => i.field)).toEqual(['status', 'purchasePrice']);
  });

  it('records the read position separately from identity', () => {
    const raw = rawData();
    const reversed = { ...raw, rows: [...raw.rows].reverse() };
    expect(findRecord(decodeTable(equipment, reversed), 'eq-001').index).toBe(7);
  });
});

describe('canonicalRecord', () => {
  it('is independent of key order', () => {
    expect(canonicalRecord({ b: 1, a: 'x' })).toBe(canonicalRecord({ a: 'x', b: 1 }));
    expect(canonicalRecord({ a: null })).not.toBe(canonicalRecord({ a: '' }));
  });
});
