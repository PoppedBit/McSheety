import { describe, expect, it } from 'vitest';
import { fixtureModel, rawData, table, users } from './__fixtures__/equipment.js';
import { McSheetyError } from './errors.js';
import { projectRecord } from './policy.js';
import { applyListQuery, parseListQuery } from './query.js';
import { decodeTable } from './records.js';
import type { Identity, RecordValues } from './types.js';

const equipment = table(fixtureModel());
const records = decodeTable(equipment, rawData()).records;
const visible = (who: Identity): RecordValues[] =>
  records.map((r) => projectRecord(equipment, who, r.values)).filter((r) => r !== undefined);

const ids = (items: readonly RecordValues[]) => items.map((r) => r['id']);

describe('list queries', () => {
  it('filters with eq, ne and in', () => {
    const q = (where: unknown) =>
      ids(
        applyListQuery(
          visible(users.viewer),
          parseListQuery(equipment, users.viewer, { where: JSON.stringify(where) }),
        ).items,
      );
    expect(q({ status: 'Inactive' })).toEqual(['eq-002', 'eq-006']);
    expect(q({ departmentId: { ne: 'dept-ops' }, status: 'Active' })).toEqual([
      'eq-004',
      'eq-007',
      'eq-008',
    ]);
    expect(q({ locationId: { in: ['loc-2', 'loc-3'] } })).toEqual([
      'eq-002',
      'eq-004',
      'eq-005',
      'eq-006',
      'eq-008',
    ]);
  });

  it('sorts with nulls last and ties broken by id', () => {
    const query = parseListQuery(equipment, users.manager, { sort: '-purchasePrice' });
    const items = applyListQuery(visible(users.manager), query).items;
    expect(ids(items).slice(0, 2)).toEqual(['eq-005', 'eq-004']);
    expect(ids(items).at(-1)).toBe('eq-006');
  });

  it('pages with opaque cursors bound to the query', () => {
    const first = parseListQuery(equipment, users.viewer, { limit: '3', sort: 'name' });
    const page1 = applyListQuery(visible(users.viewer), first);
    expect(page1.items).toHaveLength(3);
    expect(page1.nextCursor).toEqual(expect.any(String));

    const second = parseListQuery(equipment, users.viewer, {
      limit: '3',
      sort: 'name',
      cursor: page1.nextCursor ?? '',
    });
    const page2 = applyListQuery(visible(users.viewer), second);
    expect(ids(page2.items)).not.toContain(page1.items[0]?.['id']);

    expect(() =>
      parseListQuery(equipment, users.viewer, { sort: '-name', cursor: page1.nextCursor ?? '' }),
    ).toThrow(McSheetyError);
    expect(() => parseListQuery(equipment, users.viewer, { cursor: 'garbage' })).toThrow(
      McSheetyError,
    );
  });

  it('ends paging with a null cursor', () => {
    const query = parseListQuery(equipment, users.viewer, { limit: '100' });
    expect(applyListQuery(visible(users.viewer), query).nextCursor).toBeNull();
  });

  it('rejects filtering or sorting on columns the caller cannot read', () => {
    expect(() =>
      parseListQuery(equipment, users.viewer, {
        where: JSON.stringify({ purchasePrice: { ne: null } }),
      }),
    ).toThrow('Unknown field purchasePrice.');
    expect(() => parseListQuery(equipment, users.viewer, { sort: 'purchasePrice' })).toThrow(
      McSheetyError,
    );
  });

  it('validates filter values, operators and limits', () => {
    const bad = (params: Record<string, string>) => () =>
      parseListQuery(equipment, users.manager, params);
    expect(bad({ where: '{"purchasePrice":"12"}' })).toThrow(McSheetyError);
    expect(bad({ where: '{"status":{"like":"A%"}}' })).toThrow(McSheetyError);
    expect(bad({ where: '[1]' })).toThrow(McSheetyError);
    expect(bad({ limit: '101' })).toThrow(McSheetyError);
    expect(bad({ limit: '0' })).toThrow(McSheetyError);
  });
});
