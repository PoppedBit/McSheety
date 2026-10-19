import { describe, expect, it } from 'vitest';
import type { ColumnSchema, FieldType } from './types.js';
import { decodeCell, parseInput } from './values.js';

const col = (type: FieldType, extra: Partial<ColumnSchema> = {}): ColumnSchema => ({
  name: 'f',
  type,
  required: false,
  unique: false,
  default: undefined,
  enumValues: undefined,
  min: undefined,
  max: undefined,
  maxLength: undefined,
  readRoles: { any: true, roles: new Set() },
  writeRoles: { any: true, roles: new Set() },
  ...extra,
});

describe('decodeCell (Excel -> API)', () => {
  it('treats blank cells as null for every type', () => {
    for (const t of ['string', 'number', 'boolean', 'date', 'time', 'datetime', 'uuid'] as const) {
      expect(decodeCell(col(t), '')).toEqual({ ok: true, value: null });
    }
  });

  it('converts Excel date and time serials explicitly', () => {
    expect(decodeCell(col('date'), 46311)).toEqual({ ok: true, value: '2026-10-16' });
    expect(decodeCell(col('time'), 0.5)).toEqual({ ok: true, value: '12:00:00' });
  });

  it('refuses to guess a timezone for datetime serials or zone-less text', () => {
    expect(decodeCell(col('datetime'), 46311.5).ok).toBe(false);
    expect(decodeCell(col('datetime'), '2026-10-16T14:05:00').ok).toBe(false);
  });

  it('normalizes datetimes with offsets to UTC', () => {
    expect(decodeCell(col('datetime'), '2026-10-16T16:05:00+02:00')).toEqual({
      ok: true,
      value: '2026-10-16T14:05:00Z',
    });
  });

  it('rejects impossible dates', () => {
    expect(decodeCell(col('date'), '2026-02-30').ok).toBe(false);
    expect(decodeCell(col('datetime'), '2026-02-30T00:00:00Z').ok).toBe(false);
  });

  it('accepts numeric text in number columns and numbers in string columns', () => {
    expect(decodeCell(col('number'), ' 12.5 ')).toEqual({ ok: true, value: 12.5 });
    expect(decodeCell(col('number'), 'twelve').ok).toBe(false);
    expect(decodeCell(col('string'), 42)).toEqual({ ok: true, value: '42' });
  });

  it('flags values outside the enum', () => {
    expect(decodeCell(col('enum', { enumValues: ['A', 'B'] }), 'C').ok).toBe(false);
  });
});

describe('parseInput (API JSON -> Excel)', () => {
  it('does not coerce between JSON types', () => {
    expect(parseInput(col('number'), '12').ok).toBe(false);
    expect(parseInput(col('boolean'), 'true').ok).toBe(false);
    expect(parseInput(col('string'), 12).ok).toBe(false);
  });

  it('rejects formula-like leading characters in text but allows negative numbers', () => {
    for (const bad of ['=1+1', '+SUM(A1)', '-2+3', '@cmd', '\t=x']) {
      expect(parseInput(col('string'), bad).ok).toBe(false);
    }
    expect(parseInput(col('number'), -5)).toEqual({ ok: true, value: -5 });
    expect(parseInput(col('string'), '=x', { forFilter: true }).ok).toBe(true);
  });

  it('enforces numeric bounds and text length', () => {
    expect(parseInput(col('number', { min: 0 }), -1).ok).toBe(false);
    expect(parseInput(col('string', { maxLength: 3 }), 'abcd').ok).toBe(false);
  });

  it('normalizes empty strings to null and validates uuids', () => {
    expect(parseInput(col('string'), '')).toEqual({ ok: true, value: null });
    expect(parseInput(col('uuid'), 'not-a-uuid').ok).toBe(false);
    expect(parseInput(col('uuid'), 'A0EEBC99-9C0B-4EF8-BB6D-6BB9BD380A11')).toEqual({
      ok: true,
      value: 'a0eebc99-9c0b-4ef8-bb6d-6bb9bd380a11',
    });
  });
});
