import type { CellValue, FieldType } from '@mcsheety/core';

/**
 * How text is protected from Excel's input parsing when written through Graph.
 * - 'apostrophe': prefix "'" so Excel stores the text literally (as typing 'text does in Excel).
 * - 'raw': write as-is (Excel may turn "2026-10-16" into a date serial or "1/2" into a date).
 *
 * TODO(MC-002 spike): confirm Graph honours the apostrophe prefix for range/rows writes and
 * returns the value without it. If not, switch to setting numberFormat "@" on text cells.
 */
export type TextEncoding = 'apostrophe' | 'raw';

const TEXT_TYPES: ReadonlySet<FieldType> = new Set([
  'string',
  'enum',
  'uuid',
  'date',
  'time',
  'datetime',
]);

export function encodeCell(
  value: CellValue,
  type: FieldType | undefined,
  encoding: TextEncoding,
): string | number | boolean {
  if (value === null) return '';
  if (
    typeof value === 'string' &&
    encoding === 'apostrophe' &&
    (type === undefined || TEXT_TYPES.has(type))
  ) {
    return `'${value}`;
  }
  return value;
}
