/** JSON shapes of the REST API v1, shared by @mcsheety/api and @mcsheety/client. */
import type { FieldIssue } from './records.js';
import type { CellValue, FieldType, RecordValues } from './types.js';

export interface RecordEnvelope<T extends RecordValues = RecordValues> {
  readonly id: string;
  /** Opaque revision of the caller's view of the record; send back as If-Match. */
  readonly revision: string;
  readonly fields: T;
  /** Problems with readable cells (e.g. manual edits that break the schema). */
  readonly issues?: readonly FieldIssue[];
}

export interface ListResult<T extends RecordValues = RecordValues> {
  readonly data: readonly RecordEnvelope<T>[];
  readonly page: { readonly nextCursor: string | null };
  readonly meta: { readonly table: string };
}

export interface ColumnInfo {
  readonly name: string;
  readonly type: FieldType;
  readonly required: boolean;
  readonly unique: boolean;
  readonly enumValues?: readonly string[];
  readonly min?: number;
  readonly max?: number;
  readonly maxLength?: number;
  /** 'server' = filled by McSheety (generated id, timestamp or trusted attribute). */
  readonly default?: 'server' | { readonly value: Exclude<CellValue, null> };
  readonly creatable: boolean;
  readonly updatable: boolean;
}

export interface SchemaResult {
  readonly table: string;
  readonly columns: readonly ColumnInfo[];
}
