/** A single Excel cell value as Microsoft Graph returns it. Graph reports blank cells as "". */
export type CellValue = string | number | boolean | null;

/** Contents of one formal Excel Table: header names plus body rows, in sheet order. */
export interface RawTable {
  readonly columns: readonly string[];
  readonly rows: readonly (readonly CellValue[])[];
  /**
   * Optional per-cell formulas for the body rows (same shape as `rows`). A cell whose
   * formula starts with "=" is formula-bearing; McSheety rejects writes to it.
   */
  readonly formulas?: readonly (readonly CellValue[])[];
}

/** One record keyed by column name. Values are normalized for the column type; blank = null. */
export type RecordValues = Readonly<Record<string, CellValue>>;

export const FIELD_TYPES = [
  'string',
  'number',
  'boolean',
  'date',
  'time',
  'datetime',
  'uuid',
  'enum',
] as const;
export type FieldType = (typeof FIELD_TYPES)[number];

export const TABLE_ACTIONS = ['read', 'create', 'update', 'delete'] as const;
export type TableAction = (typeof TABLE_ACTIONS)[number];

/**
 * The verified caller. Only build this from a validated Entra access token and trusted
 * server-side mappings; never from request bodies or headers the client controls.
 */
export interface Identity {
  readonly tenantId: string;
  /** Entra object ID (`oid`). `tenantId + userId` is the stable principal key. */
  readonly userId: string;
  /** Verified Entra app roles (or roles mapped server-side from group object IDs). */
  readonly roles: readonly string[];
  /** Trusted attributes for row policies, e.g. `departmentId`. Absent = rule fails closed. */
  readonly attributes: Readonly<Record<string, string>>;
}

/** `*` (any role that has the relevant table action) or an explicit role list. Empty = nobody. */
export interface RoleSet {
  readonly any: boolean;
  readonly roles: ReadonlySet<string>;
}

export type DefaultSpec =
  | { readonly kind: 'generated' }
  | { readonly kind: 'now' }
  | { readonly kind: 'claim'; readonly claim: string }
  | { readonly kind: 'literal'; readonly value: Exclude<CellValue, null> };

export interface ColumnSchema {
  readonly name: string;
  readonly type: FieldType;
  readonly required: boolean;
  readonly unique: boolean;
  readonly default: DefaultSpec | undefined;
  readonly enumValues: readonly string[] | undefined;
  readonly min: number | undefined;
  readonly max: number | undefined;
  readonly maxLength: number | undefined;
  readonly readRoles: RoleSet;
  readonly writeRoles: RoleSet;
}

export type RowPolicy =
  | {
      readonly action: TableAction;
      readonly role: string;
      readonly operator: 'equals';
      readonly column: string;
      readonly claim: string;
    }
  | { readonly action: TableAction; readonly role: string; readonly operator: 'allowAll' };

export interface TableSchema {
  /** Logical table name (the worksheet name). */
  readonly name: string;
  /** Name of the formal Excel Table holding the data, e.g. `tbl_Equipment`. */
  readonly excelTable: string;
  /** Columns in Excel header order. */
  readonly columns: readonly ColumnSchema[];
  readonly columnsByName: ReadonlyMap<string, ColumnSchema>;
  /** role -> actions granted on this table. */
  readonly permissions: ReadonlyMap<string, ReadonlySet<TableAction>>;
  readonly rowPolicies: readonly RowPolicy[];
}

export interface WorkbookModel {
  readonly formatVersion: number;
  readonly schemaVersion: string;
  readonly policyVersion: string;
  readonly applicationId: string | undefined;
  /** Only tables whose metadata parsed without errors. Broken tables are absent (fail closed). */
  readonly tables: ReadonlyMap<string, TableSchema>;
  /** Every role named in mc_permissions. */
  readonly roles: ReadonlySet<string>;
}
