/**
 * Parses the McSheety metadata tables (implementation.md §4) into a WorkbookModel.
 *
 * Fail-closed rules:
 * - A malformed metadata header, or an unsupported formatVersion, yields a model with no tables.
 * - Any error attributable to one logical table removes that table from the model.
 * - Missing column policies are warnings and deny the column.
 */
import type {
  CellValue,
  ColumnSchema,
  DefaultSpec,
  FieldType,
  RawTable,
  RoleSet,
  RowPolicy,
  TableAction,
  TableSchema,
  WorkbookModel,
} from './types.js';
import { FIELD_TYPES, TABLE_ACTIONS } from './types.js';
import { decodeCell } from './values.js';

export const METADATA_TABLES = {
  config: 'mc_config',
  schema: 'mc_schema',
  permissions: 'mc_permissions',
  columns: 'mc_columns',
  rowPolicies: 'mc_rowPolicies',
} as const;

export const SUPPORTED_FORMAT_VERSION = 1;

/** Header names of one data table, as discovered in the workbook. */
export interface DataTableInfo {
  readonly columns: readonly string[];
}

export interface WorkbookSource {
  readonly config?: RawTable | undefined;
  readonly schema: RawTable | undefined;
  readonly permissions: RawTable | undefined;
  readonly columns: RawTable | undefined;
  readonly rowPolicies: RawTable | undefined;
  /** Data tables present in the workbook, keyed by Excel table name. */
  readonly dataTables: ReadonlyMap<string, DataTableInfo>;
}

export interface ParseOptions {
  /** Logical table name -> Excel table name, where `tbl_<Table>` is not usable. */
  readonly excelTableNames?: Readonly<Record<string, string>>;
}

export interface ModelIssue {
  readonly severity: 'error' | 'warning';
  /** Metadata or Excel table the issue was found in. */
  readonly source: string;
  /** Logical table affected, if any. */
  readonly table?: string;
  /** 1-based body row within `source`. */
  readonly row?: number;
  readonly message: string;
}

export interface ParseResult {
  readonly model: WorkbookModel;
  readonly issues: readonly ModelIssue[];
}

const NAME_RE = /^[A-Za-z][A-Za-z0-9_]*$/;
const ROLE_RE = /^[A-Za-z][A-Za-z0-9_.-]*$/;
const ID_COLUMN = 'id';

interface MetaRow {
  readonly row: number;
  get(column: string): CellValue;
  text(column: string): string;
}

class IssueLog {
  readonly issues: ModelIssue[] = [];
  readonly brokenTables = new Set<string>();
  fatal = false;

  error(source: string, message: string, at: { table?: string; row?: number } = {}): void {
    this.issues.push({ severity: 'error', source, message, ...at });
    if (at.table) this.brokenTables.add(at.table);
  }

  warn(source: string, message: string, at: { table?: string; row?: number } = {}): void {
    this.issues.push({ severity: 'warning', source, message, ...at });
  }
}

function readMeta(
  raw: RawTable | undefined,
  source: string,
  required: readonly string[],
  log: IssueLog,
): MetaRow[] | undefined {
  if (!raw) {
    log.error(source, `Metadata table ${source} is missing.`);
    log.fatal = true;
    return undefined;
  }
  const index = new Map(raw.columns.map((name, i) => [name, i]));
  const missing = required.filter((c) => !index.has(c));
  if (missing.length > 0) {
    log.error(source, `Metadata table ${source} is missing columns: ${missing.join(', ')}.`);
    log.fatal = true;
    return undefined;
  }
  for (const extra of raw.columns.filter((c) => !required.includes(c))) {
    log.warn(source, `Ignoring unknown column "${extra}" in ${source}.`);
  }
  return raw.rows.map((cells, i) => {
    const get = (column: string): CellValue => cells[index.get(column) ?? -1] ?? null;
    return {
      row: i + 1,
      get,
      text: (column) => {
        const v = get(column);
        return v === null ? '' : String(v).trim();
      },
    };
  });
}

function parseBool(value: CellValue): boolean | undefined {
  if (value === null || value === '') return false;
  if (typeof value === 'boolean') return value;
  if (typeof value === 'string' && /^(true|false)$/i.test(value.trim())) {
    return value.trim().toLowerCase() === 'true';
  }
  return undefined;
}

function parseRoleSet(text: string): RoleSet | string {
  if (text === '') return { any: false, roles: new Set() };
  if (text === '*') return { any: true, roles: new Set() };
  const roles = text.split(',').map((r) => r.trim());
  const bad = roles.find((r) => !ROLE_RE.test(r));
  if (bad !== undefined)
    return `invalid role name "${bad}" (use "*" alone or a comma-separated list)`;
  return { any: false, roles: new Set(roles) };
}

interface ColumnDraft {
  name: string;
  type: FieldType;
  required: boolean;
  unique: boolean;
  defaultText: CellValue;
  enumValues: string[] | undefined;
  min: number | undefined;
  max: number | undefined;
  maxLength: number | undefined;
  row: number;
}

function parseOptions(draft: ColumnDraft, options: string): string | undefined {
  if (draft.type === 'enum') {
    const values = options.split('|').map((v) => v.trim());
    if (options === '' || values.some((v) => v === '')) {
      return 'enum columns need options like "Active|Inactive"';
    }
    if (new Set(values).size !== values.length) return 'enum options contain duplicates';
    draft.enumValues = values;
    return undefined;
  }
  if (options === '') return undefined;
  const allowed: Partial<Record<FieldType, readonly string[]>> = {
    number: ['min', 'max'],
    string: ['maxLength'],
  };
  for (const part of options.split(';').map((p) => p.trim())) {
    const m = /^([A-Za-z]+)=(-?\d+(?:\.\d+)?)$/.exec(part);
    const key = m?.[1];
    if (!m || key === undefined || !allowed[draft.type]?.includes(key)) {
      return `unsupported option "${part}" for type ${draft.type}`;
    }
    const n = Number(m[2]);
    if (key === 'min') draft.min = n;
    else if (key === 'max') draft.max = n;
    else if (Number.isInteger(n) && n > 0) draft.maxLength = n;
    else return 'maxLength must be a positive integer';
  }
  if (draft.min !== undefined && draft.max !== undefined && draft.min > draft.max) {
    return 'min is greater than max';
  }
  return undefined;
}

function parseDefault(column: ColumnSchema, text: CellValue): DefaultSpec | string | undefined {
  if (text === null || text === '') return undefined;
  if (text === 'generated') {
    return column.type === 'string' || column.type === 'uuid'
      ? { kind: 'generated' }
      : '"generated" defaults are only valid for string or uuid columns';
  }
  if (text === 'now') {
    return column.type === 'datetime'
      ? { kind: 'now' }
      : '"now" defaults are only valid for datetime columns';
  }
  if (typeof text === 'string' && text.startsWith('claim:')) {
    const claim = text.slice('claim:'.length);
    if (!NAME_RE.test(claim)) return `invalid claim name "${claim}"`;
    return column.type === 'string' || column.type === 'enum' || column.type === 'uuid'
      ? { kind: 'claim', claim }
      : 'claim defaults are only valid for string, enum or uuid columns';
  }
  const decoded = decodeCell(column, text);
  if (!decoded.ok || decoded.value === null) {
    return `default ${decoded.ok ? 'is blank' : decoded.reason}`;
  }
  return { kind: 'literal', value: decoded.value };
}

export function parseWorkbookModel(
  source: WorkbookSource,
  options: ParseOptions = {},
): ParseResult {
  const log = new IssueLog();
  const { config: cfgName, schema: schemaName, permissions: permName } = METADATA_TABLES;
  const { columns: colName, rowPolicies: rpName } = METADATA_TABLES;

  // --- mc_config (optional) ---------------------------------------------------------
  const config = new Map<string, string>();
  if (source.config) {
    const rows = readMeta(source.config, cfgName, ['key', 'value'], log);
    for (const r of rows ?? []) config.set(r.text('key'), r.text('value'));
  } else {
    log.warn(cfgName, 'No mc_config table; assuming formatVersion 1.');
  }
  const formatVersion = Number(config.get('formatVersion') ?? SUPPORTED_FORMAT_VERSION);
  if (formatVersion !== SUPPORTED_FORMAT_VERSION) {
    log.error(cfgName, `Unsupported formatVersion "${config.get('formatVersion')}".`);
    log.fatal = true;
  }

  const schemaRows = readMeta(
    source.schema,
    schemaName,
    ['table', 'column', 'type', 'required', 'default', 'unique', 'options'],
    log,
  );
  const permRows = readMeta(source.permissions, permName, ['table', 'role', ...TABLE_ACTIONS], log);
  const colRows = readMeta(
    source.columns,
    colName,
    ['table', 'column', 'readRoles', 'writeRoles'],
    log,
  );
  const rpRows = readMeta(
    source.rowPolicies,
    rpName,
    ['table', 'action', 'role', 'column', 'operator', 'claim'],
    log,
  );

  const emptyModel = (): ParseResult => ({
    model: {
      formatVersion,
      schemaVersion: config.get('schemaVersion') ?? '0',
      policyVersion: config.get('policyVersion') ?? '0',
      applicationId: config.get('applicationId'),
      tables: new Map(),
      roles: new Set(),
    },
    issues: log.issues,
  });
  if (log.fatal || !schemaRows || !permRows || !colRows || !rpRows) return emptyModel();

  // --- mc_schema ----------------------------------------------------------------------
  const drafts = new Map<string, Map<string, ColumnDraft>>();
  for (const r of schemaRows) {
    const table = r.text('table');
    const column = r.text('column');
    if (!NAME_RE.test(table)) {
      log.error(schemaName, `Invalid table name "${table}".`, { row: r.row });
      continue;
    }
    const at = { table, row: r.row };
    if (!NAME_RE.test(column)) {
      log.error(schemaName, `Invalid column name "${column}".`, at);
      continue;
    }
    const type = r.text('type') as FieldType;
    if (!FIELD_TYPES.includes(type)) {
      log.error(schemaName, `Unknown type "${r.text('type')}" for ${table}.${column}.`, at);
      continue;
    }
    const required = parseBool(r.get('required'));
    const unique = parseBool(r.get('unique'));
    if (required === undefined || unique === undefined) {
      log.error(schemaName, `required/unique must be TRUE or FALSE for ${table}.${column}.`, at);
      continue;
    }
    const draft: ColumnDraft = {
      name: column,
      type,
      required,
      unique,
      defaultText: r.get('default'),
      enumValues: undefined,
      min: undefined,
      max: undefined,
      maxLength: undefined,
      row: r.row,
    };
    const optionError = parseOptions(draft, r.text('options'));
    if (optionError) {
      log.error(schemaName, `${table}.${column}: ${optionError}.`, at);
      continue;
    }
    const tableDrafts = drafts.get(table) ?? new Map<string, ColumnDraft>();
    if (tableDrafts.has(column)) {
      log.error(schemaName, `Duplicate schema row for ${table}.${column}.`, at);
      continue;
    }
    tableDrafts.set(column, draft);
    drafts.set(table, tableDrafts);
  }

  // --- mc_permissions -----------------------------------------------------------------
  const permissions = new Map<string, Map<string, Set<TableAction>>>();
  const knownRoles = new Set<string>();
  for (const r of permRows) {
    const table = r.text('table');
    const role = r.text('role');
    const at = { table, row: r.row };
    if (!drafts.has(table)) {
      log.warn(permName, `Permissions for unknown table "${table}" are ignored.`, { row: r.row });
      continue;
    }
    if (!ROLE_RE.test(role)) {
      log.error(permName, `Invalid role name "${role}".`, at);
      continue;
    }
    const tablePerms = permissions.get(table) ?? new Map<string, Set<TableAction>>();
    if (tablePerms.has(role)) {
      log.error(permName, `Duplicate permissions for ${table}/${role}.`, at);
      continue;
    }
    const actions = new Set<TableAction>();
    let valid = true;
    for (const action of TABLE_ACTIONS) {
      const granted = parseBool(r.get(action));
      if (granted === undefined) {
        log.error(permName, `${action} must be TRUE or FALSE for ${table}/${role}.`, at);
        valid = false;
      } else if (granted) actions.add(action);
    }
    if (!valid) continue;
    tablePerms.set(role, actions);
    permissions.set(table, tablePerms);
    knownRoles.add(role);
  }

  // --- mc_columns ---------------------------------------------------------------------
  const columnPolicies = new Map<string, Map<string, { read: RoleSet; write: RoleSet }>>();
  const referencedRoles: { role: string; source: string; row: number; table: string }[] = [];
  for (const r of colRows) {
    const table = r.text('table');
    const column = r.text('column');
    const at = { table, row: r.row };
    if (!drafts.has(table)) {
      log.warn(colName, `Column policy for unknown table "${table}" is ignored.`, { row: r.row });
      continue;
    }
    if (!drafts.get(table)?.has(column)) {
      log.error(colName, `Column policy for unknown column ${table}.${column}.`, at);
      continue;
    }
    const read = parseRoleSet(r.text('readRoles'));
    const write = parseRoleSet(r.text('writeRoles'));
    if (typeof read === 'string' || typeof write === 'string') {
      const reason = typeof read === 'string' ? read : (write as string);
      log.error(colName, `${table}.${column}: ${reason}.`, at);
      continue;
    }
    const tablePolicies =
      columnPolicies.get(table) ?? new Map<string, { read: RoleSet; write: RoleSet }>();
    if (tablePolicies.has(column)) {
      log.error(colName, `Duplicate column policy for ${table}.${column}.`, at);
      continue;
    }
    tablePolicies.set(column, { read, write });
    columnPolicies.set(table, tablePolicies);
    for (const role of [...read.roles, ...write.roles]) {
      referencedRoles.push({ role, source: colName, row: r.row, table });
    }
  }

  // --- mc_rowPolicies -----------------------------------------------------------------
  const rowPolicies = new Map<string, RowPolicy[]>();
  for (const r of rpRows) {
    const table = r.text('table');
    const at = { table, row: r.row };
    if (!drafts.has(table)) {
      log.warn(rpName, `Row policy for unknown table "${table}" is ignored.`, { row: r.row });
      continue;
    }
    const action = r.text('action') as TableAction;
    const role = r.text('role');
    const operator = r.text('operator');
    const column = r.text('column');
    const claim = r.text('claim');
    if (!TABLE_ACTIONS.includes(action)) {
      log.error(rpName, `Unknown action "${r.text('action')}".`, at);
      continue;
    }
    if (!ROLE_RE.test(role)) {
      log.error(rpName, `Invalid role name "${role}".`, at);
      continue;
    }
    let policy: RowPolicy;
    if (operator === 'equals') {
      if (!drafts.get(table)?.has(column)) {
        log.error(rpName, `equals policy references unknown column "${column}".`, at);
        continue;
      }
      if (!NAME_RE.test(claim)) {
        log.error(rpName, `equals policy needs a claim name, got "${claim}".`, at);
        continue;
      }
      policy = { action, role, operator, column, claim };
    } else if (operator === 'allowAll') {
      if (column !== '' || claim !== '') {
        log.error(rpName, 'allowAll policies must leave column and claim blank.', at);
        continue;
      }
      policy = { action, role, operator };
    } else {
      log.error(rpName, `Unsupported operator "${operator}" (supported: equals, allowAll).`, at);
      continue;
    }
    rowPolicies.set(table, [...(rowPolicies.get(table) ?? []), policy]);
    referencedRoles.push({ role, source: rpName, row: r.row, table });
  }

  for (const ref of referencedRoles) {
    if (!knownRoles.has(ref.role)) {
      log.warn(ref.source, `Role "${ref.role}" has no entry in mc_permissions.`, {
        table: ref.table,
        row: ref.row,
      });
    }
  }

  // --- assemble tables ----------------------------------------------------------------
  const tables = new Map<string, TableSchema>();
  for (const [table, tableDrafts] of drafts) {
    const excelTable = options.excelTableNames?.[table] ?? `tbl_${table}`;
    const info = source.dataTables.get(excelTable);
    if (!info) {
      log.error(schemaName, `Excel table "${excelTable}" for ${table} was not found.`, { table });
      continue;
    }
    for (const name of tableDrafts.keys()) {
      if (!info.columns.includes(name)) {
        log.error(excelTable, `Schema column ${table}.${name} is not a header in ${excelTable}.`, {
          table,
        });
      }
    }
    for (const header of info.columns) {
      if (!tableDrafts.has(header)) {
        log.warn(excelTable, `Header "${header}" has no schema row and is not exposed.`, { table });
      }
    }

    const policies = columnPolicies.get(table);
    const columns: ColumnSchema[] = [];
    for (const header of info.columns) {
      const draft = tableDrafts.get(header);
      if (!draft) continue;
      const policy = policies?.get(header);
      if (!policy) {
        log.warn(colName, `No column policy for ${table}.${header}; it is hidden and read-only.`, {
          table,
        });
      }
      const none: RoleSet = { any: false, roles: new Set() };
      const base: ColumnSchema = {
        name: draft.name,
        type: draft.type,
        required: draft.required,
        unique: draft.unique,
        default: undefined,
        enumValues: draft.enumValues,
        min: draft.min,
        max: draft.max,
        maxLength: draft.maxLength,
        readRoles: policy?.read ?? none,
        writeRoles: policy?.write ?? none,
      };
      const def = parseDefault(base, draft.defaultText);
      if (typeof def === 'string') {
        log.error(schemaName, `${table}.${header}: ${def}.`, { table, row: draft.row });
        continue;
      }
      columns.push({ ...base, default: def });
    }

    const id = columns.find((c) => c.name === ID_COLUMN);
    if (!id) {
      log.error(schemaName, `${table} must define an "${ID_COLUMN}" column.`, { table });
    } else {
      const problems: string[] = [];
      if (id.type !== 'string' && id.type !== 'uuid') problems.push('type string or uuid');
      if (!id.required || !id.unique) problems.push('required and unique');
      if (id.writeRoles.any || id.writeRoles.roles.size > 0) problems.push('blank writeRoles');
      if (id.default?.kind !== 'generated') problems.push('default "generated"');
      if (problems.length > 0) {
        log.error(schemaName, `${table}.${ID_COLUMN} must have ${problems.join(', ')}.`, { table });
      }
    }

    const tablePerms = permissions.get(table);
    if (!tablePerms || tablePerms.size === 0) {
      log.warn(permName, `${table} has no permissions; nobody can access it.`, { table });
    }

    if (log.brokenTables.has(table)) continue;
    tables.set(table, {
      name: table,
      excelTable,
      columns,
      columnsByName: new Map(columns.map((c) => [c.name, c])),
      permissions: tablePerms ?? new Map(),
      rowPolicies: rowPolicies.get(table) ?? [],
    });
  }

  for (const table of log.brokenTables) {
    if (drafts.has(table)) {
      log.error(schemaName, `${table} is disabled until its metadata errors are fixed.`, { table });
    }
  }

  return {
    model: { ...emptyModel().model, tables, roles: knownRoles },
    issues: log.issues,
  };
}
