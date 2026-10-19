/**
 * Default-deny authorization (implementation.md §4, §7).
 *
 * Every decision is made per role and then combined with "any role may grant", so a user
 * with several roles never mixes one role's columns with another role's rows. For example,
 * an update is allowed only if a single role grants the action, passes the row rule for
 * both the old and new record, and may write every submitted field.
 */
import type {
  ColumnSchema,
  Identity,
  RecordValues,
  RoleSet,
  TableAction,
  TableSchema,
  WorkbookModel,
} from './types.js';
import { TABLE_ACTIONS } from './types.js';

const allows = (set: RoleSet, role: string): boolean => set.any || set.roles.has(role);

/** The caller's roles that grant `action` on the table, in the caller's role order. */
export function rolesGranting(
  table: TableSchema,
  identity: Identity,
  action: TableAction,
): string[] {
  return identity.roles.filter((role) => table.permissions.get(role)?.has(action) === true);
}

/**
 * Row rule for one role/action. No rules = unrestricted beyond table permission; an
 * `allowAll` rule = unrestricted; otherwise every `equals` rule must pass. A missing trusted
 * attribute or a blank record value fails the rule (fail closed).
 */
export function rowRuleAllows(
  table: TableSchema,
  identity: Identity,
  role: string,
  action: TableAction,
  record: RecordValues,
): boolean {
  const rules = table.rowPolicies.filter((p) => p.role === role && p.action === action);
  if (rules.length === 0 || rules.some((p) => p.operator === 'allowAll')) return true;
  return rules.every((p) => {
    if (p.operator !== 'equals') return false;
    const expected = identity.attributes[p.claim];
    const actual = record[p.column];
    return (
      expected !== undefined &&
      actual !== null &&
      actual !== undefined &&
      String(actual) === expected
    );
  });
}

/** Roles that grant `action` and whose row rule admits `record`. */
export function qualifyingRoles(
  table: TableSchema,
  identity: Identity,
  action: TableAction,
  record: RecordValues,
): string[] {
  return rolesGranting(table, identity, action).filter((role) =>
    rowRuleAllows(table, identity, role, action, record),
  );
}

export function readableColumnsForRole(table: TableSchema, role: string): ColumnSchema[] {
  return table.columns.filter((c) => allows(c.readRoles, role));
}

export function writableColumnsForRole(table: TableSchema, role: string): ColumnSchema[] {
  return table.columns.filter((c) => allows(c.writeRoles, role));
}

function unionColumns(
  table: TableSchema,
  roles: readonly string[],
  pick: (role: string) => ColumnSchema[],
): ColumnSchema[] {
  const names = new Set(roles.flatMap((role) => pick(role).map((c) => c.name)));
  return table.columns.filter((c) => names.has(c.name));
}

/** Columns the caller can read on at least some rows (used for schema, filters and sorts). */
export function readableColumns(table: TableSchema, identity: Identity): ColumnSchema[] {
  return unionColumns(table, rolesGranting(table, identity, 'read'), (role) =>
    readableColumnsForRole(table, role),
  );
}

/**
 * The caller's view of one record: undefined when no role may read it, otherwise only the
 * columns readable by the roles that admit this particular row.
 */
export function projectRecord(
  table: TableSchema,
  identity: Identity,
  record: RecordValues,
): RecordValues | undefined {
  const roles = qualifyingRoles(table, identity, 'read', record);
  if (roles.length === 0) return undefined;
  const projected: Record<string, RecordValues[string]> = {};
  for (const column of unionColumns(table, roles, (role) => readableColumnsForRole(table, role))) {
    projected[column.name] = record[column.name] ?? null;
  }
  return projected;
}

export interface TableCapabilities {
  readonly table: string;
  readonly actions: readonly TableAction[];
  readonly readableColumns: readonly string[];
  readonly creatableColumns: readonly string[];
  readonly updatableColumns: readonly string[];
  /** 'all' when some granting role is not row-restricted for the action. UI hint only. */
  readonly rowScope: Readonly<Partial<Record<TableAction, 'all' | 'restricted'>>>;
}

/**
 * What the caller can do through McSheety, for UI hints. Exposes only the caller's own
 * effective rights, never other roles' policies or claim values. The API re-checks every request.
 */
export function capabilities(model: WorkbookModel, identity: Identity): TableCapabilities[] {
  const result: TableCapabilities[] = [];
  for (const table of model.tables.values()) {
    const actions = TABLE_ACTIONS.filter((a) => rolesGranting(table, identity, a).length > 0);
    if (actions.length === 0) continue;
    const rowScope: Partial<Record<TableAction, 'all' | 'restricted'>> = {};
    for (const action of actions) {
      const unrestricted = rolesGranting(table, identity, action).some((role) => {
        const rules = table.rowPolicies.filter((p) => p.role === role && p.action === action);
        return rules.length === 0 || rules.some((p) => p.operator === 'allowAll');
      });
      rowScope[action] = unrestricted ? 'all' : 'restricted';
    }
    const writable = (action: TableAction): string[] =>
      unionColumns(table, rolesGranting(table, identity, action), (role) =>
        writableColumnsForRole(table, role),
      ).map((c) => c.name);
    result.push({
      table: table.name,
      actions,
      readableColumns: readableColumns(table, identity).map((c) => c.name),
      creatableColumns: writable('create'),
      updatableColumns: writable('update'),
      rowScope,
    });
  }
  return result;
}
