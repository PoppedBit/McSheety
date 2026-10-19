/**
 * McSheety API operations (implementation.md §7, §9, §10), independent of HTTP and Lambda.
 * Every call reloads metadata and re-checks permissions; nothing is trusted from the client.
 */
import { createHash, randomUUID } from 'node:crypto';
import type {
  ColumnInfo,
  DecodedRecord,
  Identity,
  ListQueryParams,
  ListResult,
  ParseOptions,
  QueryLimits,
  RecordEnvelope,
  RecordValues,
  SchemaResult,
  TableAction,
  TableCapabilities,
  TableSchema,
  TableSnapshot,
  WorkbookAdapter,
  WorkbookContext,
  WorkbookModel,
} from '@mcsheety/core';
import {
  DEFAULT_QUERY_LIMITS,
  McSheetyError,
  applyListQuery,
  authorizeDelete,
  canonicalRecord,
  capabilities,
  decodeTable,
  findRecord,
  loadWorkbookModel,
  parseListQuery,
  prepareCreate,
  prepareUpdate,
  projectRecord,
  readableColumns,
  rolesGranting,
  writableColumnsForRole,
} from '@mcsheety/core';

export interface AuditEvent {
  readonly type: 'audit';
  readonly actor: { readonly tenantId: string; readonly userId: string };
  readonly action: Exclude<TableAction, 'read'>;
  readonly table: string;
  readonly recordId: string;
  readonly timestamp: string;
  readonly outcome: 'success' | 'denied' | 'failed';
  readonly correlationId: string;
  /** Field names only; values are not audited by default. */
  readonly changedFields: readonly string[];
  readonly errorCode?: string;
}

export interface ServiceOptions {
  readonly adapter: WorkbookAdapter;
  readonly now?: () => Date;
  readonly newId?: () => string;
  readonly queryLimits?: QueryLimits;
  readonly parseOptions?: ParseOptions;
  readonly audit?: (event: AuditEvent) => void;
  /** Receives metadata problems found while loading the workbook (never sent to clients). */
  readonly onModelIssues?: (
    issues: readonly { severity: string; message: string; table?: string }[],
    correlationId: string,
  ) => void;
}

/** The verified caller plus a workbook context carrying their delegated Graph access. */
export interface RequestScope {
  readonly identity: Identity;
  readonly ctx: WorkbookContext;
}

export function computeRevision(view: RecordValues): string {
  return createHash('sha256').update(canonicalRecord(view)).digest('base64url').slice(0, 27);
}

/** Accepts `"rev"`, `rev` or `W/"rev"`; rejects missing values and `*`. */
export function parseIfMatch(header: string | undefined): string {
  const value = header
    ?.trim()
    .replace(/^W\//, '')
    .replace(/^"(.*)"$/, '$1');
  if (!value || value === '*') throw new McSheetyError('PRECONDITION_REQUIRED');
  return value;
}

const auditable = (code: string): 'denied' | 'failed' =>
  code === 'FORBIDDEN' || code === 'NOT_FOUND' ? 'denied' : 'failed';

export class McSheetyService {
  private readonly adapter: WorkbookAdapter;
  private readonly now: () => Date;
  private readonly newId: () => string;
  private readonly limits: QueryLimits;
  private readonly options: ServiceOptions;

  constructor(options: ServiceOptions) {
    this.options = options;
    this.adapter = options.adapter;
    this.now = options.now ?? (() => new Date());
    this.newId = options.newId ?? randomUUID;
    this.limits = options.queryLimits ?? DEFAULT_QUERY_LIMITS;
  }

  private async model(ctx: WorkbookContext): Promise<WorkbookModel> {
    const { model, issues } = await loadWorkbookModel(this.adapter, ctx, this.options.parseOptions);
    if (issues.length > 0) this.options.onModelIssues?.(issues, ctx.correlationId);
    return model;
  }

  /** Resolves a table the caller may use at all. Unknown and inaccessible look the same. */
  private table(model: WorkbookModel, identity: Identity, name: string): TableSchema {
    const table = model.tables.get(name);
    const anyAction =
      table &&
      (['read', 'create', 'update', 'delete'] as const).some(
        (a) => rolesGranting(table, identity, a).length > 0,
      );
    if (!table || !anyAction) throw new McSheetyError('NOT_FOUND');
    return table;
  }

  private envelope(record: DecodedRecord, view: RecordValues): RecordEnvelope {
    const issues = record.issues.filter((i) => i.field in view);
    return {
      id: record.id,
      revision: computeRevision(view),
      fields: view,
      ...(issues.length > 0 ? { issues } : {}),
    };
  }

  /** Locates a record the caller can see; anything else is NOT_FOUND (no existence leak). */
  private visible(table: TableSchema, identity: Identity, snapshot: TableSnapshot, id: string) {
    const decoded = decodeTable(table, snapshot);
    const record = findRecord(decoded, id);
    const view = projectRecord(table, identity, record.values);
    if (!view) throw new McSheetyError('NOT_FOUND');
    return { decoded, record, view };
  }

  private recordAudit(
    scope: RequestScope,
    action: AuditEvent['action'],
    table: string,
    recordId: string,
    outcome: AuditEvent['outcome'],
    changedFields: readonly string[] = [],
    errorCode?: string,
  ): void {
    this.options.audit?.({
      type: 'audit',
      actor: { tenantId: scope.identity.tenantId, userId: scope.identity.userId },
      action,
      table,
      recordId,
      timestamp: this.now().toISOString(),
      outcome,
      correlationId: scope.ctx.correlationId,
      changedFields,
      ...(errorCode ? { errorCode } : {}),
    });
  }

  private async audited<T>(
    scope: RequestScope,
    action: AuditEvent['action'],
    table: string,
    recordId: string,
    fn: () => Promise<{ result: T; recordId?: string; changedFields?: readonly string[] }>,
  ): Promise<T> {
    try {
      const { result, recordId: finalId, changedFields } = await fn();
      this.recordAudit(scope, action, table, finalId ?? recordId, 'success', changedFields);
      return result;
    } catch (error) {
      const code = error instanceof McSheetyError ? error.code : 'INTERNAL_ERROR';
      this.recordAudit(scope, action, table, recordId, auditable(code), [], code);
      throw error;
    }
  }

  async listTables({ identity, ctx }: RequestScope): Promise<{ data: { name: string }[] }> {
    const model = await this.model(ctx);
    const data = [...model.tables.values()]
      .filter((t) => rolesGranting(t, identity, 'read').length > 0)
      .map((t) => ({ name: t.name }));
    return { data };
  }

  async capabilities({ identity, ctx }: RequestScope): Promise<{ data: TableCapabilities[] }> {
    return { data: capabilities(await this.model(ctx), identity) };
  }

  async schema({ identity, ctx }: RequestScope, tableName: string): Promise<SchemaResult> {
    const table = this.table(await this.model(ctx), identity, tableName);
    const writable = (action: TableAction) =>
      new Set(
        rolesGranting(table, identity, action).flatMap((r) =>
          writableColumnsForRole(table, r).map((c) => c.name),
        ),
      );
    const creatable = writable('create');
    const updatable = writable('update');
    const columns = readableColumns(table, identity).map((c): ColumnInfo => ({
      name: c.name,
      type: c.type,
      required: c.required,
      unique: c.unique,
      ...(c.enumValues ? { enumValues: c.enumValues } : {}),
      ...(c.min === undefined ? {} : { min: c.min }),
      ...(c.max === undefined ? {} : { max: c.max }),
      ...(c.maxLength === undefined ? {} : { maxLength: c.maxLength }),
      ...(c.default
        ? {
            default:
              c.default.kind === 'literal' ? { value: c.default.value } : ('server' as const),
          }
        : {}),
      creatable: creatable.has(c.name),
      updatable: updatable.has(c.name),
    }));
    return { table: table.name, columns };
  }

  async listRows(
    { identity, ctx }: RequestScope,
    tableName: string,
    params: ListQueryParams,
  ): Promise<ListResult> {
    const model = await this.model(ctx);
    const table = this.table(model, identity, tableName);
    if (rolesGranting(table, identity, 'read').length === 0) throw new McSheetyError('FORBIDDEN');
    const query = parseListQuery(table, identity, params, this.limits);
    const decoded = decodeTable(table, await this.adapter.readTable(ctx, table.excelTable));
    const views = new Map<RecordValues, DecodedRecord>();
    for (const record of decoded.records) {
      const view = projectRecord(table, identity, record.values);
      if (view) views.set(view, record);
    }
    const page = applyListQuery([...views.keys()], query);
    return {
      data: page.items.map((view) => this.envelope(views.get(view) as DecodedRecord, view)),
      page: { nextCursor: page.nextCursor },
      meta: { table: table.name },
    };
  }

  async getRow(
    { identity, ctx }: RequestScope,
    tableName: string,
    id: string,
  ): Promise<RecordEnvelope> {
    const table = this.table(await this.model(ctx), identity, tableName);
    const { record, view } = this.visible(
      table,
      identity,
      await this.adapter.readTable(ctx, table.excelTable),
      id,
    );
    return this.envelope(record, view);
  }

  async createRow(scope: RequestScope, tableName: string, body: unknown): Promise<RecordEnvelope> {
    return this.audited(scope, 'create', tableName, '', () =>
      this.adapter.withSession(scope.ctx, async (ctx) => {
        const { identity } = scope;
        const table = this.table(await this.model(ctx), identity, tableName);
        const snapshot = await this.adapter.readTable(ctx, table.excelTable);
        const candidate = prepareCreate(table, identity, body, decodeTable(table, snapshot), {
          now: this.now,
          newId: this.newId,
        });
        const id = String(candidate['id']);
        const row = snapshot.columns.map((c) => candidate[c] ?? null);
        const types = snapshot.columns.map((c) => table.columnsByName.get(c)?.type);

        try {
          await this.adapter.appendRow(ctx, table.excelTable, row, types);
        } catch (error) {
          // The append may have landed (timeout / 5xx after send). Never blindly retry:
          // look for the generated id instead.
          const ambiguous =
            error instanceof McSheetyError &&
            (error.code === 'WORKBOOK_UNAVAILABLE' || error.code === 'GRAPH_ERROR');
          if (!ambiguous) throw error;
          const check = decodeTable(table, await this.adapter.readTable(ctx, table.excelTable));
          if (!check.records.some((r) => r.id === id)) throw error;
        }

        // Confirm the observed result and return the caller's view of it.
        const after = await this.adapter.readTable(ctx, table.excelTable);
        const decoded = decodeTable(table, after);
        const record = findRecord(decoded, id);
        const view = projectRecord(table, identity, record.values);
        const result: RecordEnvelope = view
          ? this.envelope(record, view)
          : { id, revision: '', fields: { id } }; // created but outside the caller's read scope
        return { result, recordId: id, changedFields: Object.keys(body as object) };
      }),
    );
  }

  async updateRow(
    scope: RequestScope,
    tableName: string,
    id: string,
    body: unknown,
    ifMatch: string | undefined,
  ): Promise<RecordEnvelope> {
    return this.audited(scope, 'update', tableName, id, () => {
      const expected = parseIfMatch(ifMatch);
      return this.adapter.withSession(scope.ctx, async (ctx) => {
        const { identity } = scope;
        const table = this.table(await this.model(ctx), identity, tableName);
        const snapshot = await this.adapter.readTable(ctx, table.excelTable);
        const { decoded, record, view } = this.visible(table, identity, snapshot, id);
        if (computeRevision(view) !== expected) throw new McSheetyError('CONFLICT');

        const { changes } = prepareUpdate(table, identity, record, body, decoded);
        const changed = Object.keys(changes);
        if (changed.length === 0) return { result: this.envelope(record, view), changedFields: [] };

        const indexOf = new Map(snapshot.columns.map((c, i) => [c, i]));
        const cellChanges = new Map(
          changed.map((c) => [indexOf.get(c) as number, changes[c] ?? null]),
        );
        const types = snapshot.columns.map((c) => table.columnsByName.get(c)?.type);
        const observed = await this.adapter.updateCells(
          ctx,
          snapshot,
          record.index,
          cellChanges,
          types,
        );

        const after = decodeTable(table, { columns: snapshot.columns, rows: [observed] })
          .records[0];
        if (!after || after.id !== id) throw new McSheetyError('CONFLICT');
        const afterView = projectRecord(table, identity, after.values);
        const result: RecordEnvelope = afterView
          ? this.envelope(after, afterView)
          : { id, revision: '', fields: { id } }; // moved outside the caller's read scope
        return { result, changedFields: changed };
      });
    });
  }

  async deleteRow(
    scope: RequestScope,
    tableName: string,
    id: string,
    ifMatch: string | undefined,
  ): Promise<void> {
    await this.audited(scope, 'delete', tableName, id, () => {
      const expected = parseIfMatch(ifMatch);
      return this.adapter.withSession(scope.ctx, async (ctx) => {
        const table = this.table(await this.model(ctx), scope.identity, tableName);
        const snapshot = await this.adapter.readTable(ctx, table.excelTable);
        const { record, view } = this.visible(table, scope.identity, snapshot, id);
        if (computeRevision(view) !== expected) throw new McSheetyError('CONFLICT');
        authorizeDelete(table, scope.identity, record);
        await this.adapter.deleteRow(ctx, snapshot, record.index);
        return { result: undefined };
      });
    });
  }
}
