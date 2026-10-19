import type {
  McSheetyClient,
  SchemaResult,
  TableCapabilities,
  TableRecord,
} from '@mcsheety/client';
import { McSheetyClientError } from '@mcsheety/client';
import { useCallback, useEffect, useState } from 'react';
import { RecordForm } from './RecordForm';

type Value = string | number | boolean | null;
type Editing = { mode: 'create' } | { mode: 'edit'; record: TableRecord };

interface Props {
  readonly db: McSheetyClient;
  readonly table: string;
  readonly capability: TableCapabilities;
}

function describe(error: unknown): string {
  if (error instanceof McSheetyClientError) {
    const hint =
      error.code === 'CONFLICT'
        ? ' Reload to see the latest version, then try again.'
        : error.code === 'RATE_LIMITED'
          ? ' Wait a moment and try again.'
          : '';
    return `${error.message}${hint}${error.requestId ? ` (request ${error.requestId})` : ''}`;
  }
  return 'Something went wrong.';
}

export function TableView({ db, table, capability }: Props) {
  const [schema, setSchema] = useState<SchemaResult | null>(null);
  const [records, setRecords] = useState<readonly TableRecord[]>([]);
  const [nextCursor, setNextCursor] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [editing, setEditing] = useState<Editing | null>(null);
  const [busy, setBusy] = useState(false);
  const [fieldErrors, setFieldErrors] = useState<Record<string, string>>({});
  const [statusFilter, setStatusFilter] = useState('');

  const can = (action: 'create' | 'update' | 'delete') => capability.actions.includes(action);

  const load = useCallback(
    async (cursor: string | null = null) => {
      setLoading(true);
      setError(null);
      try {
        const [s, page] = await Promise.all([
          cursor ? Promise.resolve(null) : db.schema(table),
          db.table(table).findMany({
            limit: 25,
            sort: ['name'],
            cursor,
            ...(statusFilter ? { where: { status: statusFilter } } : {}),
          }),
        ]);
        if (s) setSchema(s);
        setRecords((prev) => (cursor ? [...prev, ...page.records] : page.records));
        setNextCursor(page.nextCursor);
      } catch (e) {
        setError(describe(e));
      } finally {
        setLoading(false);
      }
    },
    [db, table, statusFilter],
  );

  useEffect(() => {
    void load();
  }, [load]);

  async function save(values: Record<string, Value>) {
    if (!editing) return;
    setBusy(true);
    setFieldErrors({});
    try {
      if (editing.mode === 'create') await db.table(table).create(values);
      else if (Object.keys(values).length > 0) {
        await db
          .table(table)
          .update(editing.record.id, values, { ifMatch: editing.record.revision });
      }
      setEditing(null);
      await load();
    } catch (e) {
      if (e instanceof McSheetyClientError && e.details.length > 0) {
        setFieldErrors(
          Object.fromEntries(
            e.details.filter((d) => d.field).map((d) => [d.field as string, d.message]),
          ),
        );
      }
      setError(describe(e));
    } finally {
      setBusy(false);
    }
  }

  async function remove(record: TableRecord) {
    if (!window.confirm(`Delete ${record.id}? This removes the row from the Excel workbook.`))
      return;
    try {
      await db.table(table).delete(record.id, { ifMatch: record.revision });
      await load();
    } catch (e) {
      setError(describe(e));
    }
  }

  const columns = schema?.columns ?? [];
  const statusColumn = columns.find((c) => c.name === 'status' && c.type === 'enum');
  const formColumns = columns.filter((c) =>
    editing?.mode === 'create' ? c.creatable : c.updatable,
  );

  return (
    <section>
      <div className="toolbar">
        <h2>{table}</h2>
        {statusColumn ? (
          <select
            aria-label="Filter by status"
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value)}
          >
            <option value="">All statuses</option>
            {statusColumn.enumValues?.map((v) => (
              <option key={v} value={v}>
                {v}
              </option>
            ))}
          </select>
        ) : null}
        <button className="secondary" onClick={() => void load()} disabled={loading}>
          Reload
        </button>
        {can('create') ? (
          <button onClick={() => setEditing({ mode: 'create' })}>New record</button>
        ) : null}
      </div>

      <p className="permissions">
        You can <strong>{capability.actions.join(', ')}</strong>
        {capability.rowScope.read === 'restricted' ? ' — only rows in your scope' : ''}. Columns you
        cannot read are not sent to your browser.
      </p>

      {error ? (
        <div role="alert" className="error">
          {error}
        </div>
      ) : null}

      {editing ? (
        <RecordForm
          key={editing.mode === 'edit' ? editing.record.id : 'new'}
          title={editing.mode === 'create' ? `New ${table} record` : `Edit ${editing.record.id}`}
          columns={formColumns}
          initial={editing.mode === 'edit' ? editing.record.fields : {}}
          busy={busy}
          fieldErrors={fieldErrors}
          onSubmit={(v) => void save(v)}
          onCancel={() => {
            setEditing(null);
            setFieldErrors({});
          }}
        />
      ) : null}

      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              {columns.map((c) => (
                <th key={c.name}>{c.name}</th>
              ))}
              {can('update') || can('delete') ? <th aria-label="Actions" /> : null}
            </tr>
          </thead>
          <tbody>
            {records.map((r) => (
              <tr key={r.id} className={r.issues ? 'has-issues' : undefined}>
                {columns.map((c) => (
                  <td key={c.name} title={r.issues?.find((i) => i.field === c.name)?.message}>
                    {String(r.fields[c.name] ?? '')}
                  </td>
                ))}
                {can('update') || can('delete') ? (
                  <td className="row-actions">
                    {can('update') ? (
                      <button
                        className="link"
                        onClick={() => setEditing({ mode: 'edit', record: r })}
                      >
                        Edit
                      </button>
                    ) : null}
                    {can('delete') ? (
                      <button className="link danger" onClick={() => void remove(r)}>
                        Delete
                      </button>
                    ) : null}
                  </td>
                ) : null}
              </tr>
            ))}
            {!loading && records.length === 0 ? (
              <tr>
                <td colSpan={columns.length + 1} className="empty">
                  No records you can see.
                </td>
              </tr>
            ) : null}
          </tbody>
        </table>
      </div>
      {nextCursor ? (
        <button className="secondary" onClick={() => void load(nextCursor)} disabled={loading}>
          Load more
        </button>
      ) : null}
      {loading ? <p className="muted">Loading…</p> : null}
    </section>
  );
}
