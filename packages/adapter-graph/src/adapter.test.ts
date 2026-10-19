import type { TableSnapshot, WorkbookContext } from '@mcsheety/core';
import type { McSheetyError } from '@mcsheety/core';
import { describe, expect, it } from 'vitest';
import { GraphWorkbookAdapter, parseAddress } from './adapter.js';
import { isAmbiguousWrite } from './http.js';

interface Recorded {
  method: string;
  url: string;
  headers: Record<string, string>;
  body: unknown;
}

type Reply = { status?: number; body?: unknown; headers?: Record<string, string> };
type Handler = (req: Recorded) => Reply;

const BASE = 'https://graph.microsoft.com/v1.0';
const ROOT = `${BASE}/drives/drive-1/items/item-1/workbook`;

function fakeGraph(routes: [method: string, match: string | RegExp, handler: Handler | Reply[]][]) {
  const calls: Recorded[] = [];
  const fetchImpl = (input: string | URL | Request, init?: RequestInit): Promise<Response> => {
    const req: Recorded = {
      method: init?.method ?? 'GET',
      url: input instanceof Request ? input.url : input.toString(),
      headers: (init?.headers ?? {}) as Record<string, string>,
      body: typeof init?.body === 'string' ? JSON.parse(init.body) : undefined,
    };
    calls.push(req);
    const route = routes.find(
      ([m, match]) =>
        m === req.method && (typeof match === 'string' ? req.url === match : match.test(req.url)),
    );
    if (!route)
      return Promise.resolve(
        new Response(JSON.stringify({ error: { code: 'itemNotFound' } }), { status: 404 }),
      );
    const handler = route[2];
    const reply = Array.isArray(handler) ? (handler.shift() ?? { status: 500 }) : handler(req);
    const status = reply.status ?? 200;
    return Promise.resolve(
      new Response(status === 204 ? null : JSON.stringify(reply.body ?? {}), {
        status,
        headers: reply.headers ?? {},
      }),
    );
  };
  return { calls, fetchImpl };
}

const ctx: WorkbookContext = {
  correlationId: 'corr-1',
  getGraphToken: () => Promise.resolve('graph-token'),
};
const sleeps: number[] = [];
const makeAdapter = (fetchImpl: typeof fetch, extra = {}) =>
  new GraphWorkbookAdapter({
    driveId: 'drive-1',
    itemId: 'item-1',
    fetch: fetchImpl,
    sleep: (ms) => {
      sleeps.push(ms);
      return Promise.resolve();
    },
    random: () => 0.5,
    ...extra,
  });

const snapshot: TableSnapshot = {
  table: 'tbl_Equipment',
  columns: ['id', 'name', 'status', 'locationId', 'departmentId', 'purchasePrice', 'createdAt'],
  rows: [
    ['eq-001', 'Pump A', 'Active', 'loc-1', 'dept-ops', 123.5, '2026-10-16T14:05:00Z'],
    ['eq-002', 'Valve B', 'Inactive', 'loc-2', 'dept-eng', 80, '2026-10-16T14:10:00Z'],
  ],
  location: { sheet: 'Equipment', startColumn: 0, startRow: 1 },
};

describe('parseAddress', () => {
  it('handles plain and quoted sheet names', () => {
    expect(parseAddress('Equipment!A1:G9')).toEqual({
      sheet: 'Equipment',
      startColumn: 0,
      startRow: 1,
    });
    expect(parseAddress("'Bob''s Sheet'!C4:D9")).toEqual({
      sheet: "Bob's Sheet",
      startColumn: 2,
      startRow: 4,
    });
  });
});

describe('GraphWorkbookAdapter reads', () => {
  it('reads a table snapshot with delegated token, correlation and session headers', async () => {
    const { calls, fetchImpl } = fakeGraph([
      [
        'GET',
        `${ROOT}/tables/tbl_Equipment/headerRowRange?$select=address,values`,
        () => ({
          body: { address: 'Equipment!B2:H2', values: [snapshot.columns] },
        }),
      ],
      [
        'GET',
        `${ROOT}/tables/tbl_Equipment/dataBodyRange?$select=address,values,formulas`,
        () => ({
          body: { address: 'Equipment!B3:H4', values: snapshot.rows, formulas: snapshot.rows },
        }),
      ],
    ]);
    const result = await makeAdapter(fetchImpl).readTable(
      { ...ctx, sessionId: 's-1' },
      'tbl_Equipment',
    );
    expect(result.location).toEqual({ sheet: 'Equipment', startColumn: 1, startRow: 2 });
    expect(result.rows).toHaveLength(2);
    expect(calls[0]?.headers).toMatchObject({
      Authorization: 'Bearer graph-token',
      'Workbook-Session-Id': 's-1',
      'client-request-id': 'corr-1',
    });
  });

  it('enforces the row limit', async () => {
    const { fetchImpl } = fakeGraph([
      ['GET', /headerRowRange/, () => ({ body: { address: 'S!A1:A1', values: [['id']] } })],
      [
        'GET',
        /dataBodyRange/,
        () => ({ body: { address: 'S!A2:A4', values: [['a'], ['b'], ['c']] } }),
      ],
    ]);
    await expect(makeAdapter(fetchImpl, { maxRows: 2 }).readTable(ctx, 't')).rejects.toMatchObject({
      code: 'WORKBOOK_UNAVAILABLE',
    });
  });

  it('lists tables across pages and refuses nextLinks that leave Graph', async () => {
    const { fetchImpl } = fakeGraph([
      [
        'GET',
        `${ROOT}/tables?$select=name`,
        () => ({
          body: { value: [{ name: 'tbl_A' }], '@odata.nextLink': `${ROOT}/tables?$skiptoken=2` },
        }),
      ],
      ['GET', `${ROOT}/tables?$skiptoken=2`, () => ({ body: { value: [{ name: 'mc_schema' }] } })],
      [
        'GET',
        /headerRowRange/,
        (r) => ({ body: { values: [[r.url.includes('tbl_A') ? 'id' : 'table']] } }),
      ],
    ]);
    const tables = await makeAdapter(fetchImpl).listTables(ctx);
    expect([...tables.keys()]).toEqual(['tbl_A', 'mc_schema']);

    const evil = fakeGraph([
      [
        'GET',
        `${ROOT}/tables?$select=name`,
        () => ({
          body: { value: [], '@odata.nextLink': 'https://attacker.example/steal' },
        }),
      ],
    ]);
    await expect(makeAdapter(evil.fetchImpl).listTables(ctx)).rejects.toMatchObject({
      code: 'GRAPH_ERROR',
    });
    expect(evil.calls.some((c) => c.url.includes('attacker'))).toBe(false);
  });
});

describe('GraphWorkbookAdapter writes', () => {
  const rowRange = `${ROOT}/worksheets/Equipment/range(address='A2:G2')?$select=values`;

  it('patches only changed cells, grouped into contiguous runs, after verifying the row', async () => {
    const after = ['eq-001', 'Pump Z', 'Inactive', 'loc-1', 'dept-ops', 99, '2026-10-16T14:05:00Z'];
    const { calls, fetchImpl } = fakeGraph([
      ['GET', rowRange, [{ body: { values: [snapshot.rows[0]] } }, { body: { values: [after] } }]],
      ['PATCH', /range\(address=/, () => ({ body: {} })],
    ]);
    const observed = await makeAdapter(fetchImpl).updateCells(
      ctx,
      snapshot,
      0,
      new Map<number, string | number>([
        [1, 'Pump Z'],
        [2, 'Inactive'],
        [5, 99],
      ]),
      ['string', 'string', 'enum', 'string', 'string', 'number', 'datetime'],
    );
    expect(observed).toEqual(after);
    const patches = calls.filter((c) => c.method === 'PATCH');
    expect(patches.map((p) => decodeURIComponent(p.url))).toEqual([
      `${ROOT}/worksheets/Equipment/range(address='B2:C2')`,
      `${ROOT}/worksheets/Equipment/range(address='F2:F2')`,
    ]);
    expect(patches.map((p) => p.body)).toEqual([
      { values: [["'Pump Z", "'Inactive"]] },
      { values: [[99]] },
    ]);
  });

  it('throws CONFLICT without writing when the row changed since the snapshot', async () => {
    const { calls, fetchImpl } = fakeGraph([
      ['GET', rowRange, () => ({ body: { values: [snapshot.rows[1]] } })],
    ]);
    await expect(
      makeAdapter(fetchImpl).updateCells(ctx, snapshot, 0, new Map([[2, 'Inactive']]), []),
    ).rejects.toMatchObject({ code: 'CONFLICT' });
    expect(calls.some((c) => c.method === 'PATCH')).toBe(false);
  });

  it('deletes the verified row with a shift-up range delete', async () => {
    const { calls, fetchImpl } = fakeGraph([
      ['GET', /A3:G3/, () => ({ body: { values: [snapshot.rows[1]] } })],
      ['POST', /\/delete$/, () => ({ status: 204 })],
    ]);
    await makeAdapter(fetchImpl).deleteRow(ctx, snapshot, 1);
    const del = calls.find((c) => c.method === 'POST');
    expect(decodeURIComponent(del?.url ?? '')).toBe(
      `${ROOT}/worksheets/Equipment/range(address='A3:G3')/delete`,
    );
    expect(del?.body).toEqual({ shift: 'Up' });
  });

  it('never retries appends; failures after send are flagged ambiguous', async () => {
    const { calls, fetchImpl } = fakeGraph([['POST', /rows\/add$/, () => ({ status: 503 })]]);
    const error = await makeAdapter(fetchImpl)
      .appendRow(ctx, 'tbl_Equipment', ['eq-9', 'X'], ['string', 'string'])
      .catch((e: unknown) => e);
    expect(calls).toHaveLength(1);
    expect(isAmbiguousWrite(error)).toBe(true);
    expect(calls[0]?.body).toEqual({ values: [["'eq-9", "'X"]] });
  });
});

describe('GraphHttp retry and error mapping', () => {
  it('retries throttled reads, honouring Retry-After', async () => {
    sleeps.length = 0;
    const { calls, fetchImpl } = fakeGraph([
      [
        'GET',
        /dataBodyRange/,
        [
          { status: 429, headers: { 'Retry-After': '2' } },
          { body: { address: 'S!A2:A2', values: [['x']] } },
        ],
      ],
      ['GET', /headerRowRange/, () => ({ body: { address: 'S!A1:A1', values: [['id']] } })],
    ]);
    const result = await makeAdapter(fetchImpl).readTable(ctx, 't');
    expect(result.rows).toEqual([['x']]);
    expect(sleeps).toEqual([2000]);
    expect(calls.filter((c) => c.url.includes('dataBodyRange'))).toHaveLength(2);
  });

  it('gives up with RATE_LIMITED after bounded attempts', async () => {
    const { calls, fetchImpl } = fakeGraph([
      ['GET', /headerRowRange/, () => ({ status: 429, headers: { 'Retry-After': '1' } })],
    ]);
    await expect(makeAdapter(fetchImpl).readTable(ctx, 't')).rejects.toMatchObject({
      code: 'RATE_LIMITED',
      retryAfterSeconds: 1,
    });
    expect(calls).toHaveLength(4);
  });

  it('maps SharePoint denial to a sanitized FORBIDDEN', async () => {
    const { fetchImpl } = fakeGraph([
      [
        'GET',
        /headerRowRange/,
        () => ({
          status: 403,
          body: {
            error: {
              code: 'accessDenied',
              message: 'User bob@contoso.com lacks access to /sites/hr/x.xlsx',
            },
          },
        }),
      ],
    ]);
    const error = (await makeAdapter(fetchImpl)
      .readTable(ctx, 't')
      .catch((e: unknown) => e)) as McSheetyError;
    expect(error.code).toBe('FORBIDDEN');
    expect(error.message).not.toContain('contoso');
  });
});

describe('withSession', () => {
  it('opens a persistent session, passes its id, and closes it even if closing fails', async () => {
    const { calls, fetchImpl } = fakeGraph([
      ['POST', `${ROOT}/createSession`, () => ({ status: 201, body: { id: 'sess-9' } })],
      ['POST', `${ROOT}/closeSession`, () => ({ status: 500 })],
    ]);
    const seen = await makeAdapter(fetchImpl, { retry: { maxAttempts: 1 } }).withSession(ctx, (c) =>
      Promise.resolve(c.sessionId),
    );
    expect(seen).toBe('sess-9');
    expect(calls[0]?.body).toEqual({ persistChanges: true });
    expect(calls[1]?.headers['Workbook-Session-Id']).toBe('sess-9');
  });
});
