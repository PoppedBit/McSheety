import { describe, expect, it } from 'vitest';
import type { TokenRequest } from './index.js';
import { McSheetyClientError, createClient } from './index.js';

interface Sent {
  method: string;
  url: string;
  headers: Record<string, string>;
  body?: unknown;
}

function fakeApi(replies: { status: number; body?: unknown; headers?: Record<string, string> }[]) {
  const sent: Sent[] = [];
  const fetchImpl = ((url: string, init: RequestInit) => {
    sent.push({
      method: init.method ?? 'GET',
      url,
      headers: init.headers as Record<string, string>,
      ...(typeof init.body === 'string' ? { body: JSON.parse(init.body) } : {}),
    });
    const r = replies.shift() ?? { status: 500 };
    return Promise.resolve(
      new Response(r.status === 204 ? null : JSON.stringify(r.body ?? {}), {
        status: r.status,
        headers: r.headers ?? {},
      }),
    );
  }) as typeof fetch;
  return { sent, fetchImpl };
}

type Equipment = { id: string; name: string; status: 'Active' | 'Inactive'; locationId: string };
const envelope = (fields: Partial<Equipment>, revision = 'rev-1') => ({
  id: fields.id,
  revision,
  fields,
});

describe('McSheetyClient', () => {
  it('sends the API token and builds list queries', async () => {
    const { sent, fetchImpl } = fakeApi([
      {
        status: 200,
        body: {
          data: [envelope({ id: 'eq-001', name: 'Pump A' })],
          page: { nextCursor: 'c2' },
          meta: { table: 'Equipment' },
        },
      },
    ]);
    const db = createClient({
      apiUrl: 'https://api.example/',
      getToken: () => Promise.resolve('api-token'),
      fetch: fetchImpl,
    });
    const page = await db.table<Equipment>('Equipment').findMany({
      where: { status: 'Active', locationId: { in: ['loc-1'] } },
      sort: ['-name'],
      limit: 50,
    });
    expect(page.records[0]).toEqual({
      id: 'eq-001',
      revision: 'rev-1',
      fields: { id: 'eq-001', name: 'Pump A' },
    });
    expect(page.nextCursor).toBe('c2');
    const url = new URL(sent[0]?.url ?? '');
    expect(url.pathname).toBe('/v1/tables/Equipment/rows');
    expect(JSON.parse(url.searchParams.get('where') ?? '')).toEqual({
      status: 'Active',
      locationId: { in: ['loc-1'] },
    });
    expect(url.searchParams.get('sort')).toBe('-name');
    expect(sent[0]?.headers['Authorization']).toBe('Bearer api-token');
  });

  it('sends If-Match on update and delete, and keeps revision outside fields', async () => {
    const { sent, fetchImpl } = fakeApi([
      { status: 200, body: { data: envelope({ id: 'eq-001', status: 'Inactive' }, 'rev-2') } },
      { status: 204 },
    ]);
    const table = createClient({
      apiUrl: 'https://api.example',
      getToken: () => Promise.resolve('t'),
      fetch: fetchImpl,
    }).table<Equipment>('Equipment');
    const updated = await table.update('eq-001', { status: 'Inactive' }, { ifMatch: 'rev-1' });
    expect(updated.revision).toBe('rev-2');
    expect(updated.fields).not.toHaveProperty('revision');
    await table.delete('eq/002', { ifMatch: 'rev-9' });
    expect(sent.map((s) => [s.method, s.headers['If-Match']])).toEqual([
      ['PATCH', '"rev-1"'],
      ['DELETE', '"rev-9"'],
    ]);
    expect(sent[0]?.body).toEqual({ status: 'Inactive' });
    expect(sent[1]?.url).toBe('https://api.example/v1/tables/Equipment/rows/eq%2F002');
  });

  it('raises typed errors with request ids and conflict detection', async () => {
    const { fetchImpl } = fakeApi([
      {
        status: 409,
        body: {
          error: {
            code: 'CONFLICT',
            message: 'The record changed since it was read.',
            requestId: 'r-1',
            retryable: false,
          },
        },
      },
    ]);
    const table = createClient({
      apiUrl: 'https://api.example',
      getToken: () => Promise.resolve('t'),
      fetch: fetchImpl,
    }).table('Equipment');
    const error = (await table
      .update('eq-001', { name: 'x' }, { ifMatch: 'old' })
      .catch((e: unknown) => e)) as McSheetyClientError;
    expect(error).toBeInstanceOf(McSheetyClientError);
    expect(error.isConflict).toBe(true);
    expect(error.requestId).toBe('r-1');
  });

  it('retries once with the claims challenge for Conditional Access', async () => {
    const requests: (TokenRequest | undefined)[] = [];
    const { fetchImpl } = fakeApi([
      {
        status: 401,
        body: {
          error: {
            code: 'UNAUTHENTICATED',
            message: 'x',
            requestId: 'r',
            retryable: false,
            claimsChallenge: '{"c":1}',
          },
        },
      },
      { status: 200, body: { data: [{ name: 'Equipment' }] } },
    ]);
    const db = createClient({
      apiUrl: 'https://api.example',
      getToken: (req) => {
        requests.push(req);
        return Promise.resolve('t');
      },
      fetch: fetchImpl,
    });
    expect(await db.tables()).toEqual(['Equipment']);
    expect(requests).toEqual([undefined, { claims: '{"c":1}' }]);
  });

  it('retries throttled reads but never writes', async () => {
    const sleeps: number[] = [];
    const throttled = {
      status: 429,
      body: {
        error: { code: 'RATE_LIMITED', message: 'slow down', requestId: 'r', retryable: true },
      },
      headers: { 'Retry-After': '1' },
    };
    const { sent, fetchImpl } = fakeApi([
      throttled,
      { status: 200, body: { data: [] } },
      throttled,
    ]);
    const db = createClient({
      apiUrl: 'https://api.example',
      getToken: () => Promise.resolve('t'),
      fetch: fetchImpl,
      sleep: (ms) => {
        sleeps.push(ms);
        return Promise.resolve();
      },
    });
    await db.tables();
    expect(sleeps).toEqual([1000]);
    await expect(db.table('Equipment').create({ name: 'x' })).rejects.toMatchObject({
      code: 'RATE_LIMITED',
      retryAfterSeconds: 1,
    });
    expect(sent).toHaveLength(3);
  });

  it('reports network failures distinctly', async () => {
    const db = createClient({
      apiUrl: 'https://api.example',
      getToken: () => Promise.resolve('t'),
      fetch: () => Promise.reject(new TypeError('offline')),
    });
    await expect(db.tables()).rejects.toMatchObject({ code: 'NETWORK_ERROR', status: 0 });
  });
});
