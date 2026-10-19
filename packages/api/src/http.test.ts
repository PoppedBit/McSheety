import type { APIGatewayProxyEventV2 } from 'aws-lambda';
import { McSheetyError } from '@mcsheety/core';
import { staticAttributes } from '@mcsheety/auth-msal';
import { describe, expect, it } from 'vitest';
import { memoryWorkbook } from './__fixtures__/memory.js';
import { ConfigError, loadConfig, parseCredentialSecret } from './config.js';
import type { HttpRequest } from './http.js';
import { createHttpHandler } from './http.js';
import { fromLambdaEvent } from './lambda.js';
import type { LogFields } from './logging.js';
import { jsonLogger } from './logging.js';
import { McSheetyService } from './service.js';

const TENANT = 'tenant-1';
const tokens: Record<string, Record<string, unknown>> = {
  'tok-viewer': { tid: TENANT, oid: 'u-viewer', scp: 'access_as_user', roles: ['viewer'] },
  'tok-editor': { tid: TENANT, oid: 'u-ed-ops', scp: 'access_as_user', roles: ['editor'] },
  'tok-app': { tid: TENANT, oid: 'svc', idtyp: 'app', roles: ['admin'] },
};

function setup(overrides: { graphToken?: () => Promise<string> } = {}) {
  const lines: string[] = [];
  const adapter = memoryWorkbook();
  const handler = createHttpHandler({
    service: new McSheetyService({ adapter }),
    verifyToken: (t) => {
      const claims = tokens[t];
      return claims
        ? Promise.resolve(claims)
        : Promise.reject(new McSheetyError('UNAUTHENTICATED'));
    },
    identity: { tenantId: TENANT, requiredScope: 'access_as_user' },
    graphToken: overrides.graphToken ?? (() => Promise.resolve('graph-token')),
    attributes: staticAttributes({ 'u-ed-ops': { departmentId: 'dept-ops' } }),
    logger: jsonLogger((l) => lines.push(l)),
  });
  const call = (req: Partial<HttpRequest> & { path: string }, token?: string) =>
    handler({
      method: 'GET',
      headers: token ? { authorization: `Bearer ${token}` } : {},
      query: {},
      body: undefined,
      ...req,
    });
  return { call, lines, adapter };
}

const json = (body: string) =>
  JSON.parse(body) as Record<string, unknown> & { error?: { code: string } };

describe('HTTP routing and authentication', () => {
  it('serves health without authentication or workbook access', async () => {
    const { call, adapter } = setup();
    const res = await call({ path: '/v1/health' });
    expect(res.status).toBe(200);
    expect(adapter.calls).toEqual([]);
  });

  it('rejects missing, invalid and app-only tokens with the standard error shape', async () => {
    const { call } = setup();
    for (const [token, status] of [
      [undefined, 401],
      ['forged', 401],
      ['tok-app', 403],
    ] as const) {
      const res = await call({ path: '/v1/tables' }, token);
      expect(res.status).toBe(status);
      const body = json(res.body);
      expect(body.error).toMatchObject({
        requestId: res.headers['x-request-id'],
        retryable: false,
      });
    }
  });

  it('lists rows for a valid user with safe headers', async () => {
    const { call } = setup();
    const res = await call(
      { path: '/v1/tables/Equipment/rows', query: { limit: '2', sort: 'name' } },
      'tok-viewer',
    );
    expect(res.status).toBe(200);
    expect(res.headers).toMatchObject({
      'Cache-Control': 'no-store',
      'X-Content-Type-Options': 'nosniff',
    });
    const body = json(res.body) as { data: { id: string }[]; page: { nextCursor: string | null } };
    expect(body.data.map((r) => r.id)).toEqual(['eq-003', 'eq-005']); // Compressor C, Forklift E
    expect(body.page.nextCursor).toEqual(expect.any(String));
  });

  it('round-trips ETag / If-Match for updates', async () => {
    const { call } = setup();
    const got = await call({ path: '/v1/tables/Equipment/rows/eq%2D001' }, 'tok-editor');
    expect(got.status).toBe(200);
    const etag = got.headers['ETag'] as string;
    const patch = (ifMatch?: string) =>
      call({
        method: 'PATCH',
        path: '/v1/tables/Equipment/rows/eq-001',
        headers: {
          authorization: 'Bearer tok-editor',
          'content-type': 'application/json',
          ...(ifMatch ? { 'if-match': ifMatch } : {}),
        },
        body: JSON.stringify({ status: 'Inactive' }),
      });
    expect((await patch()).status).toBe(428);
    const ok = await patch(etag);
    expect(ok.status).toBe(200);
    expect(ok.headers['ETag']).not.toBe(etag);
    expect((await patch(etag)).status).toBe(409);
  });

  it('validates bodies and routes', async () => {
    const { call } = setup();
    const post = (headers: Record<string, string>, body: string) =>
      call({
        method: 'POST',
        path: '/v1/tables/Equipment/rows',
        headers: { authorization: 'Bearer tok-editor', ...headers },
        body,
      });
    expect((await post({}, '{}')).status).toBe(400);
    expect((await post({ 'content-type': 'application/json' }, '{nope')).status).toBe(400);
    expect(
      (
        await post(
          { 'content-type': 'application/json' },
          JSON.stringify({ name: 'x'.repeat(70_000) }),
        )
      ).status,
    ).toBe(400);
    expect((await call({ path: '/v1/nope' }, 'tok-viewer')).status).toBe(404);
    expect((await call({ path: '/v1/tables/../../etc/rows' }, 'tok-viewer')).status).toBe(404);
    expect((await call({ method: 'PUT', path: '/v1/tables' }, 'tok-viewer')).status).toBe(400);
    const created = await post(
      { 'content-type': 'application/json' },
      JSON.stringify({ name: 'Pump Q', locationId: 'loc-3' }),
    );
    expect(created.status).toBe(201);
  });

  it('passes Conditional Access challenges and throttling hints to the client', async () => {
    const challenge = setup({
      graphToken: () =>
        Promise.reject(
          new McSheetyError('UNAUTHENTICATED', 'Additional sign-in is required.', {
            claimsChallenge: '{"x":1}',
          }),
        ),
    });
    const res = await challenge.call({ path: '/v1/tables' }, 'tok-editor');
    expect(res.status).toBe(401);
    expect(res.headers['WWW-Authenticate']).toContain('insufficient_claims');
    expect(json(res.body).error).toMatchObject({ claimsChallenge: '{"x":1}' });

    const throttled = setup({
      graphToken: () =>
        Promise.reject(new McSheetyError('RATE_LIMITED', undefined, { retryAfterSeconds: 2.5 })),
    });
    const limited = await throttled.call({ path: '/v1/tables' }, 'tok-editor');
    expect(limited.status).toBe(429);
    expect(limited.headers['Retry-After']).toBe('3');
  });

  it('never logs bearer tokens and hides internal error messages', async () => {
    const { call, lines } = setup({
      graphToken: () => Promise.reject(new Error('boom tok-editor secret')),
    });
    const res = await call({ path: '/v1/tables' }, 'tok-editor');
    expect(res.status).toBe(500);
    expect(res.body).not.toContain('boom');
    const logs = lines.join('\n');
    expect(logs).not.toContain('tok-editor');
    expect(logs).not.toContain('boom');
    const entries = lines.map((l) => JSON.parse(l) as LogFields);
    expect(entries.find((e) => e['msg'] === 'request')).toMatchObject({
      status: 500,
      route: 'tables',
    });
  });
});

describe('fromLambdaEvent', () => {
  it('normalizes headers, strips a named stage and decodes base64 bodies', () => {
    const event = {
      rawPath: '/prod/v1/tables/Equipment/rows',
      headers: { Authorization: 'Bearer t', 'Content-Type': 'application/json' },
      queryStringParameters: { limit: '5' },
      body: Buffer.from('{"a":1}').toString('base64'),
      isBase64Encoded: true,
      requestContext: { stage: 'prod', requestId: 'req-9', http: { method: 'post' } },
    } as unknown as APIGatewayProxyEventV2;
    expect(fromLambdaEvent(event)).toEqual({
      method: 'POST',
      path: '/v1/tables/Equipment/rows',
      headers: { authorization: 'Bearer t', 'content-type': 'application/json' },
      query: { limit: '5' },
      body: '{"a":1}',
      requestId: 'req-9',
    });
  });
});

describe('loadConfig', () => {
  const valid = {
    MCSHEETY_TENANT_ID: '11111111-1111-1111-1111-111111111111',
    MCSHEETY_API_CLIENT_ID: '22222222-2222-2222-2222-222222222222',
    MCSHEETY_DRIVE_ID: 'b!drive',
    MCSHEETY_ITEM_ID: '01ITEM',
    MCSHEETY_CREDENTIAL_SECRET_ARN: 'arn:aws:secretsmanager:us-east-1:123:secret:x',
  };

  it('accepts a complete configuration with defaults', () => {
    const config = loadConfig(valid);
    expect(config).toMatchObject({
      apiScope: 'access_as_user',
      policyMode: 'co-located',
      maxRows: 2000,
    });
  });

  it('lists every problem at once', () => {
    expect(() =>
      loadConfig({ MCSHEETY_TENANT_ID: 'contoso', MCSHEETY_POLICY_MODE: 'separate' }),
    ).toThrow(ConfigError);
    try {
      loadConfig({ MCSHEETY_TENANT_ID: 'contoso', MCSHEETY_POLICY_MODE: 'separate' });
    } catch (e) {
      const message = (e as Error).message;
      for (const part of [
        'MCSHEETY_TENANT_ID must be a GUID',
        'MCSHEETY_API_CLIENT_ID is required',
        'MCSHEETY_POLICY_MODE',
        'MCSHEETY_CREDENTIAL_SECRET_ARN',
      ]) {
        expect(message).toContain(part);
      }
    }
  });

  it('rejects malformed JSON settings', () => {
    expect(() => loadConfig({ ...valid, MCSHEETY_ATTRIBUTES: '{bad' })).toThrow(
      'MCSHEETY_ATTRIBUTES must be valid JSON',
    );
  });

  it('parses credential secrets', () => {
    expect(parseCredentialSecret('s3cret')).toEqual({ clientSecret: 's3cret' });
    expect(parseCredentialSecret('{"clientSecret":"x"}')).toEqual({ clientSecret: 'x' });
    expect(
      parseCredentialSecret('{"certificate":{"thumbprintSha256":"ab","privateKey":"pk"}}'),
    ).toEqual({
      certificate: { thumbprintSha256: 'ab', privateKey: 'pk' },
    });
    expect(() => parseCredentialSecret('{}')).toThrow(ConfigError);
  });
});
