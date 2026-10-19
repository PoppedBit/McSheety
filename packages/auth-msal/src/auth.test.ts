import { InteractionRequiredAuthError, ServerError } from '@azure/msal-node';
import type { McSheetyError } from '@mcsheety/core';
import { toErrorResponse } from '@mcsheety/core';
import { SignJWT, createLocalJWKSet, exportJWK, generateKeyPair } from 'jose';
import { describe, expect, it } from 'vitest';
import { graphProfileAttributes, staticAttributes } from './attributes.js';
import { normalizeIdentity } from './identity.js';
import { createAccessTokenVerifier } from './jwt.js';
import { createGraphTokenProvider, mapOboError } from './obo.js';

const TENANT = '11111111-1111-1111-1111-111111111111';
const API = '22222222-2222-2222-2222-222222222222';

describe('createAccessTokenVerifier', async () => {
  const { privateKey, publicKey } = await generateKeyPair('RS256');
  const jwk = { ...(await exportJWK(publicKey)), kid: 'k1', alg: 'RS256' };
  const verify = createAccessTokenVerifier({
    tenantId: TENANT,
    apiClientId: API,
    keys: createLocalJWKSet({ keys: [jwk] }),
  });

  const sign = (
    claims: Record<string, unknown>,
    opts: { iss?: string; aud?: string; exp?: string } = {},
  ) =>
    new SignJWT(claims)
      .setProtectedHeader({ alg: 'RS256', kid: 'k1' })
      .setIssuer(opts.iss ?? `https://login.microsoftonline.com/${TENANT}/v2.0`)
      .setAudience(opts.aud ?? API)
      .setIssuedAt()
      .setExpirationTime(opts.exp ?? '5m')
      .sign(privateKey);

  it('accepts a valid API token', async () => {
    const payload = await verify(await sign({ tid: TENANT, oid: 'u1', scp: 'access_as_user' }));
    expect(payload['oid']).toBe('u1');
  });

  it('rejects wrong audience (e.g. a Graph token), wrong issuer and expired tokens', async () => {
    for (const token of [
      await sign({}, { aud: '00000003-0000-0000-c000-000000000000' }),
      await sign({}, { iss: `https://sts.windows.net/${TENANT}/` }),
      await sign({}, { iss: 'https://login.microsoftonline.com/other-tenant/v2.0' }),
      await sign({}, { exp: '-10m' }),
    ]) {
      await expect(verify(token)).rejects.toMatchObject({ code: 'UNAUTHENTICATED' });
    }
  });

  it('rejects tokens signed by an unknown key and unsigned tokens', async () => {
    const other = await generateKeyPair('RS256');
    const forged = await new SignJWT({ tid: TENANT })
      .setProtectedHeader({ alg: 'RS256', kid: 'k1' })
      .setIssuer(`https://login.microsoftonline.com/${TENANT}/v2.0`)
      .setAudience(API)
      .setExpirationTime('5m')
      .sign(other.privateKey);
    await expect(verify(forged)).rejects.toMatchObject({ code: 'UNAUTHENTICATED' });
    const unsigned = `${btoa(JSON.stringify({ alg: 'none' }))}.${btoa(JSON.stringify({ tid: TENANT }))}.`;
    await expect(verify(unsigned)).rejects.toMatchObject({ code: 'UNAUTHENTICATED' });
  });
});

describe('normalizeIdentity', () => {
  const opts = { tenantId: TENANT, requiredScope: 'access_as_user' };
  const base = { tid: TENANT, oid: 'user-1', scp: 'access_as_user', roles: ['editor'] };

  it('uses tid + oid and token app roles', () => {
    expect(normalizeIdentity({ ...base, preferred_username: 'x@y.com' }, opts)).toEqual({
      tenantId: TENANT,
      userId: 'user-1',
      roles: ['editor'],
    });
  });

  it('rejects other tenants, app-only tokens and missing scope', () => {
    expect(() => normalizeIdentity({ ...base, tid: 'evil' }, opts)).toThrow(
      expect.objectContaining({ code: 'UNAUTHENTICATED' }) as Error,
    );
    expect(() => normalizeIdentity({ ...base, scp: undefined, idtyp: 'app' }, opts)).toThrow(
      expect.objectContaining({ code: 'FORBIDDEN' }) as Error,
    );
    expect(() => normalizeIdentity({ ...base, scp: 'User.Read' }, opts)).toThrow(
      expect.objectContaining({ code: 'FORBIDDEN' }) as Error,
    );
    expect(() => normalizeIdentity({ ...base, oid: undefined }, opts)).toThrow(
      expect.objectContaining({ code: 'UNAUTHENTICATED' }) as Error,
    );
  });

  it('maps only configured immutable group IDs to roles', () => {
    const identity = normalizeIdentity(
      { ...base, roles: undefined, groups: ['g-1', 'g-unknown'] },
      { ...opts, groupRoles: { 'g-1': 'manager' } },
    );
    expect(identity.roles).toEqual(['manager']);
  });
});

describe('OBO token provider', () => {
  it('passes the inbound API token as the user assertion and returns the Graph token', async () => {
    const seen: unknown[] = [];
    const provider = createGraphTokenProvider({
      tenantId: TENANT,
      clientId: API,
      credential: { clientSecret: 'unused-in-test' },
      client: {
        acquireTokenOnBehalfOf: (req) => {
          seen.push(req);
          return Promise.resolve({ accessToken: 'graph-token' });
        },
      },
    });
    await expect(provider('api-token', 'corr-1')).resolves.toBe('graph-token');
    expect(seen).toEqual([
      {
        oboAssertion: 'api-token',
        scopes: ['https://graph.microsoft.com/.default'],
        correlationId: 'corr-1',
      },
    ]);
  });

  it('turns a Conditional Access challenge into a 401 carrying the claims', () => {
    const error = mapOboError(
      new InteractionRequiredAuthError(
        'interaction_required',
        'c',
        'AADSTS50076',
        '',
        '',
        '',
        '{"access_token":{"acrs":{"essential":true}}}',
      ),
    );
    const response = toErrorResponse(error, 'r');
    expect(response.status).toBe(401);
    expect(response.body.error.claimsChallenge).toContain('acrs');
  });

  it('distinguishes missing admin consent from other failures', () => {
    const consent = mapOboError(
      new ServerError(
        'invalid_grant',
        'corr',
        'redacted by MSAL',
        'consent_required',
        '65001',
        400,
      ),
    );
    expect(consent.code).toBe('FORBIDDEN');
    expect(consent.details?.[0]?.code).toBe('consent_required');
    const other = mapOboError(new Error('socket hang up'));
    expect(other.code).toBe('GRAPH_ERROR');
    expect(JSON.stringify(toErrorResponse(other, 'r'))).not.toContain('socket');
  });
});

describe('attribute resolvers', () => {
  const identity = { tenantId: TENANT, userId: 'user-1', roles: ['editor'] };
  const ctx = { identity, correlationId: 'c', getGraphToken: () => Promise.resolve('graph-token') };

  it('static resolver keys by object id', async () => {
    await expect(
      staticAttributes({ 'user-1': { departmentId: 'dept-ops' } }).resolve(ctx),
    ).resolves.toEqual({
      departmentId: 'dept-ops',
    });
    await expect(staticAttributes({}).resolve(ctx)).resolves.toEqual({});
  });

  it('graph resolver maps admin-managed profile values and fails closed on unmapped ones', async () => {
    const urls: string[] = [];
    const fake = (department: string): typeof fetch =>
      ((url: string) => {
        urls.push(url);
        return Promise.resolve(new Response(JSON.stringify({ department })));
      }) as typeof fetch;
    const make = (department: string) =>
      graphProfileAttributes({
        properties: { departmentId: 'department' },
        valueMaps: { departmentId: { Operations: 'dept-ops' } },
        fetch: fake(department),
      });
    await expect(make('Operations').resolve(ctx)).resolves.toEqual({ departmentId: 'dept-ops' });
    await expect(make('Marketing').resolve(ctx)).resolves.toEqual({});
    expect(urls[0]).toBe('https://graph.microsoft.com/v1.0/me?$select=department');
  });

  it('refuses user-editable or sensitive profile properties', () => {
    expect(() => graphProfileAttributes({ properties: { x: 'aboutMe' } })).toThrow();
  });

  it('surfaces Graph failures as sanitized errors', async () => {
    const resolver = graphProfileAttributes({
      properties: { departmentId: 'department' },
      fetch: () => Promise.resolve(new Response('nope', { status: 403 })),
    });
    const error = (await resolver.resolve(ctx).catch((e: unknown) => e)) as McSheetyError;
    expect(error.code).toBe('GRAPH_ERROR');
  });
});
