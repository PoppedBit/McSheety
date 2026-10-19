/**
 * Transport-agnostic HTTP layer for the REST API v1 (implementation.md §10). The Lambda
 * handler and the local dev server both translate to/from these shapes.
 */
import { createHash, randomUUID } from 'node:crypto';
import type { AttributeResolver, GraphTokenProvider, IdentityOptions } from '@mcsheety/auth-msal';
import { normalizeIdentity } from '@mcsheety/auth-msal';
import type { Identity } from '@mcsheety/core';
import { McSheetyError, toErrorResponse } from '@mcsheety/core';
import type { JWTPayload } from 'jose';
import type { Logger } from './logging.js';
import type { McSheetyService, RequestScope } from './service.js';

export interface HttpRequest {
  readonly method: string;
  /** Path without stage prefix, e.g. /v1/tables/Equipment/rows. */
  readonly path: string;
  /** Lower-cased header names. */
  readonly headers: Readonly<Record<string, string | undefined>>;
  readonly query: Readonly<Record<string, string | undefined>>;
  readonly body: string | undefined;
  /** Upstream request id (e.g. API Gateway), used as correlation id when present. */
  readonly requestId?: string;
}

export interface HttpResponse {
  readonly status: number;
  readonly headers: Record<string, string>;
  readonly body: string;
}

export interface HttpHandlerOptions {
  readonly service: McSheetyService;
  readonly verifyToken: (token: string) => Promise<JWTPayload>;
  readonly identity: IdentityOptions;
  readonly graphToken: GraphTokenProvider;
  readonly attributes: AttributeResolver;
  readonly logger: Logger;
  readonly maxBodyBytes?: number;
}

export type HttpHandler = (req: HttpRequest) => Promise<HttpResponse>;

const TABLE_RE = /^[A-Za-z][A-Za-z0-9_]{0,99}$/;
// eslint-disable-next-line no-control-regex -- deliberately rejects control characters in record ids
const ID_RE = /^[^\u0000-\u001f]{1,256}$/;

const BASE_HEADERS = {
  'Content-Type': 'application/json; charset=utf-8',
  'Cache-Control': 'no-store',
  'X-Content-Type-Options': 'nosniff',
};

type RouteName =
  | 'health'
  | 'tables'
  | 'schema'
  | 'capabilities'
  | 'listRows'
  | 'getRow'
  | 'createRow'
  | 'updateRow'
  | 'deleteRow';

interface Route {
  readonly name: RouteName;
  readonly table?: string;
  readonly id?: string;
}

function matchRoute(method: string, path: string): Route | 'method' | undefined {
  const parts = path
    .replace(/\/+$/, '')
    .split('/')
    .slice(1)
    .map((p) => {
      try {
        return decodeURIComponent(p);
      } catch {
        return '\u0000'; // invalid escape -> fails validation below
      }
    });
  if (parts[0] !== 'v1') return undefined;
  const [, a, b, c, d, extra] = parts;
  const want = (m: string, route: Route): Route | 'method' => (method === m ? route : 'method');
  const table = (t: string | undefined) => {
    if (t === undefined || !TABLE_RE.test(t)) throw new McSheetyError('NOT_FOUND');
    return t;
  };
  if (extra !== undefined) return undefined;
  if (a === 'health' && b === undefined) return want('GET', { name: 'health' });
  if (a === 'capabilities' && b === undefined) return want('GET', { name: 'capabilities' });
  if (a === 'tables' && b === undefined) return want('GET', { name: 'tables' });
  if (a === 'schema' && b !== undefined && c === undefined)
    return want('GET', { name: 'schema', table: table(b) });
  if (a === 'tables' && b !== undefined && c === 'rows') {
    if (d === undefined) {
      if (method === 'GET') return { name: 'listRows', table: table(b) };
      return want('POST', { name: 'createRow', table: table(b) });
    }
    if (!ID_RE.test(d)) throw new McSheetyError('NOT_FOUND');
    const route = { table: table(b), id: d };
    if (method === 'GET') return { name: 'getRow', ...route };
    if (method === 'PATCH') return { name: 'updateRow', ...route };
    return want('DELETE', { name: 'deleteRow', ...route });
  }
  return undefined;
}

function bearer(headers: HttpRequest['headers']): string {
  const value = headers['authorization'];
  const m = value ? /^Bearer\s+(\S+)$/i.exec(value) : null;
  if (!m?.[1]) throw new McSheetyError('UNAUTHENTICATED');
  return m[1];
}

/** Stable, non-reversible actor reference for logs (not for authorization). */
export function actorRef(identity: Pick<Identity, 'tenantId' | 'userId'>): string {
  return createHash('sha256')
    .update(`${identity.tenantId}:${identity.userId}`)
    .digest('hex')
    .slice(0, 16);
}

export function createHttpHandler(options: HttpHandlerOptions): HttpHandler {
  const maxBody = options.maxBodyBytes ?? 64 * 1024;

  function parseBody(req: HttpRequest): unknown {
    const type = req.headers['content-type'] ?? '';
    if (!/^application\/json\b/i.test(type)) {
      throw new McSheetyError('INVALID_REQUEST', 'Content-Type must be application/json.');
    }
    if (!req.body || Buffer.byteLength(req.body) > maxBody) {
      throw new McSheetyError(
        'INVALID_REQUEST',
        `The body must be JSON of at most ${maxBody} bytes.`,
      );
    }
    try {
      return JSON.parse(req.body);
    } catch {
      throw new McSheetyError('INVALID_REQUEST', 'The body is not valid JSON.');
    }
  }

  async function authenticate(req: HttpRequest, correlationId: string): Promise<RequestScope> {
    const token = bearer(req.headers);
    const verified = normalizeIdentity(await options.verifyToken(token), options.identity);
    let graphToken: Promise<string> | undefined;
    const getGraphToken = () => (graphToken ??= options.graphToken(token, correlationId));
    const attributes = await options.attributes.resolve({
      identity: verified,
      correlationId,
      getGraphToken,
    });
    return {
      identity: { ...verified, attributes },
      ctx: { correlationId, getGraphToken },
    };
  }

  return async (req) => {
    const started = Date.now();
    const correlationId = req.requestId ?? randomUUID();
    const headers: Record<string, string> = { ...BASE_HEADERS, 'x-request-id': correlationId };
    let route: Route | undefined;
    let actor: string | undefined;

    const respond = (
      status: number,
      payload?: unknown,
      extra: Record<string, string> = {},
    ): HttpResponse => {
      options.logger.info({
        msg: 'request',
        requestId: correlationId,
        method: req.method,
        route: route?.name,
        table: route?.table,
        actor,
        status,
        durationMs: Date.now() - started,
      });
      return {
        status,
        headers: { ...headers, ...extra },
        body: payload === undefined ? '' : JSON.stringify(payload),
      };
    };

    try {
      const matched = matchRoute(req.method, req.path);
      if (matched === undefined) throw new McSheetyError('NOT_FOUND');
      if (matched === 'method')
        throw new McSheetyError('INVALID_REQUEST', 'Method not allowed for this route.');
      route = matched;
      if (route.name === 'health') return respond(200, { status: 'ok' });

      const scope = await authenticate(req, correlationId);
      actor = actorRef(scope.identity);
      const { service } = options;
      const t = route.table as string;
      const id = route.id as string;

      switch (route.name) {
        case 'tables':
          return respond(200, await service.listTables(scope));
        case 'capabilities':
          return respond(200, await service.capabilities(scope));
        case 'schema':
          return respond(200, await service.schema(scope, t));
        case 'listRows':
          return respond(
            200,
            await service.listRows(scope, t, {
              where: req.query['where'],
              sort: req.query['sort'],
              limit: req.query['limit'],
              cursor: req.query['cursor'],
            }),
          );
        case 'getRow': {
          const record = await service.getRow(scope, t, id);
          return respond(200, { data: record }, { ETag: `"${record.revision}"` });
        }
        case 'createRow': {
          const record = await service.createRow(scope, t, parseBody(req));
          return respond(201, { data: record }, { ETag: `"${record.revision}"` });
        }
        case 'updateRow': {
          const record = await service.updateRow(
            scope,
            t,
            id,
            parseBody(req),
            req.headers['if-match'],
          );
          return respond(200, { data: record }, { ETag: `"${record.revision}"` });
        }
        case 'deleteRow':
          await service.deleteRow(scope, t, id, req.headers['if-match']);
          return respond(204);
      }
    } catch (error) {
      const { status, body, retryAfterSeconds } = toErrorResponse(error, correlationId);
      if (status >= 500 || !(error instanceof McSheetyError)) {
        options.logger.error({
          msg: 'request failed',
          requestId: correlationId,
          route: route?.name,
          errorCode: body.error.code,
          // Safe diagnostics only: the cause carries Graph status/codes, never tokens or bodies.
          cause:
            error instanceof McSheetyError ? describeCause(error.cause) : (error as Error)?.name,
        });
      }
      const extra: Record<string, string> = {};
      if (retryAfterSeconds !== undefined)
        extra['Retry-After'] = String(Math.ceil(retryAfterSeconds));
      if (body.error.claimsChallenge) {
        extra['WWW-Authenticate'] =
          `Bearer error="insufficient_claims", claims="${btoa(body.error.claimsChallenge)}"`;
      }
      return respond(status, body, extra);
    }
  };
}

function describeCause(cause: unknown): unknown {
  if (typeof cause !== 'object' || cause === null) return undefined;
  const c = cause as Record<string, unknown>;
  const pick = [
    'status',
    'graphCode',
    'innerCode',
    'graphRequestId',
    'ambiguous',
    'errorCode',
    'errorNo',
  ];
  return Object.fromEntries(pick.filter((k) => c[k] !== undefined).map((k) => [k, c[k]]));
}
