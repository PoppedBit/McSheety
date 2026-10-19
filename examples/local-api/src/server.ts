/**
 * Local McSheety API server (development only; binds to 127.0.0.1).
 *
 *   MODE=demo   In-memory copy of fixtures/equipment and fixed demo users. No Entra, no Graph.
 *               Tokens are "demo:<user>" and are accepted ONLY by this server.
 *   MODE=entra  The real stack: Entra JWT verification, OBO and Microsoft Graph against the
 *               workbook in .env (MCSHEETY_CLIENT_SECRET). Used for the MC-002 spike.
 */
import { createServer } from 'node:http';
import { staticAttributes } from '@mcsheety/auth-msal';
import type { HttpHandler, HttpRequest } from '@mcsheety/api';
import {
  McSheetyService,
  createHttpHandler,
  createMcSheetyApp,
  jsonLogger,
  loadConfig,
  parseCredentialSecret,
} from '@mcsheety/api';
import { McSheetyError } from '@mcsheety/core';
import { InMemoryWorkbookAdapter } from '@mcsheety/core/testing';
import { equipmentWorkbook } from '../../../fixtures/equipment/index.js';
import { createProbeHandler } from './probe.js';

const PORT = Number(process.env['PORT'] ?? 3000);
const ORIGIN = process.env['MCSHEETY_DEV_ORIGIN'] ?? 'http://localhost:5173';
const MODE = process.env['MODE'] ?? 'demo';
const logger = jsonLogger();

export const DEMO_USERS: Record<string, { roles: string[]; departmentId?: string; label: string }> =
  {
    viewer: { roles: ['viewer'], label: 'Vera Viewer' },
    'editor-ops': {
      roles: ['editor'],
      departmentId: 'dept-ops',
      label: 'Eddie Editor (Operations)',
    },
    'editor-eng': {
      roles: ['editor'],
      departmentId: 'dept-eng',
      label: 'Erin Editor (Engineering)',
    },
    manager: { roles: ['manager'], departmentId: 'dept-eng', label: 'Morgan Manager' },
    admin: { roles: ['admin'], label: 'Ada Admin' },
  };

function demoHandler(): HttpHandler {
  const adapter = new InMemoryWorkbookAdapter(
    equipmentWorkbook.sheets.map((s) => ({
      sheet: s.sheet,
      table: s.table,
      columns: [...s.columns],
      rows: s.rows.map((r) => r.map((v) => v ?? '')),
    })),
  );
  return createHttpHandler({
    service: new McSheetyService({ adapter, audit: (e) => logger.info({ msg: 'audit', ...e }) }),
    verifyToken: (token) => {
      const user = token.startsWith('demo:') ? DEMO_USERS[token.slice(5)] : undefined;
      if (!user) return Promise.reject(new McSheetyError('UNAUTHENTICATED'));
      return Promise.resolve({
        tid: 'demo-tenant',
        oid: token.slice(5),
        scp: 'access_as_user',
        roles: user.roles,
      });
    },
    identity: { tenantId: 'demo-tenant', requiredScope: 'access_as_user' },
    graphToken: () => Promise.resolve('unused-in-demo'),
    attributes: staticAttributes(
      Object.fromEntries(
        Object.entries(DEMO_USERS)
          .filter(([, u]) => u.departmentId)
          .map(([id, u]) => [id, { departmentId: u.departmentId as string }]),
      ),
    ),
    logger,
  });
}

function entraHandler(): { api: HttpHandler; probe: HttpHandler } {
  const config = loadConfig();
  if (!config.clientSecret)
    throw new Error('MODE=entra needs MCSHEETY_CLIENT_SECRET in .env for local runs.');
  const credential = parseCredentialSecret(config.clientSecret);
  return {
    api: createMcSheetyApp({ config, credential, logger }),
    probe: createProbeHandler(config, credential),
  };
}

const handlers = MODE === 'entra' ? entraHandler() : { api: demoHandler(), probe: undefined };

const server = createServer((req, res) => {
  const chunks: Buffer[] = [];
  req.on('data', (c: Buffer) => chunks.push(c));
  req.on('end', () => {
    void (async () => {
      const url = new URL(req.url ?? '/', `http://127.0.0.1:${PORT}`);
      const cors: Record<string, string> = {
        'Access-Control-Allow-Origin': ORIGIN,
        'Access-Control-Allow-Headers': 'Authorization, Content-Type, If-Match',
        'Access-Control-Allow-Methods': 'GET, POST, PATCH, DELETE',
        'Access-Control-Expose-Headers': 'ETag, Retry-After, WWW-Authenticate, x-request-id',
        Vary: 'Origin',
      };
      if (req.method === 'OPTIONS') {
        res.writeHead(204, cors).end();
        return;
      }
      const headers: Record<string, string> = {};
      for (const [k, v] of Object.entries(req.headers)) if (typeof v === 'string') headers[k] = v;
      const request: HttpRequest = {
        method: req.method ?? 'GET',
        path: url.pathname,
        headers,
        query: Object.fromEntries(url.searchParams),
        body: chunks.length > 0 ? Buffer.concat(chunks).toString('utf8') : undefined,
      };
      const handler =
        url.pathname.startsWith('/dev/probe') && handlers.probe ? handlers.probe : handlers.api;
      const response = await handler(request);
      res.writeHead(response.status, { ...response.headers, ...cors }).end(response.body);
    })();
  });
});

server.listen(PORT, '127.0.0.1', () => {
  console.log(
    `McSheety local API (${MODE} mode) on http://127.0.0.1:${PORT} (CORS origin ${ORIGIN})`,
  );
  if (MODE === 'demo')
    console.log(`Demo users: ${Object.keys(DEMO_USERS).join(', ')} (token "demo:<user>")`);
});
