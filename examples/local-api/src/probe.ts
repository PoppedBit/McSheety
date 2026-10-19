/**
 * MC-002 spike probe (local entra mode only): exercises the Graph behaviours McSheety relies on
 * but cannot verify with mocks, against a dedicated scratch table `tbl_McSheetyProbe`
 * (columns: id, text, num) in the configured workbook. Requires the `admin` app role.
 *
 *   GET /dev/probe   -> JSON report; paste it into docs/decisions/0002-graph-workbook-behaviour.md
 */
import { randomUUID } from 'node:crypto';
import { GraphWorkbookAdapter } from '@mcsheety/adapter-graph';
import type { McSheetyConfig, HttpHandler } from '@mcsheety/api';
import type { OboCredential } from '@mcsheety/auth-msal';
import {
  createAccessTokenVerifier,
  createGraphTokenProvider,
  normalizeIdentity,
} from '@mcsheety/auth-msal';
import type { CellValue, WorkbookContext } from '@mcsheety/core';
import { McSheetyError, toErrorResponse } from '@mcsheety/core';

const PROBE_TABLE = 'tbl_McSheetyProbe';

type Step = { step: string; ok: boolean; observed?: unknown; note?: string | undefined };

export function createProbeHandler(config: McSheetyConfig, credential: OboCredential): HttpHandler {
  const verify = createAccessTokenVerifier({
    tenantId: config.tenantId,
    apiClientId: config.apiClientId,
  });
  const obo = createGraphTokenProvider({
    tenantId: config.tenantId,
    clientId: config.apiClientId,
    credential,
  });
  const adapter = (textEncoding: 'apostrophe' | 'raw') =>
    new GraphWorkbookAdapter({ driveId: config.driveId, itemId: config.itemId, textEncoding });

  return async (req) => {
    const requestId = randomUUID();
    const json = (status: number, body: unknown) => ({
      status,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: JSON.stringify(body, null, 2),
    });
    try {
      const token = /^Bearer\s+(\S+)$/i.exec(req.headers['authorization'] ?? '')?.[1];
      if (!token) throw new McSheetyError('UNAUTHENTICATED');
      const identity = normalizeIdentity(await verify(token), {
        tenantId: config.tenantId,
        requiredScope: config.apiScope,
      });
      if (!identity.roles.includes('admin'))
        throw new McSheetyError('FORBIDDEN', 'The probe requires the admin app role.');

      let graphToken: Promise<string> | undefined;
      const ctx: WorkbookContext = {
        correlationId: requestId,
        getGraphToken: () => (graphToken ??= obo(token, requestId)),
      };
      const steps: Step[] = [];
      const record = async (step: string, fn: () => Promise<Omit<Step, 'step'>>) => {
        try {
          steps.push({ step, ...(await fn()) });
        } catch (e) {
          steps.push({
            step,
            ok: false,
            observed: toErrorResponse(e, requestId).body.error,
            note: (e as Error).cause ? JSON.stringify((e as Error).cause) : undefined,
          });
        }
      };

      const apostrophe = adapter('apostrophe');
      const raw = adapter('raw');
      const types = ['string', 'string', 'number'] as const;
      const suffix = Date.now().toString(36);
      const ids = { a: `probe-a-${suffix}`, r: `probe-r-${suffix}` };

      await record('identity + OBO token acquired', async () => ({
        ok: (await ctx.getGraphToken()).length > 0,
        observed: { roles: identity.roles },
      }));

      await apostrophe.withSession(ctx, async (s) => {
        await record('persistent session created', () =>
          Promise.resolve({ ok: s.sessionId !== undefined }),
        );

        const before = await apostrophe.readTable(s, PROBE_TABLE);
        await record('read probe table (headers, location, placeholder rows)', () =>
          Promise.resolve({
            ok: before.columns.join(',') === 'id,text,num',
            observed: {
              columns: before.columns,
              location: before.location,
              bodyRows: before.rows.length,
              allBlankRows: before.rows.filter((r) => r.every((c) => c === '' || c === null))
                .length,
            },
            note: 'An empty Excel table usually reports one blank body row; McSheety ignores rows without an id.',
          }),
        );

        await apostrophe.appendRow(s, PROBE_TABLE, [ids.a, '2026-10-16', 5], types);
        await raw.appendRow(s, PROBE_TABLE, [ids.r, '2026-10-16', 5], types);
        const afterAppend = await apostrophe.readTable(s, PROBE_TABLE);
        const rowOf = (id: string) => afterAppend.rows.find((r) => r[0] === id);
        await record("append with apostrophe encoding keeps '2026-10-16' as text", () => {
          const text = rowOf(ids.a)?.[1];
          return Promise.resolve({
            ok: text === '2026-10-16',
            observed: {
              value: text,
              type: typeof text,
              formula:
                afterAppend.formulas?.[afterAppend.rows.findIndex((r) => r[0] === ids.a)]?.[1],
            },
            note:
              text === "'2026-10-16"
                ? 'Apostrophe stored literally: switch encoding to numberFormat "@".'
                : undefined,
          });
        });
        await record('raw append of 2026-10-16 (shows Excel coercion without encoding)', () =>
          Promise.resolve({
            ok: true,
            observed: { value: rowOf(ids.r)?.[1], type: typeof rowOf(ids.r)?.[1] },
          }),
        );

        const index = afterAppend.rows.findIndex((r) => r[0] === ids.a);
        await record('updateCells patches only the changed cell', async () => {
          const observed: readonly CellValue[] = await apostrophe.updateCells(
            s,
            afterAppend,
            index,
            new Map([[1, '1/2']]),
            types,
          );
          return { ok: observed[1] === '1/2' && observed[2] === 5, observed };
        });

        await record('updateCells detects a stale snapshot (expect CONFLICT)', async () => {
          try {
            await apostrophe.updateCells(s, afterAppend, index, new Map([[1, 'stale']]), types);
            return { ok: false, note: 'Stale write was NOT detected.' };
          } catch (e) {
            return {
              ok: e instanceof McSheetyError && e.code === 'CONFLICT',
              observed: (e as McSheetyError).code,
            };
          }
        });

        await record('range delete removes exactly the target rows', async () => {
          for (const id of [ids.a, ids.r]) {
            const snap = await apostrophe.readTable(s, PROBE_TABLE);
            await apostrophe.deleteRow(
              s,
              snap,
              snap.rows.findIndex((r) => r[0] === id),
            );
          }
          const after = await apostrophe.readTable(s, PROBE_TABLE);
          const leftovers = after.rows.filter((r) => r[0] === ids.a || r[0] === ids.r).length;
          return {
            ok: leftovers === 0 && after.rows.length === before.rows.length,
            observed: { before: before.rows.length, after: after.rows.length },
          };
        });
      });

      await record('GET /me department (User.Read) for trusted attributes', async () => {
        const res = await fetch('https://graph.microsoft.com/v1.0/me?$select=department', {
          headers: { Authorization: `Bearer ${await ctx.getGraphToken()}` },
        });
        const body = (await res.json()) as { department?: string | null };
        return {
          ok: res.ok,
          observed: { status: res.status, department: body.department ?? null },
        };
      });

      return json(200, { requestId, table: PROBE_TABLE, passed: steps.every((s) => s.ok), steps });
    } catch (error) {
      const { status, body } = toErrorResponse(error, requestId);
      return json(status, body);
    }
  };
}
