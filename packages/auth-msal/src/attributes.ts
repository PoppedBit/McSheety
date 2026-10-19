/**
 * Trusted attributes for row policies (e.g. departmentId). Entra does not put a department in
 * access tokens by default, so attributes come from a server-side source, never the client.
 * A missing attribute makes dependent row rules fail closed.
 */
import { McSheetyError } from '@mcsheety/core';
import type { VerifiedIdentity } from './identity.js';

export interface AttributeContext {
  readonly identity: VerifiedIdentity;
  readonly correlationId: string;
  readonly getGraphToken: () => Promise<string>;
}

export interface AttributeResolver {
  resolve(ctx: AttributeContext): Promise<Readonly<Record<string, string>>>;
}

export const noAttributes: AttributeResolver = { resolve: () => Promise.resolve({}) };

/** Fixed mapping keyed by Entra object ID, configured by an administrator. */
export function staticAttributes(
  byUserId: Readonly<Record<string, Readonly<Record<string, string>>>>,
): AttributeResolver {
  return { resolve: ({ identity }) => Promise.resolve(byUserId[identity.userId] ?? {}) };
}

/** Admin-managed Entra user properties that may feed attributes. */
const ALLOWED_PROPERTIES = new Set([
  'department',
  'companyName',
  'officeLocation',
  'jobTitle',
  'employeeType',
]);

export interface GraphProfileAttributesOptions {
  /** attribute name -> Entra user property, e.g. { departmentId: 'department' }. */
  readonly properties: Readonly<Record<string, string>>;
  /**
   * Optional translation of property values, e.g. { departmentId: { Operations: 'dept-ops' } }.
   * When a map is given, unmapped values yield no attribute (fail closed).
   */
  readonly valueMaps?: Readonly<Record<string, Readonly<Record<string, string>>>>;
  readonly baseUrl?: string;
  readonly fetch?: typeof fetch;
}

/**
 * Reads the caller's own profile (`GET /me`) with their delegated Graph token. Requires the
 * delegated `User.Read` permission. Only reads properties administrators control.
 */
export function graphProfileAttributes(options: GraphProfileAttributesOptions): AttributeResolver {
  const props = [...new Set(Object.values(options.properties))];
  const bad = props.find((p) => !ALLOWED_PROPERTIES.has(p));
  if (bad) throw new Error(`Entra property "${bad}" is not allowed as a trusted attribute source.`);
  const base = (options.baseUrl ?? 'https://graph.microsoft.com/v1.0').replace(/\/$/, '');
  const fetchImpl = options.fetch ?? globalThis.fetch.bind(globalThis);

  return {
    async resolve({ getGraphToken, correlationId }) {
      const response = await fetchImpl(`${base}/me?$select=${props.join(',')}`, {
        headers: {
          Authorization: `Bearer ${await getGraphToken()}`,
          'client-request-id': correlationId,
        },
        signal: AbortSignal.timeout(5_000),
      }).catch((cause: unknown) => {
        throw new McSheetyError('GRAPH_ERROR', undefined, { cause });
      });
      if (!response.ok) {
        throw new McSheetyError('GRAPH_ERROR', 'Could not read your Microsoft 365 profile.', {
          cause: { status: response.status },
        });
      }
      const profile = (await response.json()) as Record<string, unknown>;
      const attributes: Record<string, string> = {};
      for (const [attribute, property] of Object.entries(options.properties)) {
        const raw = profile[property];
        if (typeof raw !== 'string' || raw === '') continue;
        const map = options.valueMaps?.[attribute];
        const value = map ? map[raw] : raw;
        if (value !== undefined) attributes[attribute] = value;
      }
      return attributes;
    },
  };
}
