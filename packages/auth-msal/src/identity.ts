import type { Identity } from '@mcsheety/core';
import { McSheetyError } from '@mcsheety/core';
import type { JWTPayload } from 'jose';

export interface IdentityOptions {
  readonly tenantId: string;
  /** Delegated scope the SPA requests for the API, e.g. "access_as_user". */
  readonly requiredScope: string;
  /**
   * Optional server-side mapping from immutable Entra group object IDs to McSheety roles.
   * Group names are never used. Group-overage tokens contribute no group roles (fail closed).
   */
  readonly groupRoles?: Readonly<Record<string, string>>;
}

export type VerifiedIdentity = Omit<Identity, 'attributes'>;

const strings = (value: unknown): string[] =>
  Array.isArray(value) ? value.filter((v): v is string => typeof v === 'string') : [];

/**
 * Builds the caller identity from an already *verified* access token payload. Rejects other
 * tenants, app-only tokens and tokens without the API's delegated scope. Uses tid + oid as the
 * principal; never email, UPN or display name.
 */
export function normalizeIdentity(claims: JWTPayload, options: IdentityOptions): VerifiedIdentity {
  const tid = claims['tid'];
  const oid = claims['oid'];
  if (tid !== options.tenantId) {
    throw new McSheetyError('UNAUTHENTICATED', 'Tokens from this tenant are not accepted.');
  }
  if (typeof oid !== 'string' || oid === '') throw new McSheetyError('UNAUTHENTICATED');

  const scp = claims['scp'];
  if (claims['idtyp'] === 'app' || typeof scp !== 'string') {
    throw new McSheetyError('FORBIDDEN', 'Application tokens are not accepted; sign in as a user.');
  }
  if (!scp.split(' ').includes(options.requiredScope)) {
    throw new McSheetyError('FORBIDDEN', 'The token is missing the required API scope.');
  }

  const roles = new Set(strings(claims['roles']));
  if (options.groupRoles) {
    for (const group of strings(claims['groups'])) {
      const role = options.groupRoles[group];
      if (role) roles.add(role);
    }
  }
  return { tenantId: tid, userId: oid, roles: [...roles] };
}
