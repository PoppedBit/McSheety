import { McSheetyError } from '@mcsheety/core';
import type { JWTPayload, JWTVerifyGetKey } from 'jose';
import { createRemoteJWKSet, jwtVerify } from 'jose';

export interface AccessTokenVerifierOptions {
  readonly tenantId: string;
  /** Application (client) ID of the McSheety API registration. */
  readonly apiClientId: string;
  readonly authorityHost?: string;
  /** Override key lookup (tests). Defaults to the tenant's published JWKS. */
  readonly keys?: JWTVerifyGetKey;
  readonly clockToleranceSeconds?: number;
}

export type AccessTokenVerifier = (token: string) => Promise<JWTPayload>;

/**
 * Verifies an Entra v2.0 access token issued for the McSheety API: RS256 signature against the
 * tenant JWKS, tenant-specific issuer, API audience and lifetime. Never decode-without-verify.
 *
 * Requires the API app registration to issue v2 tokens (`requestedAccessTokenVersion: 2`).
 */
export function createAccessTokenVerifier(
  options: AccessTokenVerifierOptions,
): AccessTokenVerifier {
  const host = (options.authorityHost ?? 'https://login.microsoftonline.com').replace(/\/$/, '');
  const keys =
    options.keys ?? createRemoteJWKSet(new URL(`${host}/${options.tenantId}/discovery/v2.0/keys`));
  const issuer = `${host}/${options.tenantId}/v2.0`;
  const audience = [options.apiClientId, `api://${options.apiClientId}`];

  return async (token) => {
    try {
      const { payload } = await jwtVerify(token, keys, {
        issuer,
        audience,
        algorithms: ['RS256'],
        clockTolerance: options.clockToleranceSeconds ?? 60,
      });
      return payload;
    } catch (cause) {
      throw new McSheetyError('UNAUTHENTICATED', undefined, { cause });
    }
  };
}
