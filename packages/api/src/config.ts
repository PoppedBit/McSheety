import type { AttributeResolver, OboCredential } from '@mcsheety/auth-msal';
import { graphProfileAttributes, noAttributes, staticAttributes } from '@mcsheety/auth-msal';

export interface McSheetyConfig {
  readonly tenantId: string;
  readonly apiClientId: string;
  /** Short scope name exposed by the API registration, e.g. "access_as_user". */
  readonly apiScope: string;
  readonly driveId: string;
  readonly itemId: string;
  readonly policyMode: 'co-located';
  readonly maxRows: number;
  readonly groupRoles: Readonly<Record<string, string>> | undefined;
  readonly attributes: AttributeResolver;
  /** Secrets Manager secret holding the OBO credential (deployed). */
  readonly credentialSecretArn: string | undefined;
  /** Local development only. Never set this in a deployed environment. */
  readonly clientSecret: string | undefined;
}

export class ConfigError extends Error {
  override readonly name = 'ConfigError';
}

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function parseJson<T>(name: string, value: string | undefined): T | undefined {
  if (value === undefined || value.trim() === '') return undefined;
  try {
    return JSON.parse(value) as T;
  } catch {
    throw new ConfigError(`${name} must be valid JSON.`);
  }
}

type AttributeConfig =
  | {
      source: 'graph';
      properties: Record<string, string>;
      valueMaps?: Record<string, Record<string, string>>;
    }
  | { source: 'static'; users: Record<string, Record<string, string>> };

function attributeResolver(raw: AttributeConfig | undefined): AttributeResolver {
  if (!raw) return noAttributes;
  if (raw.source === 'graph') {
    return graphProfileAttributes({
      properties: raw.properties,
      ...(raw.valueMaps ? { valueMaps: raw.valueMaps } : {}),
    });
  }
  if (raw.source === 'static') return staticAttributes(raw.users);
  throw new ConfigError('MCSHEETY_ATTRIBUTES.source must be "graph" or "static".');
}

/** Reads MCSHEETY_* settings. Throws ConfigError listing every problem (fail fast at cold start). */
export function loadConfig(
  env: Readonly<Record<string, string | undefined>> = process.env,
): McSheetyConfig {
  const problems: string[] = [];
  const required = (name: string, guid = false): string => {
    const value = env[name]?.trim() ?? '';
    if (value === '') problems.push(`${name} is required`);
    else if (guid && !GUID_RE.test(value)) problems.push(`${name} must be a GUID`);
    return value;
  };

  const tenantId = required('MCSHEETY_TENANT_ID', true);
  const apiClientId = required('MCSHEETY_API_CLIENT_ID', true);
  const driveId = required('MCSHEETY_DRIVE_ID');
  const itemId = required('MCSHEETY_ITEM_ID');
  const apiScope = env['MCSHEETY_API_SCOPE']?.trim() || 'access_as_user';
  if (apiScope.includes('/'))
    problems.push('MCSHEETY_API_SCOPE is the short scope name, e.g. access_as_user');

  const policyMode = env['MCSHEETY_POLICY_MODE']?.trim() || 'co-located';
  if (policyMode !== 'co-located') {
    problems.push(
      'MCSHEETY_POLICY_MODE must be "co-located" (separate policy workbooks are post-MVP)',
    );
  }

  const maxRows = Number(env['MCSHEETY_MAX_ROWS'] ?? '2000');
  if (!Number.isInteger(maxRows) || maxRows < 1)
    problems.push('MCSHEETY_MAX_ROWS must be a positive integer');

  const credentialSecretArn = env['MCSHEETY_CREDENTIAL_SECRET_ARN']?.trim() || undefined;
  const clientSecret = env['MCSHEETY_CLIENT_SECRET'] || undefined;
  if (!credentialSecretArn && !clientSecret) {
    problems.push(
      'Set MCSHEETY_CREDENTIAL_SECRET_ARN (deployed) or MCSHEETY_CLIENT_SECRET (local only)',
    );
  }

  let groupRoles: Record<string, string> | undefined;
  let attributes: AttributeResolver = noAttributes;
  try {
    groupRoles = parseJson('MCSHEETY_GROUP_ROLES', env['MCSHEETY_GROUP_ROLES']);
    attributes = attributeResolver(parseJson('MCSHEETY_ATTRIBUTES', env['MCSHEETY_ATTRIBUTES']));
  } catch (e) {
    problems.push((e as Error).message);
  }

  if (problems.length > 0)
    throw new ConfigError(`Invalid McSheety configuration: ${problems.join('; ')}.`);
  return {
    tenantId,
    apiClientId,
    apiScope,
    driveId,
    itemId,
    policyMode: 'co-located',
    maxRows,
    groupRoles,
    attributes,
    credentialSecretArn,
    clientSecret,
  };
}

/** Parses a Secrets Manager secret: plain client secret, or JSON with clientSecret / certificate. */
export function parseCredentialSecret(secret: string): OboCredential {
  const trimmed = secret.trim();
  if (!trimmed.startsWith('{')) return { clientSecret: trimmed };
  const parsed = JSON.parse(trimmed) as {
    clientSecret?: string;
    certificate?: { thumbprintSha256?: string; privateKey?: string };
  };
  if (parsed.certificate?.thumbprintSha256 && parsed.certificate.privateKey) {
    return {
      certificate: {
        thumbprintSha256: parsed.certificate.thumbprintSha256,
        privateKey: parsed.certificate.privateKey,
      },
    };
  }
  if (parsed.clientSecret) return { clientSecret: parsed.clientSecret };
  throw new ConfigError(
    'The credential secret must contain clientSecret or certificate.{thumbprintSha256,privateKey}.',
  );
}
