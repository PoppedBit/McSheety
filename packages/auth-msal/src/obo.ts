import { McSheetyError } from '@mcsheety/core';
import type { Configuration } from '@azure/msal-node';
import { ConfidentialClientApplication, InteractionRequiredAuthError } from '@azure/msal-node';

export type OboCredential =
  | { readonly clientSecret: string }
  | { readonly certificate: { readonly thumbprintSha256: string; readonly privateKey: string } };

export interface OboClient {
  acquireTokenOnBehalfOf(request: {
    oboAssertion: string;
    scopes: string[];
    correlationId?: string;
  }): Promise<{ accessToken: string } | null>;
}

export interface GraphTokenProviderOptions {
  readonly tenantId: string;
  /** McSheety API app registration (the confidential client). */
  readonly clientId: string;
  readonly credential: OboCredential;
  readonly authorityHost?: string;
  /** Defaults to Graph `.default`, i.e. the delegated permissions consented for the API app. */
  readonly scopes?: readonly string[];
  /** Inject a client (tests). */
  readonly client?: OboClient;
}

/** Exchanges the inbound API token (user assertion) for a delegated Graph token. */
export type GraphTokenProvider = (userAssertion: string, correlationId: string) => Promise<string>;

/** Error code text used for classification. MSAL 7 keeps the AADSTS number in `errorNo`. */
function messageOf(error: unknown): string {
  if (!(error instanceof Error)) return '';
  const e = error as { errorCode?: string; subError?: string; errorNo?: string };
  return [e.errorCode, e.subError, e.errorNo ? `AADSTS${e.errorNo}` : '', error.message].join(' ');
}

/**
 * MSAL Node On-Behalf-Of. One ConfidentialClientApplication per process; MSAL keys cached OBO
 * tokens by the user assertion, so cached tokens are never served to another principal.
 * Tokens are never logged or returned to the browser.
 */
export function createGraphTokenProvider(options: GraphTokenProviderOptions): GraphTokenProvider {
  const host = (options.authorityHost ?? 'https://login.microsoftonline.com').replace(/\/$/, '');
  const scopes = [...(options.scopes ?? ['https://graph.microsoft.com/.default'])];
  const auth: Configuration['auth'] = {
    clientId: options.clientId,
    authority: `${host}/${options.tenantId}`,
    ...('clientSecret' in options.credential
      ? { clientSecret: options.credential.clientSecret }
      : { clientCertificate: options.credential.certificate }),
  };
  const client: OboClient = options.client ?? new ConfidentialClientApplication({ auth });

  return async (userAssertion, correlationId) => {
    try {
      const result = await client.acquireTokenOnBehalfOf({
        oboAssertion: userAssertion,
        scopes,
        correlationId,
      });
      if (!result?.accessToken) throw new Error('OBO returned no access token');
      return result.accessToken;
    } catch (error) {
      throw mapOboError(error);
    }
  };
}

export function mapOboError(error: unknown): McSheetyError {
  const text = messageOf(error);
  if (error instanceof InteractionRequiredAuthError && error.claims) {
    return new McSheetyError(
      'UNAUTHENTICATED',
      'Additional sign-in is required by your organization.',
      {
        claimsChallenge: error.claims,
        cause: error,
      },
    );
  }
  if (/AADSTS65001|consent_required/i.test(text)) {
    return new McSheetyError(
      'FORBIDDEN',
      'An administrator must grant consent before McSheety can access Microsoft 365.',
      {
        details: [{ code: 'consent_required', message: 'Admin consent is required.' }],
        cause: error,
      },
    );
  }
  if (
    error instanceof InteractionRequiredAuthError ||
    /AADSTS50013|AADSTS500133|invalid_grant/i.test(text)
  ) {
    return new McSheetyError('UNAUTHENTICATED', 'Your sign-in has expired. Sign in again.', {
      cause: error,
    });
  }
  return new McSheetyError('GRAPH_ERROR', 'Could not obtain delegated Microsoft 365 access.', {
    cause: error,
  });
}
