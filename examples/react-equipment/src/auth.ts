/**
 * Two sign-in modes behind one interface:
 * - entra: MSAL Browser, authorization code + PKCE, access token for the McSheety API scope.
 * - demo:  pick a demo user; only the local demo API (examples/local-api, MODE=demo) accepts it.
 */
import type { AccountInfo } from '@azure/msal-browser';
import { InteractionRequiredAuthError, PublicClientApplication } from '@azure/msal-browser';
import type { TokenRequest } from '@mcsheety/client';
import { config } from './config';

export interface Session {
  readonly label: string;
  readonly getToken: (request?: TokenRequest) => Promise<string>;
  readonly signOut: () => Promise<void>;
}

export interface AuthMode {
  readonly kind: 'entra' | 'demo';
  /** Resolves an existing session (e.g. after the login redirect) or null. */
  restore(): Promise<Session | null>;
}

export const DEMO_USERS = [
  { id: 'viewer', label: 'Vera Viewer — viewer' },
  { id: 'editor-ops', label: 'Eddie Editor — editor, Operations' },
  { id: 'editor-eng', label: 'Erin Editor — editor, Engineering' },
  { id: 'manager', label: 'Morgan Manager — manager' },
  { id: 'admin', label: 'Ada Admin — admin' },
] as const;

export function demoSession(userId: string, onSignOut: () => void): Session {
  const user = DEMO_USERS.find((u) => u.id === userId);
  return {
    label: user?.label ?? userId,
    getToken: () => Promise.resolve(`demo:${userId}`),
    signOut: () => {
      onSignOut();
      return Promise.resolve();
    },
  };
}

let msal: PublicClientApplication | undefined;

async function client(): Promise<PublicClientApplication> {
  if (!msal) {
    msal = new PublicClientApplication({
      auth: {
        clientId: config.spaClientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        redirectUri: window.location.origin,
      },
      cache: { cacheLocation: 'sessionStorage' },
    });
    await msal.initialize();
  }
  return msal;
}

function entraSession(app: PublicClientApplication, account: AccountInfo): Session {
  const scopes = [config.apiScope];
  return {
    label: account.name ?? account.username,
    async getToken(request) {
      const claims = request?.claims;
      try {
        // Never request Graph scopes here: the browser only ever holds a McSheety API token.
        const result = await app.acquireTokenSilent({
          scopes,
          account,
          ...(claims ? { claims } : {}),
        });
        return result.accessToken;
      } catch (error) {
        if (error instanceof InteractionRequiredAuthError || claims) {
          await app.acquireTokenRedirect({ scopes, account, ...(claims ? { claims } : {}) });
        }
        throw error;
      }
    },
    signOut: () => app.logoutRedirect({ account }),
  };
}

export const entraMode: AuthMode = {
  kind: 'entra',
  async restore() {
    const app = await client();
    const result = await app.handleRedirectPromise();
    const account = result?.account ?? app.getActiveAccount() ?? app.getAllAccounts()[0];
    if (!account) return null;
    app.setActiveAccount(account);
    return entraSession(app, account);
  },
};

export async function signInWithEntra(): Promise<void> {
  const app = await client();
  await app.loginRedirect({ scopes: [config.apiScope] });
}
