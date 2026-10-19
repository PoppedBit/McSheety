const env = import.meta.env;

export const config = {
  apiUrl: env.VITE_MCSHEETY_API_URL ?? 'http://localhost:3000',
  demo: env.VITE_MCSHEETY_DEMO === '1',
  tenantId: env.VITE_ENTRA_TENANT_ID ?? '',
  spaClientId: env.VITE_ENTRA_SPA_CLIENT_ID ?? '',
  /** Full scope URI, e.g. api://<API client id>/access_as_user. */
  apiScope: env.VITE_MCSHEETY_API_SCOPE ?? '',
};

export function missingEntraConfig(): string[] {
  return [
    ['VITE_ENTRA_TENANT_ID', config.tenantId],
    ['VITE_ENTRA_SPA_CLIENT_ID', config.spaClientId],
    ['VITE_MCSHEETY_API_SCOPE', config.apiScope],
  ]
    .filter(([, value]) => !value)
    .map(([name]) => name as string);
}
