import { GraphWorkbookAdapter } from '@mcsheety/adapter-graph';
import type { OboCredential } from '@mcsheety/auth-msal';
import { createAccessTokenVerifier, createGraphTokenProvider } from '@mcsheety/auth-msal';
import type { WorkbookAdapter } from '@mcsheety/core';
import type { McSheetyConfig } from './config.js';
import type { HttpHandler } from './http.js';
import { createHttpHandler } from './http.js';
import type { Logger } from './logging.js';
import { McSheetyService } from './service.js';

export interface AppDependencies {
  readonly config: McSheetyConfig;
  readonly credential: OboCredential;
  readonly logger: Logger;
  /** Override storage (tests / local demos). Defaults to the Graph adapter for the configured workbook. */
  readonly adapter?: WorkbookAdapter;
}

/** Wires verifier, OBO, attributes, Graph adapter and service into one HTTP handler. */
export function createMcSheetyApp({
  config,
  credential,
  logger,
  adapter,
}: AppDependencies): HttpHandler {
  const storage =
    adapter ??
    new GraphWorkbookAdapter({
      driveId: config.driveId,
      itemId: config.itemId,
      maxRows: config.maxRows,
      onEvent: (event) => logger.info({ msg: `graph ${event.type}`, ...event }),
    });

  const service = new McSheetyService({
    adapter: storage,
    audit: (event) => logger.info({ msg: 'audit', ...event }),
    onModelIssues: (issues, requestId) => {
      const errors = issues.filter((i) => i.severity === 'error');
      if (errors.length > 0)
        logger.error({ msg: 'workbook metadata errors', requestId, issues: errors });
    },
  });

  return createHttpHandler({
    service,
    verifyToken: createAccessTokenVerifier({
      tenantId: config.tenantId,
      apiClientId: config.apiClientId,
    }),
    identity: {
      tenantId: config.tenantId,
      requiredScope: config.apiScope,
      ...(config.groupRoles ? { groupRoles: config.groupRoles } : {}),
    },
    graphToken: createGraphTokenProvider({
      tenantId: config.tenantId,
      clientId: config.apiClientId,
      credential,
    }),
    attributes: config.attributes,
    logger,
  });
}
