/**
 * AWS Lambda entry point for API Gateway HTTP API (payload format 2.0).
 * API Gateway's JWT authorizer rejects bad tokens at the edge; the handler verifies the token
 * again itself (defense in depth) and needs the raw token anyway for the OBO exchange.
 */
import { GetSecretValueCommand, SecretsManagerClient } from '@aws-sdk/client-secrets-manager';
import type { APIGatewayProxyEventV2, APIGatewayProxyStructuredResultV2 } from 'aws-lambda';
import { createMcSheetyApp } from './app.js';
import { loadConfig, parseCredentialSecret } from './config.js';
import type { HttpHandler, HttpRequest } from './http.js';
import { jsonLogger } from './logging.js';

/** Credentials are re-read periodically so secret rotation takes effect without a redeploy. */
const CREDENTIAL_TTL_MS = 60 * 60 * 1000;

export function fromLambdaEvent(event: APIGatewayProxyEventV2): HttpRequest {
  const stage = event.requestContext.stage;
  let path = event.rawPath;
  if (stage && stage !== '$default' && path.startsWith(`/${stage}/`))
    path = path.slice(stage.length + 1);
  const headers: Record<string, string> = {};
  for (const [k, v] of Object.entries(event.headers ?? {}))
    if (v !== undefined) headers[k.toLowerCase()] = v;
  return {
    method: event.requestContext.http.method.toUpperCase(),
    path,
    headers,
    query: event.queryStringParameters ?? {},
    body:
      event.body === undefined
        ? undefined
        : event.isBase64Encoded
          ? Buffer.from(event.body, 'base64').toString('utf8')
          : event.body,
    requestId: event.requestContext.requestId,
  };
}

let cached: { handler: HttpHandler; builtAt: number } | undefined;
let building: Promise<HttpHandler> | undefined;

async function build(): Promise<HttpHandler> {
  const config = loadConfig();
  const logger = jsonLogger();
  let secret = config.clientSecret;
  if (config.credentialSecretArn) {
    const result = await new SecretsManagerClient({}).send(
      new GetSecretValueCommand({ SecretId: config.credentialSecretArn }),
    );
    secret = result.SecretString;
  }
  if (!secret) throw new Error('The OBO credential secret is empty.');
  return createMcSheetyApp({ config, credential: parseCredentialSecret(secret), logger });
}

async function current(): Promise<HttpHandler> {
  if (cached && Date.now() - cached.builtAt < CREDENTIAL_TTL_MS) return cached.handler;
  building ??= build()
    .then((handler) => {
      cached = { handler, builtAt: Date.now() };
      return handler;
    })
    .finally(() => {
      building = undefined;
    });
  return building;
}

export async function handler(
  event: APIGatewayProxyEventV2,
): Promise<APIGatewayProxyStructuredResultV2> {
  const response = await (await current())(fromLambdaEvent(event));
  return { statusCode: response.status, headers: response.headers, body: response.body };
}
