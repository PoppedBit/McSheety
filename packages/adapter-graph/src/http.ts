/**
 * Minimal Microsoft Graph HTTP transport: auth header, session header, bounded retries with
 * jitter for idempotent calls, and mapping of Graph failures to sanitized McSheety errors.
 */
import type { WorkbookContext } from '@mcsheety/core';
import { McSheetyError } from '@mcsheety/core';

export interface RetryOptions {
  /** Total attempts for idempotent requests, including the first. */
  readonly maxAttempts: number;
  readonly baseDelayMs: number;
  readonly maxDelayMs: number;
  /** Give up (and report RATE_LIMITED / WORKBOOK_UNAVAILABLE) rather than exceed this. */
  readonly budgetMs: number;
}

export interface GraphHttpOptions {
  readonly baseUrl?: string;
  readonly fetch?: typeof fetch;
  readonly retry?: Partial<RetryOptions>;
  readonly requestTimeoutMs?: number;
  readonly sleep?: (ms: number) => Promise<void>;
  readonly random?: () => number;
  /** Receives safe diagnostics (no tokens, no response bodies). */
  readonly onEvent?: (event: GraphEvent) => void;
}

export interface GraphEvent {
  readonly type: 'retry' | 'failure';
  readonly correlationId: string;
  readonly method: string;
  readonly status: number | undefined;
  readonly graphCode: string | undefined;
  readonly attempt: number;
  readonly delayMs?: number;
}

/** Internal diagnostic attached as `cause`; never serialized to clients. */
export interface GraphFailure {
  readonly status: number | undefined;
  readonly graphCode: string | undefined;
  readonly innerCode: string | undefined;
  readonly graphRequestId: string | undefined;
  /** True when a non-idempotent request may or may not have been applied. */
  readonly ambiguous: boolean;
}

export const DEFAULT_RETRY: RetryOptions = {
  maxAttempts: 4,
  baseDelayMs: 400,
  maxDelayMs: 8_000,
  budgetMs: 20_000,
};

export interface GraphRequest {
  readonly method: 'GET' | 'POST' | 'PATCH' | 'DELETE';
  /** Path under the base URL, or an absolute @odata.nextLink. */
  readonly path: string;
  readonly body?: unknown;
  /** Only idempotent requests are retried. */
  readonly idempotent: boolean;
}

const RETRYABLE_STATUS = new Set([429, 503, 504]);

export function isAmbiguousWrite(error: unknown): boolean {
  return (
    error instanceof McSheetyError &&
    typeof error.cause === 'object' &&
    error.cause !== null &&
    (error.cause as Partial<GraphFailure>).ambiguous === true
  );
}

function mapFailure(failure: GraphFailure, retryAfterSeconds: number | undefined): McSheetyError {
  const options = {
    cause: failure,
    ...(retryAfterSeconds === undefined ? {} : { retryAfterSeconds }),
  };
  const { status, innerCode, graphCode } = failure;
  if (status === undefined) {
    return new McSheetyError(
      'WORKBOOK_UNAVAILABLE',
      'The workbook service did not respond in time.',
      options,
    );
  }
  if (status === 401) {
    return new McSheetyError(
      'UNAUTHENTICATED',
      'Your Microsoft 365 session needs to be renewed.',
      options,
    );
  }
  if (status === 403) {
    return new McSheetyError(
      'FORBIDDEN',
      'Your Microsoft 365 account does not have access to this workbook.',
      options,
    );
  }
  if (status === 429) return new McSheetyError('RATE_LIMITED', undefined, options);
  if (status === 404 || status === 409 || status === 423 || status >= 500) {
    const sessionLost = /session/i.test(innerCode ?? graphCode ?? '');
    return new McSheetyError(
      'WORKBOOK_UNAVAILABLE',
      sessionLost ? 'The workbook session expired. Try again.' : undefined,
      { ...options, retryable: status !== 404 },
    );
  }
  return new McSheetyError('GRAPH_ERROR', undefined, options);
}

function parseRetryAfter(header: string | null): number | undefined {
  if (header === null) return undefined;
  const seconds = Number(header);
  return Number.isFinite(seconds) && seconds >= 0 ? seconds : undefined;
}

export class GraphHttp {
  private readonly baseUrl: string;
  private readonly fetchImpl: typeof fetch;
  private readonly retry: RetryOptions;
  private readonly timeoutMs: number;
  private readonly sleep: (ms: number) => Promise<void>;
  private readonly random: () => number;
  private readonly onEvent: ((event: GraphEvent) => void) | undefined;

  constructor(options: GraphHttpOptions = {}) {
    this.baseUrl = (options.baseUrl ?? 'https://graph.microsoft.com/v1.0').replace(/\/$/, '');
    this.fetchImpl = options.fetch ?? globalThis.fetch.bind(globalThis);
    this.retry = { ...DEFAULT_RETRY, ...options.retry };
    this.timeoutMs = options.requestTimeoutMs ?? 10_000;
    this.sleep = options.sleep ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms)));
    this.random = options.random ?? Math.random;
    this.onEvent = options.onEvent;
  }

  async request<T>(ctx: WorkbookContext, req: GraphRequest): Promise<T> {
    const started = Date.now();
    const url = req.path.startsWith('https://') ? req.path : `${this.baseUrl}${req.path}`;
    if (!url.startsWith(`${this.baseUrl}/`)) {
      // Never follow links (e.g. a tampered nextLink) away from Graph with the user's token.
      throw new McSheetyError('GRAPH_ERROR');
    }

    for (let attempt = 1; ; attempt++) {
      const token = await ctx.getGraphToken();
      const headers: Record<string, string> = {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'client-request-id': ctx.correlationId,
      };
      if (ctx.sessionId) headers['Workbook-Session-Id'] = ctx.sessionId;
      if (req.body !== undefined) headers['Content-Type'] = 'application/json';

      let response: Response | undefined;
      try {
        response = await this.fetchImpl(url, {
          method: req.method,
          headers,
          ...(req.body === undefined ? {} : { body: JSON.stringify(req.body) }),
          signal: AbortSignal.timeout(this.timeoutMs),
        });
      } catch {
        response = undefined; // network error or timeout
      }

      if (response?.ok) {
        if (response.status === 204) return undefined as T;
        const text = await response.text();
        return (text === '' ? undefined : JSON.parse(text)) as T;
      }

      const status = response?.status;
      let graphCode: string | undefined;
      let innerCode: string | undefined;
      if (response) {
        try {
          const body = (await response.json()) as {
            error?: { code?: string; innerError?: { code?: string } };
          };
          graphCode = body.error?.code;
          innerCode = body.error?.innerError?.code;
        } catch {
          // non-JSON error body
        }
      }
      const retryAfter = parseRetryAfter(response?.headers.get('Retry-After') ?? null);
      const transient = status === undefined || RETRYABLE_STATUS.has(status);
      const failure: GraphFailure = {
        status,
        graphCode,
        innerCode,
        graphRequestId: response?.headers.get('request-id') ?? undefined,
        ambiguous: !req.idempotent && (status === undefined || status >= 500),
      };

      if (req.idempotent && transient && attempt < this.retry.maxAttempts) {
        const backoff = Math.min(
          this.retry.maxDelayMs,
          this.retry.baseDelayMs * 2 ** (attempt - 1),
        );
        const delayMs = Math.max(
          retryAfter === undefined ? 0 : retryAfter * 1000,
          this.random() * backoff,
        );
        if (Date.now() - started + delayMs <= this.retry.budgetMs) {
          this.onEvent?.({
            type: 'retry',
            correlationId: ctx.correlationId,
            method: req.method,
            status,
            graphCode,
            attempt,
            delayMs,
          });
          await this.sleep(delayMs);
          continue;
        }
      }
      this.onEvent?.({
        type: 'failure',
        correlationId: ctx.correlationId,
        method: req.method,
        status,
        graphCode,
        attempt,
      });
      throw mapFailure(failure, retryAfter);
    }
  }
}
