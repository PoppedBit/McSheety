/**
 * @mcsheety/client — typed REST client for the McSheety API (implementation.md §11).
 * Framework-agnostic; works in browsers and Node.js 22+. It only ever sends a token for the
 * McSheety API; it never sees Microsoft Graph tokens.
 */
import type {
  CellValue,
  ErrorCode,
  ErrorDetail,
  FieldIssue,
  ListResult,
  RecordEnvelope,
  RecordValues,
  SchemaResult,
  TableCapabilities,
} from '@mcsheety/core';

export type { ColumnInfo, FieldIssue, SchemaResult, TableCapabilities } from '@mcsheety/core';

export interface TokenRequest {
  /** Conditional Access claims challenge to satisfy (pass to MSAL `claims`). */
  readonly claims?: string;
}

export interface ClientOptions {
  /** Base URL of the McSheety API, without /v1. */
  readonly apiUrl: string;
  /** Returns an Entra access token for the McSheety API scope (not Microsoft Graph). */
  readonly getToken: (request?: TokenRequest) => Promise<string>;
  readonly fetch?: typeof fetch;
  /** Retries for throttled/unavailable GETs, honouring Retry-After. Default 2. */
  readonly maxRetries?: number;
  readonly sleep?: (ms: number) => Promise<void>;
}

/** A record as returned by the API. `revision` lives outside `fields`, so it is never persisted. */
export interface TableRecord<T extends RecordValues = RecordValues> {
  readonly id: string;
  readonly revision: string;
  readonly fields: T;
  readonly issues?: readonly FieldIssue[];
}

export type Condition<V> = V | { eq: V } | { ne: V } | { in: readonly V[] };
export type Where<T extends RecordValues> = { [K in keyof T]?: Condition<T[K]> };
export type SortSpec<T extends RecordValues> = (keyof T & string) | `-${keyof T & string}`;

export interface FindManyOptions<T extends RecordValues> {
  readonly where?: Where<T>;
  readonly sort?: readonly SortSpec<T>[];
  readonly limit?: number;
  readonly cursor?: string | null;
}

export interface Page<T extends RecordValues> {
  readonly records: readonly TableRecord<T>[];
  readonly nextCursor: string | null;
}

export interface WriteOptions {
  /** The record's `revision` from a prior read. Required by the API. */
  readonly ifMatch: string;
}

export class McSheetyClientError extends Error {
  override readonly name = 'McSheetyClientError';
  constructor(
    readonly status: number,
    readonly code: ErrorCode | 'NETWORK_ERROR',
    message: string,
    readonly requestId: string | undefined,
    readonly retryable: boolean,
    readonly details: readonly ErrorDetail[] = [],
    readonly retryAfterSeconds?: number,
    readonly claimsChallenge?: string,
  ) {
    super(message);
  }

  /** Stale revision or concurrent edit: re-read the record and retry deliberately. */
  get isConflict(): boolean {
    return this.code === 'CONFLICT';
  }
}

interface ErrorPayload {
  error?: {
    code?: ErrorCode;
    message?: string;
    requestId?: string;
    retryable?: boolean;
    details?: ErrorDetail[];
    claimsChallenge?: string;
  };
}

const toRecord = <T extends RecordValues>(e: RecordEnvelope): TableRecord<T> => ({
  id: e.id,
  revision: e.revision,
  fields: e.fields as T,
  ...(e.issues ? { issues: e.issues } : {}),
});

export class McSheetyClient {
  private readonly base: string;
  private readonly fetchImpl: typeof fetch;
  private readonly maxRetries: number;
  private readonly sleep: (ms: number) => Promise<void>;

  constructor(private readonly options: ClientOptions) {
    this.base = `${options.apiUrl.replace(/\/+$/, '')}/v1`;
    this.fetchImpl = options.fetch ?? globalThis.fetch.bind(globalThis);
    this.maxRetries = options.maxRetries ?? 2;
    this.sleep = options.sleep ?? ((ms) => new Promise((r) => setTimeout(r, ms)));
  }

  /** @internal */
  async request<T>(
    method: string,
    path: string,
    init: { body?: unknown; ifMatch?: string } = {},
  ): Promise<{ data: T; etag: string | null }> {
    let claims: string | undefined;
    for (let attempt = 0; ; attempt++) {
      const token = await this.options.getToken(claims ? { claims } : undefined);
      const headers: Record<string, string> = {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
      };
      if (init.body !== undefined) headers['Content-Type'] = 'application/json';
      if (init.ifMatch !== undefined) headers['If-Match'] = `"${init.ifMatch}"`;

      let response: Response;
      try {
        response = await this.fetchImpl(`${this.base}${path}`, {
          method,
          headers,
          ...(init.body === undefined ? {} : { body: JSON.stringify(init.body) }),
        });
      } catch {
        throw new McSheetyClientError(
          0,
          'NETWORK_ERROR',
          'Could not reach the McSheety API.',
          undefined,
          true,
        );
      }

      if (response.ok) {
        const text = await response.text();
        return {
          data: (text ? JSON.parse(text) : undefined) as T,
          etag: response.headers.get('ETag'),
        };
      }

      const payload = (await response.json().catch(() => ({}))) as ErrorPayload;
      const e = payload.error ?? {};
      const retryAfter = Number(response.headers.get('Retry-After') ?? NaN);
      const error = new McSheetyClientError(
        response.status,
        e.code ?? 'INTERNAL_ERROR',
        e.message ?? `Request failed with status ${response.status}.`,
        e.requestId ?? response.headers.get('x-request-id') ?? undefined,
        e.retryable ?? false,
        e.details ?? [],
        Number.isFinite(retryAfter) ? retryAfter : undefined,
        e.claimsChallenge,
      );

      // Conditional Access: ask the token provider once for a token satisfying the claims.
      if (response.status === 401 && e.claimsChallenge && claims === undefined) {
        claims = e.claimsChallenge;
        continue;
      }
      const transient = response.status === 429 || response.status === 503;
      if (method === 'GET' && transient && attempt < this.maxRetries) {
        await this.sleep(Math.min((error.retryAfterSeconds ?? 2 ** attempt) * 1000, 10_000));
        continue;
      }
      throw error;
    }
  }

  table<T extends RecordValues = RecordValues>(name: string): TableClient<T> {
    return new TableClient<T>(this, name);
  }

  async tables(): Promise<string[]> {
    const { data } = await this.request<{ data: { name: string }[] }>('GET', '/tables');
    return data.data.map((t) => t.name);
  }

  async capabilities(): Promise<TableCapabilities[]> {
    return (await this.request<{ data: TableCapabilities[] }>('GET', '/capabilities')).data.data;
  }

  async schema(table: string): Promise<SchemaResult> {
    return (await this.request<SchemaResult>('GET', `/schema/${encodeURIComponent(table)}`)).data;
  }
}

export class TableClient<T extends RecordValues> {
  private readonly path: string;

  constructor(
    private readonly client: McSheetyClient,
    readonly name: string,
  ) {
    this.path = `/tables/${encodeURIComponent(name)}/rows`;
  }

  async findMany(options: FindManyOptions<T> = {}): Promise<Page<T>> {
    const params = new URLSearchParams();
    if (options.where && Object.keys(options.where).length > 0)
      params.set('where', JSON.stringify(options.where));
    if (options.sort && options.sort.length > 0) params.set('sort', options.sort.join(','));
    if (options.limit !== undefined) params.set('limit', String(options.limit));
    if (options.cursor) params.set('cursor', options.cursor);
    const query = params.size > 0 ? `?${params.toString()}` : '';
    const { data } = await this.client.request<ListResult>('GET', `${this.path}${query}`);
    return { records: data.data.map((e) => toRecord<T>(e)), nextCursor: data.page.nextCursor };
  }

  /** Follows cursors until exhausted. Mind the API's table-size limits. */
  async *findAll(options: Omit<FindManyOptions<T>, 'cursor'> = {}): AsyncGenerator<TableRecord<T>> {
    let cursor: string | null = null;
    do {
      const page: Page<T> = await this.findMany({ ...options, cursor });
      yield* page.records;
      cursor = page.nextCursor;
    } while (cursor);
  }

  async findById(id: string): Promise<TableRecord<T>> {
    const { data } = await this.client.request<{ data: RecordEnvelope }>(
      'GET',
      `${this.path}/${encodeURIComponent(id)}`,
    );
    return toRecord<T>(data.data);
  }

  /** The server assigns `id` and any server-owned defaults. */
  async create(fields: Partial<Record<keyof T, CellValue>>): Promise<TableRecord<T>> {
    const { data } = await this.client.request<{ data: RecordEnvelope }>('POST', this.path, {
      body: fields,
    });
    return toRecord<T>(data.data);
  }

  async update(
    id: string,
    patch: Partial<Record<keyof T, CellValue>>,
    options: WriteOptions,
  ): Promise<TableRecord<T>> {
    const { data } = await this.client.request<{ data: RecordEnvelope }>(
      'PATCH',
      `${this.path}/${encodeURIComponent(id)}`,
      {
        body: patch,
        ifMatch: options.ifMatch,
      },
    );
    return toRecord<T>(data.data);
  }

  async delete(id: string, options: WriteOptions): Promise<void> {
    await this.client.request('DELETE', `${this.path}/${encodeURIComponent(id)}`, {
      ifMatch: options.ifMatch,
    });
  }
}

export function createClient(options: ClientOptions): McSheetyClient {
  return new McSheetyClient(options);
}
