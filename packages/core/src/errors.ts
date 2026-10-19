/**
 * Shared McSheety error contract (implementation.md §9).
 *
 * Convention: `message` and `details` are safe to send to API clients. Anything
 * sensitive (raw Graph errors, file paths, tokens, hidden field values) belongs in
 * `cause`, which is never serialized.
 */

export const ERROR_STATUS = {
  INVALID_REQUEST: 400,
  UNAUTHENTICATED: 401,
  FORBIDDEN: 403,
  NOT_FOUND: 404,
  CONFLICT: 409,
  PRECONDITION_REQUIRED: 428,
  VALIDATION_FAILED: 422,
  RATE_LIMITED: 429,
  INTERNAL_ERROR: 500,
  GRAPH_ERROR: 502,
  WORKBOOK_UNAVAILABLE: 503,
} as const;

export type ErrorCode = keyof typeof ERROR_STATUS;

const RETRYABLE_BY_DEFAULT: ReadonlySet<ErrorCode> = new Set([
  'RATE_LIMITED',
  'WORKBOOK_UNAVAILABLE',
]);

const GENERIC_MESSAGES: Record<ErrorCode, string> = {
  INVALID_REQUEST: 'The request is invalid.',
  UNAUTHENTICATED: 'Authentication is required.',
  FORBIDDEN: 'You do not have permission to perform this action.',
  NOT_FOUND: 'The requested resource was not found.',
  CONFLICT: 'The record changed since it was read.',
  PRECONDITION_REQUIRED: 'An If-Match revision is required for this request.',
  VALIDATION_FAILED: 'One or more fields are invalid.',
  RATE_LIMITED: 'Too many requests. Try again later.',
  INTERNAL_ERROR: 'An unexpected error occurred.',
  GRAPH_ERROR: 'The workbook service returned an error.',
  WORKBOOK_UNAVAILABLE: 'The workbook is temporarily unavailable.',
};

/** A client-safe detail, e.g. one failing field in a VALIDATION_FAILED response. */
export interface ErrorDetail {
  readonly field?: string;
  readonly code: string;
  readonly message: string;
}

export interface McSheetyErrorOptions {
  readonly retryable?: boolean;
  readonly details?: readonly ErrorDetail[];
  /** Seconds a client should wait before retrying (RATE_LIMITED / WORKBOOK_UNAVAILABLE). */
  readonly retryAfterSeconds?: number;
  /**
   * Entra Conditional Access claims challenge (from MSAL InteractionRequiredAuthError). Safe
   * to return: the browser passes it back to MSAL to re-authenticate with the extra claims.
   */
  readonly claimsChallenge?: string;
  /** Internal diagnostic only; never serialized to clients. */
  readonly cause?: unknown;
}

export class McSheetyError extends Error {
  override readonly name = 'McSheetyError';
  readonly code: ErrorCode;
  readonly status: number;
  readonly retryable: boolean;
  readonly details: readonly ErrorDetail[] | undefined;
  readonly retryAfterSeconds: number | undefined;
  readonly claimsChallenge: string | undefined;

  constructor(
    code: ErrorCode,
    message: string = GENERIC_MESSAGES[code],
    options: McSheetyErrorOptions = {},
  ) {
    super(message, options.cause === undefined ? undefined : { cause: options.cause });
    this.code = code;
    this.status = ERROR_STATUS[code];
    this.retryable = options.retryable ?? RETRYABLE_BY_DEFAULT.has(code);
    this.details = options.details;
    this.retryAfterSeconds = options.retryAfterSeconds;
    this.claimsChallenge = options.claimsChallenge;
  }
}

export function isMcSheetyError(value: unknown): value is McSheetyError {
  return value instanceof McSheetyError;
}

export interface ErrorBody {
  readonly error: {
    readonly code: ErrorCode;
    readonly message: string;
    readonly requestId: string;
    readonly retryable: boolean;
    readonly details?: readonly ErrorDetail[];
    readonly claimsChallenge?: string;
  };
}

export interface ErrorResponse {
  readonly status: number;
  readonly body: ErrorBody;
  readonly retryAfterSeconds?: number;
}

/**
 * Converts any thrown value into the standard error response. Unknown errors become a
 * generic INTERNAL_ERROR so their messages can never leak to the client.
 */
export function toErrorResponse(error: unknown, requestId: string): ErrorResponse {
  const known = isMcSheetyError(error) ? error : new McSheetyError('INTERNAL_ERROR');
  const body: ErrorBody = {
    error: {
      code: known.code,
      message: known.message,
      requestId,
      retryable: known.retryable,
      ...(known.details && known.details.length > 0 ? { details: known.details } : {}),
      ...(known.claimsChallenge === undefined ? {} : { claimsChallenge: known.claimsChallenge }),
    },
  };
  return {
    status: known.status,
    body,
    ...(known.retryAfterSeconds === undefined
      ? {}
      : { retryAfterSeconds: known.retryAfterSeconds }),
  };
}
