import { describe, expect, it } from 'vitest';
import { McSheetyError, toErrorResponse } from './errors.js';

describe('McSheetyError', () => {
  it('maps codes to HTTP status and default retryability', () => {
    expect(new McSheetyError('CONFLICT').status).toBe(409);
    expect(new McSheetyError('CONFLICT').retryable).toBe(false);
    expect(new McSheetyError('RATE_LIMITED').retryable).toBe(true);
    expect(new McSheetyError('WORKBOOK_UNAVAILABLE').status).toBe(503);
  });

  it('uses a generic client-safe message when none is given', () => {
    expect(new McSheetyError('FORBIDDEN').message).toBe(
      'You do not have permission to perform this action.',
    );
  });
});

describe('toErrorResponse', () => {
  it('serializes McSheetyError without its cause', () => {
    const err = new McSheetyError('GRAPH_ERROR', undefined, {
      cause: new Error('https://contoso.sharepoint.com/sites/secret/Equipment.xlsx 500'),
    });
    const res = toErrorResponse(err, 'req-1');
    expect(res.status).toBe(502);
    expect(res.body).toEqual({
      error: {
        code: 'GRAPH_ERROR',
        message: 'The workbook service returned an error.',
        requestId: 'req-1',
        retryable: false,
      },
    });
    expect(JSON.stringify(res)).not.toContain('sharepoint');
  });

  it('hides messages of unknown errors behind INTERNAL_ERROR', () => {
    const res = toErrorResponse(new Error('token=eyJ...'), 'req-2');
    expect(res.status).toBe(500);
    expect(res.body.error.code).toBe('INTERNAL_ERROR');
    expect(JSON.stringify(res)).not.toContain('eyJ');
  });

  it('includes validation details and retry-after when present', () => {
    const res = toErrorResponse(
      new McSheetyError('VALIDATION_FAILED', undefined, {
        details: [{ field: 'name', code: 'required', message: 'name is required.' }],
      }),
      'req-3',
    );
    expect(res.body.error.details).toHaveLength(1);

    const limited = toErrorResponse(
      new McSheetyError('RATE_LIMITED', undefined, { retryAfterSeconds: 5 }),
      'req-4',
    );
    expect(limited.retryAfterSeconds).toBe(5);
    expect(limited.body.error.retryable).toBe(true);
  });
});
