/**
 * Structured JSON logs for CloudWatch. Callers pass only safe fields: never bearer tokens,
 * cookies, record bodies, secrets or full Graph URLs.
 */
export type LogFields = Record<string, unknown>;

export interface Logger {
  info(fields: LogFields): void;
  error(fields: LogFields): void;
}

const REDACT_KEYS = /authorization|token|secret|password|cookie|assertion/i;

function redact(fields: LogFields): LogFields {
  return Object.fromEntries(
    Object.entries(fields)
      .filter(([, v]) => v !== undefined)
      .map(([k, v]) => [k, REDACT_KEYS.test(k) ? '[redacted]' : v]),
  );
}

export function jsonLogger(
  write: (line: string) => void = (line) => process.stdout.write(`${line}\n`),
): Logger {
  const emit = (level: string, fields: LogFields) =>
    write(JSON.stringify({ level, time: new Date().toISOString(), ...redact(fields) }));
  return { info: (f) => emit('info', f), error: (f) => emit('error', f) };
}

export const silentLogger: Logger = { info: () => undefined, error: () => undefined };
