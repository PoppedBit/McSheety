import type { ColumnInfo } from '@mcsheety/client';
import { useState } from 'react';

type Value = string | number | boolean | null;

interface Props {
  readonly title: string;
  /** Columns the user may set in this form (from the schema endpoint). */
  readonly columns: readonly ColumnInfo[];
  readonly initial: Readonly<Record<string, Value>>;
  readonly busy: boolean;
  readonly fieldErrors: Readonly<Record<string, string>>;
  readonly onSubmit: (values: Record<string, Value>) => void;
  readonly onCancel: () => void;
}

function toInput(value: Value | undefined): string {
  return value === null || value === undefined ? '' : String(value);
}

function fromInput(column: ColumnInfo, text: string): Value {
  if (text === '') return null;
  if (column.type === 'number') return Number(text);
  if (column.type === 'boolean') return text === 'true';
  return text;
}

export function RecordForm({
  title,
  columns,
  initial,
  busy,
  fieldErrors,
  onSubmit,
  onCancel,
}: Props) {
  const [values, setValues] = useState<Record<string, string>>(() =>
    Object.fromEntries(columns.map((c) => [c.name, toInput(initial[c.name])])),
  );

  return (
    <form
      className="record-form"
      onSubmit={(e) => {
        e.preventDefault();
        const changed: Record<string, Value> = {};
        for (const c of columns) {
          const next = values[c.name] ?? '';
          if (next !== toInput(initial[c.name])) changed[c.name] = fromInput(c, next);
        }
        onSubmit(changed);
      }}
    >
      <h3>{title}</h3>
      {columns.map((c) => {
        const id = `field-${c.name}`;
        const value = values[c.name] ?? '';
        const set = (v: string) => setValues((prev) => ({ ...prev, [c.name]: v }));
        return (
          <label key={c.name} htmlFor={id}>
            <span>
              {c.name}
              {c.required ? ' *' : ''}
            </span>
            {c.type === 'enum' ? (
              <select id={id} value={value} onChange={(e) => set(e.target.value)}>
                <option value="">—</option>
                {c.enumValues?.map((v) => (
                  <option key={v} value={v}>
                    {v}
                  </option>
                ))}
              </select>
            ) : c.type === 'boolean' ? (
              <select id={id} value={value} onChange={(e) => set(e.target.value)}>
                <option value="">—</option>
                <option value="true">true</option>
                <option value="false">false</option>
              </select>
            ) : (
              <input
                id={id}
                type={c.type === 'number' ? 'number' : c.type === 'date' ? 'date' : 'text'}
                step="any"
                min={c.min}
                max={c.max}
                maxLength={c.maxLength}
                value={value}
                placeholder={c.type === 'datetime' ? '2026-10-16T14:05:00Z' : undefined}
                onChange={(e) => set(e.target.value)}
              />
            )}
            {fieldErrors[c.name] ? (
              <small className="field-error">{fieldErrors[c.name]}</small>
            ) : null}
          </label>
        );
      })}
      <div className="actions">
        <button type="submit" disabled={busy}>
          {busy ? 'Saving…' : 'Save'}
        </button>
        <button type="button" className="secondary" onClick={onCancel} disabled={busy}>
          Cancel
        </button>
      </div>
    </form>
  );
}
