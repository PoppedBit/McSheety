import type { ModelIssue, RawTable } from '@mcsheety/core';
import { METADATA_TABLES, decodeTable, parseWorkbookModel } from '@mcsheety/core';
import type { LocalTable } from './xlsx.js';

export interface ValidationReport {
  readonly ok: boolean;
  readonly tables: readonly string[];
  readonly issues: readonly ModelIssue[];
}

/**
 * `mcsheety validate`: metadata problems (via the same parser the API uses) plus data
 * problems in each managed table: missing/duplicate ids, undecodable cells and formula cells.
 */
export function validateTables(tables: readonly LocalTable[]): ValidationReport {
  const byName = new Map(tables.map((t) => [t.name, t]));
  const issues: ModelIssue[] = [];
  for (const t of tables) {
    if (!t.name.startsWith('mc_') && t.sheet.startsWith('_')) {
      issues.push({
        severity: 'warning',
        source: t.name,
        message: `Data table ${t.name} is on metadata-style sheet "${t.sheet}".`,
      });
    }
  }
  const meta = (name: string): RawTable | undefined => byName.get(name);
  const { model, issues: modelIssues } = parseWorkbookModel({
    config: meta(METADATA_TABLES.config),
    schema: meta(METADATA_TABLES.schema),
    permissions: meta(METADATA_TABLES.permissions),
    columns: meta(METADATA_TABLES.columns),
    rowPolicies: meta(METADATA_TABLES.rowPolicies),
    dataTables: new Map(
      tables.filter((t) => !t.name.startsWith('mc_')).map((t) => [t.name, { columns: t.columns }]),
    ),
  });
  issues.push(...modelIssues);

  for (const schema of model.tables.values()) {
    const data = byName.get(schema.excelTable);
    if (!data) continue;
    const decoded = decodeTable(schema, data);
    const at = { table: schema.name };
    for (const row of decoded.problems.rowsWithoutId) {
      issues.push({
        severity: 'error',
        source: data.name,
        row,
        message: 'Row has a blank id and is unreachable through the API.',
        ...at,
      });
    }
    for (const id of decoded.problems.duplicateIds) {
      issues.push({
        severity: 'error',
        source: data.name,
        message: `Duplicate id "${id}"; these rows are hidden until fixed.`,
        ...at,
      });
    }
    for (const record of decoded.records) {
      for (const issue of record.issues) {
        issues.push({
          severity: 'warning',
          source: data.name,
          row: record.index + 1,
          message: `${record.id}: ${issue.message}`,
          ...at,
        });
      }
      for (const column of record.formulaColumns) {
        issues.push({
          severity: 'warning',
          source: data.name,
          row: record.index + 1,
          message: `${record.id}: ${column} contains a formula; McSheety will refuse to write it.`,
          ...at,
        });
      }
    }
  }
  return {
    ok: !issues.some((i) => i.severity === 'error'),
    tables: [...model.tables.keys()],
    issues,
  };
}

export function formatReport(file: string, report: ValidationReport): string {
  const lines = [
    `${file}: ${report.ok ? 'OK' : 'ERRORS'} — valid tables: ${report.tables.join(', ') || '(none)'}`,
  ];
  for (const i of report.issues) {
    const where = [i.source, i.row === undefined ? undefined : `row ${i.row}`]
      .filter(Boolean)
      .join(' ');
    lines.push(`  ${i.severity.toUpperCase().padEnd(7)} ${where}: ${i.message}`);
  }
  return lines.join('\n');
}
