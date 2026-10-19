#!/usr/bin/env node
/** mcsheety CLI. Currently: `mcsheety validate <workbook.xlsx>` for local files. */
import { formatReport, validateTables } from './validate.js';
import { readWorkbookTables } from './xlsx.js';

const [command, file] = process.argv.slice(2);

if (command !== 'validate' || !file) {
  console.error('Usage: mcsheety validate <workbook.xlsx>');
  process.exit(2);
}

try {
  const report = validateTables(await readWorkbookTables(file));
  console.log(formatReport(file, report));
  process.exit(report.ok ? 0 : 1);
} catch (error) {
  console.error(`Could not read ${file}: ${(error as Error).message}`);
  process.exit(2);
}
