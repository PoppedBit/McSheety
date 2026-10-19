import type { WorkbookAdapter, WorkbookContext } from './adapter.js';
import type { ParseOptions, ParseResult } from './metadata.js';
import { METADATA_TABLES, parseWorkbookModel } from './metadata.js';
import type { RawTable } from './types.js';

/**
 * Reads the metadata tables through the caller's delegated access and parses them.
 * Co-located metadata (Mode A) is only an application convenience: anyone who can edit the
 * workbook can edit these tables (implementation.md §8).
 */
export async function loadWorkbookModel(
  adapter: WorkbookAdapter,
  ctx: WorkbookContext,
  options: ParseOptions = {},
): Promise<ParseResult> {
  const tables = await adapter.listTables(ctx);
  const read = async (name: string): Promise<RawTable | undefined> =>
    tables.has(name) ? adapter.readTable(ctx, name) : undefined;

  // Sequential on purpose: Graph recommends against concurrent requests to one workbook.
  const config = await read(METADATA_TABLES.config);
  const schema = await read(METADATA_TABLES.schema);
  const permissions = await read(METADATA_TABLES.permissions);
  const columns = await read(METADATA_TABLES.columns);
  const rowPolicies = await read(METADATA_TABLES.rowPolicies);

  const dataTables = new Map([...tables].filter(([name]) => !name.startsWith('mc_')));
  return parseWorkbookModel(
    {
      config,
      schema,
      permissions,
      columns,
      rowPolicies,
      dataTables,
    },
    options,
  );
}
