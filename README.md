# McSheety

McSheety lets an application use **Excel workbooks in SharePoint as its data store**:

- each worksheet's Excel Table is a logical table
- headers are fields
- rows are records

Apps talk to a typed client and a secure REST API instead of Microsoft Graph. The workbook itself describes the schema and the app's table, column and row permissions.

- **Identity:** Microsoft Entra ID, delegated. The browser gets a token for the McSheety API only, and the API reaches Graph on the user's behalf through On-Behalf-Of (OBO).
- **Runtime:** AWS Lambda + API Gateway HTTP API, deployed with a CDK construct.
- **Scope:** small internal apps with low write volume. Not a database: no transactions, and direct Excel editors bypass McSheety's rules (see [docs/security.md](docs/security.md)).

> **Status: pre-release (v0.1 in progress).** Everything is unit-tested with mocks. Live Entra/Graph verification (MC-002) is pending; see [notes.md](notes.md).

## Try it locally (no Microsoft 365 needed)

```bash
corepack enable && pnpm install && pnpm build
pnpm --filter local-api demo                              # in-memory workbook, demo users
VITE_MCSHEETY_DEMO=1 pnpm --filter react-equipment dev    # http://localhost:5173
```

## Using the client

```ts
import { createClient } from '@mcsheety/client';

const db = createClient({
  apiUrl: import.meta.env.VITE_MCSHEETY_API_URL,
  getToken: () => auth.acquireMcSheetyApiToken(),
});

const { records, nextCursor } = await db
  .table('Equipment')
  .findMany({ where: { status: 'Active' }, sort: ['name'], limit: 50 });
const pump = await db.table('Equipment').findById('eq-001');
await db.table('Equipment').update('eq-001', { status: 'Inactive' }, { ifMatch: pump.revision });
```

## Packages

| Package                                             | Purpose                                                                                      |
| --------------------------------------------------- | -------------------------------------------------------------------------------------------- |
| [`@mcsheety/core`](packages/core)                   | Schema and policy parsing, default-deny authorization, validation, queries, adapter contract |
| [`@mcsheety/adapter-graph`](packages/adapter-graph) | Microsoft Graph v1.0 workbook adapter (sessions, retries, row verification)                  |
| [`@mcsheety/auth-msal`](packages/auth-msal)         | JWT verification, identity normalization, MSAL Node OBO, trusted attributes                  |
| [`@mcsheety/api`](packages/api)                     | REST API v1 service, HTTP router, Lambda handler                                             |
| [`@mcsheety/client`](packages/client)               | Typed browser/Node client                                                                    |
| [`@mcsheety/aws-cdk`](packages/aws-cdk)             | `McSheetyApi` construct                                                                      |
| [`mcsheety-cli`](tools/mcsheety-cli)                | `mcsheety validate <workbook.xlsx>`                                                          |

Examples:

- [`local-api`](examples/local-api): demo and Entra local server, plus the spike probe
- [`react-equipment`](examples/react-equipment): MSAL + CRUD sample
- [`cdk-equipment`](examples/cdk-equipment): reference stack

## Development

```bash
pnpm lint && pnpm typecheck && pnpm test && pnpm build
pnpm fixtures:build                           # writes fixtures/equipment/out/Equipment.xlsx
pnpm --filter react-equipment build
pnpm --filter cdk-equipment synth             # needs the MCSHEETY_* values in .env
```

## Documentation

- [implementation.md](implementation.md): product and technical specification
- [docs/workbook-format.md](docs/workbook-format.md): how to lay out a workbook
- [docs/security.md](docs/security.md): trust boundaries, guarantees and limits
- [docs/deployment.md](docs/deployment.md): Entra prerequisites, the construct, deploying
- [docs/troubleshooting.md](docs/troubleshooting.md): error codes and fixes
- [docs/decisions/](docs/decisions): architecture decision records
