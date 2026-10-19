# McSheety — Implementation Plan

**Status:** Proposed technical specification / agent handoff  
**Version:** 0.2 (McSheety naming and MVP architecture)  
**Date:** 2026-10-17  
**Primary language:** TypeScript  
**Repository:** New monorepo  
**License:** TBD (suggest MIT if open source)

> **Project statement:** McSheety is a lightweight, serverless framework that lets an application use Excel workbooks stored in SharePoint as its data store. Each data worksheet represents a logical table. Column headers define fields; subsequent rows are records. Developers interact with the data through a typed client and a secure API instead of writing Microsoft Graph calls. The workbook can describe its schema and application authorization rules.
>
> **Key identity decision:** Use **Microsoft Entra ID delegated authentication** and **Microsoft Graph On-Behalf-Of (OBO)** access for version 1. AWS Cognito is a possible later adapter, not part of the MVP authentication path.

---

## 0. Instructions for implementation agents

This file is the working product/technical contract. Agents should:

1. Read the whole specification and identify dependencies before coding.
2. Implement **vertical slices**, not an enormous framework in one pass: first sign in -> API -> OBO -> Graph read; then CRUD; then policy enforcement; then CDK packaging and SDK.
3. Treat security, delegated SharePoint access, metadata trust, and concurrent writes as **blocking requirements**, not cleanup work.
4. Prefer official Microsoft Graph **v1.0** APIs. Do not silently substitute `/beta` endpoints.
5. Do not commit credentials, tokens, tenant secrets, workbook content, or employee information.
6. Use TypeScript strict mode, explicit contracts, small packages, dependency injection, and mocked Graph fixtures for unit tests.
7. Record tradeoffs and unexpected Graph behavior in `docs/decisions/` as ADRs. Update this document when findings change the contract.
8. Never claim transactional guarantees that the Excel APIs cannot provide.
9. Do not create a competing database for business records; an optional DynamoDB lock/coordination table is acceptable in later phases.
10. Before shipping any write capability, demonstrate authorization enforcement server-side and document how direct Excel editing bypasses it.

### Decisions already made

- Product name: **McSheety** (canonical capitalization; npm scope, package and domain availability to be verified before publication).
- Deployment: AWS serverless using CDK v2, Lambda, and API Gateway.
- Storage: Excel workbooks in SharePoint Online / Microsoft 365, accessed through Microsoft Graph.
- Auth for MVP: Microsoft Entra ID with delegated identity, not app-only workbook access.
- Authoritative app permission checks: the backend. UI permission hints are optional and never trusted.
- Form of distribution: npm packages and CDK construct.
- TypeScript-first, frontend-framework-agnostic core; React demo.

### Product naming and package conventions

Use **McSheety** as the visible product/brand name everywhere: docs, headings, UI copy, user-facing error messages and example applications. Consistently use the following machine-readable spellings:

| Context | Canonical identifier |
|---|---|
| Repository / workspace root | `mcsheety` |
| npm organization/scope (proposed) | `@mcsheety` |
| Core package | `@mcsheety/core` |
| Microsoft Graph adapter | `@mcsheety/adapter-graph` |
| Identity helper | `@mcsheety/auth-msal` |
| API package | `@mcsheety/api` |
| Frontend client | `@mcsheety/client` |
| AWS construct package | `@mcsheety/aws-cdk` |
| Optional CLI executable | `mcsheety` |
| TypeScript construct | `McSheetyApi` |
| Config/environment prefix | `MCSHEETY_` and `VITE_MCSHEETY_` |
| Internal Excel metadata table prefix | `mc_` |
| Internal metadata worksheet | `_mcsheety` |
| Planning / issue labels | `MC-001`, `MC-002`, etc. |

These identifiers are a **proposed contract**, not evidence that the npm scope or domains are available. Confirm publish permissions and scope ownership before package release; if the scope is unavailable, make one deliberate global namespace decision and update all examples at once. Avoid introducing mixed package names, environment variables or metadata prefixes. Treat persisted workbook metadata names as versioned conventions once adopters begin using them.

### Open decisions / research blockers

These **must be resolved by a live tenant proof of concept** before building advanced features:

- Confirm the tenant's admin-consent policy permits the required Microsoft Graph **delegated `Files.ReadWrite`** permission and API registration.
- Confirm Microsoft Graph workbook endpoints work for a workbook in the target SharePoint document library, using its **drive ID and item ID**, for each intended user type.
- Test whether a SharePoint **read-only user** can call Graph workbook read endpoints with a delegated token containing `Files.ReadWrite` scope (Graph documentation lists this scope even for read operations; SharePoint resource ACLs still apply). If not, redesign the reader path or roles; do not assume it works.
- Verify direct multi-user Excel edits, session persistence, row indexing, version metadata and conflict behavior experimentally.
- Determine whether workbook metadata sheets will be co-located with data (trusted-editor deployments only) or live in a **separate administrator-controlled workbook** (recommended for meaningful separation of duties).

---

## 1. Product vision, users, and limits

### Target uses

Small internal applications, administrative dashboards, inventory/configuration trackers, reference/lookup data, departmental workflows, and low-write-volume CRUD systems where organizations already manage SharePoint/Excel.

### Non-goals for v0.1

- Replacing PostgreSQL/SQL Server for high-volume applications.
- Relational joins executed remotely, SQL, foreign-key transactions, database migrations, or ACID transactions.
- Guaranteed safe simultaneous changes by application users and direct Excel editors.
- Real-time subscriptions or offline sync.
- Automatic enforcement of field/row permissions on users who open the workbook directly in Excel.
- A full no-code UI generator, mobile SDK, Google Sheets adapter, multi-tenant SaaS control plane, or Cognito federation.

### Operating assumptions

- One SharePoint workbook per application for the MVP; later support multiple workbooks.
- One **data worksheet** maps to one logical table.
- A worksheet contains **one actual Excel Table** with a header row, not an arbitrary used range. Use the Excel Table API for operations.
- Intended application data volumes are small. Initial configurable maximum: **2,000 records per table** and **100 records per API page** (product safeguards, not promises about Graph limits).
- Users are in the same Entra tenant, have Microsoft 365 identities, and have appropriate SharePoint access.
- App authorization restricts what users can do **through McSheety**. SharePoint/Excel access is independently administered.

---

## 2. Architecture

```text
Browser (sample React SPA)
   |
   | MSAL browser: Entra login + access token for McSheety API
   v
API Gateway HTTP API
   | JWT authorizer (issuer, audience, scopes)
   v
Node.js Lambda / McSheety API
   |- request validation
   |- authentication claims normalization (tid, oid, roles)
   |- policy engine (table, column, row)
   |- schema validation + type coercion
   |- read/filter/page + mutation orchestration
   |- Microsoft identity OBO token exchange (MSAL Node)
   |- Microsoft Graph adapter + workbook sessions
   v
Microsoft Graph v1.0
   v
SharePoint Online document library / .xlsx workbook
   |- Equipment sheet (tbl_Equipment)
   |- Employees sheet (tbl_Employees)
   |- Locations sheet (tbl_Locations)
   |- _schema sheet (mc_schema)
   |- _permissions sheet (mc_permissions)
   |- _columns sheet (mc_columns)
   `- _rowPolicies sheet (mc_rowPolicies)
```

### AWS resources (initial)

- **API Gateway HTTP API** with JWT authorizer for Entra access tokens, restrictive CORS and scoped routes.
- **Lambda, Node.js 22.x**, running the McSheety API package; no public Graph tokens exposed to browser.
- **AWS Secrets Manager** for Entra API app client secret (or certificate credential for production).
- **CloudWatch** structured logs, error/alarm metrics and request correlation.
- **CDK v2 construct** to provision/compose resources and return API URL and API audience details.
- No RDS, DynamoDB, SQS, Step Functions, or Cognito required in the first read-only vertical slice.

Optional later resources: DynamoDB coordination lease table, API throttling/abuse protection, distributed tracing, access log retention controls, and a dead-letter mechanism **only if an asynchronous design can safely preserve delegated user authority**.

### Trust boundaries

1. Browser token is intended for **McSheety API**, not Graph.
2. API validates issuer, audience, signature, lifetime, tenant, scopes and identity.
3. API exchanges the inbound access token using OBO for **Graph**; the downstream Graph token must not be returned to the browser.
4. The effective access decision requires **both** McSheety policy authorization **and** Microsoft/SharePoint authorization.
5. All permissions are checked again on each read and write, including row and column behavior.
6. A workbook that a user can edit directly is **not a reliable security boundary for policy metadata**. See Section 8.

---

## 3. Repository organization

```text
mcsheety/
├── implementation.md
├── README.md
├── LICENSE
├── package.json
├── pnpm-workspace.yaml
├── tsconfig.base.json
├── eslint.config.js
├── .gitignore
├── .env.example
├── packages/
│   ├── core/                 # domain model, parser, validators, policy engine
│   ├── adapter-graph/        # Graph workbook client and OBO-aware adapter
│   ├── api/                  # Lambda/http handlers and error mapping
│   ├── client/               # browser/Node typed REST client
│   ├── auth-msal/            # MSAL client helpers, OBO exchange
│   └── aws-cdk/             # McSheety CDK construct
├── examples/
│   ├── react-equipment/      # sample MSAL React + CRUD UI
│   └── cdk-equipment/        # deployable reference stack
├── fixtures/
│   └── equipment/            # synthetic workbook fixture + generated test data
├── tools/
│   └── mcsheety-cli/        # phase 3 CLI, optional
├── docs/
│   ├── decisions/
│   ├── security.md
│   ├── deployment.md
│   ├── workbook-format.md
│   └── troubleshooting.md
└── .github/workflows/        # lint, test, build (no real creds in CI)
```

**Suggested tooling:** pnpm workspace, TypeScript `strict: true`, Node 22, Vitest, Zod for wire/schema input validation, MSAL Browser/Node, CDK v2, `tsup` or `tsc` builds, ESLint, Prettier, and GitHub Actions. Dependencies are suggestions, not requirements to over-engineer.

### Responsibility split

| Package | Owns | Must not own |
|---|---|---|
| `@mcsheety/core` | Models, schema parsing, policy evaluation, normalization, types, errors | Graph, AWS, browser APIs |
| `@mcsheety/adapter-graph` | Graph transport, workbook table/session operations, pagination, translation | UI or app authorization decisions |
| `@mcsheety/auth-msal` | Incoming identity normalization, OBO Graph token acquisition | Table business logic |
| `@mcsheety/api` | Route handlers, authorization orchestration, request/response semantics | Frontend framework |
| `@mcsheety/client` | Typed API client, auth token provider, response/error handling | Microsoft Graph credential exchange |
| `@mcsheety/aws-cdk` | AWS deployment construct/configuration | Workbook parsing |

**Implementation order:** `core` interfaces and fixtures first, then working Graph/OBO spike, then API, client, CDK, demo.

---

## 4. Workbook conventions

Workbook example: `Equipment.xlsx`, in a specific SharePoint site/document library. Configure by **tenant ID, drive ID, item ID** after a one-time discovery step. Do not hardcode user-specific OneDrive `/me/drive` paths for site workbooks.

### Data sheet: `Equipment`

Excel worksheet `Equipment` contains a formal Excel Table named `tbl_Equipment`:

| id | name | status | locationId | departmentId | purchasePrice | createdAt |
|---|---|---|---|---|---:|---|
| eq-001 | Pump A | Active | loc-1 | dept-ops | 123.50 | 2026-10-16T14:05:00Z |
| eq-002 | Valve B | Inactive | loc-2 | dept-eng | 80.00 | 2026-10-16T14:10:00Z |

Guidelines:

- Logical table name = worksheet name; expected Excel Table name defaults to `tbl_<WorksheetName>` with configurable mapping where Excel naming rules require it.
- Exactly one data table per data worksheet for the MVP.
- Header cells are unique, case-sensitive and nonblank. Avoid spaces in field names initially (camelCase recommended).
- Every data table has an immutable, nonempty, unique **`id`** string; new IDs generated server-side (UUIDs or stable generated IDs). Never identify a logical record by Graph row index.
- Metadata worksheets start with `_` and are excluded from normal CRUD endpoints.
- Prefer raw values, not business-critical Excel formulas. Reject formula-bearing cells in managed fields until safe semantics are defined.
- Store `datetime` values as explicit ISO 8601 UTC strings; use `date` (`YYYY-MM-DD`) and `time` (`HH:mm:ss`) as timezone-free strings. Never silently convert wall-clock times to browser-local time.
- Define null/blank semantics in the schema; don't assume Excel cell blank, zero, and empty string mean the same thing.
- Manual data edits may invalidate required/unique/reference fields. Reads must surface validation problems without leaking unauthorized data.

### Metadata sheet: `_schema` (Excel Table `mc_schema`)

| table | column | type | required | default | unique | options |
|---|---|---|---|---|---|---|
| Equipment | id | uuid | true | generated | true | |
| Equipment | name | string | true | | false | |
| Equipment | status | enum | true | Active | false | Active\|Inactive |
| Equipment | locationId | string | true | | false | |
| Equipment | departmentId | string | true | | false | |
| Equipment | purchasePrice | number | false | | false | min=0 |
| Equipment | createdAt | datetime | true | now | false | |

Types for MVP: `string`, `number`, `boolean`, `date`, `time`, `datetime`, `uuid`, `enum`. Use deterministic parsing with explicit failures for invalid coercions. `options` should eventually have a structured, versioned syntax; the displayed delimiters are only an illustrative MVP convention.

### Metadata sheet: `_permissions` (Excel Table `mc_permissions`)

| table | role | read | create | update | delete |
|---|---|---|---|---|---|
| Equipment | viewer | true | false | false | false |
| Equipment | editor | true | true | true | false |
| Equipment | manager | true | true | true | false |
| Equipment | admin | true | true | true | true |

**Rules:** Default deny. No matching role/action = forbidden. Roles are verified Entra app roles or a trusted server-side mapping from Entra group IDs, not user-supplied labels. App role names are NOT the same as delegated Graph scopes.

### Metadata sheet: `_columns` (Excel Table `mc_columns`)

| table | column | readRoles | writeRoles |
|---|---|---|---|
| Equipment | id | * | |
| Equipment | name | * | editor,manager,admin |
| Equipment | status | * | editor,manager,admin |
| Equipment | locationId | * | editor,manager,admin |
| Equipment | departmentId | * | manager,admin |
| Equipment | purchasePrice | manager,admin | admin |
| Equipment | createdAt | * | |

- `*` = any authenticated role **that already has table-level read permission**.
- Blank role list = nobody can change/read that field through McSheety.
- Do not accept arbitrary column names from a PATCH without checking the schema and column policy.
- `id` cannot be modified by any client; `createdAt` is server-owned.
- Unknown or missing column policy = deny that column until configured. Keep explicit metadata for all fields.
- Redact unreadable columns on the server; never fetch-and-hide in the browser as an authorization measure.

### Metadata sheet: `_rowPolicies` (Excel Table `mc_rowPolicies`)

| table | action | role | column | operator | claim |
|---|---|---|---|---|---|
| Equipment | read | editor | departmentId | equals | departmentId |
| Equipment | update | editor | departmentId | equals | departmentId |
| Equipment | delete | admin | | allowAll | |

This declaratively expresses: editor may read/update records only when `record.departmentId === trustedUser.departmentId`. The department claim MUST be issued by a trusted Entra claim source or server-side directory mapping; do not assume such a claim exists automatically in tokens. If unavailable, disable this policy until the identity mapping is configured.

Policy semantics for MVP:

1. First grant the **table/action** permission.
2. Apply row restriction if an applicable role/action policy exists. A matching row rule must pass; an explicit `allowAll` rule bypasses row restriction for that role/action. For a role/action with *no row policy*, rows are not further restricted beyond table authorization. Document this explicitly.
3. For a user with multiple roles, a row is accessible if **any** role that grants the table action and its associated row rule permits access.
4. On create, evaluate the prospective row and prevent a user from creating records outside their allowed department. Admin-defined defaults may populate trusted fields; client values cannot override restricted server-managed values.
5. On update, check the **existing and resulting** rows; forbid a change that moves a record outside the user's authorized scope.
6. For delete, check the existing row before deleting.
7. Support `equals` and `allowAll` first. Add `in` / `notEquals` only with rigorous typed tests. Never evaluate JavaScript, spreadsheet expressions or arbitrary operators from policies.
8. Policy parsing errors are fatal for the affected resource: fail closed, don't silently skip invalid policy rows.

### Metadata versioning

Introduce optional `_mcsheety` sheet or separate trusted configuration with keys `formatVersion=1`, `schemaVersion`, `policyVersion`, and a stable workbook/application ID. A `mcsheety validate` command should identify malformed headers, duplicate IDs, unknown roles/columns, missing policies, and mismatched Excel table definitions.

---

## 5. Microsoft Entra delegated authentication

### Exact intended flow

```text
1. Browser uses MSAL Browser Authorization Code + PKCE to sign into Entra.
2. Browser requests an access token with scope for the McSheety API
   (example: api://<API_APP_ID>/access_as_user).
3. Browser calls McSheety REST API with Authorization: Bearer <API token>.
4. API Gateway validates JWT issuer, audience and configured scope.
5. Lambda additionally validates tenant and normalizes claims (oid, tid, roles).
6. Lambda uses MSAL Node confidential-client OBO with the inbound API
   access token as user assertion to obtain a Graph delegated access token.
7. Lambda calls Microsoft Graph with the Graph token on that user's behalf.
8. Both app policy AND Microsoft/SharePoint permission must permit operation.
9. API returns filtered data, never a Graph access token or credentials.
```

**Do not pass a Graph access token as the inbound McSheety API token.** It has the wrong audience for OBO. **Do not send an ID token** as the API bearer token.

### Entra app registrations (recommended)

1. **SPA client registration**: public client, SPA redirect URIs, MSAL Browser, authorization code + PKCE, configured to request the McSheety API delegated scope.
2. **McSheety API registration**: exposes the `access_as_user` scope; confidential client identity with credential held only server-side; configured with delegated Microsoft Graph permissions required for workbook operations (initially test `Files.ReadWrite`). Configure admin consent/approved app policies as appropriate.
3. **Roles**: define Entra **app roles** such as `viewer`, `editor`, `manager`, `admin`, assigned to actual users or groups for the API application. Validate role delivery in the access token intended for the API. If using group claims instead, map **immutable group object IDs** to application roles server-side and handle Entra group-overage cases.
4. **Tenant scope**: single-tenant app for MVP. Explicitly require the configured tenant ID and accepted issuer.

### API auth checks

- Prefer API Gateway HTTP API JWT authorizer for valid issuer/audience; configure route authorization scopes.
- Inside Lambda, use claims only after a validated authorizer context; for local testing use a real JWT verifier with trusted JWKS (not a decoded-but-unverified payload).
- Validate tenant, user OID, scopes, role assignment, and deny app-only/service-principal callers for these delegated routes.
- Use `tid + oid` as the stable user identifier; do not authorize by email address, display name or UPN.
- Deny by default for missing roles or unrecognized claims.
- OBO credential: secret in Secrets Manager for dev; prefer a certificate/managed rotation in production. Document IAM restrictions.
- Token caches must isolate principals and never mix users. Never log tokens or store them in DynamoDB/SQS/workbook.
- Handle conditional-access claims challenges (401/403 with safe error codes for reauthentication) and tenant consent failures distinctly.

### Graph permissions caveat

Microsoft Graph workbook **table read and write** endpoints currently document delegated `Files.ReadWrite`, and **application permissions are not supported** for those workbook endpoints. That is the principal reason the MVP uses OBO rather than a daemon with app-only permissions. Delegated Graph scopes do not override underlying SharePoint file ACLs.

### Cognito compatibility (later)

Cognito can eventually federate Entra for application sign-in, but a **Cognito JWT is not an Entra access token for the McSheety API** and cannot automatically be used as the OBO user assertion. A future Cognito mode needs its **own deliberate design** for acquisition of an appropriate Entra token; do not fake it by exchanging unrelated tokens.

---

## 6. Microsoft Graph workbook adapter

### Core interface (illustrative)

```ts
export interface WorkbookAdapter {
  listTables(ctx: RequestContext): Promise<TableDescriptor[]>;
  readRows(ctx: RequestContext, table: string): Promise<RawRecord[]>;
  appendRow(ctx: RequestContext, table: string, values: RawRecord): Promise<void>;
  replaceRow(ctx: RequestContext, table: string, id: string, values: RawRecord): Promise<void>;
  deleteRow(ctx: RequestContext, table: string, id: string): Promise<void>;
}

export interface RequestContext {
  tenantId: string;
  userId: string;
  correlationId: string;
  graphToken: string; // server-side only; never logged
}
```

The adapter receives an already acquired Graph token or a token callback so that auth and storage can be tested independently.

### Endpoint concepts

- Discover workbook: `GET /drives/{drive-id}/items/{item-id}` (metadata).
- List workbook tables: `GET /drives/{drive-id}/items/{item-id}/workbook/tables`.
- Read table rows: `GET /drives/{drive-id}/items/{item-id}/workbook/tables/{table-name}/rows` (page through results).
- Append row: `POST /drives/{drive-id}/items/{item-id}/workbook/tables/{table-name}/rows/add`.
- Update: resolve the current row position from an immutable `id`, then use the documented table-row PATCH or row-range operation after confirming the endpoint's exact behavior in a live integration test.
- Delete: resolve the current position from `id`, then use the table-row DELETE endpoint.
- Workbook session: create a **persistent** workbook session for multiple related operations; send `Workbook-Session-Id` for calls; close when possible. Never create a nonpersistent session for an intended write.

**Important:** Graph workbook row indices are zero-based positional indices. Sorting, inserting, deleting or editing inside Excel can change them. Never cache row index as a record identity. Resolve the `id` before mutation and re-check current data.

### Workbook sessions

- Start with a one-request read path for simplicity; introduce reusable persistent sessions for multi-call operations and writes, according to Graph best practices.
- Handle session expiry/recreation.
- Keep Graph operations **sequential per workbook** within each request.
- Document request timeout budgets, Graph `429` / `503`, `Retry-After`, 5xx handling and bounded exponential backoff with jitter.
- Do not blindly retry non-idempotent inserts after ambiguous timeouts: check by request/record ID before retry. A documented example mentions retrying 504, but blindly retrying can create duplicates if completion is uncertain.

### Querying behavior

- Read rows with Graph pagination (`$top` / `$skip` where supported); filter and sort **after** applying data/row-policy constraints server-side.
- Initial query operators: `eq`, `ne`, `in` for scalar types, with sorting and bounded pages. Keep a clear cap on how many raw rows can be scanned.
- Never return unauthorized records or sensitive columns in response totals, summaries, errors, or query results.
- Do not assume arbitrary Graph SQL-like server-side filtering is supported on Excel table rows.

### Spreadsheet safety

- Validate workbook file type, existence, table names, headers, max rows, unique IDs and metadata references.
- Prevent formula injection: validate unsafe leading formula characters on user-controlled strings, and confirm actual Graph cell-write semantics with tests. Escape safely or reject; do not strip legitimate negative numbers from numeric fields.
- Reject writes to cells with formulas or protected/system-owned values unless explicitly designed.
- Handle Excel raw numeric date serial values explicitly if encountered; prefer storing ISO text.
- Reject modifications that would corrupt schema/table shape or duplicate unique keys.

---

## 7. Application authorization model

**Default deny** at table and column levels. Evaluate permissions in the API for every operation.

### Order of checks for reads

1. Validate token/tenant/user/role and Graph access.
2. Load **trusted** policy and schema (see Section 8).
3. Require `table.read` for a verified role.
4. Read rows, enforce row visibility for that role/action.
5. Remove non-readable columns from each returned row.
6. Apply allowed user filter/sort/page; guard query columns used as oracles (filtering/sorting by hidden columns must be denied).
7. Return only filtered data and safe metadata.

### Order of checks for writes

1. Require appropriate table `create` / `update` / `delete` action.
2. Validate the target record exists or doesn't exist as appropriate.
3. Enforce row policy using trusted claims and old/new record values.
4. Validate every submitted field against `_schema` and `_columns` writeRoles.
5. Apply defaults and server-managed fields; verify unique/required/type rules.
6. Perform a concurrency check (Section 9).
7. Apply Graph mutation and confirm observed result where practical.
8. Emit an audit event with actor, action, table, record ID, timestamp, outcome, correlation ID, and changed **field names** (avoid sensitive before/after contents by default).

### Capabilities endpoint

`GET /v1/capabilities` returns only what the current user can do through McSheety: visible tables, allowed actions, readable/editable fields and optional row-policy descriptions suitable for UI hints. It must not expose another role's policy details or trusted claim secrets. It is informational; the API checks every actual request again.

### Role assignment example

- `viewer`: reads public columns and authorized rows.
- `editor`: reads/creates/updates permitted fields, perhaps only their department's records.
- `manager`: wider edit rights, sees limited financial fields.
- `admin`: CRUD granted by explicit config (not a magical implicit bypass of SharePoint).

---

## 8. Critical security limitation: workbook authors vs policy authors

**Excel sheet hiding, Excel worksheet protection, and application UI settings do not create a trustworthy authorization boundary against someone with file edit permissions.** A delegated user who can edit a workbook directly can potentially alter data and co-located policy sheets, bypassing McSheety API restrictions.

Consequences:

- If a user can open and edit the workbook, McSheety **cannot guarantee** row-level and column-level write restrictions on that user outside the API.
- If they can edit `_permissions` or `_schema`, they may be able to grant themselves API permissions, unless configuration is stored/verified separately.
- Even with a separate policy workbook, users who can edit data Excel directly can bypass application record/field rules. The separate policy workbook protects *API policy integrity*, not the underlying Excel data against direct editors.

### Deployment modes

**Mode A — Trusted collaborators (prototype):** Co-locate metadata sheets with data, but allow file write permission only for trusted editors/admins. Treat sheet policies as an **application convenience**, not a security boundary for direct editors. Publish the limitation prominently in docs and demo UI.

**Mode B — Trusted configuration (preferred for real role separation):** Store the `_schema`, `_permissions`, `_columns`, and `_rowPolicies` sheets in a **separate SharePoint workbook** with edit access granted only to policy administrators (or store immutable/versioned policy configuration in AWS managed by admins). Data workbook retains business records. The API uses a trusted policy-reader credential/access path, which requires an approved access design compatible with Graph and tenant policy. With the current OBO-only model, a user's delegated identity must be able to **read** that protected metadata without being able to **write** it. Administrators get separate SharePoint write access.

**Mode C — Strong API-only authorization (future architecture change):** End users do not have direct SharePoint edit rights, and a separate server-side authorized service mediates edits. This **cannot be achieved merely through the same user's OBO Graph write permissions**. It requires a different, explicitly designed storage/access mechanism (possibly non-Excel workbook APIs). Do not promise Mode C with the MVP architecture.

For MVP, implement Mode A with a clear warning and test Mode B before recommending McSheety for genuinely sensitive business data. **Never describe `_permissions` as enforcing SharePoint permissions.**

---

## 9. Concurrency, consistency, and failure handling

Excel is a shared document, not a transactional database. Graph recommends sequential operations rather than concurrent writes to one workbook.

### Initial consistency contract

- **Best effort** consistency for a single mutation performed through McSheety.
- No atomic multi-row or cross-table transactions.
- No guarantee against edits made simultaneously through Excel or other Graph clients.
- Client updates require an `If-Match` record revision (or explicit `version` field) with a conflict response when stale. A revision derived from row content is a detection mechanism, **not a true atomic compare-and-swap in Excel**.
- Immediately before update/delete, resolve row by immutable `id` and compare current values/revision. Because Graph does not offer a DB transaction across read/check/write, a race remains; surface this documented limitation.
- To serialize API-issued writes between Lambda instances, plan an **optional DynamoDB lease/lock by workbook ID** with lease expiry, owner token, and bounded wait. Do not add it until end-to-end Graph proof validates the need. It does not prevent direct Excel writes.
- If higher integrity is required, stop and reconsider the backing store rather than inventing an unreliable pseudo-transaction system.

### API error contract

Standard JSON error shape:

```json
{
  "error": {
    "code": "CONFLICT",
    "message": "The record changed since it was read.",
    "requestId": "request-correlation-id",
    "retryable": false
  }
}
```

Suggested status mapping:

| HTTP | Code | Meaning |
|---|---|---|
| 400 | `INVALID_REQUEST` | Bad payload/query |
| 401 | `UNAUTHENTICATED` | Missing/expired or invalid API identity |
| 403 | `FORBIDDEN` | App role or SharePoint permission denied |
| 404 | `NOT_FOUND` | Not visible or non-existent resource (avoid existence leak) |
| 409 | `CONFLICT` | Stale revision / in-progress edit / duplicate ID |
| 422 | `VALIDATION_FAILED` | Schema/type/required field violations |
| 429 | `RATE_LIMITED` | API or Graph throttling, bounded retry hint |
| 502 | `GRAPH_ERROR` | Safe summary of downstream error |
| 503 | `WORKBOOK_UNAVAILABLE` | Locked/unavailable/unrecoverable session |

Never expose raw Graph errors or file paths/tokens that might disclose sensitive details.

---

## 10. REST API v1

Prefix: `/v1`  
Auth: `Authorization: Bearer <Entra access token for McSheety API>`  
Content type: `application/json`  
Responses: deterministic schema, bounded pagination, request ID header, sanitized errors.

| Method | Route | Purpose |
|---|---|---|
| `GET` | `/v1/health` | Liveness; no workbook data/credentials |
| `GET` | `/v1/tables` | List authorized logical tables |
| `GET` | `/v1/schema/:table` | Return readable field metadata |
| `GET` | `/v1/capabilities` | Effective actions/columns for current user |
| `GET` | `/v1/tables/:table/rows` | List authorized rows (`limit`, `cursor`, `sort`, bounded filter) |
| `GET` | `/v1/tables/:table/rows/:id` | Fetch one record by immutable ID |
| `POST` | `/v1/tables/:table/rows` | Create row; server owns ID and defaults |
| `PATCH` | `/v1/tables/:table/rows/:id` | Update specific permitted fields |
| `DELETE` | `/v1/tables/:table/rows/:id` | Delete authorized row |

### Example list response

```json
{
  "data": [
    { "id": "eq-001", "name": "Pump A", "status": "Active" }
  ],
  "page": { "nextCursor": null },
  "meta": { "table": "Equipment" }
}
```

Avoid returning total counts that could reveal records hidden by row policy. Design cursors so they do not leak hidden record IDs. For small-table MVP, scanning the table, applying authorization, then paging results is acceptable within configured limits.

### Example update

```http
PATCH /v1/tables/Equipment/rows/eq-001
Authorization: Bearer <API_TOKEN>
Content-Type: application/json
If-Match: "revision-from-prior-read"

{ "status": "Inactive" }
```

Return updated permitted projection with a new record revision / `ETag`, subject to the consistency caveat above. Require `If-Match` on PATCH/DELETE in the first safe-write release.

### API design rules

- Do not allow arbitrary workbook IDs, Graph URLs or table paths from an untrusted HTTP request; bind routes to an admin-configured workbook and discovered/allowed table names.
- Disallow query/filter/sort against unreadable columns.
- Reject unknown request fields rather than quietly ignoring them.
- Default max request/response size, pages, filters and timeouts.
- Use opaque, non-guessable cursors or signed cursor objects if cursor semantics need state.
- Enforce CRUD field and row policies irrespective of any frontend controls.

---

## 11. Client package

The `@mcsheety/client` package is framework-agnostic and accepts an async **API token provider**, not a Microsoft Graph token.

```ts
import { createClient } from '@mcsheety/client';

const db = createClient({
  apiUrl: import.meta.env.VITE_MCSHEETY_API_URL,
  getToken: () => auth.acquireMcSheetyApiToken()
});

const equipment = await db.table('Equipment').findMany({
  where: { status: 'Active' },
  limit: 50
});

const record = await db.table('Equipment').findById('eq-001');

await db.table('Equipment').create({
  name: 'Pump D',
  status: 'Active',
  locationId: 'loc-1'
});

await db.table('Equipment').update('eq-001',
  { status: 'Inactive' },
  { ifMatch: record.revision }
);

await db.table('Equipment').delete('eq-002', {
  ifMatch: 'revision-from-prior-read'
});
```

This code illustrates desired API shape, not an implemented library. Normalize record revisions outside business data so `revision` is not accidentally persisted as an Excel column.

Future: generate TypeScript interfaces from `_schema`, introduce `@mcsheety/react` hooks, and integrate with TanStack Query/RTK Query. Do not require React as a dependency of the core or client packages.

---

## 12. CDK construct and deployment model

Desired consumer experience (illustrative, subject to feasibility):

```ts
import { McSheetyApi } from '@mcsheety/aws-cdk';

new McSheetyApi(stack, 'EquipmentBackend', {
  tenantId: process.env.ENTRA_TENANT_ID!,
  entraApiClientId: process.env.ENTRA_API_CLIENT_ID!,
  entraApiScope: 'access_as_user',
  entraCredentialSecretArn: process.env.ENTRA_SECRET_ARN!,
  workbook: {
    driveId: process.env.SHAREPOINT_DRIVE_ID!,
    itemId: process.env.SHAREPOINT_ITEM_ID!,
  },
  metadataMode: 'co-located',
  allowedOrigins: ['http://localhost:5173'],
});
```

### CDK responsibilities

- API Gateway HTTP API + JWT authorizer with **tenant-specific Entra issuer** and correct API audience.
- Lambda bundling/node runtime, minimal IAM permissions, env settings and Secrets Manager read permission for only the referenced secret.
- Correct API Gateway route mapping, CORS configuration and safe security headers where applicable.
- CloudWatch logs, structured redaction, metrics and alarms.
- Stage-specific parameters, outputs, tags, least-privilege IAM, stack removal policy choices.
- Option to import an existing secret rather than embedding credentials into synthesized CloudFormation.
- Clear README explaining the non-CDK steps: creating and consenting Entra app registrations, assigning app roles, configuring SharePoint file permissions, recording drive/item IDs.

**CDK cannot necessarily provision another organization's Entra tenant app registrations or permissions.** Treat those as documented prerequisites (future CLI/setup guide could help), not automatic AWS resources.

### Example environment configuration

```dotenv
ENTRA_TENANT_ID=...
ENTRA_SPA_CLIENT_ID=...
ENTRA_API_CLIENT_ID=...
ENTRA_API_SCOPE=api://<ENTRA_API_CLIENT_ID>/access_as_user
SHAREPOINT_DRIVE_ID=...
SHAREPOINT_ITEM_ID=...
MCSHEETY_POLICY_MODE=co-located
VITE_MCSHEETY_API_URL=...
```

`ENTRA_API_CLIENT_SECRET` must **not** be placed in a frontend `.env` or published package; store production credentials exclusively in the approved secret manager.

---

## 13. Local development, fixtures, and verification

### Expected developer commands

```bash
pnpm install
pnpm lint
pnpm typecheck
pnpm test
pnpm build
pnpm --filter react-equipment dev
pnpm --filter cdk-equipment cdk synth
```

Add commands as implementations exist; keep workspace package names consistent with scripts.

### Test harness

- `InMemoryWorkbookAdapter` for fast deterministic policy/API tests; do not confuse it with a shipping database adapter.
- Fixture workbook with Equipment, Locations, metadata and 3–10 representative rows.
- Mock MSAL OBO exchange and Graph HTTP with deterministic test data; no live credentials in CI.
- Optional separate integration tests, gated by environment, against a **dedicated non-production SharePoint site/workbook** and test Entra tenant/app registrations.
- Golden schema and permissions fixtures for default-deny behavior.
- Test negative and malicious cases: invalid JWT/audience, role spoofing, unauthorized field, unauthorized row, missing policies, formula input, duplicate id, deleted/reordered rows, conflicting writes, throttling, Graph consent failures.
- Ensure tests don't mutate user-owned production workbooks.

### Observability

Log at minimum: request/correlation ID, hashed or stable non-PII actor reference, table name, operation, duration, Graph retry count, status, and error category. Do not log bearer tokens, cookies, raw record bodies, secrets, or full Graph URLs containing tenant-specific file paths by default.

---

## 14. Delivery phases and agent task breakdown

Use independently reviewable pull requests. Agents may work in parallel only when interfaces/contracts are stabilized.

### Phase 0 — Research and executable authentication/storage spike (blocking)

**Deliverables**

- [ ] Configure test Entra SPA and API app registrations, delegated Graph permissions, admin consent, and `viewer`/`editor` roles.
- [ ] Create disposable SharePoint site/library workbook with an Excel Table and data sheet.
- [ ] Authenticate from an MSAL browser demo to an API-specific token.
- [ ] Validate inbound token and acquire Graph token through OBO in a Node Lambda/local API.
- [ ] Read actual SharePoint workbook table rows via `/drives/{driveId}/items/{itemId}/workbook`.
- [ ] Insert, update, delete a row and confirm persistence from the Excel UI.
- [ ] Repeat read test with SharePoint read-only access; record actual authorization outcomes.
- [ ] Verify per-user SharePoint permission denial and distinct API app-role denial.
- [ ] Experiment with two concurrent edits, direct Excel row sorting, and workbook sessions.
- [ ] Document exact endpoints, consent requirements, 403/429 behavior, request/response fixtures and constraints in an ADR.

**Exit criteria:** A real delegated user can read and (when authorized) write a workbook through the API with no Graph token delivered to browser. If not, stop and redesign before continuing.

### Phase 1 — Core and schema/policy engine

**Agent scope:** `packages/core`, fixtures and unit tests.

- [ ] Define typed `TableSchema`, `ColumnSchema`, `Policy`, `Identity`, `RawRecord`, `Revision`, `McSheetyError`.
- [ ] Implement spreadsheet metadata parsing with explicit format validation and errors.
- [ ] Implement type normalization and create/update validation.
- [ ] Implement table CRUD action permissions, field projection/edit permissions and row policies.
- [ ] Implement deterministic capability generation.
- [ ] Create policy fixtures for viewer/editor/manager/admin and negative security tests.

**Exit criteria:** Role+field+row permissions can be unit-tested entirely without Microsoft Graph/AWS.

### Phase 2 — Graph adapter, API, and SDK

**Agent scope:** `adapter-graph`, `auth-msal`, `api`, `client` after interfaces from Phase 1 land.

- [ ] Graph table discovery, row pagination, workbook session lifecycle, read conversion.
- [ ] Graph append/update/delete, immutable-id resolution and conflict detection.
- [ ] Robust retry and error mapping; bounded response and scan limits.
- [ ] Secure OBO middleware/token acquisition and identity normalization.
- [ ] REST routes from Section 10 with full policy enforcement and response redaction.
- [ ] Client API with injectable token provider and update revisions.
- [ ] Unit and controlled integration tests for all CRUD operations.

**Exit criteria:** Full authorized CRUD works through a local/dev API using a real non-production workbook; disallowed operations do not reach Graph mutations.

### Phase 3 — CDK, React example, operational hardening

**Agent scope:** `aws-cdk`, `examples`, docs, CI.

- [ ] CDK stack deploys API, JWT authorizer, Lambda, logging and least-privilege Secrets Manager access.
- [ ] README and `.env.example` guide for Entra setup and SharePoint discovery.
- [ ] React + MSAL example with inventory table, create/edit/delete, visible permissions and error states.
- [ ] Demonstrate field redaction and row restrictions with at least two test users.
- [ ] Set alarms, rate/size limits, CORS, smoke tests and cost notes.
- [ ] Add `mcsheety validate` and `mcsheety doctor` commands only if time allows; otherwise defer cleanly.

**Exit criteria:** Fresh developer can follow documented setup, deploy CDK, log into the sample app, and securely manage the Excel-backed records.

### Phase 4 — Post-MVP ideas (do not block v0.1)

- [ ] Separate trusted policy workbook configuration and security model (Mode B).
- [ ] TypeScript SDK code generation from workbook schema.
- [ ] Optional DynamoDB workbook write lease and enhanced conflict handling.
- [ ] Multiple workbooks, sites, and Graph adapter configuration.
- [ ] Adapter abstraction for local XLSX or Google Sheets, with explicitly documented differing auth/concurrency semantics.
- [ ] Fine-grained predicate DSL, audit sink, caching keyed by principal + workbook + policy version.
- [ ] Web UI or CLI to bootstrap workbook templates and publish an npm package.
- [ ] Cognito/Entra federation adapter, only after a tested identity/OBO design.
- [ ] MUI/React CRUD generator and starter templates.

---

## 15. Definition of done for v0.1

The MVP is complete only when all of these conditions are true:

1. A developer can configure a SharePoint workbook with at least one named Excel Table and the documented metadata tables.
2. A user signs in with Entra, acquires a token for the McSheety API, and the Lambda exchanges it through OBO for delegated Graph access.
3. Users cannot call API endpoints without valid McSheety API tokens, cannot spoof roles, and cannot access another tenant.
4. API `findMany`, `findById`, `create`, `update`, and `delete` work against a real SharePoint workbook using stable record IDs.
5. Table, column, and row permissions are enforced in Lambda. Unauthorized columns are absent from read responses; unauthorized writes fail.
6. Create/update/delete enforce type, required fields, IDs, ownership rules, uniqueness constraints within the documented consistency envelope, and policy defaults.
7. Changes made via the API become visible in SharePoint Excel; manual Excel changes are reflected in subsequent reads.
8. Stale update/delete requests return clear conflict errors, with limitations documented (not falsely described as full optimistic transactions).
9. All normal API errors are sanitized and consistent; API, Graph and tenant consent failures are diagnosable without exposing secrets.
10. CDK deploys the sample API; React sample exercises authenticated CRUD and role-driven UI affordances.
11. Unit tests cover policy boundaries; integration tests cover Graph OBO and workbook mutation with a non-production fixture.
12. Docs clearly state the implications of direct Excel access and limits of co-located policy sheets, performance, and concurrent edits.

---

## 16. Agent work protocol and recommended first tickets

To minimize wasted parallel effort, create tickets in this sequence:

**MC-001: Workspace skeleton + CI**  
Establish pnpm workspace, TypeScript strict build, test/lint/typecheck, package exports, versioning and shared error conventions.

**MC-002: Entra/Graph technical spike**  
Prove SPA -> API -> OBO -> SharePoint Excel real access. Document blockers. This is the top priority, even before extensive API work.

**MC-003: Example Excel workbook + metadata spec**  
Create fixture/test workbook, workbook validator contract and reproducible setup docs.

**MC-004: Core metadata parser + policy engine**  
Implement and test default-deny authorization, schema validation, field redaction, row policy evaluation.

**MC-005: Graph workbook adapter**  
Implement read and CRUD, immutable ID lookup, pagination, sessions, retry/error mapping and conflict tests.

**MC-006: REST Lambda API**  
Implement validation, enforcement, routes, sanitized errors, logging.

**MC-007: Client SDK**  
Implement small ergonomic API and revision handling, with mock API tests.

**MC-008: CDK construct + setup guide**  
Deploy API Gateway JWT authorizer, Lambda, secret access and examples.

**MC-009: React sample + security demonstration**  
Demonstrate two user roles and real SharePoint CRUD, forbidden fields and rows.

**MC-010: MVP acceptance and threat-model review**  
Review direct Excel ACL bypass, metadata policy trust, OBO credentials, row races, Graph throttling and hardening.

Each ticket should include: affected package, dependencies, test plan, actual verification performed, explicit non-goals, and any ADR-worthy findings. Never mark Microsoft Graph integration complete based only on mocks.

---

## 17. Reference documentation (verify again during implementation)

Official Microsoft documentation:

1. [Microsoft identity platform: OAuth 2.0 On-Behalf-Of flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-on-behalf-of-flow)
2. [Configure a web API that calls another web API](https://learn.microsoft.com/en-us/entra/identity-platform/scenario-web-api-call-api-app-configuration)
3. [Microsoft Graph: working with Excel](https://learn.microsoft.com/en-us/graph/api/resources/excel?view=graph-rest-1.0)
4. [Microsoft Graph: list table rows](https://learn.microsoft.com/en-us/graph/api/table-list-rows?view=graph-rest-1.0)
5. [Microsoft Graph: add table rows](https://learn.microsoft.com/en-us/graph/api/tablerowcollection-add?view=graph-rest-1.0)
6. [Microsoft Graph: update table row](https://learn.microsoft.com/en-us/graph/api/tablerow-update?view=graph-rest-1.0)
7. [Microsoft Graph: delete table row](https://learn.microsoft.com/en-us/graph/api/tablerow-delete?view=graph-rest-1.0)
8. [Microsoft Graph: best practices for Excel APIs](https://learn.microsoft.com/en-us/graph/workbook-best-practice)
9. [Microsoft Graph: driveItem resource](https://learn.microsoft.com/en-us/graph/api/resources/driveitem?view=graph-rest-1.0)
10. [Microsoft Graph: read driveItem metadata](https://learn.microsoft.com/en-us/graph/api/driveitem-get?view=graph-rest-1.0)
11. [MSAL Browser](https://learn.microsoft.com/en-us/entra/msal/javascript/browser/)
12. [MSAL Node](https://learn.microsoft.com/en-us/entra/msal/javascript/node/)
13. [Amazon API Gateway HTTP API JWT authorizers](https://docs.aws.amazon.com/apigateway/latest/developerguide/http-api-jwt-authorizer.html)

**Research caution:** Graph workbook API support, permission combinations, national cloud availability, row update behavior and concurrency semantics may evolve. The Phase 0 spike must validate current behavior in the actual target tenant and record exact results. This document is a plan, not proof that every endpoint works for every SharePoint/tenant configuration.

---

## 18. First executable milestone

**Milestone:** "Show me a SharePoint Excel row in React, and let an Entra-authenticated editor change it through AWS Lambda."

A successful first demonstration has:

- a workbook with `Equipment` and a formal `tbl_Equipment` Excel Table;
- one Entra-authenticated SPA user with effective viewer access, one with editor access;
- API Gateway protecting Lambda with an API-audience JWT;
- Lambda obtaining the user's Graph access token through OBO;
- a GET returning only the user's allowed records/fields;
- an editor PATCH updating a record using a stable `id` and a revision check;
- the changed record visible when opening the workbook in SharePoint;
- a viewer PATCH rejected by McSheety even if a token is valid;
- a clear demonstration that direct SharePoint edits are governed by separate SharePoint ACLs, not McSheety policies.

**Build this first. Expand only after it works end to end.**
