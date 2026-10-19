# ADR 0001: Toolchain

- **Status:** Accepted
- **Date:** 2026-10-17

## Context

implementation.md §3 suggests a pnpm workspace, strict TypeScript, Node 22, Vitest, Zod, MSAL, CDK v2, ESLint and Prettier. It also says these are "suggestions, not requirements to over-engineer".

## Decision

| Concern            | Choice                                                                                               | Notes                                                                                                                      |
| ------------------ | ---------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------- |
| Package manager    | pnpm 10.34 via Corepack (`packageManager` field)                                                     | pnpm 11 and 12 exist. 10.x is the conservative LTS-style line; upgrade deliberately later.                                 |
| Language           | TypeScript **6.0.x**, pinned `~6.0.3`                                                                | TypeScript 7 is out, but `typescript-eslint` 8.71 supports only `<6.1`. Revisit when typescript-eslint supports 7.         |
| TS settings        | `strict`, `noUncheckedIndexedAccess`, `exactOptionalPropertyTypes`, `verbatimModuleSyntax`, NodeNext | TS 6 defaults `types` to `[]`, so `types: ["node"]` is set explicitly.                                                     |
| Runtime            | Node 22 (Lambda `nodejs22.x`, arm64)                                                                 | Works locally on Node 24 too.                                                                                              |
| Tests              | Vitest 5, workspace sources aliased (no build needed)                                                | `@mcsheety/*` → `packages/*/src`.                                                                                          |
| Validation         | **No Zod.** `@mcsheety/core` validates all inputs explicitly.                                        | The core parser already rejects unknown fields and does strict, typed coercion. A second schema layer would duplicate it.  |
| JWT                | `jose` 6                                                                                             | Real signature, issuer, audience and lifetime verification in Lambda and in local runs.                                    |
| Identity           | `@azure/msal-node` 7 (OBO), `@azure/msal-browser` 5 (SPA, redirect flow)                             | MSAL 7 removes server text from error messages; classification uses `errorNo` and `subError` (see `auth-msal/src/obo.ts`). |
| Infra              | `aws-cdk-lib` 2.273, `NodejsFunction` with esbuild, CJS output, `@aws-sdk/*` external                | The bundle is about 290 KB.                                                                                                |
| Fixture generation | `exceljs` 4.4 (dev only)                                                                             | Writes real Excel Tables. Verified with openpyxl; Excel itself is still to be confirmed (notes.md).                        |

## Consequences

- Run `corepack enable` once per machine; CI uses `pnpm/action-setup`.
- CDK bundling shells out to `pnpm` (it detects `pnpm-lock.yaml`), so pnpm must be on PATH for `cdk synth`.
