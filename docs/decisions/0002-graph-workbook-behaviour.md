# ADR 0002: Microsoft Graph workbook endpoints and behaviour

- **Status:** Proposed. **Must be confirmed by the MC-002 live spike** (notes.md §2).
- **Date:** 2026-10-17

## Context

McSheety relies on Graph v1.0 Excel APIs with delegated (OBO) tokens. Several behaviours can't be established with mocks: row addressing, text coercion on write, delete semantics, session handling, and whether read-only SharePoint users can read through `Files.ReadWrite`. implementation.md says Graph integration is not complete until verified against a real workbook.

## Decision (current implementation in `packages/adapter-graph`)

| Operation       | Endpoint(s)                                                                                                                                                                         | Why                                                                                                                                                                         |
| --------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Discover tables | `GET …/workbook/tables?$select=name` (follows `@odata.nextLink`, Graph host only), then `GET …/tables/{t}/headerRowRange?$select=values` per table                                  | Header names are needed to validate the schema.                                                                                                                             |
| Read a table    | `GET …/tables/{t}/headerRowRange?$select=address,values` + `GET …/tables/{t}/dataBodyRange?$select=address,values,formulas`                                                         | Two calls give a consistent snapshot, exclude any totals row, and return formulas for formula-cell protection. The location comes from the header address.                  |
| Append          | `POST …/tables/{t}/rows/add` `{ values: [[…]] }`                                                                                                                                    | Documented v1.0. **Never retried.** On an ambiguous failure the API re-reads and looks for the generated `id`.                                                              |
| Update          | `GET …/worksheets/{ws}/range(address='A5:G5')` to verify the row is unchanged, then `PATCH …/worksheets/{ws}/range(address='B5:C5')` per run of changed cells, then re-read the row | Writes only changed cells, so formulas and concurrent edits to other columns aren't overwritten. A mismatch on verify returns 409 CONFLICT.                                 |
| Delete          | Verify as above, then `POST …/worksheets/{ws}/range(address='A5:G5')/delete` `{ "shift": "Up" }`                                                                                    | Avoids uncertainty over `rows/{index}` vs `rows/itemAt(index=)` addressing. **Not retried.**                                                                                |
| Sessions        | `POST …/workbook/createSession {persistChanges:true}` for every write request; `Workbook-Session-Id` on each call; `closeSession` in `finally` (errors ignored)                     | Reads are sessionless (one-request read path, per the spec).                                                                                                                |
| Text encoding   | String-typed cells are written with a leading `'`                                                                                                                                   | Stops Excel from turning `2026-10-16`, `1/2`, `TRUE` or `00123` into dates, numbers or booleans. Formula-leading characters (`= + - @`, tab, CR, LF) are rejected up front. |
| Retries         | GET/PATCH only: 429/503/504/network errors, max 4 attempts, full-jitter backoff, honours `Retry-After`, 20 s budget                                                                 | Non-idempotent calls are flagged `ambiguous` on 5xx or timeout.                                                                                                             |
| Errors          | 401 → UNAUTHENTICATED, 403 → FORBIDDEN (SharePoint), 404/409/423/5xx → WORKBOOK_UNAVAILABLE, 429 → RATE_LIMITED, otherwise GRAPH_ERROR                                              | Raw Graph messages never reach clients. Status, codes and the Graph request ID are logged.                                                                                  |

## Assumptions to verify (fill in from `/dev/probe` and manual tests)

| #   | Assumption                                                                                      | Result    | Evidence                                     |
| --- | ----------------------------------------------------------------------------------------------- | --------- | -------------------------------------------- |
| A1  | Read-only SharePoint user can read table ranges via OBO with `Files.ReadWrite`                  | _pending_ |                                              |
| A2  | User without site access gets 403 (mapped to FORBIDDEN "no access to this workbook")            | _pending_ |                                              |
| A3  | `'`-prefixed writes are stored as text and read back **without** the apostrophe                 | _pending_ | probe step "append with apostrophe encoding" |
| A4  | Raw write of `2026-10-16` becomes a date serial (justifies A3)                                  | _pending_ | probe                                        |
| A5  | An empty table reports one all-blank body row; `rows/add` fills or appends correctly            | _pending_ | probe "read probe table"                     |
| A6  | Range `PATCH` of a sub-range of a table row updates only those cells                            | _pending_ | probe "updateCells"                          |
| A7  | Range `delete` with `shift: Up` of a full table row removes the table row and shrinks the table | _pending_ | probe "range delete"                         |
| A8  | `dataBodyRange.formulas` returns `=…` for formula cells and values otherwise                    | _pending_ | add a formula cell to the probe table        |
| A9  | Persistent sessions work for the delegated user; `closeSession` succeeds                        | _pending_ | probe "persistent session"                   |
| A10 | Latency per request (about 7 + N sequential Graph calls for metadata) is acceptable             | _pending_ | measure with the browser network tab or logs |
| A11 | `/me?$select=department` works with `User.Read` and returns the admin-set value                 | _pending_ | probe                                        |
| A12 | Concurrent edit in Excel for the web between read and write gives CONFLICT, not a lost update   | _pending_ | manual test                                  |

## Consequences / follow-ups

- If A3 fails, switch `encodeCell` to set `numberFormat: [["@"]]` together with the values (append becomes append + PATCH).
- If A10 is too slow, add Graph JSON `$batch` (with a `dependsOn` chain to keep workbook calls sequential) or per-principal metadata caching keyed by the drive item `cTag` (implementation.md Phase 4).
- If A1 fails, read-only users cannot use the API under OBO + `Files.ReadWrite`, and the role model needs redesign before MC-005 is called complete.
