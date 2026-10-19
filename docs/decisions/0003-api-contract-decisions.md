# ADR 0003: API contract and policy decisions beyond (or differing from) implementation.md

- **Status:** Proposed. Needs product-owner review (notes.md §4).
- **Date:** 2026-10-17

Each item says what implementation.md says, what was built, and why.

## Wire contract

1. **Record shape.** The spec's list example is flat (`{ "id", "name", … }`). Built: `{ "id", "revision", "fields": {…}, "issues"?: [...] }`.
   - Why: §11 asks for revisions to stay outside business data so they are never persisted as a column. A flat shape would collide with a real `revision` column.
   - `ETag` carries the same revision on single-record responses.
2. **Missing `If-Match` returns 428 `PRECONDITION_REQUIRED`**, a code added to the §9 table. `If-Match: *` is rejected the same way.
3. **`INTERNAL_ERROR` (500)** was added for unexpected exceptions. Their messages are never returned.
4. **`claimsChallenge`** was added to the error body. A 401 from a Conditional Access challenge also sets `WWW-Authenticate: Bearer error="insufficient_claims", claims="<base64>"`. Missing admin consent is a 403 with detail code `consent_required` (§5 asks for distinct handling).
5. **Queries:** `where` is a JSON object in the query string (`{"status":"Active","locationId":{"in":[…]}}`), and `sort` is `name,-createdAt`. Cursors are opaque base64url offsets bound to a fingerprint of the query. They never contain record IDs.
6. **Environment variables use the `MCSHEETY_` prefix** (per the §0 naming table), not the §12 example names `ENTRA_*` / `SHAREPOINT_*`. The CDK props keep the §12 names.

## Authorization semantics

7. **Per-role evaluation.** For writes, a single role must grant the action, pass the row rule (old and new row) and be allowed to write every submitted field. For reads, the columns returned for a row are the union over the roles whose row rule admits that row. This tightens §4 rule 3 so that one role's rows can't be combined with another role's columns.
8. **Several `equals` rules for the same role and action are ANDed.** `allowAll` overrides them. No rules means no row restriction (as in §4).
9. **A missing trusted attribute makes the rule fail closed.** §4 says to "disable this policy until the identity mapping is configured", which, read literally, would _widen_ access. It was implemented as deny instead.
10. **Defaults:** `generated` (string/uuid), `now` (datetime), `claim:<attribute>` (new), or a literal.
    - `claim:departmentId` lets an editor create records in their own department without being allowed to write `departmentId` (§4 rule 4).
    - The fixture adds an editor `create` row rule.
11. **Hidden columns are reported as unknown fields** (400) in writes, filters and sorts, so requests can't probe for columns the caller can't read.
12. **Invisible records return 404** for GET/PATCH/DELETE. FORBIDDEN is used only once the caller can see the record but may not change it.
13. **Revisions hash the caller's projection**, so a change to a hidden column doesn't change the revision a viewer sees (no oracle). A concurrent change to a hidden column is still detected at write time, because the adapter compares the full raw row before writing.
14. **Uniqueness checks see all rows**, including ones the caller can't read. A clash returns a generic 409. This can reveal that a hidden record with that value exists; that's documented and unavoidable without a separate index.
15. **Fixture `id` type is `string`** with `generated` defaults (UUIDs). The spec's example mixed `uuid` with IDs like `eq-001`.
16. **Rows with blank or duplicate IDs** are excluded from reads. Mutating a duplicated ID returns 409. `mcsheety validate` reports both.
17. **Blank semantics:** Graph reports blank cells as `""`, and they are normalized to `null` for every type. Writing `""` stores a blank, so empty strings and blanks are the same thing.
18. **Datetime:** Must be ISO 8601 with an explicit `Z` or offset, and is normalized to UTC.
    - Numeric serials in `datetime` columns are rejected (there's no timezone to apply).
    - `date` and `time` serials are converted.

## Operational

19. **The Lambda re-verifies the JWT** with jose even though API Gateway's authorizer already did. This is defense in depth, and the raw token is needed for OBO anyway. It also avoids API Gateway flattening array claims such as `roles` into strings.
20. **Metadata is reloaded on every request** (about 7 + N sequential Graph calls). Caching is deferred to Phase 4 (see ADR 0002, A10).
