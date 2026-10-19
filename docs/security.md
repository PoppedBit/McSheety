# McSheety security model

This page is for the people who deploy McSheety and review it for MC-010.

## The one-sentence version

McSheety enforces table, column and row permissions **for requests made through its API**. Anyone who can open the workbook in SharePoint or Excel is governed only by SharePoint permissions, and can change any cell, **including the permission sheets** in co-located mode.

## Trust boundaries

1. **Browser → API.** The browser holds an Entra access token whose audience is the McSheety API (`access_as_user`). It never receives a Graph token.
2. **API Gateway.** The JWT authorizer checks the tenant-specific v2 issuer, the audience and the scope on every data route. `/v1/health` is the only unauthenticated route, and it touches no workbook.
3. **Lambda.** It re-verifies the token with jose (signature, issuer, audience, lifetime) and rejects:
   - other tenants
   - app-only tokens (`idtyp=app` or no `scp`)
   - tokens without the API scope

   The principal is `tid + oid`, never email or UPN.

4. **OBO.** The inbound token is the user assertion for MSAL Node OBO. The Graph token stays inside the Lambda:
   - It's never logged or returned.
   - MSAL caches OBO tokens keyed by assertion, so tokens are never shared across principals.
5. **Graph / SharePoint.** Every call uses the _user's_ delegated access. The effective permission is McSheety policy **and** SharePoint permission.

## What the API guarantees (through the API only)

- **Default deny.** Unknown roles, missing permission rows and missing column policies grant nothing. Role names are case-sensitive.
- **Columns** the caller can't read are removed on the server, and are treated as unknown in writes, filters and sorts.
- **Rows** outside the caller's scope are invisible (404). Totals aren't returned, and cursors are opaque offsets.
- **Writes:** one role must authorize the action, every field, and the old and new row.
  - `id` and server-owned fields can't be set by clients.
  - Formula-leading text is rejected, and cells holding formulas can't be overwritten.
- **Policy parsing fails closed:**
  - A malformed metadata header or unsupported format version disables everything.
  - A bad row in a policy table disables that logical table.
- **Errors are sanitized.** Graph messages, file paths and tokens never reach clients. Logs carry status and codes plus a hashed actor reference.

## What it cannot guarantee

| Limitation                                                    | Why                                                       | Mitigation                                                                                                                                      |
| ------------------------------------------------------------- | --------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| Direct Excel editors bypass every McSheety rule               | Delegated users with edit rights can edit the file itself | Give file edit rights only to trusted editors (Mode A). A separate policy workbook (Mode B, post-MVP) protects _policy integrity_ but not data. |
| Co-located policy sheets can be edited by any workbook editor | Same as above                                             | Mode B, or restrict workbook edit rights to admins                                                                                              |
| No transactions; check-then-write races remain                | Graph offers no compare-and-swap                          | Row re-verification right before the write narrows the window. Revisions return 409. A DynamoDB lease is optional later.                        |
| Uniqueness can reveal that a hidden duplicate exists          | The check sees all rows                                   | Generic 409 message; documented                                                                                                                 |
| Row policies depend on trusted attributes                     | Entra doesn't issue `department` by default               | Admin-managed Graph `/me` properties or a static map; a missing attribute is denied                                                             |

## Secrets and configuration

- The OBO credential lives in Secrets Manager. The construct grants read on that one ARN only, and the Lambda refreshes it hourly.
  - Prefer a certificate in production.
  - If the secret uses a customer-managed KMS key, grant `kms:Decrypt` to the function role separately.
- No secret values go into CloudFormation, environment variables or the frontend. `MCSHEETY_CLIENT_SECRET` is for local runs only.
- CORS lists exact origins only; wildcards are rejected at synth time.
- Access logs omit IPs, user agents and query strings, since filters may contain record values.

## MC-010 review checklist

- [ ] Spike results recorded (ADR 0002), especially read-only access (A1) and concurrent edits (A12)
- [ ] Workbook edit rights limited to trusted users; the Mode A banner is visible in the app
- [ ] App roles assigned only to intended users; no group-overage reliance
- [ ] Department (or other attribute) values are admin-managed in Entra
- [ ] Secret rotation procedure tested (rotate, then wait at most 1 h or redeploy)
- [ ] Throttling limits suit tenant Graph quotas; alarms routed to an owner
- [ ] Log retention set to policy; no tokens or record bodies in logs (spot-check CloudWatch)
