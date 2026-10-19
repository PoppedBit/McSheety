# McSheety — setup notes for you

Things only you can do: tenant and AWS setup, decisions, and verification against real Microsoft 365. Work through them in order. Nothing here has been committed or deployed.

**Status (2026-10-17):** MC-001 and MC-003 through MC-009 are implemented and pass every local check (lint, typecheck, 136 unit tests, package builds, the Vite build, `cdk synth`, and a local run of the bundled Lambda). **MC-002 (the live Entra/Graph spike) is still blocking** and needs the steps below. Per implementation.md, Graph integration does not count as complete until it has been verified against a real workbook.

---

## 0. Local machine (2 minutes)

```bash
corepack enable          # puts pnpm on your PATH (pinned by package.json "packageManager")
pnpm install
pnpm lint && pnpm typecheck && pnpm test && pnpm build
```

Try the whole stack now, with no tenant (in-memory workbook, demo users):

```bash
pnpm --filter local-api demo                                # API on http://127.0.0.1:3000
VITE_MCSHEETY_DEMO=1 pnpm --filter react-equipment dev      # UI on http://localhost:5173
```

Switch between demo users to see roles in action. Eddie (editor, Operations) sees only Operations rows. Only the manager and admin see `purchasePrice`, and only the admin can delete. Demo tokens (`demo:<user>`) are accepted only by the local demo server.

I couldn't open the UI in a browser here. It typechecks and builds, and the API it calls was smoke-tested over HTTP, but please click through it once.

---

## 1. Microsoft 365 tenant (MC-002 prerequisite)

Microsoft 365 Developer Program sandbox, if you qualify. If you don't, a 30-day Business Basic/Standard trial works; with that, assign app roles to **users**, not groups, because group assignment needs Entra ID P1.

### 1a. SharePoint

1. Create a team site, e.g. **McSheety Dev**.
2. Run `pnpm fixtures:build` and upload `fixtures/equipment/out/Equipment.xlsx` to the site's **Documents** library.
   - Open it in Excel for the web once and **confirm it opens without a repair prompt**. I could only check it with openpyxl.
3. For the probe (step 4), add a worksheet `Probe` containing an Excel Table named **`tbl_McSheetyProbe`** with headers `id | text | num` and no rows.
   - It's deliberately not in `_schema`, so the API ignores it.
4. Pick three test users:
   - **editor**: site _Members_ (edit)
   - **viewer**: site _Visitors_ (read-only). This is the key test.
   - **outsider**: no site access

### 1b. App registration: "McSheety API (dev)"

- Single tenant.
- **Expose an API**: set the Application ID URI to `api://<api-client-id>` and add the scope `access_as_user` ("Admins and users" can consent).
- **Manifest**: set `"requestedAccessTokenVersion": 2`. Required: the API only accepts v2 issuers (`https://login.microsoftonline.com/<tenant>/v2.0`).
- **App roles**: `viewer`, `editor`, `manager`, `admin`, with allowed member types Users/Groups.
- **API permissions**: Microsoft Graph, _delegated_, `Files.ReadWrite` and `User.Read`. Then **Grant admin consent**.
- **Certificates & secrets**: create a client secret. It goes only in your local `.env` (and later AWS Secrets Manager), never in the frontend.
- **Enterprise applications → McSheety API (dev) → Users and groups**: assign roles, e.g. viewer → viewer user, editor → editor user, admin → you.

### 1c. App registration: "McSheety SPA (dev)"

- Single tenant. **Authentication → Single-page application** redirect URI: `http://localhost:5173`.
- **API permissions**: _My APIs_ → McSheety API (dev) → `access_as_user`. Then grant admin consent.
- Optional: on the API registration, **Expose an API → Authorized client applications**, add the SPA client ID.

### 1d. Record IDs (Graph Explorer, signed in as yourself)

```text
GET https://graph.microsoft.com/v1.0/sites/<tenant>.sharepoint.com:/sites/<site-name>      -> id
GET https://graph.microsoft.com/v1.0/sites/{site-id}/drives                                -> Documents drive id
GET https://graph.microsoft.com/v1.0/drives/{drive-id}/root/children                       -> Equipment.xlsx item id
```

### 1e. `.env` (repository root, git-ignored)

Copy `.env.example` to `.env` and fill in:

| Variable                                 | Value                                  |
| ---------------------------------------- | -------------------------------------- |
| `MCSHEETY_TENANT_ID`                     | Directory (tenant) ID                  |
| `MCSHEETY_API_CLIENT_ID`                 | API registration client ID             |
| `MCSHEETY_CLIENT_SECRET`                 | API client secret (local only)         |
| `MCSHEETY_DRIVE_ID` / `MCSHEETY_ITEM_ID` | From 1d                                |
| `MCSHEETY_ATTRIBUTES`                    | See below                              |
| `VITE_ENTRA_TENANT_ID`                   | Tenant ID                              |
| `VITE_ENTRA_SPA_CLIENT_ID`               | SPA client ID                          |
| `VITE_MCSHEETY_API_SCOPE`                | `api://<api-client-id>/access_as_user` |
| `VITE_MCSHEETY_API_URL`                  | `http://localhost:3000`                |

**Row policies need a trusted department.** Entra doesn't put one in tokens, so McSheety reads the user's admin-managed `department` from Graph `/me` and maps it to the IDs used in the workbook. In Entra admin center → Users, set _Department_ on the editor (e.g. `Operations`), then set:

```dotenv
MCSHEETY_ATTRIBUTES={"source":"graph","properties":{"departmentId":"department"},"valueMaps":{"departmentId":{"Operations":"dept-ops","Engineering":"dept-eng","Finance":"dept-fin"}}}
```

Without this, editors see no Equipment rows. That's deliberate: missing attributes fail closed.

---

## 2. Run the spike locally (MC-002), no AWS needed

```bash
pnpm build
pnpm --filter local-api entra        # real JWT verification + OBO + Graph, reads ../../.env
pnpm --filter react-equipment dev    # sign in with Microsoft
```

Then work through the implementation.md §14 Phase 0 checklist and **record what actually happens** in `docs/decisions/0002-graph-workbook-behaviour.md`:

1. Sign in as each test user (editor, viewer, outsider, admin) and note what each sees and gets.
2. **Key question:** can the _read-only_ SharePoint user (Visitors) read rows through OBO with `Files.ReadWrite`? If not, stop and redesign roles (see implementation.md §0).
3. Have the outsider sign in. Expect a 403 "does not have access to this workbook" from SharePoint, distinct from an app-role denial.
4. As admin, open `http://127.0.0.1:3000/dev/probe` with your bearer token. The easiest way is from the browser devtools console on the app page:
   ```js
   // In the app's devtools: copy a token from any API request's Authorization header, then:
   fetch('http://127.0.0.1:3000/dev/probe', { headers: { Authorization: 'Bearer <token>' } })
     .then((r) => r.json())
     .then(console.log);
   ```
   Paste the JSON into ADR 0002. It checks the assumptions I couldn't verify:
   - `'`-prefix text encoding
   - empty-table placeholder rows
   - patching only the changed cells
   - stale-write detection
   - range delete with shift-up
   - `/me` department
5. Edit a row in Excel for the web while the app is open, then save from the app. Expect **409 CONFLICT**.
6. Sort the sheet in Excel, then edit a row in the app. The right row must change, because records are found by `id`, not by position.

If the probe shows the apostrophe is stored literally (`'2026-10-16`), tell me. The fix is isolated to `packages/adapter-graph/src/encoding.ts`, which would switch to setting the `@` number format.

---

## 3. AWS (MC-008). You deploy; I only ran `cdk synth`

1. Store the OBO credential in Secrets Manager. Plain secret string, or JSON `{"clientSecret":"..."}` or `{"certificate":{"thumbprintSha256":"...","privateKey":"-----BEGIN PRIVATE KEY-----..."}}`:
   ```bash
   aws secretsmanager create-secret --name mcsheety/dev/obo --secret-string '<client secret>'
   ```
   Put the **complete** ARN (with the 6-character suffix) in `.env` as `MCSHEETY_CREDENTIAL_SECRET_ARN`.
2. `pnpm build && pnpm --filter cdk-equipment cdk bootstrap` (once per account/region).
3. Review: `pnpm --filter cdk-equipment synth`. Then, when **you** decide: `pnpm --filter cdk-equipment cdk deploy`.
4. Set `VITE_MCSHEETY_API_URL` to the `ApiUrl` output. For a hosted SPA, add its origin to `MCSHEETY_ALLOWED_ORIGINS` and to the SPA registration's redirect URIs.

---

## 4. Decisions for you

- **License**: the spec says TBD/MIT. No `LICENSE` file has been added yet.
- **npm scope `@mcsheety`**: unverified. All packages are `"private": true` until you confirm.
- **Review my deviations from the spec** in `docs/decisions/0003-api-contract-decisions.md`. In short:
  - Records are returned as `{ id, revision, fields }`, not flat.
  - A missing `If-Match` returns 428.
  - Env vars use the `MCSHEETY_` prefix.
  - Editors' new records get their department from a `claim:departmentId` default, and the fixture adds a `create` row rule for editors.
- **implementation.md** is untouched (and untracked). If you accept the ADRs, I can fold them back into the spec.
- **Commit**: everything is uncommitted on `master`. Branch and commit when you're ready.

## 5. Not done yet (by design or deferred)

- MC-002 live verification (above), and MC-010 threat-model sign-off after the spike (`docs/security.md` has the checklist).
- Mode B (separate policy workbook), DynamoDB write lease, codegen, multi-workbook: post-MVP per spec.
- `mcsheety validate` CLI: done for local `.xlsx` files (`pnpm --filter mcsheety-cli validate "$PWD/path/to/Workbook.xlsx"`; use an absolute path). `mcsheety doctor` is deferred; `/dev/probe` covers its purpose for now.
