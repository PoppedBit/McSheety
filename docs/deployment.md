# Deploying McSheety

This page is for whoever deploys McSheety to AWS. The step-by-step tenant setup for the dev sandbox is in `notes.md`; this page is the reference.

## Prerequisites that CDK does not create

- **Entra app registrations:** a SPA (public client, PKCE) and an API (exposes `access_as_user`, issues **v2** tokens, defines app roles, has delegated Graph `Files.ReadWrite` and `User.Read`, with admin consent).
- **App role assignments** for users or groups.
- **A SharePoint workbook** in a document library, with metadata tables (see `workbook-format.md`), and file permissions that suit your deployment mode (see `security.md`).
- **Drive ID and item ID** for the workbook, from Graph Explorer.
- **A Secrets Manager secret** holding the API app's credential:
  - plain string = client secret
  - `{"clientSecret":"…"}`
  - `{"certificate":{"thumbprintSha256":"…","privateKey":"…"}}`

## Construct

```ts
import { McSheetyApi } from '@mcsheety/aws-cdk';

new McSheetyApi(stack, 'EquipmentBackend', {
  tenantId: '…',
  entraApiClientId: '…',
  entraApiScope: 'access_as_user',
  entraCredentialSecretArn: 'arn:aws:secretsmanager:…:secret:mcsheety/obo-AbCdEf', // complete ARN
  workbook: { driveId: '…', itemId: '…' },
  metadataMode: 'co-located',
  allowedOrigins: ['https://equipment.example.com'],
  attributes: {
    source: 'graph',
    properties: { departmentId: 'department' },
    valueMaps: { departmentId: { Operations: 'dept-ops' } },
  },
  // groupRoles: { '<group-object-id>': 'editor' },
  // throttle: { rateLimit: 20, burstLimit: 40 }, alarmAction: new SnsAction(topic),
});
```

It creates:

- an HTTP API with a JWT authorizer (tenant v2 issuer, audience = API client ID, route scope `access_as_user`) and explicit routes
- a Node.js 22 arm64 Lambda (esbuild bundle) with a log group
- read access to the one secret
- `$default` stage throttling and JSON access logs
- alarms for Lambda errors, Lambda throttles and API 5xx

## Lambda environment (set by the construct)

| Variable                                                             | Meaning                                    |
| -------------------------------------------------------------------- | ------------------------------------------ |
| `MCSHEETY_TENANT_ID`, `MCSHEETY_API_CLIENT_ID`, `MCSHEETY_API_SCOPE` | Token validation                           |
| `MCSHEETY_DRIVE_ID`, `MCSHEETY_ITEM_ID`                              | The one workbook this API serves           |
| `MCSHEETY_CREDENTIAL_SECRET_ARN`                                     | OBO credential                             |
| `MCSHEETY_POLICY_MODE`                                               | `co-located` (only mode in v0.1)           |
| `MCSHEETY_MAX_ROWS`                                                  | Rows per table safeguard (default 2000)    |
| `MCSHEETY_ATTRIBUTES`, `MCSHEETY_GROUP_ROLES`                        | JSON; trusted attributes and group mapping |

Configuration errors fail the cold start with a message listing every problem.

## Reference stack

```bash
pnpm build
pnpm --filter cdk-equipment cdk bootstrap   # once per account/region
pnpm --filter cdk-equipment synth           # review
pnpm --filter cdk-equipment cdk deploy      # deliberate, manual
```

The stack reads `.env` at the repository root and outputs `ApiUrl`, `TokenIssuer` and `TokenAudience`.

## Cost notes

At low request volumes, all components are pay-per-use: HTTP API requests, Lambda GB-seconds, CloudWatch log ingestion and storage, and one secret (about $0.40/month).

Each API request makes roughly 7 + N sequential Graph calls (see ADR 0002), so Lambda duration is dominated by Graph latency.
