/**
 * Reference stack for the Equipment sample. Configuration comes from the repository .env
 * (see .env.example / notes.md). `pnpm --filter cdk-equipment synth` builds the template;
 * deploying is a deliberate manual step: `pnpm --filter cdk-equipment cdk deploy`.
 */
import { App, CfnOutput, Stack } from 'aws-cdk-lib';
import { McSheetyApi } from '@mcsheety/aws-cdk';

function required(name: string): string {
  const value = process.env[name];
  if (!value) throw new Error(`Set ${name} in the repository .env (see notes.md).`);
  return value;
}

const app = new App();
const stack = new Stack(app, 'McSheetyEquipment', {
  env: {
    ...(process.env['CDK_DEFAULT_ACCOUNT'] ? { account: process.env['CDK_DEFAULT_ACCOUNT'] } : {}),
    region: process.env['CDK_DEFAULT_REGION'] ?? 'us-east-1',
  },
  description: 'McSheety Equipment sample API',
});

const api = new McSheetyApi(stack, 'EquipmentBackend', {
  tenantId: required('MCSHEETY_TENANT_ID'),
  entraApiClientId: required('MCSHEETY_API_CLIENT_ID'),
  entraApiScope: process.env['MCSHEETY_API_SCOPE'] ?? 'access_as_user',
  entraCredentialSecretArn: required('MCSHEETY_CREDENTIAL_SECRET_ARN'),
  workbook: { driveId: required('MCSHEETY_DRIVE_ID'), itemId: required('MCSHEETY_ITEM_ID') },
  metadataMode: 'co-located',
  allowedOrigins: (process.env['MCSHEETY_ALLOWED_ORIGINS'] ?? 'http://localhost:5173')
    .split(',')
    .map((o) => o.trim()),
  ...(process.env['MCSHEETY_ATTRIBUTES']
    ? { attributes: JSON.parse(process.env['MCSHEETY_ATTRIBUTES']) as never }
    : {}),
});

new CfnOutput(stack, 'ApiUrl', {
  value: api.url,
  description: 'Set VITE_MCSHEETY_API_URL to this value.',
});
new CfnOutput(stack, 'TokenIssuer', { value: api.issuer });
new CfnOutput(stack, 'TokenAudience', { value: api.audience });
