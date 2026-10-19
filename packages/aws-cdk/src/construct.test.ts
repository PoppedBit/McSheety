import { App, Stack } from 'aws-cdk-lib';
import { Match, Template } from 'aws-cdk-lib/assertions';
import { Code } from 'aws-cdk-lib/aws-lambda';
import { describe, expect, it } from 'vitest';
import type { McSheetyApiProps } from './index.js';
import { McSheetyApi } from './index.js';

type Resources<P> = Record<string, { Properties: P }>;

const SECRET = 'arn:aws:secretsmanager:us-east-1:123456789012:secret:mcsheety/obo-AbCdEf';
const props: McSheetyApiProps = {
  tenantId: '11111111-1111-1111-1111-111111111111',
  entraApiClientId: '22222222-2222-2222-2222-222222222222',
  entraCredentialSecretArn: SECRET,
  workbook: { driveId: 'b!drive', itemId: '01ITEM' },
  allowedOrigins: ['http://localhost:5173'],
  attributes: { source: 'graph', properties: { departmentId: 'department' } },
  code: Code.fromInline('exports.handler = async () => ({ statusCode: 200 });'),
};

function synth(overrides: Partial<McSheetyApiProps> = {}) {
  const stack = new Stack(new App(), 'Test', {
    env: { account: '123456789012', region: 'us-east-1' },
  });
  const construct = new McSheetyApi(stack, 'Equipment', { ...props, ...overrides });
  return { template: Template.fromStack(stack), construct };
}

describe('McSheetyApi', () => {
  const { template, construct } = synth();

  it('configures a tenant-specific Entra JWT authorizer with the API audience', () => {
    template.hasResourceProperties('AWS::ApiGatewayV2::Authorizer', {
      AuthorizerType: 'JWT',
      IdentitySource: ['$request.header.Authorization'],
      JwtConfiguration: {
        Issuer: 'https://login.microsoftonline.com/11111111-1111-1111-1111-111111111111/v2.0',
        Audience: ['22222222-2222-2222-2222-222222222222'],
      },
    });
    expect(construct.issuer).toContain('/v2.0');
  });

  it('protects every data route with the scope, leaving only health open', () => {
    const routes = template.findResources('AWS::ApiGatewayV2::Route') as Resources<{
      RouteKey: string;
      AuthorizationType: string;
      AuthorizationScopes?: string[];
    }>;
    const byKey = Object.fromEntries(
      Object.values(routes).map((r) => [r.Properties.RouteKey, r.Properties]),
    );
    expect(Object.keys(byKey).sort()).toEqual([
      'DELETE /v1/tables/{table}/rows/{id}',
      'GET /v1/capabilities',
      'GET /v1/health',
      'GET /v1/schema/{table}',
      'GET /v1/tables',
      'GET /v1/tables/{table}/rows',
      'GET /v1/tables/{table}/rows/{id}',
      'PATCH /v1/tables/{table}/rows/{id}',
      'POST /v1/tables/{table}/rows',
    ]);
    for (const [key, route] of Object.entries(byKey)) {
      if (key === 'GET /v1/health') expect(route.AuthorizationType).toBe('NONE');
      else {
        expect(route.AuthorizationType).toBe('JWT');
        expect(route.AuthorizationScopes).toEqual(['access_as_user']);
      }
    }
  });

  it('restricts CORS to the configured origins and exposes revision headers', () => {
    template.hasResourceProperties('AWS::ApiGatewayV2::Api', {
      CorsConfiguration: Match.objectLike({
        AllowOrigins: ['http://localhost:5173'],
        AllowHeaders: ['Authorization', 'Content-Type', 'If-Match'],
        ExposeHeaders: Match.arrayWith(['ETag']),
      }),
    });
  });

  it('runs Node.js 22 with configuration in environment variables and no secret values', () => {
    template.hasResourceProperties('AWS::Lambda::Function', {
      Runtime: 'nodejs22.x',
      Environment: {
        Variables: Match.objectLike({
          MCSHEETY_TENANT_ID: props.tenantId,
          MCSHEETY_DRIVE_ID: 'b!drive',
          MCSHEETY_CREDENTIAL_SECRET_ARN: SECRET,
          MCSHEETY_ATTRIBUTES: '{"source":"graph","properties":{"departmentId":"department"}}',
        }),
      },
    });
    expect(JSON.stringify(template.toJSON())).not.toMatch(/MCSHEETY_CLIENT_SECRET/);
  });

  it('grants read on exactly the referenced secret', () => {
    const policies = template.findResources('AWS::IAM::Policy') as Resources<{
      PolicyDocument: { Statement: { Action: unknown; Resource: unknown }[] };
    }>;
    const statements = Object.values(policies).flatMap(
      (p) => p.Properties.PolicyDocument.Statement,
    );
    const secretStatements = statements.filter((s) =>
      JSON.stringify(s.Action).includes('secretsmanager'),
    );
    expect(secretStatements).toHaveLength(1);
    expect(secretStatements[0]?.Resource).toBe(SECRET);
    expect(JSON.stringify(statements)).not.toContain('"*"');
  });

  it('throttles the stage, writes access logs and creates alarms', () => {
    template.hasResourceProperties('AWS::ApiGatewayV2::Stage', {
      StageName: '$default',
      DefaultRouteSettings: Match.objectLike({ ThrottlingRateLimit: 20, ThrottlingBurstLimit: 40 }),
      AccessLogSettings: Match.objectLike({ Format: Match.stringLikeRegexp('requestId') }),
    });
    template.resourceCountIs('AWS::CloudWatch::Alarm', 3);
    template.resourceCountIs('AWS::Logs::LogGroup', 2);
  });

  it('rejects unsafe or incomplete configuration', () => {
    expect(() => synth({ allowedOrigins: ['*'] })).toThrow(/invalid origin/);
    expect(() => synth({ allowedOrigins: [] })).toThrow(/at least one origin/);
    expect(() => synth({ tenantId: 'contoso.onmicrosoft.com' })).toThrow(/tenantId must be a GUID/);
    expect(() =>
      synth({
        entraCredentialSecretArn: 'arn:aws:secretsmanager:us-east-1:123456789012:secret:obo',
      }),
    ).toThrow(/complete Secrets Manager ARN/);
  });
});
