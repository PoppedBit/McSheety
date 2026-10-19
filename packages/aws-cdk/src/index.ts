/**
 * McSheetyApi CDK construct (implementation.md §12): HTTP API with an Entra JWT authorizer,
 * a Node.js 22 Lambda running @mcsheety/api, least-privilege access to one existing secret,
 * structured logs, access logs, throttling and alarms.
 *
 * Not provisioned here (documented prerequisites): Entra app registrations, admin consent,
 * app role assignments, SharePoint permissions and the workbook itself.
 */
import { fileURLToPath } from 'node:url';
import { Duration, RemovalPolicy, Stack } from 'aws-cdk-lib';
import { AccessLogFormat } from 'aws-cdk-lib/aws-apigateway';
import {
  CorsHttpMethod,
  HttpApi,
  HttpMethod,
  HttpNoneAuthorizer,
  LogGroupLogDestination,
} from 'aws-cdk-lib/aws-apigatewayv2';
import { HttpJwtAuthorizer } from 'aws-cdk-lib/aws-apigatewayv2-authorizers';
import { HttpLambdaIntegration } from 'aws-cdk-lib/aws-apigatewayv2-integrations';
import type { IAlarmAction } from 'aws-cdk-lib/aws-cloudwatch';
import { Alarm, ComparisonOperator, TreatMissingData } from 'aws-cdk-lib/aws-cloudwatch';
import type { Code } from 'aws-cdk-lib/aws-lambda';
import { Architecture, Function as LambdaFunction, Runtime } from 'aws-cdk-lib/aws-lambda';
import { NodejsFunction, OutputFormat } from 'aws-cdk-lib/aws-lambda-nodejs';
import { LogGroup, RetentionDays } from 'aws-cdk-lib/aws-logs';
import { Secret } from 'aws-cdk-lib/aws-secretsmanager';
import { Construct } from 'constructs';

export type McSheetyAttributeSource =
  | {
      readonly source: 'graph';
      readonly properties: Record<string, string>;
      readonly valueMaps?: Record<string, Record<string, string>>;
    }
  | { readonly source: 'static'; readonly users: Record<string, Record<string, string>> };

export interface McSheetyApiProps {
  /** Entra tenant (directory) ID. Only this tenant's tokens are accepted. */
  readonly tenantId: string;
  /** Application (client) ID of the McSheety API registration (the token audience). */
  readonly entraApiClientId: string;
  /** Delegated scope name exposed by the API registration. @default 'access_as_user' */
  readonly entraApiScope?: string;
  /** Complete ARN of an existing Secrets Manager secret holding the OBO credential. */
  readonly entraCredentialSecretArn: string;
  /** The SharePoint workbook, by Graph drive and item ID. */
  readonly workbook: { readonly driveId: string; readonly itemId: string };
  /** Only co-located metadata (Mode A) is supported in v0.1. @default 'co-located' */
  readonly metadataMode?: 'co-located';
  /** Exact browser origins allowed by CORS. No wildcards. */
  readonly allowedOrigins: readonly string[];
  /** Trusted attribute source for row policies (e.g. departmentId). @default none */
  readonly attributes?: McSheetyAttributeSource;
  /** Immutable Entra group object ID -> McSheety role. @default none (app roles only) */
  readonly groupRoles?: Record<string, string>;
  /** Product safeguard on rows per table. @default 2000 */
  readonly maxRows?: number;
  /** @default { rateLimit: 20, burstLimit: 40 } requests per second for the whole API */
  readonly throttle?: { readonly rateLimit: number; readonly burstLimit: number };
  /** @default RetentionDays.ONE_MONTH */
  readonly logRetention?: RetentionDays;
  /** @default RemovalPolicy.DESTROY for log groups */
  readonly logRemovalPolicy?: RemovalPolicy;
  /** Where alarms notify (e.g. an SNS topic action). @default alarms without actions */
  readonly alarmAction?: IAlarmAction;
  /** @default 1024 */
  readonly memorySize?: number;
  /** Lambda timeout; keep below API Gateway's 30 s integration limit. @default 25 seconds */
  readonly timeout?: Duration;
  /** Pre-built Lambda code (tests or custom builds). @default bundles @mcsheety/api/lambda with esbuild */
  readonly code?: Code;
}

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export class McSheetyApi extends Construct {
  readonly api: HttpApi;
  readonly handler: LambdaFunction;
  readonly alarms: Alarm[];
  /** Base URL of the API (no trailing slash). */
  readonly url: string;
  /** Accepted token issuer and audience, for documentation/outputs. */
  readonly issuer: string;
  readonly audience: string;

  constructor(scope: Construct, id: string, props: McSheetyApiProps) {
    super(scope, id);
    validate(props);
    const scopeName = props.entraApiScope ?? 'access_as_user';
    this.issuer = `https://login.microsoftonline.com/${props.tenantId}/v2.0`;
    this.audience = props.entraApiClientId;

    const logGroup = new LogGroup(this, 'HandlerLogs', {
      retention: props.logRetention ?? RetentionDays.ONE_MONTH,
      removalPolicy: props.logRemovalPolicy ?? RemovalPolicy.DESTROY,
    });

    const environment: Record<string, string> = {
      NODE_OPTIONS: '--enable-source-maps',
      MCSHEETY_TENANT_ID: props.tenantId,
      MCSHEETY_API_CLIENT_ID: props.entraApiClientId,
      MCSHEETY_API_SCOPE: scopeName,
      MCSHEETY_DRIVE_ID: props.workbook.driveId,
      MCSHEETY_ITEM_ID: props.workbook.itemId,
      MCSHEETY_POLICY_MODE: props.metadataMode ?? 'co-located',
      MCSHEETY_MAX_ROWS: String(props.maxRows ?? 2000),
      MCSHEETY_CREDENTIAL_SECRET_ARN: props.entraCredentialSecretArn,
      ...(props.attributes ? { MCSHEETY_ATTRIBUTES: JSON.stringify(props.attributes) } : {}),
      ...(props.groupRoles ? { MCSHEETY_GROUP_ROLES: JSON.stringify(props.groupRoles) } : {}),
    };

    const common = {
      runtime: Runtime.NODEJS_22_X,
      architecture: Architecture.ARM_64,
      memorySize: props.memorySize ?? 1024,
      timeout: props.timeout ?? Duration.seconds(25),
      environment,
      logGroup,
      description: 'McSheety API',
    };
    this.handler = props.code
      ? new LambdaFunction(this, 'Handler', {
          ...common,
          code: props.code,
          handler: 'index.handler',
        })
      : new NodejsFunction(this, 'Handler', {
          ...common,
          entry: fileURLToPath(import.meta.resolve('@mcsheety/api/lambda')),
          handler: 'handler',
          bundling: {
            format: OutputFormat.CJS,
            minify: true,
            sourceMap: true,
            target: 'node22',
            externalModules: ['@aws-sdk/*'],
          },
        });

    // Read access to exactly this secret, nothing else. (A customer-managed KMS key needs a
    // separate decrypt grant; see docs/deployment.md.)
    Secret.fromSecretCompleteArn(this, 'Credential', props.entraCredentialSecretArn).grantRead(
      this.handler,
    );

    const accessLogs = new LogGroup(this, 'AccessLogs', {
      retention: props.logRetention ?? RetentionDays.ONE_MONTH,
      removalPolicy: props.logRemovalPolicy ?? RemovalPolicy.DESTROY,
    });

    this.api = new HttpApi(this, 'Api', {
      description: 'McSheety REST API',
      corsPreflight: {
        allowOrigins: [...props.allowedOrigins],
        allowMethods: [
          CorsHttpMethod.GET,
          CorsHttpMethod.POST,
          CorsHttpMethod.PATCH,
          CorsHttpMethod.DELETE,
        ],
        allowHeaders: ['Authorization', 'Content-Type', 'If-Match'],
        exposeHeaders: ['ETag', 'Retry-After', 'WWW-Authenticate', 'x-request-id'],
        maxAge: Duration.hours(1),
      },
      createDefaultStage: false,
    });
    this.api.addStage('Stage', {
      stageName: '$default',
      autoDeploy: true,
      throttle: props.throttle ?? { rateLimit: 20, burstLimit: 40 },
      accessLogSettings: {
        destination: new LogGroupLogDestination(accessLogs),
        // No IPs, user agents or query strings (filters may contain record values).
        format: AccessLogFormat.custom(
          JSON.stringify({
            requestId: '$context.requestId',
            time: '$context.requestTime',
            routeKey: '$context.routeKey',
            status: '$context.status',
            latencyMs: '$context.responseLatency',
            integrationLatencyMs: '$context.integrationLatency',
            authorizerError: '$context.authorizer.error',
          }),
        ),
      },
    });

    const authorizer = new HttpJwtAuthorizer('EntraJwt', this.issuer, {
      jwtAudience: [props.entraApiClientId],
      identitySource: ['$request.header.Authorization'],
    });
    const integration = new HttpLambdaIntegration('Lambda', this.handler);
    const secured = { authorizer, authorizationScopes: [scopeName] };

    this.api.addRoutes({
      path: '/v1/health',
      methods: [HttpMethod.GET],
      integration,
      authorizer: new HttpNoneAuthorizer(),
    });
    const routes: [string, HttpMethod[]][] = [
      ['/v1/tables', [HttpMethod.GET]],
      ['/v1/capabilities', [HttpMethod.GET]],
      ['/v1/schema/{table}', [HttpMethod.GET]],
      ['/v1/tables/{table}/rows', [HttpMethod.GET, HttpMethod.POST]],
      ['/v1/tables/{table}/rows/{id}', [HttpMethod.GET, HttpMethod.PATCH, HttpMethod.DELETE]],
    ];
    for (const [path, methods] of routes)
      this.api.addRoutes({ path, methods, integration, ...secured });

    this.url = `https://${this.api.apiId}.execute-api.${Stack.of(this).region}.${Stack.of(this).urlSuffix}`;

    const alarm = (
      name: string,
      metric: ReturnType<LambdaFunction['metricErrors']>,
      threshold: number,
    ) => {
      const a = new Alarm(this, name, {
        metric,
        threshold,
        evaluationPeriods: 1,
        comparisonOperator: ComparisonOperator.GREATER_THAN_OR_EQUAL_TO_THRESHOLD,
        treatMissingData: TreatMissingData.NOT_BREACHING,
      });
      if (props.alarmAction) a.addAlarmAction(props.alarmAction);
      return a;
    };
    const fiveMinutes = { period: Duration.minutes(5) };
    this.alarms = [
      alarm('HandlerErrors', this.handler.metricErrors(fiveMinutes), 1),
      alarm('HandlerThrottles', this.handler.metricThrottles(fiveMinutes), 1),
      alarm('Api5xx', this.api.metricServerError(fiveMinutes), 5),
    ];
  }
}

function validate(props: McSheetyApiProps): void {
  const problems: string[] = [];
  if (!GUID_RE.test(props.tenantId)) problems.push('tenantId must be a GUID');
  if (!GUID_RE.test(props.entraApiClientId)) problems.push('entraApiClientId must be a GUID');
  if (
    !/^arn:aws[\w-]*:secretsmanager:[\w-]+:\d{12}:secret:.+-[A-Za-z0-9]{6}$/.test(
      props.entraCredentialSecretArn,
    )
  ) {
    problems.push(
      'entraCredentialSecretArn must be a complete Secrets Manager ARN (including the 6-character suffix)',
    );
  }
  if (!props.workbook.driveId || !props.workbook.itemId)
    problems.push('workbook.driveId and workbook.itemId are required');
  if (props.allowedOrigins.length === 0)
    problems.push('allowedOrigins must list at least one origin');
  for (const origin of props.allowedOrigins) {
    if (origin.includes('*') || !/^https?:\/\/[^/]+$/.test(origin))
      problems.push(`invalid origin "${origin}" (exact scheme://host[:port], no wildcards)`);
  }
  if (props.metadataMode && props.metadataMode !== 'co-located')
    problems.push('only metadataMode "co-located" is supported');
  if (problems.length > 0) throw new Error(`McSheetyApi: ${problems.join('; ')}.`);
}
