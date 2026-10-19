import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vitest/config';

const packagesDir = fileURLToPath(new URL('./packages/', import.meta.url));

export default defineConfig({
  resolve: {
    // Tests run against workspace sources, so no build is needed first.
    alias: [
      { find: /^@mcsheety\/([a-z-]+)\/testing$/, replacement: `${packagesDir}$1/src/testing.ts` },
      { find: /^@mcsheety\/([a-z-]+)$/, replacement: `${packagesDir}$1/src/index.ts` },
    ],
  },
  test: {
    include: ['packages/*/src/**/*.test.ts', 'fixtures/**/*.test.ts', 'tools/*/src/**/*.test.ts'],
    environment: 'node',
  },
});
