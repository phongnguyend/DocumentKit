import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: './e2e',
  use: { baseURL: 'http://localhost:5177', channel: process.env.CI ? undefined : 'chrome', screenshot: 'only-on-failure' },
  webServer: { command: 'npm run dev', url: 'http://localhost:5177', reuseExistingServer: !process.env.CI },
});
