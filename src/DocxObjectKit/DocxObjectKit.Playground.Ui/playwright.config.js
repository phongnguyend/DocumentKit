import { defineConfig } from '@playwright/test';
export default defineConfig({
  testDir: './e2e',
  timeout: 60000,
  use: { baseURL: 'http://localhost:5182', channel: process.env.CI ? undefined : 'chrome', screenshot: 'only-on-failure', trace: 'retain-on-failure' },
  webServer: process.env.DOCX_TEST_EXTERNAL ? [] : [
    { command: 'dotnet run --project ../DocxObjectKit.Playground.Api', url: 'http://localhost:5092/api/health', reuseExistingServer: !process.env.CI, timeout: 120000 },
    { command: 'npm run dev', url: 'http://localhost:5182', reuseExistingServer: !process.env.CI, timeout: 120000 }
  ]
});
