import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: './e2e',
  workers: 1,
  use: { baseURL: 'http://localhost:5180', channel: process.env.CI ? undefined : 'chrome', viewport: { width: 1440, height: 1000 }, screenshot: 'only-on-failure', trace: 'retain-on-failure' },
  webServer: [
    { command: 'dotnet run --project ../PptxObjectKit.Playground.Api -c Release --no-launch-profile --urls http://localhost:5090', url: 'http://localhost:5090/api/health', reuseExistingServer: !process.env.CI, timeout: 120000 },
    { command: 'npm run dev', url: 'http://localhost:5180', reuseExistingServer: !process.env.CI },
  ],
});
