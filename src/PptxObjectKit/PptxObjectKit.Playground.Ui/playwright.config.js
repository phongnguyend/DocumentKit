import { defineConfig } from '@playwright/test';

const apiUrl = process.env.PPTX_API_URL ?? 'http://localhost:5090';
const uiUrl = `http://localhost:${process.env.PPTX_UI_PORT ?? 5180}`;

export default defineConfig({
  testDir: './e2e',
  workers: 1,
  use: { baseURL: uiUrl, channel: process.env.CI ? undefined : 'chrome', viewport: { width: 1440, height: 1000 }, screenshot: 'only-on-failure', trace: 'retain-on-failure' },
  webServer: [
    { command: `dotnet run --project ../PptxObjectKit.Playground.Api -c Release --no-launch-profile --urls ${apiUrl}`, url: `${apiUrl}/api/health`, reuseExistingServer: !process.env.CI, timeout: 120000 },
    { command: 'npm run dev', url: uiUrl, reuseExistingServer: !process.env.CI },
  ],
});
