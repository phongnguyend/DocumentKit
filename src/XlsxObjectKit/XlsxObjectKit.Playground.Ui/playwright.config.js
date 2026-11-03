import { defineConfig } from '@playwright/test';
const api = process.env.XLSX_API_URL ?? 'http://localhost:5094';
const ui = `http://localhost:${process.env.XLSX_UI_PORT ?? 5184}`;
export default defineConfig({
  testDir: './e2e', workers: 1, timeout: 60000,
  use: { baseURL: ui, channel: process.env.CI ? undefined : 'chrome', viewport: { width: 1440, height: 1000 }, screenshot: 'only-on-failure', trace: 'retain-on-failure' },
  webServer: [
    { command: `dotnet run --project ../XlsxObjectKit.Playground.Api -c Release --no-launch-profile --urls ${api}`, url: `${api}/api/health`, reuseExistingServer: !process.env.CI, timeout: 120000 },
    { command: 'npm run dev', url: ui, reuseExistingServer: !process.env.CI, timeout: 120000 }
  ]
});
