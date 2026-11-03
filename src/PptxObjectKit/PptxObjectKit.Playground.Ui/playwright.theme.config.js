import { defineConfig } from '@playwright/test';

// Theme checks need only the frontend, so no conversion backend is started.
export default defineConfig({
  testDir: './e2e', outputDir: './test-results/themes', testMatch: 'theme.spec.js', workers: 1,
  use: { baseURL: 'http://localhost:5292', channel: process.env.CI ? undefined : 'chrome', screenshot: 'only-on-failure' },
  webServer: { command: 'npm run dev -- --port 5292 --strictPort', url: 'http://localhost:5292', reuseExistingServer: false, timeout: 120000 }
});
