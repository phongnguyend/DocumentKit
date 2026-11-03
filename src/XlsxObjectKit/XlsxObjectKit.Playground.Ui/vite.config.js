import { defineConfig, loadEnv } from 'vite';

export default defineConfig(({ mode }) => {
  const environment = { ...loadEnv(mode, process.cwd(), ''), ...process.env };
  return {
    server: {
      host: 'localhost',
      port: Number(environment.XLSX_UI_PORT ?? 5184),
      strictPort: true,
      proxy: { '/api': environment.XLSX_API_URL ?? 'http://localhost:5094' }
    }
  };
});
