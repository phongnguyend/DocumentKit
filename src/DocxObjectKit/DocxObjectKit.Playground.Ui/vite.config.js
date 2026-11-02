import { defineConfig, loadEnv } from 'vite';
export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), 'DOCX_');
  return { server: { port: Number(process.env.DOCX_UI_PORT ?? env.DOCX_UI_PORT ?? 5182), strictPort: true, watch: { ignored: ['**/public/ocr/**'] }, proxy: { '/api': process.env.DOCX_API_URL ?? env.DOCX_API_URL ?? 'http://localhost:5092' } } };
});
