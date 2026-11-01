import { defineConfig, loadEnv } from 'vite';
export default defineConfig(({mode}) => {
 const env = loadEnv(mode, process.cwd(), 'PPTX_');
 return { server: { port: Number(process.env.PPTX_UI_PORT ?? env.PPTX_UI_PORT ?? 5180), strictPort: true, proxy: { '/api': process.env.PPTX_API_URL ?? env.PPTX_API_URL ?? 'http://localhost:5090' } } };
});
