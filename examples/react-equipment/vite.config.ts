import react from '@vitejs/plugin-react';
import { defineConfig } from 'vite';

export default defineConfig({
  plugins: [react()],
  // Share the repository-root .env; Vite only exposes VITE_-prefixed values to the browser.
  envDir: '../..',
  server: { port: 5173, strictPort: true },
});
