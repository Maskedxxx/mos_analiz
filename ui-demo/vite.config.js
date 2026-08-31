import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// Проксирование двух бэкендов:
//   /api -> :8081 (аудит, как было)
//   /gen -> :8090 (генерация; префикс срезается rewrite'ом, т.к. оба бэкенда отдают /api/types/*)
const proxyConfig = {
  '/api': {
    target: 'http://localhost:8081',
    changeOrigin: true,
  },
  '/gen': {
    target: 'http://localhost:8090',
    changeOrigin: true,
    rewrite: (p) => p.replace(/^\/gen/, ''),
  },
}

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  server: {
    allowedHosts: true,
    proxy: proxyConfig,
  },
  // прод-сборка через `vite preview`: те же allowedHosts и проксирование, что и в dev
  preview: {
    allowedHosts: true,
    proxy: proxyConfig,
  },
})
