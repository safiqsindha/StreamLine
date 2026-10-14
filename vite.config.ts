/// <reference types="vitest/config" />
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

const API_PORT = Number(process.env.STREAMLINE_PORT ?? 4317)

export default defineConfig({
  root: 'src/web',
  plugins: [react(), tailwindcss()],
  build: { outDir: '../../dist/web', emptyOutDir: true },
  server: {
    host: '127.0.0.1',
    proxy: { '/api': `http://127.0.0.1:${API_PORT}` },
  },
  test: {
    root: '.',
    include: ['test/**/*.test.ts'],
  },
})
