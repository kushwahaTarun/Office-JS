import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import fs from 'fs'
import path from 'path'

// Check if certificates exist
const certKeyPath = path.resolve(__dirname, 'certs/localhost-key.pem')
const certPath = path.resolve(__dirname, 'certs/localhost.pem')
const hasCerts = fs.existsSync(certKeyPath) && fs.existsSync(certPath)

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    strictPort: true, // Fail if port is in use instead of trying another
    // Only use HTTPS if certificates are available
    https: hasCerts ? {
      key: fs.readFileSync(certKeyPath),
      cert: fs.readFileSync(certPath),
    } : undefined,
    headers: {
      'Access-Control-Allow-Origin': '*',
    },
  },
  build: {
    outDir: 'dist',
    sourcemap: true,
  },
})