import { defineConfig, loadEnv, type Plugin } from 'vite'
import vue from '@vitejs/plugin-vue'
import { resolve } from 'path'
import type { IncomingMessage, ServerResponse } from 'node:http'

function supportWebhook(webhook: string): Plugin {
  async function handle(req: IncomingMessage, res: ServerResponse, next: () => void) {
    const path = (req.url || '').split('?')[0]
    if (path !== '/api/support') return next()
    if (req.method !== 'POST') {
      res.statusCode = 405
      res.end()
      return
    }
    if (!webhook) {
      res.statusCode = 500
      res.end('missing webhook')
      return
    }
    const chunks: Buffer[] = []
    for await (const chunk of req) chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk))
    const response = await fetch(webhook, {
      method: 'POST',
      headers: { 'content-type': req.headers['content-type'] || '' },
      body: Buffer.concat(chunks),
    })
    res.statusCode = response.status
    res.end(await response.text())
  }

  return {
    name: 'support-webhook',
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        void handle(req, res, next)
      })
    },
    configurePreviewServer(server) {
      server.middlewares.use((req, res, next) => {
        void handle(req, res, next)
      })
    },
  }
}

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), '')
  return {
    plugins: [vue(), supportWebhook(env.DISCORD_SUPPORT_WEBHOOK || '')],
    resolve: {
      alias: {
        '@': resolve(__dirname, 'src'),
      },
    },
    server: {
      port: 5173,
    },
    build: {
      outDir: 'dist',
      rollupOptions: {
        output: {
          manualChunks: {
            'excel-engine': ['exceljs', 'xlsx'],
            'vue-vendor': ['vue', 'pinia'],
          },
        },
      },
    },
  }
})
