import { resolve } from 'node:path'
import { fileURLToPath, pathToFileURL, URL } from 'node:url'
import { defineConfig, loadEnv, type Plugin } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

const localApiModules: Record<string, string> = {
  '/api/auth': 'api/auth.js',
  '/api/data': 'api/data.js',
  '/api/users': 'api/users.js',
  '/api/friends': 'api/friends.js',
  '/api/shopping': 'api/shopping.js',
}

function localApiPlugin(): Plugin {
  return {
    name: 'local-serverless-api',
    configureServer(server) {
      server.middlewares.use(async (req, res, next) => {
        const path = req.url?.split('?')[0] ?? ''
        const modulePath = localApiModules[path]
        if (!modulePath) return next()

        try {
          if (req.method !== 'GET') {
            const chunks: Buffer[] = []
            let size = 0
            for await (const chunk of req) {
              const buffer = Buffer.from(chunk)
              size += buffer.length
              if (size > 2 * 1024 * 1024) throw new Error('Request body too large')
              chunks.push(buffer)
            }
            const raw = Buffer.concat(chunks).toString('utf8')
            ;(req as typeof req & { body?: unknown }).body = raw ? JSON.parse(raw) : {}
          }

          const moduleFile = pathToFileURL(resolve(process.cwd(), modulePath)).href
          const module = await import(/* @vite-ignore */ moduleFile + '?t=' + Date.now())
          await module.default(req, res)
        } catch (error) {
          console.error('Local API failure', error)
          if (!res.headersSent) res.setHeader('Content-Type', 'application/json; charset=utf-8')
          res.statusCode = 500
          res.end(JSON.stringify({ error: 'Local API request failed' }))
        }
      })
    },
  }
}

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, '.', '')
  for (const [key, value] of Object.entries(env)) {
    if (process.env[key] === undefined) process.env[key] = value
  }

  return {
    plugins: [react(), tailwindcss(), localApiPlugin()],
    build: {
      sourcemap: false,
    },
    resolve: {
      alias: {
        '@': fileURLToPath(new URL('./src', import.meta.url)),
      },
    },
  }
})
