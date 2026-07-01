import { fileURLToPath, URL } from 'node:url'
import { defineConfig, loadEnv, type Plugin } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { VitePWA } from 'vite-plugin-pwa'

function localApiPlugin(): Plugin {
  return {
    name: 'local-serverless-api',
    configureServer(server) {
      server.middlewares.use(async (req, res, next) => {
        const path = req.url?.split('?')[0]
        if (path !== '/api/auth' && path !== '/api/data' && path !== '/api/push' && path !== '/api/users') return next()

        try {
          if (req.method === 'POST') {
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

          // @ts-ignore Local development loads the same JavaScript handlers used by Vercel.
          const module = path === '/api/auth' ? await import('./api/auth.js') : path === '/api/push' ? await import('./api/push.js') : path === '/api/users' ? await import('./api/users.js') : await import('./api/data.js')
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

// https://vite.dev/config/
export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, '.', '')
  for (const [key, value] of Object.entries(env)) {
    if (process.env[key] === undefined) process.env[key] = value
  }

  return {
    plugins: [
      react(),
      tailwindcss(),
      localApiPlugin(),
      VitePWA({
        strategies: 'injectManifest',
        srcDir: 'src',
        filename: 'push-sw.ts',
        registerType: 'autoUpdate',
        includeAssets: ['favicon.svg', 'apple-touch-icon.png', 'icon-192.png', 'icon-512.png'],
        manifest: {
          name: 'RubyLife',
          short_name: 'RubyLife',
          description: 'Controle financeiro pessoal',
          start_url: '/',
          scope: '/',
          display: 'standalone',
          background_color: '#ffffff',
          theme_color: '#f5f5f7',
          icons: [
            { src: '/icon-192.png', sizes: '192x192', type: 'image/png', purpose: 'any' },
            { src: '/icon-512.png', sizes: '512x512', type: 'image/png', purpose: 'any' },
            { src: '/icon-512.png', sizes: '512x512', type: 'image/png', purpose: 'maskable' },
          ],
        },
        injectManifest: {
          globPatterns: ['**/*.{js,css,html,ico,png,svg,webmanifest}'],
        },
        workbox: {
          runtimeCaching: [
            {
              urlPattern: /.*supabase\.co.*/i,
              handler: 'NetworkFirst',
              options: {
                cacheName: 'supabase-api-cache',
                expiration: {
                  maxEntries: 100,
                  maxAgeSeconds: 60 * 60 * 24,
                },
                cacheableResponse: {
                  statuses: [0, 200],
                },
              },
            },
            {
              urlPattern: /\.(?:png|jpg|jpeg|svg|gif|webp|ico|woff|woff2|ttf|eot)$/i,
              handler: 'CacheFirst',
              options: {
                cacheName: 'static-assets-cache',
                expiration: {
                  maxEntries: 200,
                  maxAgeSeconds: 60 * 60 * 24 * 30,
                },
                cacheableResponse: {
                  statuses: [0, 200],
                },
              },
            },
          ],
        },
      }),
    ],
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

