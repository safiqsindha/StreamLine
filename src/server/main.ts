import { createServer, type IncomingMessage, type ServerResponse } from 'node:http'
import { readFile, stat } from 'node:fs/promises'
import { extname, join, resolve, sep } from 'node:path'
import { createApp } from './app'
import { FileStore } from './fileStore'

const LOOPBACK_HOSTS = new Set(['127.0.0.1', 'localhost', '::1', '[::1]'])

const CONTENT_TYPES: Record<string, string> = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.ico': 'image/x-icon',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.ttf': 'font/ttf',
  '.map': 'application/json; charset=utf-8',
  '.txt': 'text/plain; charset=utf-8',
}

function hostIsLoopback(hostHeader: string | undefined): boolean {
  if (!hostHeader) return false
  const host = hostHeader.startsWith('[') ? hostHeader.slice(0, hostHeader.indexOf(']') + 1) : hostHeader.split(':')[0]
  return LOOPBACK_HOSTS.has(host)
}

async function fileIfExists(path: string): Promise<string | null> {
  try {
    return (await stat(path)).isFile() ? path : null
  } catch {
    return null
  }
}

function staticHandler(webRoot: string) {
  const root = resolve(webRoot)
  return async (req: IncomingMessage, res: ServerResponse): Promise<void> => {
    const pathname = new URL(req.url ?? '/', 'http://localhost').pathname
    let decoded: string
    try {
      decoded = decodeURIComponent(pathname)
    } catch {
      res.writeHead(400).end()
      return
    }
    if (decoded.includes('\0')) {
      res.writeHead(400).end()
      return
    }
    const candidate = resolve(join(root, decoded))
    let file: string | null = null
    if (candidate === root || candidate.startsWith(root + sep)) file = await fileIfExists(candidate)
    if (!file) {
      if (extname(decoded) && decoded !== '/') {
        res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' }).end('Not found')
        return
      }
      file = await fileIfExists(join(root, 'index.html'))
    }
    if (!file) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' }).end('Not found')
      return
    }
    const body = await readFile(file)
    res.writeHead(200, {
      'Content-Type': CONTENT_TYPES[extname(file).toLowerCase()] ?? 'application/octet-stream',
      'Content-Length': body.length,
    })
    res.end(body)
  }
}

async function main(): Promise<void> {
  const host = process.env.STREAMLINE_HOST ?? '127.0.0.1'
  if (!LOOPBACK_HOSTS.has(host)) {
    console.error(`Streamline only listens on this computer. "${host}" is not a loopback address; use 127.0.0.1.`)
    process.exit(1)
  }
  const port = Number(process.env.STREAMLINE_PORT ?? 4317)
  if (!Number.isInteger(port) || port < 0 || port > 65535) {
    console.error('STREAMLINE_PORT must be a port number between 0 and 65535.')
    process.exit(1)
  }
  const dataDir = process.env.STREAMLINE_DATA ?? './data'

  const api = createApp(new FileStore(dataDir))
  const webRoot = resolve('dist/web')
  const serveWeb = (await fileIfExists(join(webRoot, 'index.html'))) ? staticHandler(webRoot) : null

  const server = createServer((req, res) => {
    if (!hostIsLoopback(req.headers.host)) {
      res.writeHead(403, { 'Content-Type': 'application/json; charset=utf-8' })
      res.end(JSON.stringify({ error: { message: 'Streamline only answers requests from this computer.' } }))
      return
    }
    const pathname = (req.url ?? '/').split('?')[0]
    if (pathname === '/api' || pathname.startsWith('/api/')) return api(req, res)
    if (serveWeb && (req.method === 'GET' || req.method === 'HEAD')) {
      serveWeb(req, res).catch(() => {
        if (!res.headersSent) res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' })
        res.end('Something went wrong.')
      })
      return
    }
    res.writeHead(404, { 'Content-Type': 'application/json; charset=utf-8' })
    res.end(JSON.stringify({ error: { message: 'There is nothing at this address.' } }))
  })

  server.on('error', (e: NodeJS.ErrnoException) => {
    console.error(e.code === 'EADDRINUSE' ? `Port ${port} is already in use. Set STREAMLINE_PORT to another port.` : `Streamline could not start: ${e.message}`)
    process.exit(1)
  })
  server.listen(port, host === 'localhost' ? '127.0.0.1' : host, () => {
    const addr = server.address()
    const actual = typeof addr === 'object' && addr ? addr.port : port
    console.log(`Streamline running at http://127.0.0.1:${actual}`)
  })
}

void main()
