import fs from 'node:fs';
import https from 'node:https';
import path from 'node:path';
import { WebSocketServer } from 'ws';

const PORT = Number(process.env.PORT || 3443);
const DIST_DIR = path.resolve(
  process.env.ADDIN_DIST_DIR || path.join(import.meta.dirname, '..', 'dist', 'excel-market-data-addin', 'browser')
);
const CERT_DIR = path.resolve(process.env.CERT_DIR || path.join(import.meta.dirname, 'certs'));

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.mjs': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
  '.svg': 'image/svg+xml',
  '.xml': 'application/xml; charset=utf-8'
};

/**
 * Starts the local HTTPS static server (for Angular 19 dist) and WSS WebSocket server.
 * Call this from your Electron main process.
 */
export function startMarketAddinServer({ enableMockFeed = false } = {}) {
  const keyPath = path.join(CERT_DIR, 'localhost.key');
  const certPath = path.join(CERT_DIR, 'localhost.crt');

  if (!fs.existsSync(keyPath) || !fs.existsSync(certPath)) {
    throw new Error(`Missing TLS certs in ${CERT_DIR}. Run dev-certs.sh first.`);
  }

  const server = https.createServer(
    {
      key: fs.readFileSync(keyPath),
      cert: fs.readFileSync(certPath)
    },
    (req, res) => {
      const reqPath = (req.url || '/').split('?')[0];
      const safePath = path.normalize(reqPath === '/' ? '/index.html' : reqPath).replace(/^(\.\.[/\\])+/, '');
      let filePath = path.join(DIST_DIR, safePath);

      if (!fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
        filePath = path.join(DIST_DIR, 'index.html');
      }

      const ext = path.extname(filePath).toLowerCase();
      res.writeHead(200, {
        'Content-Type': MIME_TYPES[ext] || 'application/octet-stream',
        'Cache-Control': 'no-cache'
      });
      fs.createReadStream(filePath).pipe(res);
    }
  );

  const wss = new WebSocketServer({ server, path: '/ws' });
  const subscriptions = new Map(); // ws -> Set<string>

  wss.on('connection', (ws) => {
    subscriptions.set(ws, new Set());

    ws.on('message', (data) => {
      try {
        const msg = JSON.parse(data.toString());
        const symbol = typeof msg.symbol === 'string' ? msg.symbol.trim().toUpperCase() : '';
        if (!symbol) return;

        const clientSubs = subscriptions.get(ws);
        if (msg.type === 'subscribe') {
          clientSubs?.add(symbol);
        } else if (msg.type === 'unsubscribe') {
          clientSubs?.delete(symbol);
        }
      } catch {
        // Ignore malformed JSON
      }
    });

    ws.on('close', () => {
      subscriptions.delete(ws);
    });
  });

  function broadcastTick(tick) {
    const symbol = String(tick.symbol || '').toUpperCase();
    const payload = JSON.stringify(tick);
    for (const [client, subs] of subscriptions.entries()) {
      if (client.readyState === 1 && subs.has(symbol)) {
        client.send(payload);
      }
    }
  }

  let mockTimer = null;
  if (enableMockFeed) {
    const prices = new Map();
    mockTimer = setInterval(() => {
      const allSymbols = new Set();
      for (const subs of subscriptions.values()) {
        for (const s of subs) allSymbols.add(s);
      }
      for (const symbol of allSymbols) {
        const prev = prices.get(symbol) ?? 100 + Math.random() * 100;
        const next = Number((prev + (Math.random() - 0.5) * 0.5).toFixed(2));
        prices.set(symbol, next);
        broadcastTick({
          symbol,
          price: next,
          bid: Number((next - 0.02).toFixed(2)),
          ask: Number((next + 0.02).toFixed(2)),
          volume: Math.floor(1000 + Math.random() * 50000)
        });
      }
    }, 250);
  }

  server.listen(PORT, '127.0.0.1', () => {
    console.log(`Excel Add-in HTTPS + WSS server running at https://localhost:${PORT}`);
  });

  return {
    server,
    wss,
    broadcastTick,
    stop: () => {
      if (mockTimer) clearInterval(mockTimer);
      wss.close();
      server.close();
    }
  };
}

// Allow standalone execution for testing: node local-server.mjs --mock
if (process.argv.includes('--mock')) {
  startMarketAddinServer({ enableMockFeed: true });
}
