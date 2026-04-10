/**
 * Web 用: Excel 読込 API + 本番では静的ファイル配信。
 * サーバーが参照できるパス（UNC 含む）を dataDir に指定する。
 *
 * 環境変数:
 *   PORT — 待受ポート（既定: SERVE_STATIC 時 8080、それ以外 3001）
 *   SERVE_STATIC=1 — build/ を配信（本番）
 *   LEASE_REPORT_CACHE_DIR — キャッシュ保存先
 *   LEASE_REPORT_API_TOKEN — 設定時は Authorization: Bearer <token> 必須
 *   LEASE_REPORT_CORS_ORIGIN — CORS（カンマ区切り複数可）。未設定時は *
 *   EXCEL_PASSWORD — 暗号化 Excel 用（既定は excelBackend と同じ）
 */
const path = require('path');
const express = require('express');
const { createExcelBackend } = require('./excelBackend');

const serveStatic = process.env.SERVE_STATIC === '1' || process.argv.includes('--static');
const defaultPort = serveStatic ? 8080 : 3001;
const PORT = parseInt(process.env.PORT, 10) || defaultPort;

const cacheDir = process.env.LEASE_REPORT_CACHE_DIR
  || path.join(__dirname, '.lease-report-cache');

const excelBackend = createExcelBackend({ cacheDir });
const apiToken = process.env.LEASE_REPORT_API_TOKEN;
const corsOrigins = (process.env.LEASE_REPORT_CORS_ORIGIN || '')
  .split(',')
  .map((s) => s.trim())
  .filter(Boolean);

const app = express();
app.use(express.json({ limit: '2mb' }));

function corsMiddleware(req, res, next) {
  const origin = req.headers.origin;
  if (corsOrigins.length > 0) {
    if (origin && corsOrigins.includes(origin)) {
      res.setHeader('Access-Control-Allow-Origin', origin);
    }
  } else {
    res.setHeader('Access-Control-Allow-Origin', '*');
  }
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
}

app.use(corsMiddleware);

app.get('/api/health', (req, res) => {
  res.json({ ok: true, service: 'lease-report-api' });
});

function apiAuth(req, res, next) {
  if (!apiToken) return next();
  const hdr = req.headers.authorization || '';
  const ok = hdr === `Bearer ${apiToken}`;
  if (!ok) return res.status(401).json({ success: false, error: 'Unauthorized' });
  next();
}

app.post('/api/load-excel', apiAuth, async (req, res) => {
  try {
    const dirPath = req.body?.dirPath;
    const forceRefresh = Boolean(req.body?.forceRefresh);
    if (dirPath === undefined || dirPath === null || String(dirPath).trim() === '') {
      return res.status(400).json({ success: false, error: 'dirPath が必要です' });
    }
    const result = await excelBackend.loadExcelData(String(dirPath), forceRefresh);
    res.json(result);
  } catch (err) {
    res.status(500).json({ success: false, error: err.message || String(err) });
  }
});

app.get('/api/check-path', apiAuth, (req, res) => {
  const dirPath = req.query.path;
  if (dirPath === undefined || dirPath === null) {
    return res.status(400).json({ exists: false, error: 'path クエリが必要です' });
  }
  res.json(excelBackend.checkPath(String(dirPath)));
});

if (serveStatic) {
  const buildDir = path.join(__dirname, 'build');
  app.use(express.static(buildDir));
  app.get('*', (req, res, next) => {
    if (req.path.startsWith('/api')) return next();
    res.sendFile(path.join(buildDir, 'index.html'), (err) => {
      if (err) next(err);
    });
  });
}

app.listen(PORT, '0.0.0.0', () => {
  // eslint-disable-next-line no-console
  console.log(`[lease-report] listening on http://0.0.0.0:${PORT}${serveStatic ? ' (static + API)' : ' (API only)'}`);
});
