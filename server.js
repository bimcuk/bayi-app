const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = 3000;
const PUBLIC_DIR = __dirname;

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon'
};

const server = http.createServer((req, res) => {
  // Parsing URL dan menghapus query string
  let reqUrl = req.url.split('?')[0];

  // Routing Clean URLs (mirip Vercel)
  if (reqUrl === '/' || reqUrl === '') {
    reqUrl = '/index.html';
  } else if (reqUrl === '/admin') {
    reqUrl = '/admin.html';
  }

  let filePath = path.join(PUBLIC_DIR, reqUrl);

  // Pastikan path tetap berada di dalam PUBLIC_DIR (keamanan path traversal)
  if (!filePath.startsWith(PUBLIC_DIR)) {
    res.writeHead(403, { 'Content-Type': 'text/plain' });
    return res.end('403 Forbidden');
  }

  fs.stat(filePath, (err, stats) => {
    if (err || !stats.isFile()) {
      // Coba cari file dengan ekstensi .html
      const htmlPath = filePath + '.html';
      if (fs.existsSync(htmlPath)) {
        filePath = htmlPath;
      } else {
        res.writeHead(404, { 'Content-Type': 'text/html; charset=utf-8' });
        return res.end('<h1>404 Not Found</h1><p>Halaman tidak ditemukan di server lokal.</p>');
      }
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';

    fs.readFile(filePath, (readErr, content) => {
      if (readErr) {
        res.writeHead(500, { 'Content-Type': 'text/plain' });
        return res.end('500 Internal Server Error');
      }

      res.writeHead(200, {
        'Content-Type': contentType,
        'Cache-Control': 'no-cache, no-store, must-revalidate',
        'Access-Control-Allow-Origin': '*'
      });
      res.end(content);
    });
  });
});

server.listen(PORT, () => {
  console.log('========================================================');
  console.log(`🚀 LOCAL PREVIEW SERVER AKTIF: http://localhost:${PORT}`);
  console.log('========================================================');
  console.log(`👉 Aplikasi Utama : http://localhost:${PORT}`);
  console.log(`👉 Halaman Admin  : http://localhost:${PORT}/admin`);
  console.log('--------------------------------------------------------');
  console.log('Buka browser kamu dan akses alamat di atas untuk preview.');
  console.log('Tekan Ctrl + C di terminal jika ingin menghentikan server.');
  console.log('========================================================');
});
