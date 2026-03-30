// Vercel entrypoint
const previewHandler = require('./api/preview');
const convertHandler = require('./api/convert');
const path = require('path');
const fs = require('fs');

// 读取静态文件
const staticFiles = {
  '/': 'public/index.html'
};

async function handler(req, res) {
  const urlPath = req.url.split('?')[0];

  try {
    // API 路由
    if (urlPath === '/api/health') {
      return res.json({ status: 'ok', timestamp: new Date().toISOString() });
    }

    if (urlPath === '/api/preview' || urlPath === '/api/preview/') {
      return previewHandler(req, res);
    }

    if (urlPath === '/api/convert' || urlPath === '/api/convert/') {
      return convertHandler(req, res);
    }

    // 静态文件
    if (urlPath === '/' || urlPath === '/index.html') {
      const indexPath = path.join(__dirname, 'public', 'index.html');
      const content = fs.readFileSync(indexPath, 'utf-8');
      res.setHeader('Content-Type', 'text/html');
      return res.send(content);
    }

    // 404
    res.status(404).json({ error: 'Not found' });

  } catch (error) {
    console.error('Handler error:', error);
    res.status(500).json({ error: error.message });
  }
}

module.exports = handler;