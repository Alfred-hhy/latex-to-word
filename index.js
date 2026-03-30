// Vercel entrypoint - 无需 express
const previewHandler = require('./api/preview');
const convertHandler = require('./api/convert');
const path = require('path');
const fs = require('fs');

const staticFiles = {
  '/': 'public/index.html'
};

async function handler(req, res) {
  const urlPath = req.url.split('?')[0];

  try {
    if (urlPath === '/api/health') {
      res.setHeader('Access-Control-Allow-Origin', '*');
      return res.json({ status: 'ok', timestamp: new Date().toISOString() });
    }

    if (urlPath === '/api/preview' || urlPath === '/api/preview/') {
      return previewHandler(req, res);
    }

    if (urlPath === '/api/convert' || urlPath === '/api/convert/') {
      return convertHandler(req, res);
    }

    if (urlPath === '/' || urlPath === '/index.html') {
      const indexPath = path.join(__dirname, 'public', 'index.html');
      const content = fs.readFileSync(indexPath, 'utf-8');
      res.setHeader('Content-Type', 'text/html');
      return res.send(content);
    }

    res.status(404).json({ error: 'Not found' });

  } catch (error) {
    console.error('Handler error:', error);
    res.status(500).json({ error: error.message });
  }
}

module.exports = handler;