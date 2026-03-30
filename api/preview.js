const { latexToOMML } = require('latex-to-omml');

/**
 * 预处理 LaTeX，处理 AI 生成公式常见错误
 */
function preprocessLatex(latex) {
  let processed = latex;
  processed = processed.replace(/\s+\\$/g, ' \\;');
  processed = processed.replace(/\\\s+/g, ' ');
  return processed;
}

async function handler(req, res) {
  try {
    // CORS headers
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
      return res.status(200).end();
    }

    if (req.method !== 'POST') {
      return res.status(405).json({ error: '只支持 POST 请求' });
    }

    const { latex } = req.body || {};

    if (!latex || typeof latex !== 'string') {
      return res.status(400).json({ error: '请提供有效的 LaTeX 公式' });
    }

    const preprocessed = preprocessLatex(latex);
    const omml = await latexToOMML(preprocessed);

    res.json({
      success: true,
      latex: latex,
      preprocessed: preprocessed,
      message: '公式已转换，可在 Word 中查看完整效果'
    });

  } catch (error) {
    console.error('预览接口错误:', error);
    res.status(500).json({ error: '转换失败', details: error.message });
  }
}

module.exports = handler;