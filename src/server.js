const express = require('express');
const multer = require('multer');
const path = require('path');
const { latexToOMML } = require('latex-to-omml');
const { Document, Packer, Paragraph, TextRun, ImportedXmlComponent } = require('docx');

const app = express();
const PORT = 3000;

// 中间件
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, '../public')));

// 文件上传配置
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 5 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const allowedTypes = ['.tex', '.txt', '.md'];
    const ext = path.extname(file.originalname).toLowerCase();
    if (allowedTypes.includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error('只支持 .tex, .txt 或 .md 文件'));
    }
  }
});

// LaTeX 公式匹配模式
const inlineMathPattern = /\$([^$]+)\$/g;
const blockMathPattern = /\$\$([^$]+)\$\$/g;

/**
 * 预处理 LaTeX，处理 AI 生成公式常见错误
 */
function preprocessLatex(latex) {
  let processed = latex;
  // 修复常见的 AI 生成公式错误
  processed = processed.replace(/\s+\\$/g, ' \\;');
  processed = processed.replace(/\\\s+/g, ' ');
  return processed;
}

/**
 * 从文本中提取所有公式
 */
function extractFormulas(text) {
  const formulas = [];

  // 提取块级公式 $$...$$
  let match;
  while ((match = blockMathPattern.exec(text)) !== null) {
    formulas.push({
      type: 'block',
      original: match[0],
      content: match[1].trim(),
      index: match.index
    });
  }

  // 提取行内公式 $...$
  inlineMathPattern.lastIndex = 0;
  while ((match = inlineMathPattern.exec(text)) !== null) {
    const isInBlock = formulas.some(f =>
      match.index >= f.index && match.index < f.index + f.original.length
    );
    if (!isInBlock) {
      formulas.push({
        type: 'inline',
        original: match[0],
        content: match[1].trim(),
        index: match.index
      });
    }
  }

  return formulas.sort((a, b) => a.index - b.index);
}

/**
 * 转换混合文本为 Word 文档
 */
async function convertTextToDocx(text) {
  const formulas = extractFormulas(text);
  const children = [];

  if (formulas.length === 0) {
    // 没有公式，纯文本
    const lines = text.split('\n');
    for (const line of lines) {
      if (line.trim()) {
        children.push(new Paragraph({
          children: [new TextRun(line)],
          spacing: { after: 200 }
        }));
      }
    }
  } else {
    // 处理文本和公式混合
    let lastIndex = 0;

    for (const formula of formulas) {
      // 处理公式前的文本
      if (formula.index > lastIndex) {
        const textBefore = text.substring(lastIndex, formula.index);
        const lines = textBefore.split('\n');

        for (let i = 0; i < lines.length; i++) {
          if (lines[i].trim()) {
            children.push(new Paragraph({
              children: [new TextRun(lines[i])],
              spacing: { after: 200 }
            }));
          }
        }
      }

      // 转换公式
      try {
        const preprocessed = preprocessLatex(formula.content);
        const omml = await latexToOMML(preprocessed);
        const mathEl = new ImportedXmlComponent(omml);

        children.push(new Paragraph({
          children: [mathEl],
          spacing: {
            before: formula.type === 'block' ? 400 : 100,
            after: formula.type === 'block' ? 400 : 100
          },
          alignment: formula.type === 'block' ? 'center' : 'left'
        }));
      } catch (error) {
        console.error('公式转换错误:', error.message);
        children.push(new Paragraph({
          children: [new TextRun(`[公式错误: ${formula.content}]`)],
          spacing: { after: 200 }
        }));
      }

      lastIndex = formula.index + formula.original.length;
    }

    // 处理最后剩余的文本
    if (lastIndex < text.length) {
      const textAfter = text.substring(lastIndex);
      if (textAfter.trim()) {
        children.push(new Paragraph({
          children: [new TextRun(textAfter)],
          spacing: { after: 200 }
        }));
      }
    }
  }

  const doc = new Document({
    sections: [{
      properties: {},
      children: children
    }]
  });

  return Packer.toBuffer(doc);
}

/**
 * 公式预览
 */
app.post('/api/preview', async (req, res) => {
  try {
    const { latex } = req.body;

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
});

/**
 * 文本转换接口
 */
app.post('/api/convert', async (req, res) => {
  try {
    const { text } = req.body;

    if (!text || typeof text !== 'string') {
      return res.status(400).json({ error: '请提供有效的文本内容' });
    }

    const buffer = await convertTextToDocx(text);

    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': 'attachment; filename="latex-output.docx"',
      'Content-Length': buffer.length
    });

    res.send(buffer);

  } catch (error) {
    console.error('转换接口错误:', error);
    res.status(500).json({ error: '转换失败', details: error.message });
  }
});

/**
 * 文件上传转换接口
 */
app.post('/api/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: '请上传文件' });
    }

    const text = req.file.buffer.toString('utf8');
    const originalName = path.basename(req.file.originalname, path.extname(req.file.originalname));
    const buffer = await convertTextToDocx(text);

    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename="${originalName}.docx"`,
      'Content-Length': buffer.length
    });

    res.send(buffer);

  } catch (error) {
    console.error('上传转换错误:', error);
    res.status(500).json({ error: '转换失败', details: error.message });
  }
});

/**
 * 健康检查
 */
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// 错误处理
app.use((err, req, res, next) => {
  console.error('服务器错误:', err);
  res.status(500).json({ error: err.message || '服务器内部错误' });
});

// 启动服务器
app.listen(PORT, () => {
  console.log(`
╔════════════════════════════════════════════════════════════╗
║     LaTeX to Word 转换服务已启动                            ║
╠════════════════════════════════════════════════════════════╣
║  本地访问: http://localhost:${PORT}                            ║
║                                                            ║
║  API 接口:                                                  ║
║  - POST /api/preview   - 公式预览                           ║
║  - POST /api/convert   - 文本转换                          ║
║  - POST /api/upload    - 文件上传转换                       ║
╚════════════════════════════════════════════════════════════╝
  `);
});