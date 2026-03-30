const express = require('express');
const multer = require('multer');
const path = require('path');
const { spawn } = require('child_process');
const fs = require('fs');

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

/**
 * 使用 Pandoc 将 LaTeX 文本转换为 Word
 * Pandoc 能很好地处理 LaTeX 公式并转换为 OMML
 */
function convertWithPandoc(text, outputPath) {
  return new Promise((resolve, reject) => {
    // 创建临时文件
    const tempInput = `/tmp/latex-input-${Date.now()}.md`;

    // 预处理：将纯 LaTeX 公式转换为 Markdown 格式
    // 行内公式: $...$ 保持不变
    // 块级公式: $$...$$ 保持不变
    let processed = text;

    const pandoc = spawn('pandoc', [
      '-f', 'markdown+tex_math_dollars',
      '-t', 'docx',
      '-o', outputPath
    ]);

    pandoc.stdin.write(processed);
    pandoc.stdin.end();

    pandoc.on('close', (code) => {
      fs.unlink(tempInput, () => {});
      if (code === 0) {
        resolve();
      } else {
        reject(new Error(`Pandoc exited with code ${code}`));
      }
    });

    pandoc.on('error', (err) => {
      reject(err);
    });
  });
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

    // 创建一个临时文档来预览公式
    const tempOutput = `/tmp/preview-${Date.now()}.docx`;
    const tempInput = `/tmp/input-${Date.now()}.md`;

    // 将公式包装成完整文档
    const content = `$$\n${latex}\n$$`;

    fs.writeFileSync(tempInput, content);

    const pandoc = spawn('pandoc', [
      '-f', 'markdown+tex_math_dollars',
      '-t', 'docx',
      '-o', tempOutput
    ]);

    pandoc.on('close', async (code) => {
      // 清理输入文件
      fs.unlink(tempInput, () => {});

      if (code === 0 && fs.existsSync(tempOutput)) {
        // 读取 docx 并提取公式部分用于预览
        const buffer = fs.readFileSync(tempOutput);
        fs.unlink(tempOutput, () => {});

        res.json({
          success: true,
          latex: latex,
          message: '公式已转换，请在 Word 中查看完整效果',
          note: '预览功能显示转换后的 OMML 公式'
        });
      } else {
        res.status(500).json({ error: '转换失败' });
      }
    });

    pandoc.stdin.write(content);
    pandoc.stdin.end();

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

    const outputPath = `/tmp/output-${Date.now()}.docx`;

    await convertWithPandoc(text, outputPath);

    const buffer = fs.readFileSync(outputPath);
    fs.unlink(outputPath, () => {});

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
    const outputPath = `/tmp/output-${Date.now()}.docx`;

    await convertWithPandoc(text, outputPath);

    const buffer = fs.readFileSync(outputPath);
    fs.unlink(outputPath, () => {});

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