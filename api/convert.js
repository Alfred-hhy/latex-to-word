const { latexToOMML } = require('latex-to-omml');
const { Document, Packer, Paragraph, TextRun, ImportedXmlComponent } = require('docx');

// LaTeX 公式匹配模式
const inlineMathPattern = /\$([^$]+)\$/g;
const blockMathPattern = /\$\$([^$]+)\$\$/g;

/**
 * 预处理 LaTeX
 */
function preprocessLatex(latex) {
  let processed = latex;
  processed = processed.replace(/\s+\\$/g, ' \\;');
  processed = processed.replace(/\\\s+/g, ' ');
  return processed;
}

/**
 * 从文本中提取所有公式
 */
function extractFormulas(text) {
  const formulas = [];

  let match;
  while ((match = blockMathPattern.exec(text)) !== null) {
    formulas.push({
      type: 'block',
      original: match[0],
      content: match[1].trim(),
      index: match.index
    });
  }

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

    const { text } = req.body || {};

    if (!text || typeof text !== 'string') {
      return res.status(400).json({ error: '请提供有效的文本内容' });
    }

    const formulas = extractFormulas(text);
    const children = [];

    if (formulas.length === 0) {
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
      let lastIndex = 0;

      for (const formula of formulas) {
        if (formula.index > lastIndex) {
          const textBefore = text.substring(lastIndex, formula.index);
          const lines = textBefore.split('\n');

          for (const line of lines) {
            if (line.trim()) {
              children.push(new Paragraph({
                children: [new TextRun(line)],
                spacing: { after: 200 }
              }));
            }
          }
        }

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

    const buffer = await Packer.toBuffer(doc);

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
}

module.exports = handler;