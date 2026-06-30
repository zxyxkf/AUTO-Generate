import express from 'express';
import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { documentTypes, existingPath, rootDir, themes } from './config.js';

const app = express();
const port = Number(process.env.PORT || 5174);

app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});
app.use(express.json({ limit: '20mb' }));
app.use('/generated', express.static(path.join(rootDir, 'generated')));
app.use('/assets', express.static(path.join(rootDir, 'assets')));

const distDir = path.join(rootDir, 'dist');
if (fs.existsSync(distDir)) {
  app.use(express.static(distDir));
}

app.get('/api/templates', (_req, res) => {
  res.json({
    documentTypes,
    themes: Object.entries(themes).map(([id, theme]) => ({
      id,
      label: theme.label,
      quotationLabel: theme.quotationLabel || theme.label,
      company: theme.company,
      supplierInfo: theme.supplierInfo || '',
      hasQuotation: theme.hasQuotation !== false,
      hasSupply: theme.hasSupply === true && Boolean(theme.supplyTemplate),
      hasContract: theme.hasContract !== false && Boolean(theme.contractTemplate),
      stamp: publicAssetUrl(preferredOfficialStamp(theme)) || publicAssetUrl(theme.fallbackStamp),
      contractStamp: publicAssetUrl(preferredContractStamp(theme)) || publicAssetPath(theme.contractStamp),
      contractStampExists: Boolean(existingPath(preferredContractStamp(theme))),
    })),
  });
});

app.post('/api/generate/:type', async (req, res) => {
  try {
    const type = req.params.type;
    if (!documentTypes[type]) return res.status(404).json({ error: `未知文档类型：${type}` });
    const payload = normalizePayload(type, req.body || {});
    const result = await runPowerShell(path.join(rootDir, 'scripts', 'office', 'generate.ps1'), {
      rootDir,
      type,
      payload,
    });
    res.json(result);
  } catch (error) {
    res.status(500).json({ error: error.message, details: error.details || '' });
  }
});

app.post('/api/preview/:type', async (req, res) => {
  try {
    const type = req.params.type;
    if (!documentTypes[type]) return res.status(404).json({ error: `未知文档类型：${type}` });
    const payload = normalizePayload(type, req.body || {});
    payload.preview = true;
    const result = await runPowerShell(path.join(rootDir, 'scripts', 'office', 'generate.ps1'), {
      rootDir,
      type,
      payload,
    });
    res.json(result);
  } catch (error) {
    res.status(500).json({ error: error.message, details: error.details || '' });
  }
});

if (fs.existsSync(distDir)) {
  app.get(/.*/, (req, res, next) => {
    if (req.path.startsWith('/api/') || req.path.startsWith('/generated/') || req.path.startsWith('/assets/')) {
      next();
      return;
    }
    res.sendFile(path.join(distDir, 'index.html'));
  });
}

app.listen(port, '127.0.0.1', () => {
  console.log(`Office generator listening on http://127.0.0.1:${port}`);
});

function normalizePayload(type, body) {
  const themeId = body.themeId || defaultThemeFor(type);
  const theme = themes[themeId];
  if (!theme) throw new Error(`未知主题：${themeId}`);
  if (type === 'supply' && !theme.supplyTemplate) throw new Error(`${theme.label} 暂无供货清单模板`);
  if (type === 'contract' && !theme.contractTemplate) throw new Error(`${theme.label} 暂无合同模板`);

  const templatePath = templateFor(type, theme);
  const stampType = body.stampType === 'official' ? 'official' : 'contract';
  const stampPath = stampPathFor(theme, stampType);
  return {
    ...body,
    type,
    themeId,
    stampType,
    themeLabel: theme.label,
    company: theme.company,
    supplierInfo: theme.supplierInfo || '',
    templatePath,
    stampPath,
    date: documentDateText(body.date),
    rows: Array.isArray(body.rows) ? body.rows : [],
    items: Array.isArray(body.items) ? body.items : Array.isArray(body.rows) ? body.rows : [],
  };
}

function defaultThemeFor(type) {
  if (type === 'contract') return 'contract_teyilai_invoice';
  if (type === 'supply') return 'contract_teyilai_invoice';
  return 'supply_teyilai_invoice';
}

function templateFor(type, theme) {
  if (type === 'quotation') return documentTypes.quotation.template;
  if (type === 'supply') return theme.supplyTemplate;
  return theme.contractTemplate;
}

function stampPathFor(theme, stampType) {
  if (stampType === 'contract') return existingPath(preferredContractStamp(theme));
  return existingPath(preferredOfficialStamp(theme)) || existingPath(theme.fallbackStamp);
}

function preferredOfficialStamp(theme) {
  return preferredVividStamp(theme.stamp);
}

function preferredContractStamp(theme) {
  return preferredVividStamp(theme.contractStamp);
}

function preferredVividStamp(pathValue) {
  const base = pathValue || '';
  const vivid = base.replace(/\.png$/i, '2.png');
  return existingPath(vivid) ? vivid : base;
}

function publicAssetUrl(relativePath) {
  const fullPath = existingPath(relativePath);
  if (!fullPath) return '';
  const assetRelative = path.relative(path.join(rootDir, 'assets'), fullPath).replaceAll(path.sep, '/');
  return `/assets/${assetRelative}`;
}

function publicAssetPath(relativePath) {
  if (!relativePath) return '';
  const normalized = relativePath.replaceAll(path.sep, '/');
  if (normalized.startsWith('assets/')) return `/${normalized}`;
  return publicAssetUrl(relativePath);
}

function documentDateText(value) {
  const text = String(value ?? '').trim();
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
  if (!match) return text;
  return `${Number(match[1])}年${Number(match[2])}月${Number(match[3])}日`;
}

function runPowerShell(script, input) {
  return new Promise((resolve, reject) => {
    const inputDir = path.join(rootDir, 'work');
    fs.mkdirSync(inputDir, { recursive: true });
    const inputPath = path.join(inputDir, `office-input-${Date.now()}-${Math.random().toString(16).slice(2)}.json`);
    fs.writeFileSync(inputPath, JSON.stringify(input), 'utf8');
    const child = spawn('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script, '-InputJsonPath', inputPath], {
      cwd: rootDir,
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    child.stdout.setEncoding('utf8');
    child.stderr.setEncoding('utf8');
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', (chunk) => { stdout += chunk; });
    child.stderr.on('data', (chunk) => { stderr += chunk; });
    child.on('error', reject);
    child.on('close', (code) => {
      fs.rmSync(inputPath, { force: true });
      if (code !== 0) {
        const error = new Error(`Office 生成失败，退出码 ${code}`);
        error.details = stderr || stdout;
        reject(error);
        return;
      }
      try {
        resolve(JSON.parse(stdout.trim()));
      } catch (error) {
        error.details = stdout || stderr;
        reject(error);
      }
    });
  });
}
