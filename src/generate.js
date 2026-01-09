import fs from 'node:fs/promises';
import path from 'node:path';

import { applyTemplate } from './core/template.js';

export { generateWorkbookFromConfig } from './core/generate.js';

export async function generateXlsxFileFromConfig(config, outPath) {
  const { generateWorkbookFromConfig } = await import('./core/generate.js');
  const workbook = generateWorkbookFromConfig(config);
  await workbook.xlsx.writeFile(outPath);
  return outPath;
}

function sanitizeFileComponent(value) {
  const str = typeof value === 'string' ? value : String(value ?? 'output');
  return str.replace(/[\\/:*?"<>|]/g, '_');
}

function ensureXlsxExtension(name) {
  return name.toLowerCase().endsWith('.xlsx') ? name : `${name}.xlsx`;
}

function formatToday() {
  const date = new Date();
  const year = String(date.getFullYear());
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return { today: `${year}-${month}-${day}`, year, month, day };
}

function buildTemplateVariables(variables, extra) {
  const base = variables && typeof variables === 'object' ? { ...variables } : {};
  return { ...base, ...extra };
}

function buildFilenameTemplateContext({ variables, fileName, baseName }) {
  const { today, year, month, day } = formatToday();
  return buildTemplateVariables(variables, {
    file: { name: fileName, baseName },
    today,
    date: {
      today,
      year,
      month,
      day,
    },
  });
}

function resolveNameFromFile({ baseName, fileName, variables }) {
  const context = buildFilenameTemplateContext({ variables, fileName, baseName });
  const resolved = applyTemplate(baseName, context);
  const trimmed = typeof resolved === 'string' ? resolved.trim() : '';
  const finalBase = trimmed || baseName || 'output';
  const sanitized = sanitizeFileComponent(finalBase);
  return ensureXlsxExtension(sanitized);
}

export async function generateXlsxFileFromConfigFile({ configPath, outPath }) {
  const absConfigPath = path.isAbsolute(configPath) ? configPath : path.join(process.cwd(), configPath);
  const absOutPath = path.isAbsolute(outPath) ? outPath : path.join(process.cwd(), outPath);
  const configStr = await fs.readFile(absConfigPath, 'utf-8');
  const config = JSON.parse(configStr);
  await generateXlsxFileFromConfig(config, absOutPath);
  return absOutPath;
}

export async function generateXlsxFilesFromConfigDir({ configDir, outDir = 'out' }) {
  if (!configDir) {
    throw new Error('configDir is required');
  }

  const absConfigDir = path.isAbsolute(configDir) ? configDir : path.join(process.cwd(), configDir);
  const absOutDir = path.isAbsolute(outDir) ? outDir : path.join(process.cwd(), outDir);

  await fs.mkdir(absOutDir, { recursive: true });
  const entries = await fs.readdir(absConfigDir, { withFileTypes: true });
  const jsonFiles = entries
    .filter((entry) => entry.isFile())
    .map((entry) => entry.name)
    .filter((name) => name.toLowerCase().endsWith('.json'));

  const results = [];
  for (const fileName of jsonFiles) {
    const configPath = path.join(absConfigDir, fileName);
    const configStr = await fs.readFile(configPath, 'utf-8');
    const config = JSON.parse(configStr);
    const variables = config?.variables;

    const baseName = path.parse(fileName).name || 'output';
    const fileWithExt = resolveNameFromFile({ baseName, fileName, variables });
    const outPath = path.join(absOutDir, fileWithExt);
    const generated = await generateXlsxFileFromConfig(config, outPath);
    results.push({ configPath, outPath: generated });
  }

  return results;
}
