import fs from 'node:fs/promises';
import path from 'node:path';

export { generateWorkbookFromConfig } from './core/generate.js';

export async function generateXlsxFileFromConfig(config, outPath) {
  const { generateWorkbookFromConfig } = await import('./core/generate.js');
  const workbook = generateWorkbookFromConfig(config);
  await workbook.xlsx.writeFile(outPath);
  return outPath;
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
    const baseName = path.parse(fileName).name || 'output';
    const outPath = path.join(absOutDir, `${baseName}.xlsx`);
    const generated = await generateXlsxFileFromConfigFile({ configPath, outPath });
    results.push({ configPath, outPath: generated });
  }

  return results;
}
