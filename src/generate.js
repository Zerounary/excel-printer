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
