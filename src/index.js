import { parseArgs } from './cli.js';
import { generateXlsxFileFromConfigFile, generateXlsxFilesFromConfigDir } from './generate.js';

async function main() {
  const { config: configPathRaw, out: outPathRaw, configDir, outDir } = parseArgs(process.argv);

  if (configDir) {
    const results = await generateXlsxFilesFromConfigDir({
      configDir,
      outDir,
    });
    results.forEach(({ configPath, outPath }) => {
      process.stdout.write(`Generated: ${outPath} (from ${configPath})\n`);
    });
    if (results.length === 0) {
      process.stdout.write('No JSON config files found.\n');
    }
    return;
  }

  const outPath = await generateXlsxFileFromConfigFile({
    configPath: configPathRaw,
    outPath: outPathRaw,
  });
  process.stdout.write(`Generated: ${outPath}\n`);
}

main().catch((err) => {
  process.stderr.write(`${err?.stack || err}\n`);
  process.exitCode = 1;
});
