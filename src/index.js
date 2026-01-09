import { parseArgs } from './cli.js';
import { generateXlsxFileFromConfigFile } from './generate.js';

async function main() {
  const { config: configPathRaw, out: outPathRaw } = parseArgs(process.argv);

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
