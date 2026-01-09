export function parseArgs(argv) {
  const args = {
    config: 'config.json',
    out: 'output.xlsx',
    configDir: undefined,
    outDir: 'out',
  };

  for (let i = 2; i < argv.length; i += 1) {
    const cur = argv[i];
    const next = argv[i + 1];

    if (cur === '--config' && next) {
      args.config = next;
      i += 1;
      continue;
    }

    if (cur === '--out' && next) {
      args.out = next;
      i += 1;
      continue;
    }

    if (cur === '--config-dir' && next) {
      args.configDir = next;
      i += 1;
      continue;
    }

    if (cur === '--out-dir' && next) {
      args.outDir = next;
      i += 1;
      continue;
    }
  }

  return args;
}
