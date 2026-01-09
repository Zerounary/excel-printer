export function normalizeMaxColumns(maxColumns) {
  if (typeof maxColumns !== 'number' || !Number.isFinite(maxColumns) || maxColumns < 1) {
    return 6;
  }
  return Math.floor(maxColumns);
}

export function toExcelColumnName(colNumber1Indexed) {
  let n = colNumber1Indexed;
  let name = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

export function getByPath(obj, pathStr) {
  if (!pathStr) return undefined;
  const parts = String(pathStr)
    .split('.')
    .map((s) => s.trim())
    .filter(Boolean);
  let cur = obj;
  for (const p of parts) {
    if (cur == null) return undefined;
    cur = cur[p];
  }
  return cur;
}
