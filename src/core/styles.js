export function setThinBorder(cell) {
  cell.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' },
  };
}

export function styleRangeBorder(worksheet, row, colStart, colEnd) {
  for (let c = colStart; c <= colEnd; c += 1) {
    setThinBorder(worksheet.getCell(row, c));
  }
}

export function isPlainObject(val) {
  return !!val && typeof val === 'object' && !Array.isArray(val);
}

export function mergeDeep(base, override) {
  if (!isPlainObject(base)) return isPlainObject(override) ? { ...override } : override;
  if (!isPlainObject(override)) return { ...base };
  const out = { ...base };
  for (const [k, v] of Object.entries(override)) {
    if (isPlainObject(v) && isPlainObject(out[k])) {
      out[k] = mergeDeep(out[k], v);
    } else {
      out[k] = v;
    }
  }
  return out;
}

export function mergeStyles(...styles) {
  let out = {};
  for (const s of styles) {
    if (!isPlainObject(s)) continue;
    out = mergeDeep(out, s);
  }
  return out;
}

export function applyCellStyle(cell, style) {
  if (!isPlainObject(style)) return;
  if (isPlainObject(style.font)) cell.font = style.font;
  if (isPlainObject(style.alignment)) cell.alignment = style.alignment;
  if (isPlainObject(style.border)) cell.border = style.border;
  if (isPlainObject(style.fill)) cell.fill = style.fill;
  if (typeof style.numFmt === 'string') cell.numFmt = style.numFmt;
}

export function applyDefaultBorders(worksheet, rowStart, rowEnd, colStart, colEnd, { skipRows } = {}) {
  for (let r = rowStart; r <= rowEnd; r += 1) {
    if (skipRows && typeof skipRows.has === 'function' && skipRows.has(r)) continue;
    for (let c = colStart; c <= colEnd; c += 1) {
      const cell = worksheet.getCell(r, c);
      if (!cell.border) {
        setThinBorder(cell);
      }
    }
  }
}
