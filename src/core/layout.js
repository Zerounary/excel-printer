const BASE_SIX_COL_PATTERN = [32, 14, 10, 12, 12, 18];
const BASE_SIX_TOTAL = BASE_SIX_COL_PATTERN.reduce((sum, width) => sum + width, 0);
const DEFAULT_MARGIN = {
  left: 0.3,
  right: 0.3,
  top: 0.4,
  bottom: 0.4,
  header: 0.2,
  footer: 0.2,
};

const PAPER_PRESETS = {
  A3: { paperSize: 8, totalWidth: 120 },
  A4: { paperSize: 9, totalWidth: BASE_SIX_TOTAL },
  A5: { paperSize: 11, totalWidth: 80 },
  A6: { paperSize: 70, totalWidth: 64 },
};

const DEFAULT_PAPER = PAPER_PRESETS.A4;

function clampTotalWidth(maxColumns, requestedTotal) {
  if (maxColumns <= 0) return BASE_SIX_TOTAL;
  const base = maxColumns === 6 ? BASE_SIX_TOTAL : maxColumns * 14;
  return Math.max(base, requestedTotal ?? base);
}

function buildColumnWidths(maxColumns, targetTotalWidth) {
  if (maxColumns <= 0) return [];
  if (maxColumns === 6) {
    const ratio = targetTotalWidth / BASE_SIX_TOTAL;
    return BASE_SIX_COL_PATTERN.map((width) => Number((width * ratio).toFixed(2)));
  }

  const evenWidth = targetTotalWidth / maxColumns;
  return Array.from({ length: maxColumns }, () => Number(evenWidth.toFixed(2)));
}

export function resolvePaperOptions(paperSizeInput) {
  if (typeof paperSizeInput === 'number' && Number.isFinite(paperSizeInput)) {
    return { ...DEFAULT_PAPER, paperSize: paperSizeInput };
  }
  if (typeof paperSizeInput === 'string') {
    const normalized = paperSizeInput.trim().toUpperCase();
    if (PAPER_PRESETS[normalized]) {
      return PAPER_PRESETS[normalized];
    }
  }
  return DEFAULT_PAPER;
}

export function estimateRowHeight(text, colWidth, base = 20) {
  if (!text) return base;
  const str = String(text);
  const approxCharsPerLine = Math.max(10, Math.floor(colWidth * 1.6));
  const lines = str
    .split(/\r?\n/)
    .reduce((acc, part) => acc + Math.max(1, Math.ceil(part.length / approxCharsPerLine)), 0);
  return Math.min(120, base + (lines - 1) * 14);
}

export function applySheetDefaults(worksheet, maxColumns, paperOptions = {}) {
  const { paperSize, totalWidth, margins } = { ...DEFAULT_PAPER, ...paperOptions };
  const targetTotalWidth = clampTotalWidth(maxColumns, totalWidth);
  const colWidths = buildColumnWidths(maxColumns, targetTotalWidth);
  worksheet.columns = colWidths.map((width) => ({ width }));

  worksheet.pageSetup = {
    paperSize,
    orientation: 'portrait',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    horizontalCentered: true,
    margins: margins || DEFAULT_MARGIN,
  };
  worksheet.properties.defaultRowHeight = 20;
}

export function mergeAcross(worksheet, row, colStart, colEnd) {
  if (colEnd > colStart) {
    worksheet.mergeCells(row, colStart, row, colEnd);
  }
  return worksheet.getCell(row, colStart);
}
