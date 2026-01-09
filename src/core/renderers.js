import { estimateRowHeight, mergeAcross } from './layout.js';
import { applyCellStyle, mergeStyles, styleRangeBorder } from './styles.js';

export function renderTitle({ worksheet, maxColumns, cursorRow, value }) {
  const cell = mergeAcross(worksheet, cursorRow, 1, maxColumns);
  cell.value = value ?? '';
  cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  cell.font = { name: '宋体', size: 16, bold: true };
  worksheet.getRow(cursorRow).height = 30;
  return cursorRow + 1;
}

export function renderSpaceRow({ worksheet, maxColumns, cursorRow, count = 1, height, noBorderRows }) {
  const n = typeof count === 'number' && Number.isFinite(count) ? Math.max(1, Math.floor(count)) : 1;
  for (let i = 0; i < n; i += 1) {
    const row = cursorRow + i;
    const cell = mergeAcross(worksheet, row, 1, maxColumns);
    cell.value = '';
    for (let c = 1; c <= maxColumns; c += 1) {
      styleRangeBorder(worksheet, row, c, c);
    }
    if (typeof height === 'number' && Number.isFinite(height) && height > 0) {
      worksheet.getRow(row).height = height;
    }
    if (noBorderRows && typeof noBorderRows.add === 'function') {
      noBorderRows.add(row);
    }
  }
  return cursorRow + n;
}

export function renderText({ worksheet, maxColumns, cursorRow, value, style }) {
  const cell = mergeAcross(worksheet, cursorRow, 1, maxColumns);
  cell.value = value ?? '';
  cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  cell.font = { name: '宋体', size: 11 };
  applyCellStyle(cell, style);
  const totalWidth = worksheet.columns.reduce((sum, col) => sum + (col.width ?? 14), 0);
  worksheet.getRow(cursorRow).height = estimateRowHeight(cell.value, totalWidth, 18);
  return cursorRow + 1;
}

export function renderForm({ worksheet, maxColumns, cursorRow, fields, style, fieldStyle }) {
  const perRow = 2;
  const fieldWidth = Math.floor(maxColumns / perRow);
  const remainder = maxColumns - fieldWidth * perRow;

  let row = cursorRow;
  for (let i = 0; i < (fields?.length ?? 0); i += 1) {
    const field = fields[i];
    const idxInRow = i % perRow;
    if (idxInRow === 0 && i !== 0) {
      row += 1;
    }

    const leftWidth = idxInRow === 0 ? fieldWidth + remainder : fieldWidth;
    const colStart = idxInRow === 0 ? 1 : 1 + (fieldWidth + remainder);
    const colEnd = Math.min(maxColumns, colStart + leftWidth - 1);

    const cell = mergeAcross(worksheet, row, colStart, colEnd);
    const label = field?.label ?? '';
    const val = field?.value ?? '';
    cell.value = val ? `${label}：${val}` : `${label}：`;
    cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    cell.font = { name: '宋体', size: 11 };

    const mergedStyle = mergeStyles(style, fieldStyle, field?.style);
    applyCellStyle(cell, mergedStyle);
  }

  worksheet.getRow(row).height = 20;
  return row + 1;
}

export function renderTable({
  worksheet,
  maxColumns,
  cursorRow,
  headers,
  rows,
  columnWidths,
  mergeHeaderSame,
  mergeColumns,
  style,
  headerStyle,
  bodyStyle,
  columnStyles,
  rowStyles,
  cellStyles,
}) {
  let row = cursorRow;

  const headerLabels = Array.from({ length: maxColumns }, (_, i) => headers?.[i] ?? '');

  if (columnWidths && Array.isArray(worksheet.columns)) {
    if (Array.isArray(columnWidths)) {
      for (let i = 0; i < Math.min(maxColumns, columnWidths.length); i += 1) {
        const w = columnWidths[i];
        if (typeof w === 'number' && Number.isFinite(w) && w > 0) {
          if (!worksheet.columns[i]) worksheet.columns[i] = {};
          worksheet.columns[i].width = w;
        }
      }
    } else if (typeof columnWidths === 'object') {
      for (const [key, w] of Object.entries(columnWidths)) {
        if (typeof w !== 'number' || !Number.isFinite(w) || w <= 0) continue;

        let idx = -1;
        const n = Number(key);
        if (Number.isFinite(n) && String(n) === key) {
          idx = Math.floor(n) - 1;
        } else {
          idx = headerLabels.findIndex((h) => h === key);
        }

        if (idx >= 0 && idx < maxColumns) {
          if (!worksheet.columns[idx]) worksheet.columns[idx] = {};
          worksheet.columns[idx].width = w;
        }
      }
    }
  }

  const mergeColumnSet = new Set();
  if (Array.isArray(mergeColumns)) {
    for (const item of mergeColumns) {
      if (typeof item === 'number' && Number.isFinite(item)) {
        const idx = Math.floor(item);
        if (idx >= 1 && idx <= maxColumns) mergeColumnSet.add(idx);
        continue;
      }
      if (typeof item === 'string' && item) {
        const found = headerLabels.findIndex((h) => h === item);
        if (found >= 0) mergeColumnSet.add(found + 1);
      }
    }
  }

  const filledRows = (rows ?? []).map((r) => Array.from({ length: maxColumns }, (_, i) => r?.[i] ?? ''));
  if (mergeColumnSet.size > 0) {
    const last = new Map();
    for (const r of filledRows) {
      for (const c of mergeColumnSet) {
        const idx = c - 1;
        const v = r[idx];
        if (v === '' || v === null || typeof v === 'undefined') {
          if (last.has(c)) r[idx] = last.get(c);
        } else {
          last.set(c, v);
        }
      }
    }
  }

  for (let c = 1; c <= maxColumns; c += 1) {
    const cell = worksheet.getCell(row, c);
    cell.value = headerLabels[c - 1];
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.font = { name: '宋体', size: 11, bold: true };

    const key = `0,${c}`;
    const mergedStyle = mergeStyles(style, headerStyle, columnStyles?.[c - 1], cellStyles?.[key]);
    applyCellStyle(cell, mergedStyle);
  }
  styleRangeBorder(worksheet, row, 1, maxColumns);
  worksheet.getRow(row).height = 22;

  if (mergeHeaderSame) {
    let start = 1;
    while (start <= maxColumns) {
      const label = headerLabels[start - 1];
      let end = start;
      while (end + 1 <= maxColumns && headerLabels[end] === label) {
        end += 1;
      }
      if (end > start) {
        worksheet.mergeCells(row, start, row, end);
      }
      start = end + 1;
    }
  }

  row += 1;

  const colWidths = worksheet.columns.map((c) => c.width ?? 14);
  let tableRowIndex = 1;
  for (const dataRow of filledRows) {
    for (let c = 1; c <= maxColumns; c += 1) {
      const cell = worksheet.getCell(row, c);
      cell.value = dataRow?.[c - 1] ?? '';
      cell.alignment = {
        vertical: 'middle',
        horizontal: c === 1 ? 'left' : 'center',
        wrapText: true,
      };
      cell.font = { name: '宋体', size: 11 };

      const key = `${tableRowIndex},${c}`;
      const mergedStyle = mergeStyles(
        style,
        bodyStyle,
        columnStyles?.[c - 1],
        rowStyles?.[tableRowIndex - 1],
        cellStyles?.[key]
      );
      applyCellStyle(cell, mergedStyle);
    }

    styleRangeBorder(worksheet, row, 1, maxColumns);
    const firstColText = worksheet.getCell(row, 1).value;
    worksheet.getRow(row).height = estimateRowHeight(firstColText, colWidths[0] ?? 30, 20);
    row += 1;
    tableRowIndex += 1;
  }

  if (mergeColumnSet.size > 0 && filledRows.length > 0) {
    const firstDataRow = cursorRow + 1;
    for (const col of mergeColumnSet) {
      let segStart = 0;
      let segVal = filledRows[0]?.[col - 1];
      for (let i = 1; i <= filledRows.length; i += 1) {
        const curVal = i < filledRows.length ? filledRows[i]?.[col - 1] : Symbol('end');
        if (curVal !== segVal) {
          const segEnd = i - 1;
          if (segEnd > segStart && segVal !== '' && segVal !== null && typeof segVal !== 'undefined') {
            worksheet.mergeCells(firstDataRow + segStart, col, firstDataRow + segEnd, col);
          }
          segStart = i;
          segVal = curVal;
        }
      }
    }
  }

  return row;
}
