import ExcelJS from 'exceljs';

import { applySheetDefaults, resolvePaperOptions } from './layout.js';
import { applyCellStyle, applyDefaultBorders, isPlainObject, mergeStyles } from './styles.js';
import { normalizeMaxColumns, toExcelColumnName } from './utils.js';
import { renderForm, renderSpaceRow, renderTable, renderText, renderTitle } from './renderers.js';
import { applyTemplate } from './template.js';

function normalizeSheetsConfig(config) {
  if (!isPlainObject(config)) {
    return { style: undefined, variables: undefined, sheets: [] };
  }

  const globalStyle = isPlainObject(config.style) ? config.style : undefined;
  const variables = isPlainObject(config.variables) ? config.variables : isPlainObject(config.vars) ? config.vars : undefined;

  if (Array.isArray(config.sheets)) {
    const sheets = config.sheets.filter((s) => isPlainObject(s));
    return { style: globalStyle, variables, sheets };
  }

  const legacySheet = {
    name: config.name || config.sheetName || '打印',
    maxColumns: config.maxColumns,
    style: config.sheetStyle,
    rows: config.rows,
  };

  return { style: globalStyle, variables, sheets: [legacySheet] };
}

function normalizeBlock(block, variables) {
  if (!block || typeof block !== 'object') return block;
  const withVars = variables ? applyTemplate(block, variables) : block;

  if (!withVars || typeof withVars !== 'object') return withVars;

  if (withVars.type === 'title') {
    const value = typeof withVars.value !== 'undefined' ? withVars.value : withVars.val;
    return { ...withVars, value };
  }

  if (withVars.type === 'text') {
    const value = typeof withVars.value !== 'undefined' ? withVars.value : withVars.val;
    return { ...withVars, value };
  }

  return withVars;
}

export function generateWorkbookFromConfig(config) {
  const { style: globalStyle, variables, sheets } = normalizeSheetsConfig(config);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'excel-printer';
  workbook.created = new Date();

  for (const sheet of sheets) {
    const maxColumns = normalizeMaxColumns(sheet.maxColumns);
    const sheetStyle = isPlainObject(sheet.style) ? sheet.style : undefined;
    const mergedGlobalStyle = mergeStyles(globalStyle, sheetStyle);

    const worksheet = workbook.addWorksheet(sheet.name || 'Sheet', {
      views: [{ showGridLines: false }],
    });

    const paperOptions = resolvePaperOptions(sheet.paper || sheet.paperSize);
    const totalWidthOverride =
      typeof sheet.totalWidth === 'number' && Number.isFinite(sheet.totalWidth) ? sheet.totalWidth : undefined;
    const marginsOverride = isPlainObject(sheet.margins) ? sheet.margins : undefined;

    applySheetDefaults(worksheet, maxColumns, {
      ...paperOptions,
      totalWidth: totalWidthOverride ?? paperOptions.totalWidth,
      margins: marginsOverride,
    });

    let cursorRow = 1;
    const noBorderRows = new Set();

    for (const rawBlock of sheet.rows ?? []) {
      const block = normalizeBlock(rawBlock, variables);
      if (!block || typeof block !== 'object') continue;

      if (block.type === 'title') {
        const base = {
          font: { name: '宋体', size: 16, bold: true },
          alignment: { vertical: 'middle', horizontal: 'center', wrapText: true },
        };
        cursorRow = renderTitle({ worksheet, maxColumns, cursorRow, value: block.value });
        const cell = worksheet.getCell(cursorRow - 1, 1);
        applyCellStyle(cell, mergeStyles(base, mergedGlobalStyle, block.style));
        continue;
      }

      if (block.type === 'space-row') {
        cursorRow = renderSpaceRow({
          worksheet,
          maxColumns,
          cursorRow,
          count: block.count,
          height: block.height,
          noBorderRows,
        });
        continue;
      }

      if (block.type === 'form') {
        cursorRow = renderForm({
          worksheet,
          maxColumns,
          cursorRow,
          fields: block.fields,
          style: mergeStyles(mergedGlobalStyle, block.style),
          fieldStyle: block.fieldStyle,
        });
        continue;
      }

      if (block.type === 'table') {
        cursorRow = renderTable({
          worksheet,
          maxColumns,
          cursorRow,
          headers: block.headers,
          rows: block.rows,
          columnWidths: block.columnWidths,
          mergeHeaderSame: block.mergeHeaderSame,
          mergeColumns: block.mergeColumns,
          style: mergeStyles(mergedGlobalStyle, block.style),
          headerStyle: block.headerStyle,
          bodyStyle: block.bodyStyle,
          columnStyles: block.columnStyles,
          rowStyles: block.rowStyles,
          cellStyles: block.cellStyles,
        });
        continue;
      }

      if (block.type === 'text') {
        cursorRow = renderText({
          worksheet,
          maxColumns,
          cursorRow,
          value: block.value,
          style: mergeStyles(mergedGlobalStyle, block.style),
        });
        continue;
      }
    }

    const lastRow = Math.max(1, cursorRow - 1);
    const lastColName = toExcelColumnName(maxColumns);
    worksheet.pageSetup.printArea = `A1:${lastColName}${lastRow}`;

    applyDefaultBorders(worksheet, 1, lastRow, 1, maxColumns, { skipRows: noBorderRows });
  }

  return workbook;
}
