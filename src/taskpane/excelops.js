/* ===================================================================
   excelops.js
   All direct interaction with the Excel workbook via the Office.js
   Excel JavaScript API: reading the active cell, creating/resizing/
   moving/deleting native Excel Tables, and reading current table
   geometry (used for overlap detection on refresh).
   =================================================================== */

function normalizeValue(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") return JSON.stringify(v);
  return v;
}

function rowsToValues(fields, rows) {
  return rows.map(r => fields.map(f => normalizeValue(r[f])));
}

/** Returns { sheetName, address } of the currently selected single cell. */
async function getActiveCellInfo() {
  let info = null;
  await Excel.run(async context => {
    const cell = context.workbook.getActiveCell();
    cell.load("address,rowIndex,columnIndex");
    const sheet = cell.worksheet;
    sheet.load("name");
    await context.sync();
    info = {
      sheetName: sheet.name,
      address: cell.address.includes("!") ? cell.address.split("!")[1] : cell.address,
      rowIndex: cell.rowIndex,
      columnIndex: cell.columnIndex
    };
  });
  return info;
}

/** True if a table with this name already exists anywhere in the workbook. */
async function tableNameExists(tableName) {
  let exists = false;
  await Excel.run(async context => {
    const table = context.workbook.tables.getItemOrNullObject(tableName);
    table.load("name");
    await context.sync();
    exists = !table.isNullObject;
  });
  return exists;
}

/** Creates a brand-new Excel Table with header row + data rows at the anchor cell. */
async function createPlaceholderTable(sheetName, anchorAddress, tableName, fields, rows) {
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const anchor = sheet.getRange(anchorAddress);
    anchor.load("rowIndex,columnIndex");
    await context.sync();

    const numRows = rows.length + 1;
    const numCols = fields.length;
    const range = sheet.getRangeByIndexes(anchor.rowIndex, anchor.columnIndex, numRows, numCols);
    range.values = [fields, ...rowsToValues(fields, rows)];

    const table = sheet.tables.add(range, true);
    table.name = tableName;
    sheet.getRange(anchorAddress).select();
    await context.sync();
  });
}

/** Reads the current occupied rectangle of a registered table, or null if it no longer exists. */
async function getCurrentTableRect(sheetName, tableName) {
  let rect = null;
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const table = sheet.tables.getItemOrNullObject(tableName);
    table.load("name");
    await context.sync();
    if (table.isNullObject) {
      rect = null;
      return;
    }
    const range = table.getRange();
    range.load("rowIndex,columnIndex,rowCount,columnCount,address");
    await context.sync();
    rect = {
      sheetName,
      tableName,
      startRow: range.rowIndex,
      startCol: range.columnIndex,
      endRow: range.rowIndex + range.rowCount - 1,
      endCol: range.columnIndex + range.columnCount - 1,
      address: range.address
    };
  });
  return rect;
}

/** Simple axis-aligned rectangle overlap test (same sheet only). */
function rectsOverlap(a, b) {
  if (a.sheetName !== b.sheetName) return false;
  return !(a.endRow < b.startRow || a.startRow > b.endRow || a.endCol < b.startCol || a.startCol > b.endCol);
}

/**
 * Re-writes a table's data in place, growing/shrinking its range as
 * needed, anchored at its original top-left cell. Safe to call when
 * the new row count differs from the old one.
 */
async function refreshPlaceholderTable(tableDef, rows) {
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItem(tableDef.sheetName);
    const oldTable = sheet.tables.getItemOrNullObject(tableDef.tableName);
    oldTable.load("name");
    await context.sync();

    let oldRect = null;
    if (!oldTable.isNullObject) {
      const r = oldTable.getRange();
      r.load("rowIndex,columnIndex,rowCount,columnCount");
      await context.sync();
      oldRect = {
        startRow: r.rowIndex,
        startCol: r.columnIndex,
        endRow: r.rowIndex + r.rowCount - 1,
        endCol: r.columnIndex + r.columnCount - 1
      };
      oldTable.delete();
      await context.sync();
    }

    const anchor = sheet.getRange(tableDef.anchorAddress);
    anchor.load("rowIndex,columnIndex");
    await context.sync();

    const startRow = anchor.rowIndex;
    const startCol = anchor.columnIndex;
    const numRows = rows.length + 1;
    const numCols = tableDef.fields.length;

    if (oldRect) {
      const clearRowCount = Math.max(oldRect.endRow, startRow + numRows - 1) - startRow + 1;
      const clearColCount = Math.max(oldRect.endCol, startCol + numCols - 1) - startCol + 1;
      const clearRange = sheet.getRangeByIndexes(startRow, startCol, clearRowCount, clearColCount);
      clearRange.clear(Excel.ClearApplyTo.contents);
      await context.sync();
    }

    const newRange = sheet.getRangeByIndexes(startRow, startCol, numRows, numCols);
    newRange.values = [tableDef.fields, ...rowsToValues(tableDef.fields, rows)];
    const newTable = sheet.tables.add(newRange, true);
    newTable.name = tableDef.tableName;
    await context.sync();
  });
}

/**
 * Moves an existing table (its current data, as-is) to a new anchor
 * cell, used to resolve a refresh conflict. Returns the new anchor
 * address so the caller can update the registry.
 */
async function moveTableToNewAnchor(sheetName, tableName, newAnchorAddress) {
  let result = null;
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const table = sheet.tables.getItem(tableName);
    const range = table.getRange();
    range.load("values,rowIndex,columnIndex,rowCount,columnCount");
    await context.sync();

    const values = range.values;
    table.delete();
    const oldClear = sheet.getRangeByIndexes(range.rowIndex, range.columnIndex, range.rowCount, range.columnCount);
    oldClear.clear(Excel.ClearApplyTo.contents);
    await context.sync();

    const anchor = sheet.getRange(newAnchorAddress);
    anchor.load("rowIndex,columnIndex");
    await context.sync();

    const newRange = sheet.getRangeByIndexes(anchor.rowIndex, anchor.columnIndex, values.length, values[0].length);
    newRange.values = values;
    const newTable = sheet.tables.add(newRange, true);
    newTable.name = tableName;
    await context.sync();

    result = { anchorAddress: newAnchorAddress };
  });
  return result;
}

/** Deletes a table and clears its cells (used when the user removes a placeholder). */
async function deleteTableAndClear(sheetName, tableName) {
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const table = sheet.tables.getItemOrNullObject(tableName);
    table.load("name");
    await context.sync();
    if (table.isNullObject) return;

    const range = table.getRange();
    range.load("rowIndex,columnIndex,rowCount,columnCount");
    await context.sync();
    table.delete();
    const clearRange = sheet.getRangeByIndexes(range.rowIndex, range.columnIndex, range.rowCount, range.columnCount);
    clearRange.clear(Excel.ClearApplyTo.contents);
    await context.sync();
  });
}
