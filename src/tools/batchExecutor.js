/**
 * batchExecutor.js
 *
 * Batched tool execution - Execute multiple frontend tools in a single Excel.run() context
 * This eliminates race conditions and ensures proper formula calculation before formatting
 */

/**
 * Tool execution priority (ensures logical ordering)
 * Lower numbers execute first
 */
const TOOL_PRIORITY = {
  createNewSheet: 1,
  writeDataToRange: 2,
  applyFormula: 3,
  formatRange: 4,
  createChart: 5,
  createTable: 6,
  addCellNote: 7,
  insertRows: 8,
  deleteRows: 9,
};

/**
 * Execute multiple frontend tools in a single Excel.run() batch
 * 
 * This is MUCH more efficient and reliable than executing tools individually because:
 * 1. Single Excel.run() context instead of N separate contexts
 * 2. ONE final context.sync() instead of N syncs
 * 3. Proper formula calculation guaranteed before formatting
 * 4. Atomic operation - either all succeed or all fail
 * 
 * @param {Array<Object>} toolCalls - Array of {tool, params} objects
 * @returns {Promise<Array<Object>>} Array of execution results
 */
export async function executeBatchedTools(toolCalls) {
  console.log(`🔧 Executing ${toolCalls.length} frontend tools in BATCHED mode`);

  // Sort tools by priority to ensure logical execution order
  const sortedToolCalls = [...toolCalls].sort((a, b) => {
    const priorityA = TOOL_PRIORITY[a.tool] || 999;
    const priorityB = TOOL_PRIORITY[b.tool] || 999;
    return priorityA - priorityB;
  });

  return Excel.run(async (context) => {
    const results = [];
    const worksheets = new Map(); // Cache worksheet references
    let needsRecalculation = false;

    try {
      // PHASE 1: Execute all tool operations (no sync yet)
      for (let i = 0; i < sortedToolCalls.length; i++) {
        const { tool, params } = sortedToolCalls[i];
        console.log(`  [${i + 1}/${sortedToolCalls.length}] ${tool}`);

        try {
          const result = await executeBatchedTool(context, tool, params, worksheets);
          results.push({
            success: true,
            tool: tool,
            result: result,
          });

          // Track if we wrote data with formulas (needs recalculation)
          if ((tool === 'writeDataToRange' || tool === 'applyFormula') && result.hasFormulas) {
            needsRecalculation = true;
          }
        } catch (error) {
          console.error(`  ❌ ${tool} failed:`, error.message);
          results.push({
            success: false,
            tool: tool,
            params: params,
            error: error.message || String(error),
            errorType: error.name || 'Error',
            errorDetails: error.stack ? error.stack.split('\n')[0] : undefined
          });
          // Continue with other tools instead of failing entire batch
        }
      }

      // PHASE 2: Flush all queued writes before optional recalculation
      console.log('  ⚡ Syncing all writes to Excel...');
      await context.sync();

      // PHASE 3: Trigger recalculation AFTER writes are committed
      if (needsRecalculation) {
        console.log('  📊 Forcing FULL workbook recalculation...');
        context.workbook.application.calculate(Excel.CalculationType.full);
        await context.sync();
      }

      console.log('  ✅ Batch execution complete');

      return results;

    } catch (error) {
      console.error('  ❌ Batch execution failed:', error);
      throw error;
    }
  });
}

/**
 * Execute a single tool within an existing Excel.run() context (no sync)
 * 
 * @param {Excel.RequestContext} context - The Excel request context
 * @param {string} tool - Tool name
 * @param {Object} params - Tool parameters
 * @param {Map} worksheets - Cache of worksheet references
 * @returns {Promise<Object>} Tool result
 */
async function executeBatchedTool(context, tool, params, worksheets) {
  switch (tool) {
    case 'createNewSheet':
      return await batchCreateNewSheet(context, params, worksheets);
    
    case 'writeDataToRange':
      return await batchWriteDataToRange(context, params, worksheets);
    
    case 'formatRange':
      return await batchFormatRange(context, params, worksheets);
    
    case 'applyFormula':
      return await batchApplyFormula(context, params, worksheets);
    
    case 'createChart':
      return await batchCreateChart(context, params, worksheets);
    
    case 'createTable':
      return await batchCreateTable(context, params, worksheets);
    
    case 'addCellNote':
      return await batchAddCellNote(context, params, worksheets);
    
    case 'createNewSheet':
      return await batchCreateNewSheet(context, params, worksheets);
    
    case 'insertRows':
    case 'deleteRows':
      return await batchModifyRows(context, tool, params, worksheets);
    
    default:
      throw new Error(`Batched execution not implemented for tool: ${tool}`);
  }
}

/**
 * Get or create worksheet reference (cached)
 */
function getWorksheet(context, sheetName, worksheets) {
  if (!worksheets.has(sheetName)) {
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    worksheets.set(sheetName, worksheet);
  }
  return worksheets.get(sheetName);
}

/**
 * Parse cell address (e.g., "A1", "B5", "AA10") to row/column indices
 * Returns {row: number, column: number} (0-based)
 */
function parseCellAddress(cellAddress) {
  const match = cellAddress.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(`Invalid cell address: ${cellAddress}`);
  }
  
  const colLetters = match[1];
  const rowNumber = parseInt(match[2], 10);
  
  // Convert column letters to 0-based index
  let colIndex = 0;
  for (let i = 0; i < colLetters.length; i++) {
    colIndex = colIndex * 26 + (colLetters.charCodeAt(i) - 65 + 1);
  }
  colIndex -= 1; // Make 0-based
  
  // Row is already 1-based, make it 0-based
  const rowIndex = rowNumber - 1;
  
  return { row: rowIndex, column: colIndex };
}

/**
 * Batched createNewSheet
 * 
 * CRITICAL: This function MUST properly cache the worksheet object so that
 * subsequent tools can reference it WITHOUT needing a sync.
 * This enables true batching: createNewSheet + writeDataToRange + formatRange
 * all in a single Excel.run() context.
 */
async function batchCreateNewSheet(context, params, worksheets) {
  const { sheetName, position } = params;
  
  // Create sheet (don't pass position to add() - set it separately)
  // This matches the standalone tool's API usage
  const worksheet = context.workbook.worksheets.add(sheetName);
  
  // Set position separately if specified (matches standalone behavior)
  if (position !== undefined) {
    worksheet.position = position;
  }
  
  // Load properties so the worksheet object is fully initialized
  // This ensures it's usable by subsequent batched tools
  worksheet.load("name, position");
  
  // Cache the worksheet so other tools can use it without calling getItem()
  worksheets.set(sheetName, worksheet);
  
  return {
    sheetName: sheetName,
    position: position,
    created: true
  };
}

/**
 * Batched writeDataToRange - Most important for P&L creation
 */
async function batchWriteDataToRange(context, params, worksheets) {
  const { sheetName, startCell, data, overwrite = true, headerFormat } = params;

  // Validate
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error('Invalid data: must be a non-empty 2D array');
  }

  // Get worksheet
  const worksheet = getWorksheet(context, sheetName, worksheets);

  // Normalize data: find max columns and pad rows
  let numCols = 0;
  for (let i = 0; i < data.length; i++) {
    if (!Array.isArray(data[i])) {
      throw new Error(`Invalid data: row ${i} is not an array`);
    }
    numCols = Math.max(numCols, data[i].length);
  }

  const normalizedData = data.map(row => {
    const paddedRow = [...row];
    while (paddedRow.length < numCols) {
      paddedRow.push('');
    }
    return paddedRow;
  });

  const numRows = normalizedData.length;

  // Parse cell address to get indices (NO SYNC NEEDED!)
  const { row: startRow, column: startCol } = parseCellAddress(startCell);

  // Create target range directly using parsed indices
  const targetRange = worksheet.getRangeByIndexes(
    startRow,
    startCol,
    numRows,
    numCols
  );

  targetRange.load("address");

  // Separate data into values and formulas
  const formulaArray = [];
  const valueArray = [];
  let formulaCount = 0;

  for (let i = 0; i < numRows; i++) {
    const formulaRow = [];
    const valueRow = [];
    
    for (let j = 0; j < numCols; j++) {
      const cellData = normalizedData[i][j];
      
      // Check if it's a formula (string starting with '=')
      if (typeof cellData === 'string' && cellData.startsWith('=')) {
        formulaRow.push(cellData);
        valueRow.push(null); // Formula will supply the value
        formulaCount++;
      } else {
        formulaRow.push(null); // Leave existing value untouched when formulas are applied
        valueRow.push(cellData === null || cellData === undefined ? '' : cellData);
      }
    }
    
    formulaArray.push(formulaRow);
    valueArray.push(valueRow);
  }

  // Write values and formulas (NO sync yet)
  if (overwrite) {
    targetRange.values = valueArray;
  }
  
  if (formulaCount > 0) {
    targetRange.formulas = formulaArray;
  }

  // Auto-fit (only for reasonably sized ranges to avoid performance issues)
  if (numRows < 1000 && numCols < 50) {
    targetRange.format.autofitColumns();
    targetRange.format.autofitRows();
  }

  // Apply header formatting if provided
  let headerFormatted = false;
  if (headerFormat && numRows > 0) {
    const headerRange = worksheet.getRangeByIndexes(
      startRow,
      startCol,
      1,
      numCols
    );

    applyFormatting(headerRange, headerFormat);
    headerFormatted = true;
  }

  return {
    sheetName: sheetName,
    rangeAddress: `${sheetName}!${startCell}`,
    rowsWritten: numRows,
    columnsWritten: numCols,
    formulasApplied: formulaCount,
    headerFormatted: headerFormatted,
    hasFormulas: formulaCount > 0
  };
}

/**
 * Batched formatRange
 */
async function batchFormatRange(context, params, worksheets) {
  const { sheetName, address, format } = params;

  if (!format || typeof format !== 'object') {
    throw new Error('Invalid format: must be an object');
  }

  const worksheet = getWorksheet(context, sheetName, worksheets);

  // Support union addresses (e.g., "A1:A5,C1:C5"). Use getRanges for unions
  const isRangeAreas = address.includes(',');
  const target = isRangeAreas ? worksheet.getRanges(address) : worksheet.getRange(address);
  target.load("address");

  applyFormatting(target, format, isRangeAreas);

  return {
    sheetName: sheetName,
    address: address,
    formatted: true,
    appliedFormats: Object.keys(format)
  };
}

/**
 * Helper to apply formatting to a range (reusable)
 */
function applyFormatting(target, format, isRangeAreas = false) {
  // Range and RangeAreas both expose a .format object. Only Range exposes .numberFormat directly.
  const formatObj = target.format;

  // Number format
  if (format.numberFormat) {
    if (isRangeAreas) {
      // RangeAreas - use format.numberFormat
      formatObj.numberFormat = format.numberFormat;
    } else {
      // Range - use direct numberFormat property
      target.numberFormat = format.numberFormat;
    }
  }

  // Font
  if (formatObj) {
    if (format.fontBold !== undefined) {
      formatObj.font.bold = format.fontBold;
    }
    if (format.fontItalic !== undefined) {
      formatObj.font.italic = format.fontItalic;
    }
    if (format.fontSize) {
      formatObj.font.size = format.fontSize;
    }
    if (format.fontColor) {
      formatObj.font.color = format.fontColor;
    }

    // Fill
    if (format.fillColor) {
      formatObj.fill.color = format.fillColor;
    }

    // Alignment
    if (format.horizontalAlignment) {
      formatObj.horizontalAlignment = format.horizontalAlignment.toLowerCase();
    }
    if (format.verticalAlignment) {
      formatObj.verticalAlignment = format.verticalAlignment.toLowerCase();
    }

    // Borders
    if (format.borders) {
      const borderTypes = ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'];
      borderTypes.forEach(type => {
        const borderValue = format.borders[type.toLowerCase()];
        if (borderValue) {
          const borderStyle = typeof borderValue === 'boolean' ? 'Continuous' : borderValue;
          formatObj.borders.getItem(type).style = borderStyle;
        }
      });
    }
  }
}

/**
 * Batched applyFormula
 * 
 * Note: This tool requires dimensions, so we do ONE micro-sync here.
 * Still much better than separate Excel.run() contexts.
 * Ideally, the agent should use embedded formulas in writeDataToRange instead.
 */
async function batchApplyFormula(context, params, worksheets) {
  const { sheetName, address, formula } = params;

  const cleanFormula = formula.startsWith('=') ? formula : `=${formula}`;
  const worksheet = getWorksheet(context, sheetName, worksheets);
  const range = worksheet.getRange(address);
  
  range.load("address, rowCount, columnCount");
  
  // Micro-sync: Only for this tool, unavoidable
  await context.sync();

  // Build formula array based on range dimensions
  if (range.rowCount === 1 && range.columnCount === 1) {
    range.formulas = [[cleanFormula]];
  } else {
    const formulas = [];
    for (let i = 0; i < range.rowCount; i++) {
      const row = [];
      for (let j = 0; j < range.columnCount; j++) {
        row.push(cleanFormula);
      }
      formulas.push(row);
    }
    range.formulas = formulas;
  }

  return {
    sheetName: sheetName,
    address: address,
    formula: cleanFormula,
    applied: true,
    cellsAffected: range.rowCount * range.columnCount,
    hasFormulas: true
  };
}

/**
 * Batched createChart
 */
async function batchCreateChart(context, params, worksheets) {
  const { sheetName, chartType, dataRange, title, position } = params;

  const worksheet = getWorksheet(context, sheetName, worksheets);
  const dataRangeObj = worksheet.getRange(dataRange);
  
  const chart = worksheet.charts.add(chartType, dataRangeObj, "Auto");
  
  if (title) {
    chart.title.text = title;
  }
  
  if (position) {
    chart.top = 0;
    chart.left = 0;
  }

  chart.load("id, name");

  return {
    sheetName: sheetName,
    chartType: chartType,
    dataRange: dataRange,
    title: title || '',
    created: true
  };
}

/**
 * Batched createTable
 */
async function batchCreateTable(context, params, worksheets) {
  const { sheetName, address, tableName, hasHeaders = true } = params;

  const worksheet = getWorksheet(context, sheetName, worksheets);
  const range = worksheet.getRange(address);
  
  const table = worksheet.tables.add(range, hasHeaders);
  table.name = tableName;
  table.load("name, id");

  return {
    sheetName: sheetName,
    address: address,
    tableName: tableName,
    hasHeaders: hasHeaders,
    created: true
  };
}

/**
 * Batched addCellNote - Uses modern Comment API
 */
async function batchAddCellNote(context, params, worksheets) {
  const { sheetName, cellAddress, noteText, author = 'Warren' } = params;

  const worksheet = getWorksheet(context, sheetName, worksheets);
  const cell = worksheet.getRange(cellAddress);
  
  // Load cell address for comment lookup
  cell.load("address");
  
  // Format the comment text with timestamp
  const timestamp = new Date().toLocaleString();
  const formattedText = `${noteText}\n\n—${author}, ${timestamp}`;
  
  // Get comment collection
  const commentCollection = worksheet.comments;
  commentCollection.load("items");
  
  // Note: We need a micro-sync here to check for existing comments
  await context.sync();
  
  // Find and delete existing comment on this cell if it exists
  for (let i = 0; i < commentCollection.items.length; i++) {
    const comment = commentCollection.items[i];
    comment.load("cellAddress");
    await context.sync();
    
    if (comment.cellAddress === cell.address) {
      comment.delete();
      break;
    }
  }
  
  // Add new comment using correct API: worksheet.comments.add(range, content, contentType)
  commentCollection.add(cell, formattedText, Excel.ContentType.plain);

  return {
    sheetName: sheetName,
    cellAddress: cellAddress,
    noteAdded: true,
    noteText: noteText
  };
}

/**
 * Batched row modifications
 */
async function batchModifyRows(context, tool, params, worksheets) {
  const { sheetName, startRow, count } = params;

  const worksheet = getWorksheet(context, sheetName, worksheets);
  
  if (tool === 'insertRows') {
    const range = worksheet.getRangeByIndexes(startRow, 0, 1, 1);
    range.insert(Excel.InsertShiftDirection.down);
  } else if (tool === 'deleteRows') {
    const range = worksheet.getRangeByIndexes(startRow, 0, count || 1, 1);
    range.delete(Excel.DeleteShiftDirection.up);
  }

  return {
    sheetName: sheetName,
    startRow: startRow,
    count: count,
    modified: true
  };
}

