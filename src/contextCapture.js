/**
 * contextCapture.js
 *
 * Captures Initial Context from Excel workbook.
 * This includes workbook metadata, all sheets structure, and workbook-level named items.
 *
 * Key Features:
 * - Captures NO full data, only structure and metadata
 * - Lightweight and fast
 * - Handles errors gracefully (protected sheets, empty sheets, etc.)
 * - Uses Office.js Excel API efficiently
 */

/**
 * Main function to capture complete initial context
 *
 * @returns {Promise<Object>} Initial context object containing workbook, sheets, and active location
 */
export async function captureInitialContext() {
  return Excel.run(async (context) => {
    console.log("📊 Starting Initial Context Capture...");

    // Load workbook metadata
    const workbook = context.workbook;
    workbook.load("name, isDirty");

    // Load all worksheets
    const worksheets = workbook.worksheets;
    worksheets.load("items");

    // Load workbook-level named ranges
    const workbookNamedRanges = workbook.names;
    workbookNamedRanges.load("items");

    // Get active sheet and cell
    const activeSheet = workbook.worksheets.getActiveWorksheet();
    activeSheet.load("name");

    const activeCell = workbook.getActiveCell();
    activeCell.load("address");

    await context.sync();

    // Build initial context object
    const initialContext = {
      workbook: {
        name: workbook.name,
        path: "", // Not available in Office.js, would need platform-specific APIs
        isDirty: workbook.isDirty,
      },
      sheets: [],
      workbookNamedRanges: [],
      activeSheet: activeSheet.name,
      activeCell: activeCell.address,
    };

    // Process each sheet
    for (let i = 0; i < worksheets.items.length; i++) {
      const sheet = worksheets.items[i];
      const sheetData = await captureSheetMetadata(context, sheet, activeSheet.name);
      if (sheetData) {
        initialContext.sheets.push(sheetData);
      }
    }

    // Process workbook-level named ranges
    for (let i = 0; i < workbookNamedRanges.items.length; i++) {
      const namedRange = workbookNamedRanges.items[i];
      namedRange.load("name, formula");
      await context.sync();

      initialContext.workbookNamedRanges.push({
        name: namedRange.name,
        address: namedRange.formula.replace("=", ""), // Remove leading "="
        sheet: extractSheetFromFormula(namedRange.formula),
      });
    }

    console.log("✅ Initial Context Captured:", initialContext);
    return initialContext;
  });
}

/**
 * Captures metadata for a single sheet
 *
 * @param {Excel.RequestContext} context - Excel request context
 * @param {Excel.Worksheet} sheet - Worksheet to capture
 * @param {string} activeSheetName - Name of currently active sheet
 * @returns {Promise<Object|null>} Sheet metadata object or null if error
 */
async function captureSheetMetadata(context, sheet, activeSheetName) {
  try {
    // Load basic sheet properties
    sheet.load("name, position, visibility");
    await context.sync();

    console.log(`  📄 Processing sheet: ${sheet.name}`);

    const sheetData = {
      name: sheet.name,
      index: sheet.position,
      isActive: sheet.name === activeSheetName,
      usedRange: "",
      rowCount: 0,
      columnCount: 0,
      potentialHeaders: [],
      tables: [],
      namedRanges: [],
      hasFormulas: false,
      formulaRanges: [],  // Array of range addresses where formulas exist (compressed)
      formulaCount: 0,    // Total count of formula cells
      hasCharts: false,
      hasPivotTables: false,
    };

    // Get used range (the area containing data)
    let usedRange;
    try {
      usedRange = sheet.getUsedRange();
      usedRange.load("address, rowCount, columnCount");
      await context.sync();

      sheetData.usedRange = usedRange.address;
      sheetData.rowCount = usedRange.rowCount;
      sheetData.columnCount = usedRange.columnCount;

      // Get first row as potential headers (if data exists)
      if (usedRange.rowCount > 0) {
        const firstRow = usedRange.getRow(0);
        firstRow.load("values");
        await context.sync();
        sheetData.potentialHeaders = firstRow.values[0].map(val => String(val));
      }
    } catch (error) {
      // Sheet is likely empty, which is fine
      console.log(`    ℹ️ No used range found for ${sheet.name} (likely empty)`);
    }

    // Capture tables
    const tables = sheet.tables;
    tables.load("items");
    await context.sync();

    for (let i = 0; i < tables.items.length; i++) {
      const table = tables.items[i];
      table.load("name");
      const tableRange = table.getRange();
      tableRange.load("address");
      await context.sync();

      sheetData.tables.push({
        name: table.name,
        address: tableRange.address,
      });
    }

    // Capture sheet-level named ranges
    const namedRanges = sheet.names;
    namedRanges.load("items");
    await context.sync();

    for (let i = 0; i < namedRanges.items.length; i++) {
      const namedRange = namedRanges.items[i];
      namedRange.load("name, formula");
      await context.sync();

      sheetData.namedRanges.push({
        name: namedRange.name,
        address: namedRange.formula.replace("=", ""),
      });
    }

    // Check for formulas and extract formula ranges (compressed format)
    if (usedRange) {
      try {
        const formulaCells = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
        formulaCells.load("address, areas");
        await context.sync();

        sheetData.hasFormulas = true;

        // Extract formula ranges from each area (much more compact than individual cells)
        const areas = formulaCells.areas;
        areas.load("items");
        await context.sync();

        console.log(`    📐 Found formulas, extracting ranges...`);

        let totalFormulaCount = 0;
        for (let i = 0; i < areas.items.length; i++) {
          const area = areas.items[i];
          area.load("address, rowCount, columnCount");
          await context.sync();

          // Store the range address (e.g., "D2:D100" or "F5" for single cell)
          const rangeAddress = area.address.split('!')[1] || area.address;
          sheetData.formulaRanges.push(rangeAddress);
          
          // Count total formula cells in this area
          totalFormulaCount += area.rowCount * area.columnCount;
        }

        sheetData.formulaCount = totalFormulaCount;
        console.log(`    ✓ Found ${totalFormulaCount} formulas in ${sheetData.formulaRanges.length} range(s): ${sheetData.formulaRanges.join(', ')}`);

      } catch (error) {
        // No formulas found - this is expected behavior
        sheetData.hasFormulas = false;
        console.log(`    ℹ️ No formulas found in ${sheet.name}`);
      }
    }

    // Check for charts
    const charts = sheet.charts;
    charts.load("count");
    await context.sync();
    sheetData.hasCharts = charts.count > 0;

    // Check for pivot tables
    const pivotTables = sheet.pivotTables;
    pivotTables.load("count");
    await context.sync();
    sheetData.hasPivotTables = pivotTables.count > 0;

    console.log(`    ✓ Sheet processed: ${sheet.name} (${sheetData.rowCount}x${sheetData.columnCount})`);
    return sheetData;

  } catch (error) {
    console.error(`    ❌ Error processing sheet ${sheet.name}:`, error);
    // Return null to skip this sheet rather than failing entirely
    return null;
  }
}

/**
 * Helper function to extract sheet name from a named range formula
 *
 * @param {string} formula - Named range formula (e.g., "=Sheet1!A1:B10")
 * @returns {string} Sheet name or empty string
 */
function extractSheetFromFormula(formula) {
  const match = formula.match(/=([^!]+)!/);
  return match ? match[1].replace(/'/g, "") : "";
}

/**
 * Helper function to extract dependencies from a formula
 * Identifies cell references, range references, and cross-sheet references
 *
 * @param {string} formula - Excel formula (e.g., "=SUM(A1:A10)+Sheet2!B5")
 * @returns {Object} Dependencies object with cells, ranges, and sheets
 */
function extractFormulaDependencies(formula) {
  const dependencies = {
    cells: [],        // Individual cell references (e.g., "A1", "Sheet2!B5")
    ranges: [],       // Range references (e.g., "A1:B10", "Sheet2!C1:D5")
    sheets: [],       // Referenced sheet names (e.g., "Sheet2", "Data")
    functions: [],    // Functions used (e.g., "SUM", "VLOOKUP", "IF")
  };

  // Extract all functions used in the formula
  // Match function names: word characters followed by opening parenthesis
  const functionPattern = /\b([A-Z][A-Z0-9_\.]*)\s*\(/gi;
  let functionMatch;
  while ((functionMatch = functionPattern.exec(formula)) !== null) {
    const funcName = functionMatch[1].toUpperCase();
    if (!dependencies.functions.includes(funcName)) {
      dependencies.functions.push(funcName);
    }
  }

  // Extract cross-sheet references (e.g., Sheet2!A1 or 'Sheet Name'!A1:B5)
  // Pattern: optional quote, sheet name, optional quote, !, cell/range reference
  const crossSheetPattern = /(?:'([^']+)'|(\w+))!([A-Z]+\d+(?::[A-Z]+\d+)?)/gi;
  let crossSheetMatch;
  while ((crossSheetMatch = crossSheetPattern.exec(formula)) !== null) {
    const sheetName = crossSheetMatch[1] || crossSheetMatch[2];
    const reference = crossSheetMatch[3];
    const fullReference = `${sheetName}!${reference}`;

    // Add sheet name if not already in list
    if (!dependencies.sheets.includes(sheetName)) {
      dependencies.sheets.push(sheetName);
    }

    // Determine if it's a range or single cell
    if (reference.includes(":")) {
      if (!dependencies.ranges.includes(fullReference)) {
        dependencies.ranges.push(fullReference);
      }
    } else {
      if (!dependencies.cells.includes(fullReference)) {
        dependencies.cells.push(fullReference);
      }
    }
  }

  // Extract same-sheet references (e.g., A1, B5:C10)
  // Remove cross-sheet references first to avoid double-counting
  const formulaWithoutCrossSheet = formula.replace(crossSheetPattern, "");

  // Pattern for cell ranges (e.g., A1:B10)
  const rangePattern = /\b([A-Z]+\d+):([A-Z]+\d+)\b/gi;
  let rangeMatch;
  while ((rangeMatch = rangePattern.exec(formulaWithoutCrossSheet)) !== null) {
    const range = rangeMatch[0];
    if (!dependencies.ranges.includes(range)) {
      dependencies.ranges.push(range);
    }
  }

  // Pattern for individual cells (e.g., A1, B5)
  // Match only if not part of a range (already extracted above)
  const cellPattern = /\b([A-Z]+\d+)\b/g;
  let cellMatch;
  const formulaWithoutRanges = formulaWithoutCrossSheet.replace(rangePattern, "");
  while ((cellMatch = cellPattern.exec(formulaWithoutRanges)) !== null) {
    const cell = cellMatch[0];
    if (!dependencies.cells.includes(cell)) {
      dependencies.cells.push(cell);
    }
  }

  return dependencies;
}
