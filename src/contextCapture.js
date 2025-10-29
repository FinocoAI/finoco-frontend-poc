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
 *
 * Two capture modes:
 * 1. captureLightweightMap() - Ultra-lightweight workbook map (75-80% smaller)
 * 2. captureInitialContext() - Full detailed context (legacy, for backward compatibility)
 */

/**
 * Capture lightweight workbook map (RECOMMENDED for large workbooks)
 *
 * This function captures a minimal workbook structure that's 75-80% smaller than full context.
 * Agent can then use getSheetMetadata() and getRangePreview() to explore sheets as needed.
 *
 * What's included:
 * - Sheet names, sizes, and basic flags (hasFormulas, hasCharts, etc.)
 * - Table names and named range names (but not full details)
 * - Active context (active sheet, active cell)
 *
 * What's excluded (fetch on-demand with getSheetMetadata):
 * - Headers (fetch per sheet when needed)
 * - Formula ranges and details
 * - Table structures and headers
 * - Named range values
 *
 * @returns {Promise<Object>} Lightweight workbook map
 */
export async function captureLightweightMap() {
  return Excel.run(async (context) => {
    console.log("📊 Starting Lightweight Map Capture...");

    // Load workbook metadata
    const workbook = context.workbook;
    workbook.load("name, isDirty");

    // Load all worksheets
    const worksheets = workbook.worksheets;
    worksheets.load("items");

    // Load workbook-level named ranges (names only)
    const workbookNamedRanges = workbook.names;
    workbookNamedRanges.load("items");

    // Get active sheet and cell
    const activeSheet = workbook.worksheets.getActiveWorksheet();
    activeSheet.load("name");

    const activeCell = workbook.getActiveCell();
    activeCell.load("address, values");

    await context.sync();

    // Build lightweight map
    const lightweightMap = {
      workbook: {
        name: workbook.name,
        isDirty: workbook.isDirty,
        sheetCount: worksheets.items.length,
        hasMultipleSheets: worksheets.items.length > 1,
      },
      sheets: [],
      workbookNamedRanges: [],
      activeSheet: activeSheet.name,
      activeCell: activeCell.address,
      activeCellValue: activeCell.values[0][0],
    };

    // Process each sheet with optimized batching
    console.log(`📊 Processing ${worksheets.items.length} sheets with optimized batching...`);
    for (let i = 0; i < worksheets.items.length; i++) {
      const sheet = worksheets.items[i];
      console.log(`  [${i + 1}/${worksheets.items.length}] Processing sheet: ${sheet.name || 'Unnamed'}`);
      const sheetData = await captureLightweightSheetMetadataOptimized(context, sheet, activeSheet.name);
      if (sheetData) {
        lightweightMap.sheets.push(sheetData);
        console.log(`    ✓ Sheet captured (${sheetData.rowCount}x${sheetData.columnCount}, ${sheetData.formulaCount} formulas)`);
      } else {
        console.log(`    ⚠️ Sheet skipped due to error`);
      }
    }

    // Process workbook-level named ranges (names and addresses only, no values)
    // FIXED: Load all properties in batch BEFORE syncing (not inside loop)
    for (let i = 0; i < workbookNamedRanges.items.length; i++) {
      const namedRange = workbookNamedRanges.items[i];
      namedRange.load("name, formula");
    }
    
    // Single sync after loading all properties
    if (workbookNamedRanges.items.length > 0) {
      await context.sync();
      console.log(`  📝 Processing ${workbookNamedRanges.items.length} workbook-level named ranges...`);
      
      for (let i = 0; i < workbookNamedRanges.items.length; i++) {
        const namedRange = workbookNamedRanges.items[i];
        if (isValidNamedRange(namedRange.name, namedRange.formula)) {
          lightweightMap.workbookNamedRanges.push({
            name: namedRange.name,
            address: namedRange.formula.replace("=", ""),
            sheet: extractSheetFromFormula(namedRange.formula),
          });
        }
      }
    }

    const mapSize = JSON.stringify(lightweightMap).length;
    const successfulSheets = lightweightMap.sheets.filter(s => !s.error).length;
    const failedSheets = lightweightMap.sheets.filter(s => s.error).length;
    
    console.log("✅ Lightweight Map Captured Successfully!");
    console.log(`   📦 Map size: ${mapSize.toLocaleString()} characters`);
    console.log(`   ✓ Sheets captured: ${successfulSheets}/${worksheets.items.length}`);
    if (failedSheets > 0) {
      console.log(`   ⚠️ Sheets skipped: ${failedSheets} (errors encountered)`);
    }
    console.log(`   📊 Named ranges: ${lightweightMap.workbookNamedRanges.length}`);
    
    return lightweightMap;
  });
}

/**
 * OPTIMIZED: Capture lightweight metadata for a single sheet with aggressive batching
 * Reduces sync() calls from ~8-10 to ~3-4 per sheet by batching all loads
 *
 * @param {Excel.RequestContext} context - Excel request context
 * @param {Excel.Worksheet} sheet - Worksheet to capture
 * @param {string} activeSheetName - Name of currently active sheet
 * @returns {Promise<Object|null>} Lightweight sheet metadata or null if error
 */
async function captureLightweightSheetMetadataOptimized(context, sheet, activeSheetName) {
  try {
    // MEGA-BATCH 1: Load ALL sheet properties at once before any sync
    sheet.load("name, position, visibility");
    const tables = sheet.tables;
    tables.load("items");
    const namedRanges = sheet.names;
    namedRanges.load("items");
    const charts = sheet.charts;
    charts.load("count");
    const pivotTables = sheet.pivotTables;
    pivotTables.load("count");
    
    // First sync to get sheet basics
    await context.sync();

    const sheetData = {
      name: sheet.name,
      index: sheet.position,
      isActive: sheet.name === activeSheetName,
      isEmpty: true,
      rowCount: 0,
      columnCount: 0,
      usedRange: "",
      hasFormulas: false,
      formulaCount: 0,
      hasCharts: charts.count > 0,
      chartCount: charts.count,
      hasPivotTables: pivotTables.count > 0,
      pivotTableCount: pivotTables.count,
      hasTables: tables.items.length > 0,
      tableCount: tables.items.length,
      tableNames: [],
      hasNamedRanges: false,
      namedRangeCount: 0,
      namedRangeNames: [],
    };

    // MEGA-BATCH 2: Load used range + all table names + all named range details
    let usedRange;
    try {
      usedRange = sheet.getUsedRange();
      usedRange.load("address, rowCount, columnCount");
    } catch (error) {
      // Sheet is empty, that's fine
    }

    // Load all table names in batch
    for (let i = 0; i < tables.items.length; i++) {
      tables.items[i].load("name");
    }

    // Load all named range properties in batch
    for (let i = 0; i < namedRanges.items.length; i++) {
      namedRanges.items[i].load("name, formula");
    }

    // Single sync for all of batch 2
    await context.sync();

    // Process used range info (if exists)
    if (usedRange) {
      sheetData.usedRange = usedRange.address.split('!')[1] || usedRange.address;
      sheetData.rowCount = usedRange.rowCount;
      sheetData.columnCount = usedRange.columnCount;
      sheetData.isEmpty = false;
    }

    // Process tables (already loaded)
    for (let i = 0; i < tables.items.length; i++) {
      sheetData.tableNames.push(tables.items[i].name);
    }

    // Process named ranges (already loaded)
    for (let i = 0; i < namedRanges.items.length; i++) {
      const namedRange = namedRanges.items[i];
      if (isValidNamedRange(namedRange.name, namedRange.formula)) {
        sheetData.namedRangeNames.push(namedRange.name);
      }
    }
    sheetData.hasNamedRanges = sheetData.namedRangeNames.length > 0;
    sheetData.namedRangeCount = sheetData.namedRangeNames.length;

    // MEGA-BATCH 3: Formula detection (only if sheet has data)
    if (usedRange) {
      try {
        const formulaCells = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
        formulaCells.load("areas");
        await context.sync();

        const areas = formulaCells.areas;
        areas.load("items");
        await context.sync();

        // Load all area properties in one batch
        for (let i = 0; i < areas.items.length; i++) {
          areas.items[i].load("rowCount, columnCount");
        }
        
        await context.sync();
        
        let totalCount = 0;
        for (let i = 0; i < areas.items.length; i++) {
          totalCount += areas.items[i].rowCount * areas.items[i].columnCount;
        }

        sheetData.hasFormulas = true;
        sheetData.formulaCount = totalCount;
      } catch (error) {
        // No formulas or sheet is protected
        sheetData.hasFormulas = false;
        sheetData.formulaCount = 0;
      }
    }

    return sheetData;

  } catch (error) {
    console.error(`    ❌ Error processing sheet ${sheet ? sheet.name : 'unknown'}:`, error.message || error);
    return {
      name: sheet?.name || 'Unknown',
      index: sheet?.position || -1,
      isActive: false,
      isEmpty: true,
      error: error.message || 'Failed to process sheet'
    };
  }
}

/**
 * LEGACY: Capture lightweight metadata for a single sheet
 * Use captureLightweightSheetMetadataOptimized for better performance
 *
 * @param {Excel.RequestContext} context - Excel request context
 * @param {Excel.Worksheet} sheet - Worksheet to capture
 * @param {string} activeSheetName - Name of currently active sheet
 * @returns {Promise<Object|null>} Lightweight sheet metadata or null if error
 */
async function captureLightweightSheetMetadata(context, sheet, activeSheetName) {
  try {
    sheet.load("name, position, visibility");
    await context.sync();

    const sheetData = {
      name: sheet.name,
      index: sheet.position,
      isActive: sheet.name === activeSheetName,
      isEmpty: true,
      rowCount: 0,
      columnCount: 0,
      usedRange: "",
      hasFormulas: false,
      formulaCount: 0,
      hasCharts: false,
      chartCount: 0,
      hasPivotTables: false,
      pivotTableCount: 0,
      hasTables: false,
      tableCount: 0,
      tableNames: [],
      hasNamedRanges: false,
      namedRangeCount: 0,
      namedRangeNames: [],
    };

    // Get used range (size only, no data)
    try {
      const usedRange = sheet.getUsedRange();
      usedRange.load("address, rowCount, columnCount");
      await context.sync();

      sheetData.usedRange = usedRange.address.split('!')[1] || usedRange.address;
      sheetData.rowCount = usedRange.rowCount;
      sheetData.columnCount = usedRange.columnCount;
      sheetData.isEmpty = false;

      // Check for formulas (count only)
      try {
        const formulaCells = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
        formulaCells.load("areas");
        await context.sync();

        const areas = formulaCells.areas;
        areas.load("items");
        await context.sync();

        // FIXED: Load all area properties in batch, then sync once
        for (let i = 0; i < areas.items.length; i++) {
          const area = areas.items[i];
          area.load("rowCount, columnCount");
        }
        
        await context.sync();
        
        let totalCount = 0;
        for (let i = 0; i < areas.items.length; i++) {
          const area = areas.items[i];
          totalCount += area.rowCount * area.columnCount;
        }

        sheetData.hasFormulas = true;
        sheetData.formulaCount = totalCount;
      } catch (error) {
        // No formulas found or sheet is protected - this is expected
        sheetData.hasFormulas = false;
        sheetData.formulaCount = 0;
      }
    } catch (error) {
      console.log(`    ℹ️ Sheet "${sheet.name}" is empty`);
    }

    // Get table info (names and counts only)
    const tables = sheet.tables;
    tables.load("items");
    await context.sync();

    sheetData.hasTables = tables.items.length > 0;
    sheetData.tableCount = tables.items.length;

    // FIXED: Load all table names in batch, then sync once
    for (let i = 0; i < tables.items.length; i++) {
      const table = tables.items[i];
      table.load("name");
    }
    
    if (tables.items.length > 0) {
      await context.sync();
      for (let i = 0; i < tables.items.length; i++) {
        sheetData.tableNames.push(tables.items[i].name);
      }
    }

    // Get named range info (names only)
    const namedRanges = sheet.names;
    namedRanges.load("items");
    await context.sync();

    // FIXED: Load all named range properties in batch, then sync once
    for (let i = 0; i < namedRanges.items.length; i++) {
      const namedRange = namedRanges.items[i];
      namedRange.load("name, formula");
    }
    
    if (namedRanges.items.length > 0) {
      await context.sync();
      for (let i = 0; i < namedRanges.items.length; i++) {
        const namedRange = namedRanges.items[i];
        if (isValidNamedRange(namedRange.name, namedRange.formula)) {
          sheetData.namedRangeNames.push(namedRange.name);
        }
      }
    }

    sheetData.hasNamedRanges = sheetData.namedRangeNames.length > 0;
    sheetData.namedRangeCount = sheetData.namedRangeNames.length;

    // Get chart info (count only)
    const charts = sheet.charts;
    charts.load("count");
    await context.sync();
    sheetData.hasCharts = charts.count > 0;
    sheetData.chartCount = charts.count;

    // Get pivot table info (count only)
    const pivotTables = sheet.pivotTables;
    pivotTables.load("count");
    await context.sync();
    sheetData.hasPivotTables = pivotTables.count > 0;
    sheetData.pivotTableCount = pivotTables.count;

    return sheetData;

  } catch (error) {
    console.error(`    ❌ Error processing sheet ${sheet ? sheet.name : 'unknown'}:`, error.message || error);
    // Return a minimal sheet object to continue processing other sheets
    return {
      name: sheet?.name || 'Unknown',
      index: sheet?.position || -1,
      isActive: false,
      isEmpty: true,
      error: error.message || 'Failed to process sheet'
    };
  }
}

/**
 * Main function to capture complete initial context (LEGACY - for backward compatibility)
 *
 * NOTE: For large workbooks (50+ sheets), use captureLightweightMap() instead.
 * This function captures more details but results in larger payloads.
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

    // Process each sheet with optimized batching
    console.log(`📊 Processing ${worksheets.items.length} sheets with optimized batching...`);
    for (let i = 0; i < worksheets.items.length; i++) {
      const sheet = worksheets.items[i];
      const sheetData = await captureSheetMetadataOptimized(context, sheet, activeSheet.name);
      if (sheetData) {
        initialContext.sheets.push(sheetData);
      }
    }

    // Process workbook-level named ranges (with validation)
    for (let i = 0; i < workbookNamedRanges.items.length; i++) {
      const namedRange = workbookNamedRanges.items[i];
      namedRange.load("name, formula");
      await context.sync();

      // Validate and filter out problematic named ranges
      if (isValidNamedRange(namedRange.name, namedRange.formula)) {
        initialContext.workbookNamedRanges.push({
          name: namedRange.name,
          address: namedRange.formula.replace("=", ""), // Remove leading "="
          sheet: extractSheetFromFormula(namedRange.formula),
        });
      } else {
        console.log(`    ⚠️ Skipping invalid named range: ${namedRange.name} (${namedRange.formula})`);
      }
    }

    console.log("✅ Initial Context Captured:", initialContext);
    return initialContext;
  });
}

/**
 * Optimized version: Captures metadata for a single sheet with batched operations
 * This version minimizes sync() calls by batching property loads
 *
 * @param {Excel.RequestContext} context - Excel request context
 * @param {Excel.Worksheet} sheet - Worksheet to capture
 * @param {string} activeSheetName - Name of currently active sheet
 * @returns {Promise<Object|null>} Sheet metadata object or null if error
 */
async function captureSheetMetadataOptimized(context, sheet, activeSheetName) {
  try {
    // BATCH 1: Load all basic properties at once
    sheet.load("name, position, visibility");
    const tables = sheet.tables;
    tables.load("items");
    const namedRanges = sheet.names;
    namedRanges.load("items");
    const charts = sheet.charts;
    charts.load("count");
    const pivotTables = sheet.pivotTables;
    pivotTables.load("count");
    
    // Single sync for all basic properties
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
      formulaRanges: [],
      formulaCount: 0,
      hasCharts: charts.count > 0,
      hasPivotTables: pivotTables.count > 0,
    };

    // BATCH 2: Load used range and table/named range details
    let usedRange;
    let firstRow;
    try {
      usedRange = sheet.getUsedRange();
      usedRange.load("address, rowCount, columnCount");
      
      // Load table names in batch
      for (let i = 0; i < tables.items.length; i++) {
        const table = tables.items[i];
        table.load("name");
        const tableRange = table.getRange();
        tableRange.load("address");
      }
      
      // Load named range details in batch
      for (let i = 0; i < namedRanges.items.length; i++) {
        const namedRange = namedRanges.items[i];
        namedRange.load("name, formula");
      }
      
      await context.sync();

      sheetData.usedRange = usedRange.address;
      sheetData.rowCount = usedRange.rowCount;
      sheetData.columnCount = usedRange.columnCount;

      // Get first row if data exists
      if (usedRange.rowCount > 0) {
        firstRow = usedRange.getRow(0);
        firstRow.load("values");
        await context.sync();
        sheetData.potentialHeaders = firstRow.values[0].map(val => String(val));
      }
      
    } catch (error) {
      console.log(`    ℹ️ No used range found for ${sheet.name} (likely empty)`);
    }

    // Process tables (already loaded)
    for (let i = 0; i < tables.items.length; i++) {
      const table = tables.items[i];
      const tableRange = table.getRange();
      sheetData.tables.push({
        name: table.name,
        address: tableRange.address,
      });
    }

    // Process named ranges (already loaded, just validate)
    for (let i = 0; i < namedRanges.items.length; i++) {
      const namedRange = namedRanges.items[i];
      if (isValidNamedRange(namedRange.name, namedRange.formula)) {
        sheetData.namedRanges.push({
          name: namedRange.name,
          address: namedRange.formula.replace("=", ""),
        });
      }
    }

    // BATCH 3: Formula detection (most expensive operation)
    if (usedRange) {
      try {
        const formulaCells = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
        formulaCells.load("address, areas");
        await context.sync();

        sheetData.hasFormulas = true;

        const areas = formulaCells.areas;
        areas.load("items");
        await context.sync();

        console.log(`    📐 Found formulas, extracting ranges...`);

        // Load all area properties in one batch
        for (let i = 0; i < areas.items.length; i++) {
          const area = areas.items[i];
          area.load("address, rowCount, columnCount");
        }
        
        await context.sync();

        let totalFormulaCount = 0;
        for (let i = 0; i < areas.items.length; i++) {
          const area = areas.items[i];
          const rangeAddress = area.address.split('!')[1] || area.address;
          sheetData.formulaRanges.push(rangeAddress);
          totalFormulaCount += area.rowCount * area.columnCount;
        }

        sheetData.formulaCount = totalFormulaCount;
        console.log(`    ✓ Found ${totalFormulaCount} formulas in ${sheetData.formulaRanges.length} range(s)`);

      } catch (error) {
        sheetData.hasFormulas = false;
        console.log(`    ℹ️ No formulas found in ${sheet.name}`);
      }
    }

    console.log(`    ✓ Sheet processed: ${sheet.name} (${sheetData.rowCount}x${sheetData.columnCount})`);
    return sheetData;

  } catch (error) {
    console.error(`    ❌ Error processing sheet ${sheet.name}:`, error);
    return null;
  }
}

/**
 * Original version: Captures metadata for a single sheet
 * LEGACY - Use captureSheetMetadataOptimized for better performance
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

    // Capture sheet-level named ranges (with validation)
    const namedRanges = sheet.names;
    namedRanges.load("items");
    await context.sync();

    for (let i = 0; i < namedRanges.items.length; i++) {
      const namedRange = namedRanges.items[i];
      namedRange.load("name, formula");
      await context.sync();

      // Validate and filter out problematic named ranges
      if (isValidNamedRange(namedRange.name, namedRange.formula)) {
        sheetData.namedRanges.push({
          name: namedRange.name,
          address: namedRange.formula.replace("=", ""),
        });
      } else {
        console.log(`    ⚠️ Skipping invalid named range on ${sheet.name}: ${namedRange.name}`);
      }
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
 * Validates if a named range is valid and should be included in context
 * Filters out:
 * - Named ranges with #REF! errors
 * - External file references (URLs, absolute paths)
 * - Special/escape characters in names (\b, \c, etc.)
 * - Empty or invalid formulas
 *
 * @param {string} name - Named range name
 * @param {string} formula - Named range formula
 * @returns {boolean} True if valid, false if should be filtered out
 */
function isValidNamedRange(name, formula) {
  // Filter out names with backslash escape characters
  if (name.includes("\\")) {
    return false;
  }

  // Filter out if formula contains #REF! error
  if (formula.includes("#REF!")) {
    return false;
  }

  // Filter out external file references (http://, https://, file paths)
  if (formula.includes("http://") || formula.includes("https://") || formula.includes("file://")) {
    return false;
  }

  // Filter out SharePoint/OneDrive URLs (d.docs.live.net, etc.)
  if (formula.includes(".live.net") || formula.includes(".sharepoint.com")) {
    return false;
  }

  // Filter out absolute file paths (common patterns)
  if (formula.includes(":\\") || formula.includes("]/")) {
    return false;
  }

  // Filter out empty formulas
  if (!formula || formula.trim() === "=" || formula.trim() === "") {
    return false;
  }

  return true;
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
