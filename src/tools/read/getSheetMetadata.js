/**
 * getSheetMetadata.js
 *
 * Tool: Fetch detailed metadata for a specific sheet
 * Executor: Frontend (requires Office.js)
 *
 * Use Cases:
 * - Progressive data loading: fetch sheet details when needed
 * - Understand sheet structure before loading full data
 * - Discover headers, tables, formulas without loading all cells
 * - Explore sheet capabilities (charts, pivot tables, etc.)
 */

/**
 * Tool definition for LLM/Backend
 */
export const toolDefinition = {
  name: "getSheetMetadata",
  description: "Fetch detailed metadata and structure for a specific sheet. Returns headers, table info, formula summary, and structural details without loading all cell data. Use this to explore a sheet before fetching full data.",
  executor: "frontend",
  requiresContinuation: true,  // READ tool - results must be sent back to agent
  parameters: {
    sheetName: {
      type: "string",
      required: true,
      description: "Name of the sheet to fetch metadata for"
    },
    includeHeaders: {
      type: "boolean",
      required: false,
      default: true,
      description: "Include first row as potential headers"
    },
    includeFormulaRanges: {
      type: "boolean",
      required: false,
      default: true,
      description: "Include information about where formulas exist"
    },
    includeTableDetails: {
      type: "boolean",
      required: false,
      default: true,
      description: "Include detailed information about Excel Tables"
    },
    includeNamedRangeDetails: {
      type: "boolean",
      required: false,
      default: true,
      description: "Include detailed information about named ranges"
    }
  },
  returns: {
    sheetName: "string",
    usedRange: "string",
    rowCount: "number",
    columnCount: "number",
    isEmpty: "boolean",
    headers: "array",
    hasFormulas: "boolean",
    formulaCount: "number",
    formulaRanges: "array",
    formulaSummary: "object",
    tables: "array",
    namedRanges: "array",
    charts: "array",
    pivotTables: "array",
    firstDataRow: "number",
    lastDataRow: "number"
  }
};

/**
 * Helper function to validate if a named range is valid
 */
function isValidNamedRange(name, formula) {
  if (name.includes("\\")) return false;
  if (formula.includes("#REF!")) return false;
  if (formula.includes("http://") || formula.includes("https://") || formula.includes("file://")) return false;
  if (formula.includes(".live.net") || formula.includes(".sharepoint.com")) return false;
  if (formula.includes(":\\") || formula.includes("]/")) return false;
  if (!formula || formula.trim() === "=" || formula.trim() === "") return false;
  return true;
}

/**
 * Execute the tool
 *
 * @param {Object} params - Tool parameters
 * @param {string} params.sheetName - Sheet name
 * @param {boolean} [params.includeHeaders=true] - Include headers
 * @param {boolean} [params.includeFormulaRanges=true] - Include formula ranges
 * @param {boolean} [params.includeTableDetails=true] - Include table details
 * @param {boolean} [params.includeNamedRangeDetails=true] - Include named range details
 * @returns {Promise<Object>} Sheet metadata object
 */
export async function execute(params) {
  console.log(`🔧 Executing getSheetMetadata:`, params);

  return Excel.run(async (context) => {
    try {
      const {
        sheetName,
        includeHeaders = true,
        includeFormulaRanges = true,
        includeTableDetails = true,
        includeNamedRangeDetails = true
      } = params;

      // Validate parameters
      if (!sheetName || typeof sheetName !== 'string') {
        throw new Error('Invalid sheetName: must be a non-empty string');
      }

      // Get the worksheet
      const sheet = context.workbook.worksheets.getItem(sheetName);
      sheet.load("name, position, visibility");

      // Initialize metadata object
      const metadata = {
        sheetName: sheetName,
        usedRange: "",
        rowCount: 0,
        columnCount: 0,
        isEmpty: true,
        headers: [],
        hasFormulas: false,
        formulaCount: 0,
        formulaRanges: [],
        formulaSummary: {},
        tables: [],
        namedRanges: [],
        charts: [],
        pivotTables: [],
        firstDataRow: 0,
        lastDataRow: 0
      };

      await context.sync();

      // Get used range
      let usedRange;
      try {
        usedRange = sheet.getUsedRange();
        usedRange.load("address, rowCount, columnCount");
        await context.sync();

        metadata.usedRange = usedRange.address.split('!')[1] || usedRange.address;
        metadata.rowCount = usedRange.rowCount;
        metadata.columnCount = usedRange.columnCount;
        metadata.isEmpty = false;
        metadata.firstDataRow = includeHeaders && usedRange.rowCount > 1 ? 2 : 1;
        metadata.lastDataRow = usedRange.rowCount;

        console.log(`  📊 Sheet "${sheetName}" has ${metadata.rowCount}x${metadata.columnCount} used range`);

        // Get headers if requested
        if (includeHeaders && usedRange.rowCount > 0) {
          const firstRow = usedRange.getRow(0);
          firstRow.load("values");
          await context.sync();
          metadata.headers = firstRow.values[0].map(val => String(val));
          console.log(`  ✓ Captured ${metadata.headers.length} headers`);
        }

      } catch (error) {
        // Sheet is empty
        console.log(`  ℹ️ Sheet "${sheetName}" is empty`);
        metadata.isEmpty = true;
      }

      // Get table details if requested
      if (includeTableDetails) {
        const tables = sheet.tables;
        tables.load("items");
        await context.sync();

        for (let i = 0; i < tables.items.length; i++) {
          const table = tables.items[i];
          table.load("name");
          const tableRange = table.getRange();
          tableRange.load("address, rowCount, columnCount");
          const headerRow = table.getHeaderRowRange();
          headerRow.load("values");
          await context.sync();

          metadata.tables.push({
            name: table.name,
            address: tableRange.address.split('!')[1] || tableRange.address,
            headers: headerRow.values[0].map(val => String(val)),
            rowCount: tableRange.rowCount - 1,  // Exclude header
            columnCount: tableRange.columnCount,
            hasFilters: false  // Can be enhanced later
          });
        }
        console.log(`  ✓ Found ${metadata.tables.length} table(s)`);
      }

      // Get named ranges if requested
      if (includeNamedRangeDetails) {
        const namedRanges = sheet.names;
        namedRanges.load("items");
        await context.sync();

        for (let i = 0; i < namedRanges.items.length; i++) {
          const namedRange = namedRanges.items[i];
          namedRange.load("name, formula");
          await context.sync();

          if (isValidNamedRange(namedRange.name, namedRange.formula)) {
            // Try to get the value
            let value = null;
            let rangeObj = null;
            try {
              rangeObj = namedRange.getRange();
              rangeObj.load("values, address");
              await context.sync();
              
              // Get single value or indicate it's a range
              if (rangeObj.values.length === 1 && rangeObj.values[0].length === 1) {
                value = rangeObj.values[0][0];
              } else {
                value = `[Range: ${rangeObj.values.length}x${rangeObj.values[0].length}]`;
              }
            } catch (e) {
              // Named range might be invalid or external
              value = null;
            }

            metadata.namedRanges.push({
              name: namedRange.name,
              address: namedRange.formula.replace("=", ""),
              value: value
            });
          }
        }
        console.log(`  ✓ Found ${metadata.namedRanges.length} named range(s)`);
      }

      // Get formula information if requested
      if (includeFormulaRanges && usedRange) {
        try {
          const formulaCells = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
          formulaCells.load("address, areas");
          await context.sync();

          metadata.hasFormulas = true;

          // Get formula ranges
          const areas = formulaCells.areas;
          areas.load("items");
          await context.sync();

          let totalFormulaCount = 0;
          for (let i = 0; i < areas.items.length; i++) {
            const area = areas.items[i];
            area.load("address, rowCount, columnCount, formulas");
            await context.sync();

            const rangeAddress = area.address.split('!')[1] || area.address;
            metadata.formulaRanges.push(rangeAddress);
            totalFormulaCount += area.rowCount * area.columnCount;

            // Get sample formula from this range
            if (area.formulas && area.formulas.length > 0 && area.formulas[0].length > 0) {
              const sampleFormula = area.formulas[0][0];
              if (sampleFormula) {
                metadata.formulaSummary[rangeAddress] = sampleFormula;
              }
            }
          }

          metadata.formulaCount = totalFormulaCount;
          console.log(`  ✓ Found ${totalFormulaCount} formula(s) in ${metadata.formulaRanges.length} range(s)`);
        } catch (error) {
          // No formulas found
          metadata.hasFormulas = false;
          console.log(`  ℹ️ No formulas found`);
        }
      }

      // Get chart information
      const charts = sheet.charts;
      charts.load("items");
      await context.sync();

      for (let i = 0; i < charts.items.length; i++) {
        const chart = charts.items[i];
        chart.load("name, chartType");
        const chartRange = chart.getDataRange();
        chartRange.load("address");
        await context.sync();

        metadata.charts.push({
          name: chart.name,
          type: chart.chartType,
          dataRange: chartRange.address.split('!')[1] || chartRange.address
        });
      }
      console.log(`  ✓ Found ${metadata.charts.length} chart(s)`);

      // Get pivot table information
      const pivotTables = sheet.pivotTables;
      pivotTables.load("items");
      await context.sync();

      for (let i = 0; i < pivotTables.items.length; i++) {
        const pivot = pivotTables.items[i];
        pivot.load("name");
        await context.sync();

        metadata.pivotTables.push({
          name: pivot.name
        });
      }
      console.log(`  ✓ Found ${metadata.pivotTables.length} pivot table(s)`);

      console.log(`  ✅ Sheet metadata retrieved for "${sheetName}"`);

      return metadata;

    } catch (error) {
      console.error(`  ❌ Error in getSheetMetadata:`, error);
      throw {
        tool: "getSheetMetadata",
        error: error.message,
        params: params
      };
    }
  });
}

