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
    },
    includeStructure: {
      type: "boolean",
      required: false,
      default: true,
      description: "Include smart structure detection (header rows, row labels, table regions)"
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
    lastDataRow: "number",
    structure: "object - Smart structure detection including headerRows, rowLabels, and tableRegions"
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
 * Detect header rows in a sheet
 * Looks for rows that contain mostly text/labels followed by data columns
 * 
 * @param {Excel.Range} usedRange - The used range of the sheet
 * @param {Excel.RequestContext} context - Excel context
 * @returns {Promise<Array>} Array of header row objects
 */
async function detectHeaderRows(usedRange, context) {
  const headerRows = [];
  
  try {
    // Load all values and formulas to analyze structure
    usedRange.load("values, formulas, rowCount, columnCount");
    await context.sync();
    
    const values = usedRange.values;
    const formulas = usedRange.formulas;
    const rowCount = usedRange.rowCount;
    const colCount = usedRange.columnCount;
    
    // Analyze each row to find potential headers
    for (let rowIdx = 0; rowIdx < Math.min(rowCount, 100); rowIdx++) {
      const row = values[rowIdx];
      
      // Skip empty rows
      if (row.every(cell => cell === "" || cell === null)) {
        continue;
      }
      
      // Count text cells vs numeric cells
      let textCount = 0;
      let numericCount = 0;
      let nonEmptyCount = 0;
      
      for (let colIdx = 0; colIdx < colCount; colIdx++) {
        const cell = row[colIdx];
        if (cell !== "" && cell !== null) {
          nonEmptyCount++;
          if (typeof cell === 'string' && isNaN(parseFloat(cell))) {
            textCount++;
          } else if (typeof cell === 'number' || !isNaN(parseFloat(cell))) {
            numericCount++;
          }
        }
      }
      
      // A header row has mostly text (at least 40% of non-empty cells are text)
      // And has at least 2 non-empty cells
      const isLikelyHeader = nonEmptyCount >= 2 && 
                             (textCount / nonEmptyCount) >= 0.4;
      
      // Also check if next row has mostly data (numbers/formulas)
      let nextRowHasData = false;
      if (rowIdx < rowCount - 1) {
        const nextRow = values[rowIdx + 1];
        const nextFormulas = formulas[rowIdx + 1];
        let dataCount = 0;
        let nextNonEmpty = 0;
        
        for (let colIdx = 0; colIdx < colCount; colIdx++) {
          if (nextRow[colIdx] !== "" && nextRow[colIdx] !== null) {
            nextNonEmpty++;
            if (typeof nextRow[colIdx] === 'number' || nextFormulas[colIdx] !== "") {
              dataCount++;
            }
          }
        }
        
        nextRowHasData = nextNonEmpty >= 2 && (dataCount / nextNonEmpty) >= 0.5;
      }
      
      if (isLikelyHeader && (nextRowHasData || rowIdx === 0)) {
        // Extract ONLY non-empty header values with their column positions
        const nonEmptyValues = [];
        for (let colIdx = 0; colIdx < colCount; colIdx++) {
          const cell = row[colIdx];
          if (cell !== "" && cell !== null) {
            nonEmptyValues.push({
              col: colIdx,
              value: String(cell)
            });
          }
        }
        
        headerRows.push({
          rowIndex: rowIdx + 1, // 1-based for user display
          nonEmptyValues: nonEmptyValues,
          nonEmptyCount: nonEmptyCount,
          emptyCount: colCount - nonEmptyCount
        });
      }
    }
    
    console.log(`    ✓ Detected ${headerRows.length} header row(s)`);
    return headerRows;
    
  } catch (error) {
    console.log(`    ⚠️ Error detecting header rows:`, error.message);
    return [];
  }
}

/**
 * Detect row labels (left columns with text labels for each row)
 * 
 * @param {Excel.Range} usedRange - The used range of the sheet
 * @param {Excel.RequestContext} context - Excel context
 * @returns {Promise<Object>} Row labels structure
 */
async function detectRowLabels(usedRange, context) {
  try {
    usedRange.load("values, columnCount, rowCount");
    await context.sync();
    
    const values = usedRange.values;
    const colCount = usedRange.columnCount;
    const rowCount = usedRange.rowCount;
    
    // Check first 3 columns for row labels
    const labelColumns = [];
    
    for (let colIdx = 0; colIdx < Math.min(3, colCount); colIdx++) {
      let textRowCount = 0;
      
      for (let rowIdx = 0; rowIdx < rowCount; rowIdx++) {
        const cell = values[rowIdx][colIdx];
        if (cell !== "" && cell !== null && typeof cell === 'string') {
          textRowCount++;
        }
      }
      
      // If >30% of rows have text in this column, it's likely a label column
      if (textRowCount / rowCount > 0.3) {
        labelColumns.push(colIdx);
      }
    }
    
    // Extract key row labels (non-empty labels)
    const keyRows = [];
    if (labelColumns.length > 0) {
      const labelColIdx = labelColumns[labelColumns.length - 1]; // Use rightmost label column
      
      for (let rowIdx = 0; rowIdx < Math.min(rowCount, 200); rowIdx++) {
        const label = values[rowIdx][labelColIdx];
        if (label && String(label).trim() !== "") {
          keyRows.push({
            rowIndex: rowIdx + 1, // 1-based
            label: String(label).trim()
          });
        }
      }
    }
    
    console.log(`    ✓ Detected ${labelColumns.length} label column(s) with ${keyRows.length} labeled rows`);
    
    return {
      labelColumnIndices: labelColumns,
      keyRows: keyRows.slice(0, 30), // Limit to first 30 to save tokens
      totalLabeledRows: keyRows.length
    };
    
  } catch (error) {
    console.log(`    ⚠️ Error detecting row labels:`, error.message);
    return { labelColumnIndices: [], keyRows: [], totalLabeledRows: 0 };
  }
}

/**
 * Detect table-like regions (header row + data rows)
 * 
 * @param {Array} headerRows - Detected header rows
 * @param {Excel.Range} usedRange - The used range
 * @param {Excel.RequestContext} context - Excel context
 * @returns {Promise<Array>} Array of table region objects
 */
async function detectTableRegions(headerRows, usedRange, context) {
  const tableRegions = [];
  
  try {
    usedRange.load("values, formulas, rowCount, columnCount, address");
    await context.sync();
    
    const values = usedRange.values;
    const formulas = usedRange.formulas;
    const rowCount = usedRange.rowCount;
    const colCount = usedRange.columnCount;
    const sheetName = usedRange.address.split('!')[0];
    
    // For each header row, find the extent of the table below it
    for (const headerRow of headerRows) {
      const headerIdx = headerRow.rowIndex - 1; // Convert to 0-based
      
      // Find where the data block ends (next empty row or next header)
      let dataEndIdx = headerIdx + 1;
      let consecutiveEmptyRows = 0;
      
      for (let rowIdx = headerIdx + 1; rowIdx < rowCount && dataEndIdx < rowCount; rowIdx++) {
        const row = values[rowIdx];
        const isEmptyRow = row.every(cell => cell === "" || cell === null);
        
        if (isEmptyRow) {
          consecutiveEmptyRows++;
          if (consecutiveEmptyRows >= 2) {
            dataEndIdx = rowIdx - 1;
            break;
          }
        } else {
          consecutiveEmptyRows = 0;
          dataEndIdx = rowIdx;
        }
      }
      
      const dataRowCount = dataEndIdx - headerIdx;
      
      // Only consider it a table if it has at least 2 data rows
      if (dataRowCount >= 2) {
        // Find leftmost and rightmost non-empty columns in header using new format
        let leftCol = colCount;
        let rightCol = 0;
        
        if (headerRow.nonEmptyValues && headerRow.nonEmptyValues.length > 0) {
          leftCol = headerRow.nonEmptyValues[0].col;
          rightCol = headerRow.nonEmptyValues[headerRow.nonEmptyValues.length - 1].col;
        }
        
        // Extract non-empty header values (already compressed in new format!)
        const headers = headerRow.nonEmptyValues ? 
          headerRow.nonEmptyValues.map(h => h.value) : [];
        
        // Extract row labels from this table region
        const rowLabels = [];
        for (let rowIdx = headerIdx + 1; rowIdx <= dataEndIdx && rowIdx < rowCount; rowIdx++) {
          // Check first few columns for labels
          for (let colIdx = 0; colIdx < Math.min(3, colCount); colIdx++) {
            const cell = values[rowIdx][colIdx];
            if (cell && String(cell).trim() !== "" && typeof cell === 'string') {
              rowLabels.push(String(cell).trim());
              break;
            }
          }
        }
        
        // Convert to Excel address
        const startCell = `${String.fromCharCode(65 + leftCol)}${headerIdx + 1}`;
        const endCell = `${String.fromCharCode(65 + Math.min(rightCol, 25))}${dataEndIdx + 1}`;
        const range = `${startCell}:${endCell}`;
        
        tableRegions.push({
          range: range,
          headerRow: headerIdx + 1, // 1-based
          headers: headers.slice(0, 15), // Limit to 15 columns to save tokens
          headerCount: headers.length,
          dataRows: dataRowCount,
          rowLabels: rowLabels.slice(0, 15), // Limit to 15 rows to save tokens
          rowLabelCount: rowLabels.length
        });
      }
    }
    
    console.log(`    ✓ Detected ${tableRegions.length} table region(s)`);
    return tableRegions;
    
  } catch (error) {
    console.log(`    ⚠️ Error detecting table regions:`, error.message);
    return [];
  }
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
        includeNamedRangeDetails = true,
        includeStructure = true
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

          // Load all area properties in batch (without formulas - too expensive!)
          for (let i = 0; i < areas.items.length; i++) {
            const area = areas.items[i];
            area.load("address, rowCount, columnCount");
          }
          await context.sync();

          let totalFormulaCount = 0;
          for (let i = 0; i < areas.items.length; i++) {
            const area = areas.items[i];
            const rangeAddress = area.address.split('!')[1] || area.address;
            metadata.formulaRanges.push(rangeAddress);
            totalFormulaCount += area.rowCount * area.columnCount;

            // Get ONLY ONE sample formula from this range (much faster!)
            try {
              const firstCell = area.getCell(0, 0);
              firstCell.load("formulas");
              await context.sync();
              
              if (firstCell.formulas && firstCell.formulas.length > 0 && firstCell.formulas[0].length > 0) {
                const sampleFormula = firstCell.formulas[0][0];
                if (sampleFormula) {
                  metadata.formulaSummary[rangeAddress] = sampleFormula;
                }
              }
            } catch (e) {
              // If we can't get sample, skip it
              console.log(`    ⚠️ Couldn't get sample formula from ${rangeAddress}`);
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

      // Perform structure detection if requested
      if (includeStructure && usedRange) {
        console.log(`  🔍 Analyzing sheet structure...`);
        
        const structure = {
          headerRows: [],
          rowLabels: { labelColumnIndices: [], keyRows: [], totalLabeledRows: 0 },
          tableRegions: []
        };
        
        try {
          // Detect header rows
          structure.headerRows = await detectHeaderRows(usedRange, context);
          
          // Detect row labels
          structure.rowLabels = await detectRowLabels(usedRange, context);
          
          // Detect table regions based on header rows
          structure.tableRegions = await detectTableRegions(structure.headerRows, usedRange, context);
          
          metadata.structure = structure;
          console.log(`  ✓ Structure analysis complete`);
        } catch (error) {
          console.log(`  ⚠️ Structure detection failed:`, error.message);
          metadata.structure = structure; // Return partial results
        }
      }

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

