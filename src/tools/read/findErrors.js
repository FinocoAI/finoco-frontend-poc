/**
 * findErrors.js
 *
 * Tool: Find all errors in worksheet or workbook
 * Executor: Frontend (requires Office.js)
 * 
 * Detects: #REF!, #VALUE!, #N/A, #DIV/0!, #NUM!, #NAME?, #NULL!
 * Shows error propagation and affected cells
 */

export const toolDefinition = {
  name: "findErrors",
  description: "Find all formula errors in a sheet or entire workbook. Detects #REF!, #VALUE!, #N/A, #DIV/0!, #NUM!, #NAME?, #NULL! and shows error propagation.",
  executor: "frontend",
  parameters: {
    scope: {
      type: "string",
      required: false,
      description: "Search scope: 'sheet' (current sheet) or 'workbook' (all sheets). Default: 'sheet'",
      enum: ["sheet", "workbook"]
    },
    sheetName: {
      type: "string",
      required: false,
      description: "Specific sheet name to search. If not provided, uses active sheet (when scope='sheet')"
    },
    includeHidden: {
      type: "boolean",
      required: false,
      description: "Include hidden rows/columns in search. Default: false"
    },
    tracePropagation: {
      type: "boolean",
      required: false,
      description: "Trace which cells are affected by each error. Default: true"
    }
  },
  returns: {
    success: "boolean",
    errorCount: "number",
    errors: "array of error objects with location, type, formula, and affected cells",
    summary: "object with error type breakdown"
  }
};

export async function execute(params) {
  console.log('🔍 Executing findErrors:', params);

  return Excel.run(async (context) => {
    const { 
      scope = 'sheet', 
      sheetName = null,
      includeHidden = false,
      tracePropagation = true 
    } = params;

    const errors = [];
    const errorTypes = {
      '#REF!': 0,
      '#VALUE!': 0,
      '#N/A': 0,
      '#DIV/0!': 0,
      '#NUM!': 0,
      '#NAME?': 0,
      '#NULL!': 0
    };

    try {
      // Determine which sheets to scan
      let sheetsToScan = [];
      
      if (scope === 'workbook') {
        const allSheets = context.workbook.worksheets;
        allSheets.load('items/name, items/visibility');
        await context.sync();
        
        sheetsToScan = allSheets.items
          .filter(sheet => includeHidden || sheet.visibility === Excel.SheetVisibility.visible)
          .map(sheet => sheet.name);
        
        console.log(`  📊 Scanning ${sheetsToScan.length} sheets in workbook`);
      } else {
        // Single sheet
        if (sheetName) {
          sheetsToScan = [sheetName];
        } else {
          const activeSheet = context.workbook.worksheets.getActiveWorksheet();
          activeSheet.load('name');
          await context.sync();
          sheetsToScan = [activeSheet.name];
        }
        console.log(`  📄 Scanning sheet: ${sheetsToScan[0]}`);
      }

      // Scan each sheet
      for (const sheetNameToScan of sheetsToScan) {
        const worksheet = context.workbook.worksheets.getItem(sheetNameToScan);
        const usedRange = worksheet.getUsedRange();
        
        usedRange.load('address, formulas, values, rowCount, columnCount');
        await context.sync();

        console.log(`  🔎 Scanning ${sheetNameToScan}: ${usedRange.address}`);

        // Scan all cells for errors
        const formulas = usedRange.formulas;
        const values = usedRange.values;
        
        for (let row = 0; row < usedRange.rowCount; row++) {
          for (let col = 0; col < usedRange.columnCount; col++) {
            const value = values[row][col];
            const formula = formulas[row][col];
            
            // Check if value is an error
            if (typeof value === 'string' && value.startsWith('#')) {
              const errorType = value;
              
              if (errorTypes.hasOwnProperty(errorType)) {
                errorTypes[errorType]++;
                
                // Convert row/col to cell address (e.g., A1, B5)
                const cellAddress = columnToLetter(col) + (row + 1);
                const fullAddress = `${sheetNameToScan}!${cellAddress}`;
                
                const errorInfo = {
                  address: cellAddress,
                  fullAddress: fullAddress,
                  sheet: sheetNameToScan,
                  errorType: errorType,
                  formula: formula || null,
                  value: value,
                  row: row + 1,
                  column: col + 1
                };

                // Trace propagation if requested
                if (tracePropagation) {
                  try {
                    const affectedCells = await traceErrorPropagation(
                      context, 
                      sheetNameToScan, 
                      cellAddress
                    );
                    errorInfo.affectedCells = affectedCells;
                    errorInfo.propagationCount = affectedCells.length;
                  } catch (e) {
                    console.warn(`  ⚠️ Could not trace propagation for ${fullAddress}:`, e.message);
                    errorInfo.affectedCells = [];
                    errorInfo.propagationCount = 0;
                  }
                }
                
                errors.push(errorInfo);
                console.log(`  ❌ Found ${errorType} at ${fullAddress}`);
              }
            }
          }
        }
      }

      // Calculate summary
      const totalErrors = errors.length;
      const criticalErrors = errorTypes['#REF!'] + errorTypes['#VALUE!'] + errorTypes['#NAME?'];
      const dataErrors = errorTypes['#N/A'] + errorTypes['#DIV/0!'] + errorTypes['#NUM!'];

      const result = {
        success: true,
        errorCount: totalErrors,
        errors: errors,
        summary: {
          total: totalErrors,
          critical: criticalErrors,
          dataErrors: dataErrors,
          byType: errorTypes,
          sheetsScanned: sheetsToScan.length,
          scope: scope
        }
      };

      console.log(`  ✅ Scan complete: ${totalErrors} errors found`);
      console.log(`     Critical: ${criticalErrors}, Data: ${dataErrors}`);

      return result;

    } catch (error) {
      console.error(`  ❌ Error in findErrors:`, error);
      throw {
        tool: "findErrors",
        success: false,
        error: error.message,
        params: params
      };
    }
  });
}

/**
 * Trace error propagation - find cells that reference this error cell
 */
async function traceErrorPropagation(context, sheetName, cellAddress) {
  try {
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const cell = worksheet.getRange(cellAddress);
    
    // Get direct dependents
    const dependents = cell.getDirectDependents();
    dependents.load('address, values');
    await context.sync();

    const affectedCells = [];
    
    // Check if dependents contain errors (propagation)
    if (dependents.address) {
      const areas = dependents.areas;
      areas.load('items');
      await context.sync();

      for (let i = 0; i < areas.items.length; i++) {
        const area = areas.items[i];
        area.load('address, values');
        await context.sync();

        // Extract addresses from the area
        let areaAddress = area.address;
        if (areaAddress.includes('!')) {
          areaAddress = areaAddress.split('!')[1];
        }

        const values = area.values;
        const hasError = values.some(row => 
          row.some(cell => typeof cell === 'string' && cell.startsWith('#'))
        );

        affectedCells.push({
          address: areaAddress,
          hasError: hasError
        });
      }
    }

    return affectedCells;
  } catch (e) {
    // If cell has no dependents, return empty array
    return [];
  }
}

/**
 * Convert column index to letter (0 = A, 1 = B, 25 = Z, 26 = AA)
 */
function columnToLetter(column) {
  let letter = '';
  let temp = column;
  
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  
  return letter;
}

