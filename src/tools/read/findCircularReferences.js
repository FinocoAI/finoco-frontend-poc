/**
 * findCircularReferences.js
 *
 * Tool: Find circular references in worksheet or workbook
 * Executor: Frontend (requires Office.js)
 * 
 * Detects circular reference chains and shows the loop path
 */

export const toolDefinition = {
  name: "findCircularReferences",
  description: "Detect circular references in formulas. Shows the circular chain (e.g., A1→B2→C3→A1) and helps identify intentional vs problematic circular refs.",
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
    }
  },
  returns: {
    success: "boolean",
    circularCount: "number",
    circularReferences: "array of circular reference objects with chain path",
    hasCirculars: "boolean"
  }
};

/**
 * Check if a formula directly references its own cell address
 * Handles cases like =SUM(A1, B2, A1) in cell A1
 */
function checkDirectSelfReference(formula, cellAddress, sheetName) {
  // Remove the leading '=' and convert to uppercase for comparison
  const formulaUpper = formula.toUpperCase();
  const cellAddressUpper = cellAddress.toUpperCase();
  
  // Pattern to match cell references in the formula
  // Matches: A1, $A$1, $A1, A$1, Sheet1!A1, 'Sheet Name'!A1
  const cellRefPattern = /(?:(?:'[^']*'|[\w]+)!)?(\$?[A-Z]+\$?\d+)/g;
  
  let match;
  while ((match = cellRefPattern.exec(formulaUpper)) !== null) {
    const refMatch = match[0]; // Full match (e.g., "Sheet1!A1" or "A1")
    const cellPart = match[1]; // Just the cell part (e.g., "A1" or "$A$1")
    
    // Remove $ signs from absolute references
    const normalizedCell = cellPart.replace(/\$/g, '');
    
    // Check if this reference is to the same cell
    if (normalizedCell === cellAddressUpper) {
      // If there's a sheet reference, verify it's the same sheet or no sheet specified
      if (refMatch.includes('!')) {
        const sheetPart = refMatch.split('!')[0].replace(/'/g, '').toUpperCase();
        const currentSheetUpper = sheetName.toUpperCase();
        if (sheetPart === currentSheetUpper) {
          return true;
        }
      } else {
        // No sheet specified, same cell reference = circular
        return true;
      }
    }
  }
  
  return false;
}

export async function execute(params) {
  console.log('🔄 Executing findCircularReferences:', params);

  return Excel.run(async (context) => {
    const { 
      scope = 'sheet', 
      sheetName = null
    } = params;

    const circularRefs = [];
    const checkedCells = new Set(); // Avoid re-checking

    try {
      // Determine which sheets to scan
      let sheetsToScan = [];
      
      if (scope === 'workbook') {
        const allSheets = context.workbook.worksheets;
        allSheets.load('items/name');
        await context.sync();
        
        sheetsToScan = allSheets.items.map(sheet => sheet.name);
        console.log(`  📊 Scanning ${sheetsToScan.length} sheets for circular references`);
      } else {
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

      // Scan each sheet for formulas
      for (const sheetNameToScan of sheetsToScan) {
        const worksheet = context.workbook.worksheets.getItem(sheetNameToScan);
        const usedRange = worksheet.getUsedRange();
        
        usedRange.load('address, formulas, rowCount, columnCount');
        await context.sync();

        console.log(`  🔎 Scanning ${sheetNameToScan}: ${usedRange.address}`);

        const formulas = usedRange.formulas;
        
        // Find all cells with formulas
        for (let row = 0; row < usedRange.rowCount; row++) {
          for (let col = 0; col < usedRange.columnCount; col++) {
            const formula = formulas[row][col];
            
            // Only check cells with formulas
            if (formula && typeof formula === 'string' && formula.startsWith('=')) {
              const cellAddress = columnToLetter(col) + (row + 1);
              const fullAddress = `${sheetNameToScan}!${cellAddress}`;
              
              // Skip if already checked
              if (checkedCells.has(fullAddress)) {
                continue;
              }

              // FIRST: Check for direct self-reference in the formula
              // This catches cases like =SUM(A1, B2, A1) where A1 references itself
              const directSelfRef = checkDirectSelfReference(formula, cellAddress, sheetNameToScan);
              if (directSelfRef) {
                circularRefs.push({
                  startCell: fullAddress,
                  chain: [fullAddress, fullAddress],
                  chainLength: 2,
                  formula: formula
                });
                checkedCells.add(fullAddress);
                console.log(`  🔄 Found direct self-reference: ${fullAddress} → ${fullAddress}`);
                continue;
              }

              // Check if this cell has a circular reference
              try {
                const circularChain = await detectCircularChain(
                  context,
                  sheetNameToScan,
                  cellAddress,
                  new Set(),
                  [fullAddress]
                );

                if (circularChain) {
                  circularRefs.push({
                    startCell: fullAddress,
                    chain: circularChain,
                    chainLength: circularChain.length,
                    formula: formula
                  });

                  // Mark all cells in chain as checked
                  circularChain.forEach(addr => checkedCells.add(addr));
                  
                  console.log(`  🔄 Found circular: ${circularChain.join(' → ')}`);
                }
              } catch (e) {
                // Cell might not have precedents or other issues - skip
                console.warn(`  ⚠️ Could not check ${fullAddress}:`, e.message);
              }

              checkedCells.add(fullAddress);
            }
          }
        }
      }

      const result = {
        success: true,
        circularCount: circularRefs.length,
        hasCirculars: circularRefs.length > 0,
        circularReferences: circularRefs,
        sheetsScanned: sheetsToScan.length,
        scope: scope
      };

      if (circularRefs.length > 0) {
        console.log(`  ✅ Found ${circularRefs.length} circular reference(s)`);
      } else {
        console.log(`  ✅ No circular references found`);
      }

      return result;

    } catch (error) {
      console.error(`  ❌ Error in findCircularReferences:`, error);
      throw {
        tool: "findCircularReferences",
        success: false,
        error: error.message,
        params: params
      };
    }
  });
}

/**
 * Detect circular reference chain using DFS
 * Returns the chain if circular, null otherwise
 */
async function detectCircularChain(context, sheetName, cellAddress, visited, path) {
  const fullAddress = `${sheetName}!${cellAddress}`;
  
  // If we've seen this cell before in current path, we found a circle!
  if (visited.has(fullAddress)) {
    // Find where the circle starts
    const circleStartIndex = path.indexOf(fullAddress);
    if (circleStartIndex !== -1) {
      // Return the circular portion of the path
      return path.slice(circleStartIndex).concat([fullAddress]);
    }
    return null;
  }

  // Add current cell to visited set
  visited.add(fullAddress);

  try {
    // Get precedents (cells this formula depends on)
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const cell = worksheet.getRange(cellAddress);
    const precedents = cell.getDirectPrecedents();
    
    precedents.load('address');
    await context.sync();

    if (!precedents.address) {
      // No precedents, no circle
      return null;
    }

    // Parse precedent addresses
    const areas = precedents.areas;
    areas.load('items');
    await context.sync();

    for (let i = 0; i < areas.items.length; i++) {
      const area = areas.items[i];
      area.load('address');
      await context.sync();

      let areaAddress = area.address;
      
      // Parse sheet and address
      let precSheet = sheetName;
      let precAddress = areaAddress;
      
      if (areaAddress.includes('!')) {
        const parts = areaAddress.split('!');
        precSheet = parts[0].replace(/'/g, ''); // Remove quotes
        precAddress = parts[1];
      }

      // If it's a range, just check the first cell
      if (precAddress.includes(':')) {
        precAddress = precAddress.split(':')[0];
      }

      const precFullAddress = `${precSheet}!${precAddress}`;

      // Recursively check this precedent
      const newPath = [...path, precFullAddress];
      const circularChain = await detectCircularChain(
        context,
        precSheet,
        precAddress,
        new Set(visited),
        newPath
      );

      if (circularChain) {
        return circularChain;
      }
    }

    return null;

  } catch (e) {
    // No precedents or error accessing them
    return null;
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

