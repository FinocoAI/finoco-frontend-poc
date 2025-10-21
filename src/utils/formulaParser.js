/**
 * formulaParser.js
 * 
 * Utility to parse Excel formulas and extract cell/range references
 * Used as fallback when Excel API can't detect precedents
 */

/**
 * Extract all cell and range references from a formula
 * Handles: A1, Sheet1!A1, 'Sheet Name'!A1:B10, A:A, 1:1, named ranges
 * 
 * @param {string} formula - The formula string (with or without leading =)
 * @param {string} defaultSheet - The sheet where the formula exists
 * @returns {Array} Array of reference objects: { sheet, address, isFullColumn, isFullRow }
 */
export function parseFormulaReferences(formula, defaultSheet) {
  if (!formula || typeof formula !== 'string') {
    return [];
  }

  // Remove leading = if present
  formula = formula.trim();
  if (formula.startsWith('=')) {
    formula = formula.substring(1);
  }

  const references = [];
  const seen = new Set(); // Avoid duplicates

  // Pattern breakdown:
  // 1. Optional sheet name: (?:([^\s!]+|'[^']+')\!)?
  // 2. Cell/range reference: ([A-Z]+[0-9]+(?:\:[A-Z]+[0-9]+)?|[A-Z]+\:[A-Z]+|[0-9]+\:[0-9]+)
  
  // Regex to match cell references:
  // - Sheet1!A1 or 'Sheet Name'!A1
  // - A1:B10 (ranges)
  // - A:A (full columns)
  // - 1:1 (full rows)
  // - A1 (simple cell)
  const cellRefPattern = /(?:([^\s!()]+|'[^']+')\!)?([A-Z]{1,3}[0-9]{1,7}(?:\:[A-Z]{1,3}[0-9]{1,7})?|[A-Z]{1,3}\:[A-Z]{1,3}|[0-9]{1,7}\:[0-9]{1,7}|[A-Z]{1,3}[0-9]{1,7})/gi;

  let match;
  while ((match = cellRefPattern.exec(formula)) !== null) {
    let sheetName = match[1];
    let address = match[2];

    // Skip if this looks like a function name (e.g., "SUM" before parenthesis)
    const charBefore = formula[match.index - 1];
    if (charBefore && /[A-Za-z]/.test(charBefore)) {
      continue;
    }

    // Clean sheet name (remove quotes)
    if (sheetName) {
      sheetName = sheetName.replace(/^'|'$/g, '');
    } else {
      sheetName = defaultSheet;
    }

    // Check if it's a full column or full row reference
    const isFullColumn = /^[A-Z]{1,3}\:[A-Z]{1,3}$/.test(address);
    const isFullRow = /^[0-9]{1,7}\:[0-9]{1,7}$/.test(address);

    // Create unique key
    const key = `${sheetName}!${address}`;
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);

    references.push({
      sheet: sheetName,
      address: address,
      isFullColumn: isFullColumn,
      isFullRow: isFullRow,
      isRange: address.includes(':')
    });
  }

  return references;
}

/**
 * Convert full column/row references to a representative cell
 * For full columns (A:A), returns A1
 * For full rows (1:1), returns A1
 * This helps with visualization in dependency graph
 * 
 * @param {string} address - The address (e.g., "A:A", "1:1", "A1:B10")
 * @returns {string} Representative address
 */
export function getRepresentativeAddress(address) {
  // Full column (A:A, B:C, etc.)
  if (/^[A-Z]{1,3}\:[A-Z]{1,3}$/.test(address)) {
    const firstCol = address.split(':')[0];
    return `${firstCol}1`; // Return first cell of first column
  }
  
  // Full row (1:1, 2:5, etc.)
  if (/^[0-9]{1,7}\:[0-9]{1,7}$/.test(address)) {
    const firstRow = address.split(':')[0];
    return `A${firstRow}`; // Return first cell of first row
  }
  
  // Range (A1:B10) - return first cell
  if (address.includes(':')) {
    return address.split(':')[0];
  }
  
  // Single cell - return as is
  return address;
}

/**
 * Expand a range to individual cells (with limit to avoid huge ranges)
 * Used when we need actual cell addresses, not ranges
 * 
 * @param {string} address - Range address (e.g., "A1:B3")
 * @param {number} maxCells - Maximum cells to return (default 100)
 * @returns {Array<string>} Array of cell addresses
 */
export function expandRange(address, maxCells = 100) {
  // Not a range
  if (!address.includes(':')) {
    return [address];
  }

  // Full column or full row - too large, return representative
  if (/^[A-Z]{1,3}\:[A-Z]{1,3}$/.test(address) || /^[0-9]{1,7}\:[0-9]{1,7}$/.test(address)) {
    return [getRepresentativeAddress(address)];
  }

  // Parse range
  const [start, end] = address.split(':');
  const startCol = start.match(/[A-Z]+/)[0];
  const startRow = parseInt(start.match(/[0-9]+/)[0]);
  const endCol = end.match(/[A-Z]+/)[0];
  const endRow = parseInt(end.match(/[0-9]+/)[0]);

  const cells = [];
  const startColNum = columnToNumber(startCol);
  const endColNum = columnToNumber(endCol);

  // Expand with limit
  for (let row = startRow; row <= endRow; row++) {
    for (let col = startColNum; col <= endColNum; col++) {
      if (cells.length >= maxCells) {
        return cells;
      }
      cells.push(`${numberToColumn(col)}${row}`);
    }
  }

  return cells;
}

/**
 * Convert column letter to number (A=1, B=2, Z=26, AA=27, etc.)
 */
function columnToNumber(col) {
  let num = 0;
  for (let i = 0; i < col.length; i++) {
    num = num * 26 + (col.charCodeAt(i) - 64);
  }
  return num;
}

/**
 * Convert column number to letter (1=A, 2=B, 26=Z, 27=AA, etc.)
 */
function numberToColumn(num) {
  let col = '';
  while (num > 0) {
    const remainder = (num - 1) % 26;
    col = String.fromCharCode(65 + remainder) + col;
    num = Math.floor((num - 1) / 26);
  }
  return col;
}

/**
 * Check if a formula references a specific cell
 * Used for finding dependents through formula scanning
 * 
 * @param {string} formula - The formula to check
 * @param {string} targetSheet - Sheet name of target cell
 * @param {string} targetAddress - Address of target cell
 * @returns {boolean} True if formula references the target cell
 */
export function formulaReferencesCell(formula, targetSheet, targetAddress) {
  const references = parseFormulaReferences(formula, targetSheet);
  
  return references.some(ref => {
    // Exact match
    if (ref.sheet === targetSheet && ref.address === targetAddress) {
      return true;
    }
    
    // Check if target is within a range
    if (ref.sheet === targetSheet && ref.isRange) {
      return isAddressInRange(targetAddress, ref.address);
    }
    
    return false;
  });
}

/**
 * Check if an address is within a range
 * 
 * @param {string} address - Cell address (e.g., "B5")
 * @param {string} range - Range (e.g., "A1:C10", "B:B", "5:5")
 * @returns {boolean}
 */
function isAddressInRange(address, range) {
  if (!range.includes(':')) {
    return address === range;
  }

  // Full column (B:B)
  if (/^[A-Z]{1,3}\:[A-Z]{1,3}$/.test(range)) {
    const [startCol, endCol] = range.split(':');
    const addrCol = address.match(/[A-Z]+/)[0];
    const startNum = columnToNumber(startCol);
    const endNum = columnToNumber(endCol);
    const addrNum = columnToNumber(addrCol);
    return addrNum >= startNum && addrNum <= endNum;
  }

  // Full row (5:5)
  if (/^[0-9]{1,7}\:[0-9]{1,7}$/.test(range)) {
    const [startRow, endRow] = range.split(':').map(Number);
    const addrRow = parseInt(address.match(/[0-9]+/)[0]);
    return addrRow >= startRow && addrRow <= endRow;
  }

  // Regular range (A1:C10)
  const [start, end] = range.split(':');
  const startCol = columnToNumber(start.match(/[A-Z]+/)[0]);
  const startRow = parseInt(start.match(/[0-9]+/)[0]);
  const endCol = columnToNumber(end.match(/[A-Z]+/)[0]);
  const endRow = parseInt(end.match(/[0-9]+/)[0]);
  
  const addrCol = columnToNumber(address.match(/[A-Z]+/)[0]);
  const addrRow = parseInt(address.match(/[0-9]+/)[0]);

  return addrCol >= startCol && addrCol <= endCol &&
         addrRow >= startRow && addrRow <= endRow;
}

