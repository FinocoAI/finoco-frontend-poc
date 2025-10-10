/**
 * getColumnData.js
 *
 * Tool: Get all values from specific column(s)
 * Executor: Frontend (requires Office.js)
 *
 * Use Cases:
 * - Extract specific columns for analysis
 * - Get column data without full table
 * - Optimize data transfer (only needed columns)
 */

/**
 * Tool definition for LLM/Backend
 */
export const toolDefinition = {
  name: "getColumnData",
  description: "Get all values from specific column(s) in a sheet. Returns complete column data from used range.",
  executor: "frontend",
  parameters: {
    sheetName: {
      type: "string",
      required: true,
      description: "Name of the sheet containing the columns"
    },
    columns: {
      type: "array",
      required: true,
      description: "Array of column letters or indices (e.g., ['A', 'C', 'D'] or [0, 2, 3])"
    },
    includeHeaders: {
      type: "boolean",
      required: false,
      default: true,
      description: "If true, includes first row as headers"
    }
  },
  returns: {
    sheetName: "string",
    columns: "array",           // Column identifiers used
    rowCount: "number",
    data: "object",             // { columnLetter: [values...] }
  }
};

/**
 * Convert column letter to zero-based index
 * A=0, B=1, Z=25, AA=26, etc.
 *
 * @param {string} letter - Column letter (e.g., "A", "Z", "AA")
 * @returns {number} Zero-based column index
 */
function columnLetterToIndex(letter) {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index - 1;
}

/**
 * Convert zero-based index to column letter
 * 0=A, 1=B, 25=Z, 26=AA, etc.
 *
 * @param {number} index - Zero-based column index
 * @returns {string} Column letter
 */
function columnIndexToLetter(index) {
  let letter = '';
  while (index >= 0) {
    letter = String.fromCharCode((index % 26) + 65) + letter;
    index = Math.floor(index / 26) - 1;
  }
  return letter;
}

/**
 * Execute the tool
 *
 * @param {Object} params - Tool parameters
 * @param {string} params.sheetName - Sheet name
 * @param {Array<string|number>} params.columns - Column letters or indices
 * @param {boolean} [params.includeHeaders=true] - Include headers
 * @returns {Promise<Object>} Column data object
 */
export async function execute(params) {
  console.log(`🔧 Executing getColumnData:`, params);

  return Excel.run(async (context) => {
    try {
      const { sheetName, columns, includeHeaders = true } = params;

      // Validate parameters
      if (!sheetName || typeof sheetName !== 'string') {
        throw new Error('Invalid sheetName: must be a non-empty string');
      }
      if (!Array.isArray(columns) || columns.length === 0) {
        throw new Error('Invalid columns: must be a non-empty array');
      }

      // Get the worksheet
      const worksheet = context.workbook.worksheets.getItem(sheetName);

      // Get the used range to know how many rows we have
      const usedRange = worksheet.getUsedRange();
      usedRange.load("rowCount, columnCount, address");
      await context.sync();

      console.log(`  📊 Sheet ${sheetName} has ${usedRange.rowCount} rows`);

      // Normalize column identifiers to indices
      const columnIndices = columns.map(col => {
        if (typeof col === 'number') {
          return col;
        } else if (typeof col === 'string') {
          return columnLetterToIndex(col.toUpperCase());
        } else {
          throw new Error(`Invalid column identifier: ${col}`);
        }
      });

      // Validate column indices
      for (const idx of columnIndices) {
        if (idx < 0 || idx >= usedRange.columnCount) {
          throw new Error(`Column index ${idx} is out of range (0-${usedRange.columnCount - 1})`);
        }
      }

      // Fetch each column's data
      const columnData = {};
      const columnLetters = [];

      for (const colIndex of columnIndices) {
        const columnLetter = columnIndexToLetter(colIndex);
        columnLetters.push(columnLetter);

        // Get the column range (from row 1 to last used row)
        const columnRange = worksheet.getRangeByIndexes(0, colIndex, usedRange.rowCount, 1);
        columnRange.load("values");
        await context.sync();

        // Extract values (flatten 2D array to 1D)
        const values = columnRange.values.map(row => row[0]);

        columnData[columnLetter] = includeHeaders ? values : values.slice(1);

        console.log(`  ✓ Retrieved column ${columnLetter} (${values.length} values)`);
      }

      const result = {
        sheetName: sheetName,
        columns: columnLetters,
        rowCount: includeHeaders ? usedRange.rowCount : usedRange.rowCount - 1,
        includeHeaders: includeHeaders,
        data: columnData,
      };

      console.log(`  ✅ Retrieved ${columnLetters.length} column(s) from ${sheetName}`);

      return result;

    } catch (error) {
      console.error(`  ❌ Error in getColumnData:`, error);
      throw {
        tool: "getColumnData",
        error: error.message,
        params: params
      };
    }
  });
}
