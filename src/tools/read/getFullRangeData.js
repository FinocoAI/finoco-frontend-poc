/**
 * getFullRangeData.js
 *
 * Tool: Fetch complete data from a specific range address
 * Executor: Frontend (requires Office.js)
 *
 * Use Cases:
 * - Agent needs full data after seeing preview
 * - User asks to analyze specific range
 * - Backend needs complete dataset for processing
 */

/**
 * Tool definition for LLM/Backend
 */
export const toolDefinition = {
  name: "getFullRangeData",
  description: "Fetch complete data from a specific range address. Returns all cell values, formulas, and metadata.",
  executor: "frontend",
  parameters: {
    sheetName: {
      type: "string",
      required: true,
      description: "Name of the sheet containing the range"
    },
    address: {
      type: "string",
      required: true,
      description: "Range address in A1 notation (e.g., 'A1:C100', 'B5:E50')"
    },
    includeFormulas: {
      type: "boolean",
      required: false,
      default: false,
      description: "If true, returns formulas instead of values for formula cells"
    }
  },
  returns: {
    sheetName: "string",
    address: "string",
    rowCount: "number",
    columnCount: "number",
    values: "array[][]",        // 2D array of values
    formulas: "array[][]",      // 2D array of formulas (if includeFormulas=true)
    headers: "array",           // First row as headers
  }
};

/**
 * Execute the tool
 *
 * @param {Object} params - Tool parameters
 * @param {string} params.sheetName - Sheet name
 * @param {string} params.address - Range address (e.g., "A1:C100")
 * @param {boolean} [params.includeFormulas=false] - Include formulas
 * @returns {Promise<Object>} Range data object
 */
export async function execute(params) {
  console.log(`🔧 Executing getFullRangeData:`, params);

  return Excel.run(async (context) => {
    try {
      const { sheetName, address, includeFormulas = false } = params;

      // Validate parameters
      if (!sheetName || typeof sheetName !== 'string') {
        throw new Error('Invalid sheetName: must be a non-empty string');
      }
      if (!address || typeof address !== 'string') {
        throw new Error('Invalid address: must be a valid range address (e.g., "A1:C100")');
      }

      // Get the worksheet
      const worksheet = context.workbook.worksheets.getItem(sheetName);

      // Get the range
      const range = worksheet.getRange(address);
      range.load("address, values, rowCount, columnCount");

      // Load formulas if requested
      if (includeFormulas) {
        range.load("formulas");
      }

      await context.sync();

      // Build result object
      const result = {
        sheetName: sheetName,
        address: range.address,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        values: range.values,
        headers: range.rowCount > 0 ? range.values[0] : [],
      };

      // Add formulas if requested
      if (includeFormulas) {
        result.formulas = range.formulas;
      }

      console.log(`  ✅ Retrieved ${result.rowCount}x${result.columnCount} range from ${sheetName}`);

      return result;

    } catch (error) {
      console.error(`  ❌ Error in getFullRangeData:`, error);
      throw {
        tool: "getFullRangeData",
        error: error.message,
        params: params
      };
    }
  });
}
