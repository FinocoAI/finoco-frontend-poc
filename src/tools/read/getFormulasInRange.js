/**
 * getFormulasInRange.js
 *
 * Tool: Get formulas (not values) from a range
 * Executor: Frontend (requires Office.js)
 */

export const toolDefinition = {
  name: "getFormulasInRange",
  description: "Get formulas (not computed values) from a range",
  executor: "frontend",
  parameters: {
    sheetName: {
      type: "string",
      required: true,
      description: "Name of the sheet"
    },
    address: {
      type: "string",
      required: true,
      description: "Range address (e.g., 'A1:C10')"
    }
  },
  returns: {
    sheetName: "string",
    address: "string",
    formulas: "array[][]",
    values: "array[][]",
    rowCount: "number",
    columnCount: "number"
  }
};

export async function execute(params) {
  console.log(`🔧 Executing getFormulasInRange:`, params);

  return Excel.run(async (context) => {
    try {
      const { sheetName, address } = params;

      if (!sheetName || typeof sheetName !== 'string') {
        throw new Error('Invalid sheetName');
      }
      if (!address || typeof address !== 'string') {
        throw new Error('Invalid address');
      }

      // Get worksheet and range
      const worksheet = context.workbook.worksheets.getItem(sheetName);
      const range = worksheet.getRange(address);

      range.load("address, formulas, values, rowCount, columnCount");
      await context.sync();

      const result = {
        sheetName: sheetName,
        address: range.address,
        formulas: range.formulas,
        values: range.values,
        rowCount: range.rowCount,
        columnCount: range.columnCount
      };

      console.log(`  ✅ Retrieved formulas from ${address} (${result.rowCount}x${result.columnCount})`);

      return result;

    } catch (error) {
      console.error(`  ❌ Error in getFormulasInRange:`, error);
      throw {
        tool: "getFormulasInRange",
        error: error.message,
        params: params
      };
    }
  });
}
