/**
 * getNamedRangeData.js
 *
 * Tool: Fetch data from a named range
 * Executor: Frontend (requires Office.js)
 */

export const toolDefinition = {
  name: "getNamedRangeData",
  description: "Fetch data from a named range (workbook-level or sheet-level)",
  executor: "frontend",
  parameters: {
    rangeName: {
      type: "string",
      required: true,
      description: "Name of the named range"
    }
  },
  returns: {
    rangeName: "string",
    address: "string",
    sheetName: "string",
    values: "array[][]",
    rowCount: "number",
    columnCount: "number"
  }
};

export async function execute(params) {
  console.log(`🔧 Executing getNamedRangeData:`, params);

  return Excel.run(async (context) => {
    try {
      const { rangeName } = params;

      if (!rangeName || typeof rangeName !== 'string') {
        throw new Error('Invalid rangeName: must be a non-empty string');
      }

      // Get the named range
      const namedRange = context.workbook.names.getItem(rangeName);
      namedRange.load("name, formula");

      await context.sync();

      // Get the range that the name refers to
      const range = namedRange.getRange();
      range.load("address, values, rowCount, columnCount");

      const worksheet = range.worksheet;
      worksheet.load("name");

      await context.sync();

      const result = {
        rangeName: namedRange.name,
        address: range.address,
        sheetName: worksheet.name,
        values: range.values,
        rowCount: range.rowCount,
        columnCount: range.columnCount
      };

      console.log(`  ✅ Retrieved named range "${rangeName}" (${result.rowCount}x${result.columnCount})`);

      return result;

    } catch (error) {
      console.error(`  ❌ Error in getNamedRangeData:`, error);
      throw {
        tool: "getNamedRangeData",
        error: error.message,
        params: params
      };
    }
  });
}
