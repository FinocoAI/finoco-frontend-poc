/**
 * insertRows.js
 *
 * Tool: Insert new rows at a specific position
 * Executor: Frontend (requires Office.js)
 */

export const toolDefinition = {
  name: "insertRows",
  description: "Insert new empty rows at a specific position in the sheet",
  executor: "frontend",
  parameters: {
    sheetName: {
      type: "string",
      required: true,
      description: "Name of the sheet"
    },
    startRow: {
      type: "number",
      required: true,
      description: "Row number where to insert (1-based)"
    },
    rowCount: {
      type: "number",
      required: true,
      description: "Number of rows to insert"
    }
  },
  returns: {
    sheetName: "string",
    startRow: "number",
    rowCount: "number",
    inserted: "boolean"
  }
};

export async function execute(params) {
  console.log(`🔧 Executing insertRows:`, params);

  return Excel.run(async (context) => {
    try {
      const { sheetName, startRow, rowCount } = params;

      // Validate parameters
      if (!sheetName || typeof sheetName !== 'string') {
        throw new Error('Invalid sheetName');
      }
      if (typeof startRow !== 'number' || startRow < 1) {
        throw new Error('Invalid startRow: must be a positive number');
      }
      if (typeof rowCount !== 'number' || rowCount < 1) {
        throw new Error('Invalid rowCount: must be a positive number');
      }

      // Get worksheet
      const worksheet = context.workbook.worksheets.getItem(sheetName);

      // Get range representing the rows to insert
      // Convert 1-based row number to 0-based index
      const range = worksheet.getRangeByIndexes(startRow - 1, 0, rowCount, 1);

      // Insert rows by shifting down
      range.insert(Excel.InsertShiftDirection.down);

      await context.sync();

      const result = {
        sheetName: sheetName,
        startRow: startRow,
        rowCount: rowCount,
        inserted: true
      };

      console.log(`  ✅ Inserted ${rowCount} row(s) at row ${startRow} in ${sheetName}`);

      return result;

    } catch (error) {
      console.error(`  ❌ Error in insertRows:`, error);
      throw {
        tool: "insertRows",
        error: error.message,
        params: params
      };
    }
  });
}
