/**
 * deleteRows.js
 *
 * Tool: Delete rows from a sheet
 * Executor: Frontend (requires Office.js)
 */

export const toolDefinition = {
  name: "deleteRows",
  description: "Delete rows from a sheet",
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
      description: "Starting row number to delete (1-based)"
    },
    rowCount: {
      type: "number",
      required: true,
      description: "Number of rows to delete"
    }
  },
  returns: {
    sheetName: "string",
    startRow: "number",
    rowCount: "number",
    deleted: "boolean"
  }
};

export async function execute(params) {
  console.log(`🔧 Executing deleteRows:`, params);

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

      // Get range representing the rows to delete
      // Convert 1-based row number to 0-based index
      const range = worksheet.getRangeByIndexes(startRow - 1, 0, rowCount, 1);

      // Delete rows by shifting up
      range.delete(Excel.DeleteShiftDirection.up);

      await context.sync();

      const result = {
        sheetName: sheetName,
        startRow: startRow,
        rowCount: rowCount,
        deleted: true
      };

      console.log(`  ✅ Deleted ${rowCount} row(s) starting at row ${startRow} in ${sheetName}`);

      return result;

    } catch (error) {
      console.error(`  ❌ Error in deleteRows:`, error);
      throw {
        tool: "deleteRows",
        error: error.message,
        params: params
      };
    }
  });
}
