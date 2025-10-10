/**
 * getTableData.js
 *
 * Tool: Fetch all data from an Excel Table
 * Executor: Frontend (requires Office.js)
 */

export const toolDefinition = {
  name: "getTableData",
  description: "Fetch all data from an Excel Table (structured data with headers)",
  executor: "frontend",
  parameters: {
    tableName: {
      type: "string",
      required: true,
      description: "Name of the Excel Table"
    },
    includeHeaders: {
      type: "boolean",
      required: false,
      default: true,
      description: "Include header row in results"
    }
  },
  returns: {
    tableName: "string",
    sheetName: "string",
    address: "string",
    headers: "array",
    data: "array[][]",
    rowCount: "number",
    columnCount: "number"
  }
};

export async function execute(params) {
  console.log(`🔧 Executing getTableData:`, params);

  return Excel.run(async (context) => {
    try {
      const { tableName, includeHeaders = true } = params;

      if (!tableName || typeof tableName !== 'string') {
        throw new Error('Invalid tableName: must be a non-empty string');
      }

      // Get the table
      const table = context.workbook.tables.getItem(tableName);
      table.load("name");

      // Get table range
      const tableRange = table.getRange();
      tableRange.load("address, values, rowCount, columnCount");

      // Get header row
      const headerRow = table.getHeaderRowRange();
      headerRow.load("values");

      // Get worksheet
      const worksheet = tableRange.worksheet;
      worksheet.load("name");

      await context.sync();

      const headers = headerRow.values[0];
      const allData = tableRange.values;

      // Remove header row from data if includeHeaders is false
      const dataOnly = includeHeaders ? allData : allData.slice(1);

      const result = {
        tableName: table.name,
        sheetName: worksheet.name,
        address: tableRange.address,
        headers: headers,
        data: dataOnly,
        rowCount: dataOnly.length,
        columnCount: tableRange.columnCount
      };

      console.log(`  ✅ Retrieved table "${tableName}" with ${result.rowCount}x${result.columnCount} data`);

      return result;

    } catch (error) {
      console.error(`  ❌ Error in getTableData:`, error);
      throw {
        tool: "getTableData",
        error: error.message,
        params: params
      };
    }
  });
}
