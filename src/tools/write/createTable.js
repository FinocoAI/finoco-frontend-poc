/**
 * createTable.js
 *
 * Tool: Convert a range to an Excel Table
 * Executor: Frontend (requires Office.js)
 */

export const toolDefinition = {
  name: "createTable",
  description: "Convert a range to an Excel Table with structured data",
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
      description: "Range address to convert (e.g., 'A1:D10')"
    },
    tableName: {
      type: "string",
      required: true,
      description: "Name for the new table"
    },
    hasHeaders: {
      type: "boolean",
      required: false,
      default: true,
      description: "Whether the first row contains headers"
    }
  },
  returns: {
    sheetName: "string",
    tableName: "string",
    address: "string",
    created: "boolean"
  }
};

export async function execute(params) {
  console.log(`🔧 Executing createTable:`, params);

  return Excel.run(async (context) => {
    try {
      const { sheetName, address, tableName, hasHeaders = true } = params;

      // Validate parameters
      if (!sheetName || typeof sheetName !== 'string') {
        throw new Error('Invalid sheetName');
      }
      if (!address || typeof address !== 'string') {
        throw new Error('Invalid address');
      }
      if (!tableName || typeof tableName !== 'string') {
        throw new Error('Invalid tableName');
      }

      // Get worksheet
      const worksheet = context.workbook.worksheets.getItem(sheetName);

      // Get the range
      const range = worksheet.getRange(address);
      range.load("address");

      // Create table from range
      const table = worksheet.tables.add(range, hasHeaders);
      table.name = tableName;

      await context.sync();

      const result = {
        sheetName: sheetName,
        tableName: table.name,
        address: range.address,
        created: true
      };

      console.log(`  ✅ Created table "${tableName}" from ${range.address}`);

      return result;

    } catch (error) {
      console.error(`  ❌ Error in createTable:`, error);
      throw {
        tool: "createTable",
        error: error.message,
        params: params
      };
    }
  });
}
