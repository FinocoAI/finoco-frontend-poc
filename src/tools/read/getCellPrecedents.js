/**
 * getCellPrecedents.js
 *
 * Tool: Find all cells that this cell references (precedent tracing)
 * Executor: Frontend (requires Office.js)
 */

export const toolDefinition = {
  name: "getCellPrecedents",
  description: "Find all cells that this cell depends on/references (trace precedents)",
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
      description: "Cell address (e.g., 'C5')"
    }
  },
  returns: {
    sheetName: "string",
    sourceCell: "string",
    precedents: "array"  // Array of precedent cell addresses
  }
};

export async function execute(params) {
  console.log(`🔧 Executing getCellPrecedents:`, params);

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

      range.load("address");
      await context.sync();

      try {
        // Get direct precedents (cells that this cell depends on)
        const precedents = range.getDirectPrecedents();
        precedents.load("address, areas");
        await context.sync();

        // Extract precedent addresses
        const areas = precedents.areas;
        areas.load("items");
        await context.sync();

        const precedentAddresses = [];
        for (let i = 0; i < areas.items.length; i++) {
          const area = areas.items[i];
          area.load("address");
          await context.sync();
          precedentAddresses.push(area.address);
        }

        const result = {
          sheetName: sheetName,
          sourceCell: range.address,
          precedents: precedentAddresses,
          precedentCount: precedentAddresses.length
        };

        console.log(`  ✅ Found ${precedentAddresses.length} precedent(s) for ${address}`);

        return result;

      } catch (error) {
        // No precedents found (not a formula cell, or no dependencies)
        if (error.message && (error.message.includes('DirectPrecedentsNotFound') || error.message.includes('OperationCellsCannotBeReferenced'))) {
          const result = {
            sheetName: sheetName,
            sourceCell: range.address,
            precedents: [],
            precedentCount: 0
          };

          console.log(`  ℹ️ No precedents found for ${address}`);
          return result;
        }
        throw error;
      }

    } catch (error) {
      console.error(`  ❌ Error in getCellPrecedents:`, error);
      throw {
        tool: "getCellPrecedents",
        error: error.message,
        params: params
      };
    }
  });
}
