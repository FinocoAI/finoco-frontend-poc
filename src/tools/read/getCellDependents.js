/**
 * getCellDependents.js
 *
 * Tool: Find all cells that reference this cell (dependent tracing)
 * Executor: Frontend (requires Office.js)
 */

export const toolDefinition = {
  name: "getCellDependents",
  description: "Find all cells that reference/depend on this cell (trace dependents)",
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
      description: "Cell address (e.g., 'A5')"
    }
  },
  returns: {
    sheetName: "string",
    sourceCell: "string",
    dependents: "array"  // Array of dependent cell addresses
  }
};

export async function execute(params) {
  console.log(`🔧 Executing getCellDependents:`, params);

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
        // Get direct dependents (cells that depend on this cell)
        const dependents = range.getDirectDependents();
        dependents.load("address, areas");
        await context.sync();

        // Extract dependent addresses
        const areas = dependents.areas;
        areas.load("items");
        await context.sync();

        const dependentAddresses = [];
        for (let i = 0; i < areas.items.length; i++) {
          const area = areas.items[i];
          area.load("address");
          await context.sync();
          dependentAddresses.push(area.address);
        }

        const result = {
          sheetName: sheetName,
          sourceCell: range.address,
          dependents: dependentAddresses,
          dependentCount: dependentAddresses.length
        };

        console.log(`  ✅ Found ${dependentAddresses.length} dependent(s) for ${address}`);

        return result;

      } catch (error) {
        // No dependents found
        if (error.message && error.message.includes('DirectDependentsNotFound')) {
          const result = {
            sheetName: sheetName,
            sourceCell: range.address,
            dependents: [],
            dependentCount: 0
          };

          console.log(`  ℹ️ No dependents found for ${address}`);
          return result;
        }
        throw error;
      }

    } catch (error) {
      console.error(`  ❌ Error in getCellDependents:`, error);
      throw {
        tool: "getCellDependents",
        error: error.message,
        params: params
      };
    }
  });
}
