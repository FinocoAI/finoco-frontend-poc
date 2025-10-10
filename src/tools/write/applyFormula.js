/**
 * applyFormula.js
 *
 * Tool: Apply a formula to a range
 * Executor: Frontend (requires Office.js)
 */

export const toolDefinition = {
  name: "applyFormula",
  description: "Apply a formula to a cell or range",
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
      description: "Cell or range address (e.g., 'A1' or 'A1:A10')"
    },
    formula: {
      type: "string",
      required: true,
      description: "Formula to apply (e.g., '=SUM(B:B)', '=A1*2')"
    }
  },
  returns: {
    sheetName: "string",
    address: "string",
    formula: "string",
    applied: "boolean"
  }
};

export async function execute(params) {
  console.log(`🔧 Executing applyFormula:`, params);

  return Excel.run(async (context) => {
    try {
      const { sheetName, address, formula } = params;

      // Validate parameters
      if (!sheetName || typeof sheetName !== 'string') {
        throw new Error('Invalid sheetName');
      }
      if (!address || typeof address !== 'string') {
        throw new Error('Invalid address');
      }
      if (!formula || typeof formula !== 'string') {
        throw new Error('Invalid formula');
      }

      // Ensure formula starts with =
      const cleanFormula = formula.startsWith('=') ? formula : `=${formula}`;

      // Get worksheet
      const worksheet = context.workbook.worksheets.getItem(sheetName);

      // Get the range
      const range = worksheet.getRange(address);
      range.load("address, rowCount, columnCount");
      await context.sync();

      // Apply formula to all cells in range
      if (range.rowCount === 1 && range.columnCount === 1) {
        // Single cell
        range.formulas = [[cleanFormula]];
      } else {
        // Multiple cells - apply same formula to all
        const formulas = [];
        for (let i = 0; i < range.rowCount; i++) {
          const row = [];
          for (let j = 0; j < range.columnCount; j++) {
            row.push(cleanFormula);
          }
          formulas.push(row);
        }
        range.formulas = formulas;
      }

      await context.sync();

      const result = {
        sheetName: sheetName,
        address: range.address,
        formula: cleanFormula,
        applied: true,
        cellsAffected: range.rowCount * range.columnCount
      };

      console.log(`  ✅ Applied formula "${cleanFormula}" to ${range.address}`);

      return result;

    } catch (error) {
      console.error(`  ❌ Error in applyFormula:`, error);
      throw {
        tool: "applyFormula",
        error: error.message,
        params: params
      };
    }
  });
}
