/**
 * createNewSheet.js
 *
 * Tool: Add a new worksheet to the workbook
 * Executor: Frontend (requires Office.js)
 */

export const toolDefinition = {
  name: "createNewSheet",
  description: "Add a new worksheet to the workbook",
  executor: "frontend",
  parameters: {
    sheetName: {
      type: "string",
      required: true,
      description: "Name for the new sheet"
    },
    position: {
      type: "number",
      required: false,
      description: "Index position for the new sheet (0-based). If not specified, adds at end."
    }
  },
  returns: {
    sheetName: "string",
    position: "number",
    created: "boolean"
  }
};

export async function execute(params) {
  console.log(`🔧 Executing createNewSheet:`, params);

  return Excel.run(async (context) => {
    try {
      const { sheetName, position } = params;

      // Validate parameters
      if (!sheetName || typeof sheetName !== 'string') {
        throw new Error('Invalid sheetName');
      }
      if (position !== undefined && (typeof position !== 'number' || position < 0)) {
        throw new Error('Invalid position: must be a non-negative number');
      }

      // Add new worksheet
      const sheet = context.workbook.worksheets.add(sheetName);

      // Set position if specified
      if (position !== undefined) {
        sheet.position = position;
      }

      sheet.load("name, position");
      await context.sync();

      const result = {
        sheetName: sheet.name,
        position: sheet.position,
        created: true
      };

      console.log(`  ✅ Created new sheet "${sheet.name}" at position ${sheet.position}`);

      return result;

    } catch (error) {
      console.error(`  ❌ Error in createNewSheet:`, error);
      throw {
        tool: "createNewSheet",
        error: error.message,
        params: params
      };
    }
  });
}
