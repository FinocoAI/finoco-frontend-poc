/**
 * writeDataToRange.js
 *
 * Tool: Write data to a specific range
 * Executor: Frontend (requires Office.js)
 */

export const toolDefinition = {
  name: "writeDataToRange",
  description: "Write 2D array data to a specific range in Excel",
  executor: "frontend",
  parameters: {
    sheetName: {
      type: "string",
      required: true,
      description: "Name of the sheet"
    },
    startCell: {
      type: "string",
      required: true,
      description: "Starting cell address (e.g., 'A1')"
    },
    data: {
      type: "array",
      required: true,
      description: "2D array of data to write"
    },
    overwrite: {
      type: "boolean",
      required: false,
      default: true,
      description: "If false, will only write to empty cells"
    }
  },
  returns: {
    sheetName: "string",
    rangeAddress: "string",
    rowsWritten: "number",
    columnsWritten: "number"
  }
};

export async function execute(params) {
  console.log(`🔧 Executing writeDataToRange:`, params);

  return Excel.run(async (context) => {
    try {
      const { sheetName, startCell, data, overwrite = true } = params;

      // Validate parameters
      if (!sheetName || typeof sheetName !== 'string') {
        throw new Error('Invalid sheetName');
      }
      if (!startCell || typeof startCell !== 'string') {
        throw new Error('Invalid startCell');
      }
      if (!Array.isArray(data) || data.length === 0) {
        throw new Error('Invalid data: must be a non-empty 2D array');
      }

      // Normalize data: find max columns and pad rows as needed
      let numCols = 0;
      for (let i = 0; i < data.length; i++) {
        if (!Array.isArray(data[i])) {
          throw new Error(`Invalid data: row ${i} is not an array`);
        }
        numCols = Math.max(numCols, data[i].length);
      }

      // Pad rows to have consistent column count
      const normalizedData = data.map(row => {
        const paddedRow = [...row];
        while (paddedRow.length < numCols) {
          paddedRow.push(''); // Pad with empty strings
        }
        return paddedRow;
      });

      const numRows = normalizedData.length;

      // Get worksheet
      const worksheet = context.workbook.worksheets.getItem(sheetName);

      // Get starting range and calculate target range
      const startRange = worksheet.getRange(startCell);
      startRange.load("rowIndex, columnIndex");
      await context.sync();

      // Create target range from start cell with data dimensions
      const targetRange = worksheet.getRangeByIndexes(
        startRange.rowIndex,
        startRange.columnIndex,
        numRows,
        numCols
      );

      targetRange.load("address");

      // If not overwriting, check for existing data
      if (!overwrite) {
        targetRange.load("values");
        await context.sync();

        const existingValues = targetRange.values;
        const newData = [];

        for (let i = 0; i < numRows; i++) {
          const row = [];
          for (let j = 0; j < numCols; j++) {
            // Only write if cell is empty
            if (existingValues[i][j] === "" || existingValues[i][j] === null) {
              row.push(normalizedData[i][j]);
            } else {
              row.push(existingValues[i][j]);
            }
          }
          newData.push(row);
        }

        targetRange.values = newData;
      } else {
        // Overwrite mode - write directly
        targetRange.values = normalizedData;
      }

      // Auto-fit columns and rows
      targetRange.format.autofitColumns();
      targetRange.format.autofitRows();

      await context.sync();

      const result = {
        sheetName: sheetName,
        rangeAddress: targetRange.address,
        rowsWritten: numRows,
        columnsWritten: numCols
      };

      console.log(`  ✅ Wrote ${numRows}x${numCols} data to ${targetRange.address}`);

      return result;

    } catch (error) {
      console.error(`  ❌ Error in writeDataToRange:`, error);
      throw {
        tool: "writeDataToRange",
        error: error.message,
        params: params
      };
    }
  });
}
