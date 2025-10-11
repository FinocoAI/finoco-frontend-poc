/**
 * getRelatedData.js
 *
 * Tool: Follow relationships - find data in another sheet based on lookup value
 * Similar to VLOOKUP but more flexible
 * Executor: Frontend (requires Office.js)
 */

export const toolDefinition = {
  name: "getRelatedData",
  description: "Follow relationships - find data in another sheet based on lookup value (like VLOOKUP)",
  executor: "frontend",
  parameters: {
    sourceSheet: {
      type: "string",
      required: true,
      description: "Sheet containing the lookup value"
    },
    lookupColumn: {
      type: "string",
      required: true,
      description: "Column letter to search for the lookup value (e.g., 'A')"
    },
    lookupValue: {
      type: "any",
      required: true,
      description: "Value to search for"
    },
    targetSheet: {
      type: "string",
      required: true,
      description: "Sheet to return data from"
    },
    returnColumns: {
      type: "array",
      required: true,
      description: "Array of column letters to return (e.g., ['B', 'C', 'D'])"
    }
  },
  returns: {
    sourceSheet: "string",
    targetSheet: "string",
    lookupValue: "any",
    matchCount: "number",
    matches: "array"  // Array of matching rows with requested columns
  }
};

/**
 * Convert column letter to zero-based index
 */
function columnLetterToIndex(letter) {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index - 1;
}

export async function execute(params) {
  console.log(`🔧 Executing getRelatedData:`, params);

  return Excel.run(async (context) => {
    try {
      const { sourceSheet, lookupColumn, lookupValue, targetSheet, returnColumns } = params;

      // Validate parameters
      if (!sourceSheet || typeof sourceSheet !== 'string') {
        throw new Error('Invalid sourceSheet');
      }
      if (!lookupColumn || typeof lookupColumn !== 'string') {
        throw new Error('Invalid lookupColumn');
      }
      if (!targetSheet || typeof targetSheet !== 'string') {
        throw new Error('Invalid targetSheet');
      }
      if (!Array.isArray(returnColumns) || returnColumns.length === 0) {
        throw new Error('Invalid returnColumns: must be a non-empty array');
      }

      // Get source sheet and used range
      const worksheet = context.workbook.worksheets.getItem(sourceSheet);
      const usedRange = worksheet.getUsedRange();
      usedRange.load("values, rowCount, columnCount");
      await context.sync();

      // Get lookup column index
      const lookupColIndex = columnLetterToIndex(lookupColumn.toUpperCase());
      if (lookupColIndex < 0 || lookupColIndex >= usedRange.columnCount) {
        throw new Error(`Lookup column ${lookupColumn} is out of range`);
      }

      // Search for matching rows
      const matches = [];
      const values = usedRange.values;

      for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
        const cellValue = values[rowIndex][lookupColIndex];

        if (cellValue == lookupValue) {
          // Found a match, get data from target sheet
          const targetData = await getTargetSheetData(
            context,
            targetSheet,
            rowIndex,
            returnColumns
          );

          matches.push({
            sourceRow: rowIndex + 1,
            targetRow: rowIndex + 1,
            data: targetData
          });
        }
      }

      const result = {
        sourceSheet: sourceSheet,
        targetSheet: targetSheet,
        lookupColumn: lookupColumn,
        lookupValue: lookupValue,
        returnColumns: returnColumns,
        matchCount: matches.length,
        matches: matches
      };

      console.log(`  ✅ Found ${matches.length} match(es) for value "${lookupValue}"`);

      return result;

    } catch (error) {
      console.error(`  ❌ Error in getRelatedData:`, error);
      throw {
        tool: "getRelatedData",
        error: error.message,
        params: params
      };
    }
  });
}

/**
 * Get data from target sheet for specific row and columns
 */
async function getTargetSheetData(context, targetSheetName, rowIndex, columnLetters) {
  const targetSheet = context.workbook.worksheets.getItem(targetSheetName);
  const data = {};

  for (const colLetter of columnLetters) {
    const colIndex = columnLetterToIndex(colLetter.toUpperCase());

    // Get cell at row, column
    const cell = targetSheet.getRangeByIndexes(rowIndex, colIndex, 1, 1);
    cell.load("values");
    await context.sync();

    data[colLetter] = cell.values[0][0];
  }

  return data;
}
