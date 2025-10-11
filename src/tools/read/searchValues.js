/**
 * searchValues.js
 *
 * Tool: Search for values matching criteria in a sheet
 * Executor: Frontend (requires Office.js)
 */

export const toolDefinition = {
  name: "searchValues",
  description: "Search for values matching criteria in a specific column",
  executor: "frontend",
  parameters: {
    sheetName: {
      type: "string",
      required: true,
      description: "Name of the sheet to search"
    },
    column: {
      type: "string",
      required: true,
      description: "Column letter to search (e.g., 'A', 'B')"
    },
    operator: {
      type: "string",
      required: true,
      description: "Comparison operator: 'equals', 'contains', 'greater', 'less', 'greaterOrEqual', 'lessOrEqual'"
    },
    value: {
      type: "any",
      required: true,
      description: "Value to search for"
    },
    returnFullRow: {
      type: "boolean",
      required: false,
      default: false,
      description: "If true, returns entire row data for matches"
    }
  },
  returns: {
    sheetName: "string",
    column: "string",
    matchCount: "number",
    matches: "array"  // Array of {row, value} or full row data if returnFullRow=true
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

/**
 * Check if value matches criteria
 */
function matchesCriteria(cellValue, operator, searchValue) {
  switch (operator) {
    case 'equals':
      return cellValue == searchValue;
    case 'contains':
      return String(cellValue).toLowerCase().includes(String(searchValue).toLowerCase());
    case 'greater':
      return Number(cellValue) > Number(searchValue);
    case 'less':
      return Number(cellValue) < Number(searchValue);
    case 'greaterOrEqual':
      return Number(cellValue) >= Number(searchValue);
    case 'lessOrEqual':
      return Number(cellValue) <= Number(searchValue);
    default:
      throw new Error(`Unknown operator: ${operator}`);
  }
}

export async function execute(params) {
  console.log(`🔧 Executing searchValues:`, params);

  return Excel.run(async (context) => {
    try {
      const { sheetName, column, operator, value, returnFullRow = false } = params;

      // Validate parameters
      if (!sheetName || typeof sheetName !== 'string') {
        throw new Error('Invalid sheetName');
      }
      if (!column || typeof column !== 'string') {
        throw new Error('Invalid column');
      }
      if (!operator || typeof operator !== 'string') {
        throw new Error('Invalid operator');
      }

      const validOperators = ['equals', 'contains', 'greater', 'less', 'greaterOrEqual', 'lessOrEqual'];
      if (!validOperators.includes(operator)) {
        throw new Error(`Invalid operator. Must be one of: ${validOperators.join(', ')}`);
      }

      // Get worksheet
      const worksheet = context.workbook.worksheets.getItem(sheetName);
      const usedRange = worksheet.getUsedRange();
      usedRange.load("values, rowCount, columnCount");
      await context.sync();

      // Get column index
      const colIndex = columnLetterToIndex(column.toUpperCase());
      if (colIndex < 0 || colIndex >= usedRange.columnCount) {
        throw new Error(`Column ${column} is out of range`);
      }

      // Search through column
      const matches = [];
      const values = usedRange.values;

      for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
        const cellValue = values[rowIndex][colIndex];

        if (matchesCriteria(cellValue, operator, value)) {
          if (returnFullRow) {
            matches.push({
              row: rowIndex + 1,  // 1-based row number
              data: values[rowIndex]
            });
          } else {
            matches.push({
              row: rowIndex + 1,
              value: cellValue
            });
          }
        }
      }

      const result = {
        sheetName: sheetName,
        column: column,
        operator: operator,
        searchValue: value,
        matchCount: matches.length,
        matches: matches
      };

      console.log(`  ✅ Found ${matches.length} match(es) in column ${column}`);

      return result;

    } catch (error) {
      console.error(`  ❌ Error in searchValues:`, error);
      throw {
        tool: "searchValues",
        error: error.message,
        params: params
      };
    }
  });
}
