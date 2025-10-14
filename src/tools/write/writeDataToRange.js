/**
 * writeDataToRange.js
 *
 * Tool: Write data to a specific range
 * Executor: Frontend (requires Office.js)
 */

import { showOverwriteModal } from '../../ui/overwriteModal.js';

export const toolDefinition = {
  name: "writeDataToRange",
  description: "Write 2D array data (with optional formulas) to a specific range in Excel. Automatically detects and applies formulas (strings starting with '=').",
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
      description: "2D array where each cell can be a static value (string/number) or a formula (string starting with '='). Formulas are automatically detected and applied."
    },
    overwrite: {
      type: "boolean",
      required: false,
      default: true,
      description: "If false, will only write to empty cells"
    },
    headerFormat: {
      type: "object",
      required: false,
      description: "Optional formatting to apply to the first row (headers). Same properties as formatRange: fontBold, fontSize, fontColor, fillColor, horizontalAlignment, etc."
    }
  },
  returns: {
    sheetName: "string",
    rangeAddress: "string",
    rowsWritten: "number",
    columnsWritten: "number",
    formulasApplied: "number",
    headerFormatted: "boolean"
  }
};

export async function execute(params) {
  console.log(`🔧 Executing writeDataToRange:`, params);

  return Excel.run(async (context) => {
    try {
      const { sheetName, startCell, data, overwrite = true, headerFormat } = params;

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

      targetRange.load("address, values");
      await context.sync();

      // =========================================================================
      // OVERWRITE PROTECTION: Check for existing data and request user approval
      // =========================================================================
      const existingValues = targetRange.values;
      let cellsWithData = [];
      let hasExistingData = false;

      // Scan for non-empty cells
      for (let i = 0; i < numRows; i++) {
        for (let j = 0; j < numCols; j++) {
          const cellValue = existingValues[i][j];
          if (cellValue !== "" && cellValue !== null && cellValue !== undefined) {
            hasExistingData = true;
            // Calculate cell address (e.g., A1, B2)
            const colLetter = String.fromCharCode(65 + startRange.columnIndex + j);
            const rowNum = startRange.rowIndex + i + 1;
            cellsWithData.push(`${colLetter}${rowNum}`);
          }
        }
      }

      // If existing data found, request user approval
      if (hasExistingData && overwrite !== false) {
        console.log(`⚠️ Overwrite protection: ${cellsWithData.length} cells contain data`);
        
        // Show approval modal (this is async and blocks execution)
        const userApproved = await showOverwriteModal(
          targetRange.address,
          cellsWithData.length,
          cellsWithData
        );

        if (!userApproved) {
          // User cancelled - throw error with feedback for agent
          console.log('❌ User denied overwrite operation');
          throw {
            tool: "writeDataToRange",
            userCancelled: true,
            sheetName: sheetName,
            rangeAddress: targetRange.address,
            cellsAffected: cellsWithData.length,
            message: `User cancelled: ${cellsWithData.length} cells would be overwritten in ${targetRange.address}`,
            feedback: `The user declined to overwrite ${cellsWithData.length} existing cells in range ${targetRange.address}. The data was NOT written. Consider asking for a different location or approach.`
          };
        }

        console.log('✅ User approved overwrite operation');
      }

      // =========================================================================
      // Proceed with write operation
      // =========================================================================

      // Separate data into values and formulas
      const formulaArray = [];
      const valueArray = [];
      let formulaCount = 0;

      for (let i = 0; i < numRows; i++) {
        const formulaRow = [];
        const valueRow = [];
        
        for (let j = 0; j < numCols; j++) {
          const cellData = normalizedData[i][j];
          
          // Check if it's a formula (string starting with '=')
          if (typeof cellData === 'string' && cellData.startsWith('=')) {
            formulaRow.push(cellData);
            valueRow.push(''); // Placeholder for formula cells
            formulaCount++;
          } else {
            formulaRow.push('');
            valueRow.push(cellData === null || cellData === undefined ? '' : cellData);
          }
        }
        
        formulaArray.push(formulaRow);
        valueArray.push(valueRow);
      }

      // If not overwriting, check for existing data
      if (!overwrite) {
        targetRange.load("values");
        await context.sync();

        const existingValues = targetRange.values;
        const mergedValues = [];

        for (let i = 0; i < numRows; i++) {
          const row = [];
          for (let j = 0; j < numCols; j++) {
            // Only write if cell is empty
            if (existingValues[i][j] === "" || existingValues[i][j] === null) {
              row.push(valueArray[i][j]);
            } else {
              row.push(existingValues[i][j]);
            }
          }
          mergedValues.push(row);
        }

        targetRange.values = mergedValues;
      } else {
        // Overwrite mode - write values first
        targetRange.values = valueArray;
      }

      // Apply formulas (this works in both overwrite and non-overwrite mode)
      if (formulaCount > 0) {
        targetRange.formulas = formulaArray;
      }

      // Critical: Sync and allow Excel to calculate formulas BEFORE any formatting
      await context.sync();
      
      // Force Excel to recalculate (ensures formulas are computed before formatting)
      if (formulaCount > 0) {
        context.workbook.application.calculate(Excel.CalculationType.recalculate);
        await context.sync();
      }

      // Auto-fit columns and rows
      targetRange.format.autofitColumns();
      targetRange.format.autofitRows();

      await context.sync();

      // Apply header formatting if provided
      let headerFormatted = false;
      if (headerFormat && numRows > 0) {
        const headerRange = worksheet.getRangeByIndexes(
          startRange.rowIndex,
          startRange.columnIndex,
          1,  // Just first row
          numCols
        );

        // Apply header formatting
        if (headerFormat.numberFormat) {
          headerRange.numberFormat = headerFormat.numberFormat;
        }
        if (headerFormat.fontBold !== undefined) {
          headerRange.format.font.bold = headerFormat.fontBold;
        }
        if (headerFormat.fontItalic !== undefined) {
          headerRange.format.font.italic = headerFormat.fontItalic;
        }
        if (headerFormat.fontSize) {
          headerRange.format.font.size = headerFormat.fontSize;
        }
        if (headerFormat.fontColor) {
          headerRange.format.font.color = headerFormat.fontColor;
        }
        if (headerFormat.fillColor) {
          headerRange.format.fill.color = headerFormat.fillColor;
        }
        if (headerFormat.horizontalAlignment) {
          // Excel API requires lowercase alignment values
          headerRange.format.horizontalAlignment = headerFormat.horizontalAlignment.toLowerCase();
        }
        if (headerFormat.verticalAlignment) {
          // Excel API requires lowercase alignment values
          headerRange.format.verticalAlignment = headerFormat.verticalAlignment.toLowerCase();
        }
        if (headerFormat.borders) {
          const borderTypes = ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'];
          borderTypes.forEach(type => {
            const borderValue = headerFormat.borders[type.toLowerCase()];
            if (borderValue) {
              // If boolean true, use "Continuous" style; otherwise use the provided style string
              const borderStyle = typeof borderValue === 'boolean' ? 'Continuous' : borderValue;
              headerRange.format.borders.getItem(type).style = borderStyle;
            }
          });
        }

        await context.sync();
        headerFormatted = true;
        console.log(`  ✅ Applied header formatting to first row`);
      }

      const result = {
        sheetName: sheetName,
        rangeAddress: targetRange.address,
        rowsWritten: numRows,
        columnsWritten: numCols,
        formulasApplied: formulaCount,
        headerFormatted: headerFormatted
      };

      console.log(`  ✅ Wrote ${numRows}x${numCols} data to ${targetRange.address}${formulaCount > 0 ? ` (${formulaCount} formulas included)` : ''}`);

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
