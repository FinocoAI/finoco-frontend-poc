/**
 * getRangePreview.js
 *
 * Tool: Fetch a preview/sample of rows from a large range
 * Executor: Frontend (requires Office.js)
 *
 * Use Cases:
 * - Preview large datasets without loading all rows
 * - Understand data structure before full analysis
 * - Sample data for quick insights
 * - Progressive loading for huge financial models
 */

/**
 * Tool definition for LLM/Backend
 */
export const toolDefinition = {
  name: "getRangePreview",
  description: "Fetch a preview/sample of rows from a range without loading all data. Perfect for exploring large datasets. Returns sample rows, headers, and metadata.",
  executor: "frontend",
  requiresContinuation: true,  // READ tool - results must be sent back to agent
  parameters: {
    sheetName: {
      type: "string",
      required: true,
      description: "Name of the sheet containing the range"
    },
    address: {
      type: "string",
      required: true,
      description: "Range address in A1 notation (e.g., 'A1:K10000'). DO NOT include sheet name."
    },
    previewRows: {
      type: "number",
      required: false,
      default: 10,
      description: "Number of rows to include in preview (default: 10, max: 100)"
    },
    sampleStrategy: {
      type: "string",
      required: false,
      default: "first",
      description: "Sampling strategy: 'first' (default), 'evenly_spaced', or 'random'"
    },
    includeFormulas: {
      type: "boolean",
      required: false,
      default: false,
      description: "Include formulas in addition to values"
    }
  },
  returns: {
    sheetName: "string",
    fullAddress: "string",
    totalRows: "number",
    totalColumns: "number",
    previewRows: "number",
    sampleStrategy: "string",
    preview: "array[][]",
    formulas: "array[][]",
    headers: "array",
    hasMore: "boolean",
    dataTypes: "object",
    sampleIndices: "array"
  }
};

/**
 * Helper function to infer data types from preview
 */
function inferDataTypes(preview, headers) {
  const types = {};
  
  if (preview.length < 2) {
    return types;
  }

  // Skip header row, analyze first few data rows
  const dataRows = preview.slice(1, Math.min(6, preview.length));
  
  headers.forEach((header, colIndex) => {
    const values = dataRows.map(row => row[colIndex]);
    
    // Count types
    let numberCount = 0;
    let stringCount = 0;
    let dateCount = 0;
    let boolCount = 0;
    let emptyCount = 0;

    values.forEach(val => {
      if (val === null || val === undefined || val === "") {
        emptyCount++;
      } else if (typeof val === "number") {
        numberCount++;
      } else if (typeof val === "boolean") {
        boolCount++;
      } else if (typeof val === "string") {
        // Check if it looks like a date
        if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(val) || /^\d{4}-\d{2}-\d{2}/.test(val)) {
          dateCount++;
        } else {
          stringCount++;
        }
      }
    });

    // Determine predominant type
    if (numberCount > values.length / 2) {
      types[header] = "number";
    } else if (boolCount > values.length / 2) {
      types[header] = "boolean";
    } else if (dateCount > values.length / 2) {
      types[header] = "date";
    } else if (emptyCount === values.length) {
      types[header] = "empty";
    } else {
      types[header] = "string";
    }
  });

  return types;
}

/**
 * Get row indices based on sampling strategy
 */
function getSampleIndices(totalRows, previewRows, strategy, hasHeader) {
  const indices = [];
  
  // Always include header if it exists
  if (hasHeader) {
    indices.push(0);
    previewRows = Math.min(previewRows, totalRows);
  }

  const startRow = hasHeader ? 1 : 0;
  const availableRows = totalRows - startRow;
  const rowsToSample = Math.min(previewRows - (hasHeader ? 1 : 0), availableRows);

  if (strategy === "evenly_spaced") {
    // Sample evenly across the range
    const step = Math.max(1, Math.floor(availableRows / rowsToSample));
    for (let i = 0; i < rowsToSample; i++) {
      const idx = startRow + i * step;
      if (idx < totalRows) {
        indices.push(idx);
      }
    }
  } else if (strategy === "random") {
    // Random sampling
    const availableIndices = Array.from({ length: availableRows }, (_, i) => startRow + i);
    for (let i = 0; i < rowsToSample; i++) {
      const randomIdx = Math.floor(Math.random() * availableIndices.length);
      indices.push(availableIndices[randomIdx]);
      availableIndices.splice(randomIdx, 1);
    }
    indices.sort((a, b) => a - b);
  } else {
    // Default: first rows
    for (let i = startRow; i < Math.min(startRow + rowsToSample, totalRows); i++) {
      indices.push(i);
    }
  }

  return indices;
}

/**
 * Execute the tool
 *
 * @param {Object} params - Tool parameters
 * @param {string} params.sheetName - Sheet name
 * @param {string} params.address - Range address (e.g., "A1:K10000")
 * @param {number} [params.previewRows=10] - Number of preview rows
 * @param {string} [params.sampleStrategy="first"] - Sampling strategy
 * @param {boolean} [params.includeFormulas=false] - Include formulas
 * @returns {Promise<Object>} Range preview object
 */
export async function execute(params) {
  console.log(`🔧 Executing getRangePreview:`, params);

  return Excel.run(async (context) => {
    try {
      const {
        sheetName,
        address,
        previewRows = 10,
        sampleStrategy = "first",
        includeFormulas = false
      } = params;

      // Validate parameters
      if (!sheetName || typeof sheetName !== 'string') {
        throw new Error('Invalid sheetName: must be a non-empty string');
      }
      if (!address || typeof address !== 'string') {
        throw new Error('Invalid address: must be a valid range address');
      }

      // Limit preview rows
      const limitedPreviewRows = Math.min(Math.max(1, previewRows), 100);

      // Get the worksheet
      const worksheet = context.workbook.worksheets.getItem(sheetName);

      // Get the full range
      const fullRange = worksheet.getRange(address);
      fullRange.load("address, rowCount, columnCount");
      await context.sync();

      const totalRows = fullRange.rowCount;
      const totalColumns = fullRange.columnCount;

      console.log(`  📊 Full range: ${totalRows}x${totalColumns}`);

      // Determine if first row is header (simple heuristic)
      const hasHeader = true;  // Assume first row is header

      // Get sample indices
      const sampleIndices = getSampleIndices(totalRows, limitedPreviewRows, sampleStrategy, hasHeader);
      
      console.log(`  📋 Sampling ${sampleIndices.length} rows using "${sampleStrategy}" strategy`);

      // Fetch sampled rows
      const preview = [];
      const formulas = [];

      for (const rowIndex of sampleIndices) {
        const row = fullRange.getRow(rowIndex);
        row.load("values");
        if (includeFormulas) {
          row.load("formulas");
        }
        await context.sync();

        preview.push(row.values[0]);
        if (includeFormulas) {
          formulas.push(row.formulas[0]);
        }
      }

      // Extract headers (first row)
      const headers = preview.length > 0 ? preview[0].map(val => String(val)) : [];

      // Infer data types
      const dataTypes = inferDataTypes(preview, headers);

      const result = {
        sheetName: sheetName,
        fullAddress: fullRange.address,
        totalRows: totalRows,
        totalColumns: totalColumns,
        previewRows: preview.length,
        sampleStrategy: sampleStrategy,
        preview: preview,
        headers: headers,
        hasMore: totalRows > preview.length,
        dataTypes: dataTypes,
        sampleIndices: sampleIndices
      };

      if (includeFormulas && formulas.length > 0) {
        result.formulas = formulas;
      }

      console.log(`  ✅ Retrieved preview: ${preview.length}/${totalRows} rows`);

      return result;

    } catch (error) {
      console.error(`  ❌ Error in getRangePreview:`, error);
      throw {
        tool: "getRangePreview",
        error: error.message,
        params: params
      };
    }
  });
}

