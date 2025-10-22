/**
 * queryBuilder.js
 *
 * Packages user query and builds complete context for backend API.
 * Combines initial context, user selection, and query into a single request payload.
 *
 * Key Features:
 * - Simple query object construction
 * - Optional intent detection (can be expanded later)
 * - Combines all context pieces
 * - Prepares final payload for backend
 */

/**
 * Builds a complete query object with all context
 *
 * @param {string} queryText - The user's query text
 * @param {Object} initialContext - Initial context from captureInitialContext()
 * @param {Object|null} userSelection - User selection from captureUserSelection()
 * @param {string|null} intent - Optional intent classification
 * @returns {Object} Complete query payload for backend
 */
export function buildQueryPayload(queryText, initialContext, userSelection = null, intent = null) {
  console.log("📦 Building query payload...");

  const payload = {
    // User's query
    query: {
      task: queryText,
      intent: intent || detectIntent(queryText), // Auto-detect if not provided
      timestamp: new Date().toISOString(),
    },

    // Initial context (always included)
    initialContext: initialContext,

    // User selection (optional - only if user selected something)
    userSelection: userSelection,
  };

  console.log("✅ Query payload built:", {
    queryLength: queryText.length,
    hasSelection: !!userSelection,
    intent: payload.query.intent,
    sheetsCount: initialContext?.sheets?.length || 0,
  });

  return payload;
}

/**
 * Simple intent detection based on keywords
 * This is a basic implementation - can be enhanced with ML later
 *
 * Intent types:
 * - analysis: User wants to analyze data
 * - transformation: User wants to modify/transform data
 * - visualization: User wants to create charts/visuals
 * - query: User wants to find/filter data
 * - calculation: User wants to perform calculations
 * - formatting: User wants to format cells
 * - unknown: Cannot determine intent
 *
 * @param {string} queryText - The user's query text
 * @returns {string} Detected intent
 */
export function detectIntent(queryText) {
  const text = queryText.toLowerCase();

  // Analysis keywords
  const analysisKeywords = [
    "analyze",
    "analysis",
    "summarize",
    "summary",
    "statistics",
    "trend",
    "correlation",
    "distribution",
    "average",
    "mean",
    "median",
    "insights",
  ];

  // Transformation keywords
  const transformationKeywords = [
    "transform",
    "convert",
    "change",
    "modify",
    "update",
    "replace",
    "merge",
    "split",
    "combine",
    "pivot",
    "unpivot",
    "transpose",
  ];

  // Visualization keywords
  const visualizationKeywords = [
    "chart",
    "graph",
    "plot",
    "visualize",
    "visualization",
    "dashboard",
    "diagram",
  ];

  // Query/Filter keywords
  const queryKeywords = [
    "find",
    "search",
    "filter",
    "where",
    "select",
    "show",
    "get",
    "lookup",
    "match",
  ];

  // Calculation keywords
  const calculationKeywords = [
    "calculate",
    "compute",
    "sum",
    "total",
    "count",
    "formula",
    "equation",
  ];

  // Formatting keywords
  const formattingKeywords = [
    "format",
    "style",
    "color",
    "highlight",
    "conditional formatting",
    "bold",
    "italic",
  ];

  // Check each category
  if (analysisKeywords.some((keyword) => text.includes(keyword))) {
    return "analysis";
  }

  if (transformationKeywords.some((keyword) => text.includes(keyword))) {
    return "transformation";
  }

  if (visualizationKeywords.some((keyword) => text.includes(keyword))) {
    return "visualization";
  }

  if (queryKeywords.some((keyword) => text.includes(keyword))) {
    return "query";
  }

  if (calculationKeywords.some((keyword) => text.includes(keyword))) {
    return "calculation";
  }

  if (formattingKeywords.some((keyword) => text.includes(keyword))) {
    return "formatting";
  }

  // Default
  return "unknown";
}

/**
 * Helper function to validate query payload before sending
 *
 * @param {Object} payload - Query payload to validate
 * @returns {Object} Validation result { isValid: boolean, errors: string[] }
 */
export function validateQueryPayload(payload) {
  const errors = [];

  // Check query exists and is not empty
  if (!payload.query || !payload.query.task || payload.query.task.trim() === "") {
    errors.push("Query task is required and cannot be empty");
  }

  // Check initial context exists
  if (!payload.initialContext) {
    errors.push("Initial context is required");
  }

  // Check initial context has required fields
  if (payload.initialContext) {
    if (!payload.initialContext.workbook) {
      errors.push("Initial context must include workbook metadata");
    }
    if (!payload.initialContext.sheets || !Array.isArray(payload.initialContext.sheets)) {
      errors.push("Initial context must include sheets array");
    }
  }

  return {
    isValid: errors.length === 0,
    errors: errors,
  };
}

/**
 * Pretty prints the query payload for debugging
 *
 * @param {Object} payload - Query payload to print
 */
export function debugPrintPayload(payload) {
  console.group("🔍 Query Payload Debug");

  console.log("📝 Query:", {
    task: payload.query?.task,
    intent: payload.query?.intent,
    timestamp: payload.query?.timestamp,
  });

  console.log("📊 Initial Context:", {
    workbookName: payload.initialContext?.workbook?.name,
    isDirty: payload.initialContext?.workbook?.isDirty,
    sheetCount: payload.initialContext?.sheets?.length,
    activeSheet: payload.initialContext?.activeSheet,
    activeCell: payload.initialContext?.activeCell,
    sheets: payload.initialContext?.sheets?.map((s) => ({
      name: s.name,
      size: `${s.rowCount}x${s.columnCount}`,
      hasFormulas: s.hasFormulas,
      // Support both old (formulas array) and new (formulaCount) structures
      formulasCount: s.formulas?.length || s.formulaCount || 0,
      hasCharts: s.hasCharts,
      // Support both old (tables array) and new (tableNames array) structures
      tables: s.tables?.length || s.tableNames?.length || 0,
    })),
  });

  // Log formulas separately if any exist
  // Support both old structure (detailed formulas) and new structure (formula counts)
  const sheetsWithFormulas = payload.initialContext?.sheets?.filter(s => 
    (s.formulas?.length > 0) || (s.hasFormulas && s.formulaCount > 0)
  ) || [];
  
  if (sheetsWithFormulas.length > 0) {
    console.group("📐 Formulas Detected:");
    sheetsWithFormulas.forEach(sheet => {
      // Old structure: detailed formula array
      if (sheet.formulas && sheet.formulas.length > 0) {
        console.group(`  Sheet: ${sheet.name} (${sheet.formulas.length} formulas)`);
        sheet.formulas.forEach((f, index) => {
          console.log(`    ${index + 1}. ${f.cell}: ${f.formula}`);
          console.log(`       Value: ${f.value}`);
          if (f.dependencies.functions.length > 0) {
            console.log(`       Functions: ${f.dependencies.functions.join(", ")}`);
          }
          if (f.dependencies.ranges.length > 0) {
            console.log(`       Ranges: ${f.dependencies.ranges.join(", ")}`);
          }
          if (f.dependencies.cells.length > 0) {
            console.log(`       Cells: ${f.dependencies.cells.join(", ")}`);
          }
          if (f.dependencies.sheets.length > 0) {
            console.log(`       Cross-sheet refs: ${f.dependencies.sheets.join(", ")}`);
          }
        });
        console.groupEnd();
      }
      // New structure: lightweight map with counts only
      else if (sheet.hasFormulas && sheet.formulaCount > 0) {
        console.log(`  Sheet: ${sheet.name} - ${sheet.formulaCount} formulas in ranges: ${sheet.formulaRanges?.join(', ') || 'N/A'}`);
      }
    });
    console.groupEnd();
  }

  if (payload.userSelection) {
    console.log("🎯 User Selection:", {
      isMultipleAreas: payload.userSelection?.isMultipleAreas,
      regionCount: payload.userSelection?.regions?.length,
      regions: payload.userSelection?.regions?.map((r) => ({
        sheet: r.sheetName,
        address: r.address,
        size: `${r.rowCount}x${r.columnCount}`,
        previewRows: r.preview?.length,
      })),
    });
  } else {
    console.log("🎯 User Selection: None");
  }

  const validation = validateQueryPayload(payload);
  console.log("✓ Validation:", validation);

  console.groupEnd();
}
