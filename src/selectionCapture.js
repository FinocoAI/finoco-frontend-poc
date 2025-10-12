/**
 * selectionCapture.js
 *
 * Captures User Selection Context from Excel.
 * This includes any ranges the user has highlighted before making a query.
 *
 * Key Features:
 * - Captures selections from the active sheet
 * - Smart data capture strategy:
 *   • Small selections (≤100 rows): Sends FULL data immediately (no tool call needed)
 *   • Large selections (>100 rows): Sends PREVIEW only (agent calls getFullRangeData if needed)
 * - Supports multiple non-contiguous selections (Ctrl+Click)
 * - Balances immediate context vs. token usage
 * - Returns null if nothing is selected
 * - Uses 'isComplete' flag to indicate if agent has all data
 *
 * Note on Cross-Sheet Selections:
 * Excel clears selections when you switch sheets. Therefore, we can only capture
 * selections from the currently active sheet. However, multiple non-contiguous
 * areas on the same sheet are fully supported (e.g., A1:B5 and D1:E5).
 */

/**
 * Configuration for data capture strategy
 */
const PREVIEW_ROW_LIMIT = 5;           // Preview rows for large selections
const FULL_DATA_THRESHOLD = 100;       // Send full data if <= this many rows

/**
 * Main function to capture user selection context
 * Captures selections from the active sheet (including multiple non-contiguous areas)
 *
 * Note: Excel doesn't maintain selections across sheets - when you switch sheets,
 * the previous sheet's selection is cleared. Therefore we can only capture
 * selections from the currently active sheet.
 *
 * @returns {Promise<Object|null>} User selection object or null if nothing selected
 */
export async function captureUserSelection() {
  return Excel.run(async (context) => {
    console.log("🎯 Starting User Selection Capture...");

    try {
      // Get all selected ranges from the active sheet (supports multiple selections)
      const selectedRanges = context.workbook.getSelectedRanges();
      selectedRanges.load("areas, areaCount");
      await context.sync();

      // If nothing selected, return null
      if (selectedRanges.areaCount === 0) {
        console.log("  ℹ️ No selection found");
        return null;
      }

      const areas = selectedRanges.areas;
      areas.load("items");
      await context.sync();

      const userSelection = {
        isMultipleAreas: selectedRanges.areaCount > 1,
        regions: [],
      };

      let totalRegions = 0;

      // Process each selected area
      for (let i = 0; i < areas.items.length; i++) {
        const area = areas.items[i];
        area.load("rowCount, columnCount");
        await context.sync();

        // Skip single cell selections
        if (area.rowCount === 1 && area.columnCount === 1) {
          console.log(`  ℹ️ Skipping single cell selection`);
          continue;
        }

        const regionData = await captureRegionPreview(context, area);
        if (regionData) {
          userSelection.regions.push(regionData);
          totalRegions++;
        }
      }

      // If no meaningful selections found (only single cells), return null
      if (totalRegions === 0) {
        console.log("  ℹ️ Only single cell selected, not capturing");
        return null;
      }

      // Log summary
      const uniqueSheets = [...new Set(userSelection.regions.map(r => r.sheetName))];
      console.log("✅ User Selection Captured:", userSelection);
      console.log(`   📊 ${totalRegions} region(s) from: ${uniqueSheets.join(", ")}`);

      return userSelection;

    } catch (error) {
      console.error("❌ Error capturing user selection:", error);
      return null;
    }
  });
}

/**
 * Captures data for a single selected region
 * Strategy:
 * - Small selections (≤ FULL_DATA_THRESHOLD rows): Send all data immediately
 * - Large selections (> FULL_DATA_THRESHOLD rows): Send preview only (agent calls getFullRangeData if needed)
 *
 * @param {Excel.RequestContext} context - Excel request context
 * @param {Excel.Range} range - Range to capture
 * @returns {Promise<Object|null>} Region data object or null if error
 */
async function captureRegionPreview(context, range) {
  try {
    // Load basic range properties
    range.load("address, rowCount, columnCount");
    const worksheet = range.worksheet;
    worksheet.load("name");
    await context.sync();

    console.log(`  📍 Processing region: ${range.address} (${range.rowCount}x${range.columnCount})`);

    const regionData = {
      sheetName: worksheet.name,
      address: range.address,
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      headers: [],
    };

    // Decide strategy: full data vs preview
    const sendFullData = range.rowCount <= FULL_DATA_THRESHOLD;

    if (sendFullData) {
      // Small selection: send ALL data
      range.load("values");
      await context.sync();
      
      regionData.data = range.values;  // Full data
      regionData.isComplete = true;    // Flag: agent has all data
      
      // Extract headers from first row
      if (range.rowCount > 0) {
        regionData.headers = range.values[0].map(val => String(val));
      }
      
      console.log(`    ✓ Region captured: FULL data (${range.rowCount} rows)`);
      
    } else {
      // Large selection: send preview only
      const previewRowCount = Math.min(PREVIEW_ROW_LIMIT, range.rowCount);
      
      if (previewRowCount > 0) {
        const previewRange = range.getAbsoluteResizedRange(previewRowCount, range.columnCount);
        previewRange.load("values");
        await context.sync();
        
        regionData.preview = previewRange.values;  // Preview only
        regionData.isComplete = false;             // Flag: agent needs to call getFullRangeData
        
        // Extract headers from preview
        regionData.headers = previewRange.values[0].map(val => String(val));
      }
      
      console.log(`    ✓ Region captured: PREVIEW only (${previewRowCount}/${range.rowCount} rows)`);
    }

    return regionData;

  } catch (error) {
    console.error(`    ❌ Error capturing region:`, error);
    return null;
  }
}

/**
 * Helper function to check if user has a meaningful selection
 * Checks the active sheet for selections
 *
 * @returns {Promise<boolean>} True if user has selected more than a single cell
 */
export async function hasUserSelection() {
  return Excel.run(async (context) => {
    try {
      const selectedRanges = context.workbook.getSelectedRanges();
      selectedRanges.load("areaCount");
      await context.sync();

      if (selectedRanges.areaCount === 0) {
        return false;
      }

      const areas = selectedRanges.areas;
      areas.load("items");
      await context.sync();

      // Check if any area is more than a single cell
      for (let i = 0; i < areas.items.length; i++) {
        const area = areas.items[i];
        area.load("rowCount, columnCount");
        await context.sync();

        if (!(area.rowCount === 1 && area.columnCount === 1)) {
          return true;
        }
      }

      return false;
    } catch (error) {
      return false;
    }
  });
}
