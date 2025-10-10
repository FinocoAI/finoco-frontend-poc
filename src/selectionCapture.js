/**
 * selectionCapture.js
 *
 * Captures User Selection Context from Excel.
 * This includes any ranges the user has highlighted before making a query.
 *
 * Key Features:
 * - Captures selections from the active sheet
 * - Captures only PREVIEW data (first 5 rows)
 * - Supports multiple non-contiguous selections (Ctrl+Click)
 * - Keeps token usage low
 * - Returns null if nothing is selected
 * - Gives agent enough context to decide if it needs full data
 *
 * Note on Cross-Sheet Selections:
 * Excel clears selections when you switch sheets. Therefore, we can only capture
 * selections from the currently active sheet. However, multiple non-contiguous
 * areas on the same sheet are fully supported (e.g., A1:B5 and D1:E5).
 */

/**
 * Maximum number of preview rows to capture
 * This keeps token usage low while providing enough context
 */
const PREVIEW_ROW_LIMIT = 5;

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
 * Captures preview data for a single selected region
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
      preview: [],
    };

    // Get first row as headers
    if (range.rowCount > 0) {
      const headerRow = range.getRow(0);
      headerRow.load("values");
      await context.sync();
      regionData.headers = headerRow.values[0].map(val => String(val));
    }

    // Get preview data (first N rows)
    const previewRowCount = Math.min(PREVIEW_ROW_LIMIT, range.rowCount);

    if (previewRowCount > 0) {
      const previewRange = range.getAbsoluteResizedRange(previewRowCount, range.columnCount);
      previewRange.load("values");
      await context.sync();
      regionData.preview = previewRange.values;
    }

    console.log(`    ✓ Region captured: ${previewRowCount} preview rows`);
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
