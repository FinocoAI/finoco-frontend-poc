/**
 * Context Manager Module
 * Handles Excel context capture and management
 */

import { captureInitialContext } from '../contextCapture.js';
import { captureUserSelection } from '../selectionCapture.js';
import { updateContextDisplay } from '../ui/contextUI.js';

// Legacy context structure (for UI display)
export let excelContext = {
    workbookName: null,
    sheetName: null,
    allSheets: [],
    selectedRange: null,
    selectedData: null,
    headers: null,
    selectedRanges: [],
    isMultipleAreas: false
};

// Enhanced context structure (as per architecture)
export let enhancedContext = {
    initialContext: null,    // From captureInitialContext()
    userSelection: null,     // From captureUserSelection()
};

/**
 * Update enhanced context using new architecture
 */
export async function updateEnhancedContext() {
    try {
        console.log('🔄 Updating enhanced context...');

        // Capture initial context (always)
        enhancedContext.initialContext = await captureInitialContext();

        // Capture user selection (always try - it returns null if nothing selected)
        enhancedContext.userSelection = await captureUserSelection();

        console.log('✅ Enhanced context updated');
    } catch (error) {
        console.error('❌ Error updating enhanced context:', error);
    }
}

/**
 * Update Excel context (lightweight UI context)
 */
export async function updateExcelContext() {
    try {
        await Excel.run(async (context) => {
            const workbook = context.workbook;
            const worksheets = workbook.worksheets;
            const activeSheet = worksheets.getActiveWorksheet();

            workbook.load('name');
            activeSheet.load('name');
            worksheets.load('items/name');

            await context.sync();

            // Update basic context
            excelContext.workbookName = workbook.name;
            excelContext.sheetName = activeSheet.name;
            excelContext.allSheets = worksheets.items.map(sheet => sheet.name);

            // Use getSelectedRanges() which returns RangeAreas (supports multiple selections)
            const selectedRanges = context.workbook.getSelectedRanges();
            selectedRanges.load('address, areaCount');
            await context.sync();

            // Clean up the address
            let cleanAddress = selectedRanges.address;
            if (cleanAddress.includes('!')) {
                cleanAddress = cleanAddress.split('!')[1];
            }

            const areaCount = selectedRanges.areaCount;
            excelContext.isMultipleAreas = areaCount > 1;

            console.log('Selection has', areaCount, 'area(s)');

            if (excelContext.isMultipleAreas) {
                // Handle multiple areas using RangeAreas
                const areas = selectedRanges.areas;
                areas.load('items');
                await context.sync();

                excelContext.selectedRanges = [];
                let allAddresses = [];

                for (let i = 0; i < areas.items.length; i++) {
                    const area = areas.items[i];
                    area.load('address, values, rowCount, columnCount');
                    await context.sync();

                    // Clean up the address
                    let areaAddress = area.address;
                    if (areaAddress.includes('!')) {
                        areaAddress = areaAddress.split('!')[1];
                    }

                    allAddresses.push(areaAddress);

                    // Detect headers for this area
                    let areaHeaders = null;
                    if (area.rowCount > 1) {
                        areaHeaders = area.values[0];
                    }

                    excelContext.selectedRanges.push({
                        address: areaAddress,
                        data: area.values,
                        headers: areaHeaders,
                        rowCount: area.rowCount,
                        columnCount: area.columnCount
                    });

                    console.log(`  - Area ${i + 1}: ${areaAddress} (${area.rowCount}x${area.columnCount})`);
                }

                // Set combined address for display
                excelContext.selectedRange = cleanAddress;
                // For backward compatibility, use first area's data
                excelContext.selectedData = excelContext.selectedRanges[0].data;
                excelContext.headers = excelContext.selectedRanges[0].headers;

            } else {
                // Single area selection
                const areas = selectedRanges.areas;
                areas.load('items');
                await context.sync();

                const singleArea = areas.items[0];
                singleArea.load('address, values, rowCount, columnCount');
                await context.sync();

                excelContext.selectedRange = cleanAddress;
                excelContext.selectedData = singleArea.values;

                // Try to detect headers (first row of selection)
                if (singleArea.rowCount > 1) {
                    excelContext.headers = singleArea.values[0];
                } else {
                    excelContext.headers = null;
                }

                // Store as single range
                excelContext.selectedRanges = [{
                    address: cleanAddress,
                    data: singleArea.values,
                    headers: excelContext.headers,
                    rowCount: singleArea.rowCount,
                    columnCount: singleArea.columnCount
                }];

                console.log('Excel context updated:');
                console.log('  - Range:', cleanAddress);
                console.log('  - Rows:', singleArea.rowCount);
                console.log('  - Cols:', singleArea.columnCount);
                console.log('  - Data sample:', singleArea.values.slice(0, 2));
            }

            // Update UI
            updateContextDisplay(excelContext);
        });
    } catch (error) {
        console.error('Error updating Excel context:', error);
        updateContextDisplay(excelContext);
    }
}

/**
 * Set up Excel event listeners
 */
export async function setupExcelEventListeners() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();

            // Listen to selection changes - lightweight UI update only
            sheet.onSelectionChanged.add(async () => {
                await updateExcelContext();
            });

            // Listen to sheet activation - lightweight UI update only
            context.workbook.worksheets.onActivated.add(async () => {
                await updateExcelContext();
            });

            await context.sync();
            console.log('Excel event listeners set up (lightweight mode)');
        });
    } catch (error) {
        console.error('Error setting up Excel event listeners:', error);
    }
}

