/**
 * Context UI Module
 * Handles context display in the sidebar
 */

/**
 * Update context display in UI
 */
export function updateContextDisplay(excelContext) {
    document.getElementById('workbookName').textContent = excelContext.workbookName || '-';
    document.getElementById('sheetName').textContent = excelContext.sheetName || '-';

    const selectionElement = document.getElementById('selectionRange');
    if (excelContext.isMultipleAreas && excelContext.selectedRanges.length > 1) {
        // Format: Sheet2: A1:B5, Sheet2: D1:E5
        const formattedRanges = excelContext.selectedRanges.map((range, index) =>
            `${excelContext.sheetName}: ${range.address}`
        ).join(', ');
        selectionElement.textContent = formattedRanges;
        selectionElement.style.color = '#0078d4';
        selectionElement.style.fontWeight = '600';
        selectionElement.style.fontSize = '11px'; // Smaller font for multiple ranges
    } else {
        selectionElement.textContent = excelContext.selectedRange || '-';
        selectionElement.style.color = '';
        selectionElement.style.fontWeight = '';
        selectionElement.style.fontSize = '';
    }
}

/**
 * Update API connection status display
 */
export function updateAPIStatus(isConnected) {
    const statusDot = document.getElementById('statusDot');
    const statusText = document.getElementById('statusText');

    // Clear existing classes
    statusDot.classList.remove('connected', 'error');

    if (isConnected) {
        statusDot.classList.add('connected');
        statusText.textContent = 'Connected';
    } else {
        statusDot.classList.add('error');
        statusText.textContent = 'Disconnected';
    }
}

