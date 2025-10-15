/**
 * Edit Preview Modal Module
 * Shows preview of data changes before execution with approve/reject options
 * 
 * This modal intercepts ALL write operations and shows the user:
 * - What data will be written
 * - Which cells will be affected
 * - Before/after comparison (if overwriting)
 * - Option to approve or reject changes
 */

/**
 * Show edit preview modal with before/after data comparison
 * 
 * @param {Object|Array} previewData - Data about the pending write operation(s)
 *   Single sheet: Object with sheet data
 *   Multiple sheets: Array of objects, one per sheet
 * @returns {Promise<{approved: boolean, autoApprove: boolean, feedback?: string}>}
 */
export function showEditPreviewModal(previewData) {
    // Handle multi-sheet operations
    if (Array.isArray(previewData)) {
        return showMultiSheetPreviewModal(previewData);
    }
    
    // Single sheet operation (original behavior)
    return showSingleSheetPreviewModal(previewData);
}

/**
 * Show preview for single sheet operation
 */
function showSingleSheetPreviewModal(previewData) {
    return new Promise((resolve) => {
        const modal = document.getElementById('editPreviewModal');
        const operationTypeElement = document.getElementById('editOperationType');
        const sheetNameElement = document.getElementById('editSheetName');
        const rangeInfoElement = document.getElementById('editRangeInfo');
        const warningElement = document.getElementById('editWarning');
        const previewTableContainer = document.getElementById('editPreviewTable');
        const approveButton = document.getElementById('editApprove');
        const rejectButton = document.getElementById('editReject');

        // Build operation description
        const operationLabels = {
            'writeDataToRange': '⚠️ Overwriting Data',
            'applyFormula': '⚠️ Overwriting with Formula',
            'formatRange': '🎨 Formatting Range',
        };
        const operationLabel = operationLabels[previewData.operation] || '⚠️ Overwriting';
        operationTypeElement.textContent = operationLabel;

        // Show sheet name prominently
        sheetNameElement.textContent = previewData.sheetName;

        // Build range info
        const rangeInfo = `${previewData.rangeAddress} (${previewData.rowCount} rows × ${previewData.colCount} columns)`;
        rangeInfoElement.textContent = rangeInfo;

        // Always show warning (this modal only appears for overwrites)
        const overwriteCount = previewData.affectedCells?.length || (previewData.rowCount * previewData.colCount);
        warningElement.innerHTML = `
            <div style="display: flex; align-items: center; gap: 8px; padding: 12px; background: #fef3c7; border: 1px solid #f59e0b; border-radius: 6px; margin-top: 12px;">
                <span style="font-size: 20px;">⚠️</span>
                <span style="color: #92400e; font-weight: 500;">
                    This will overwrite ${overwriteCount} cell${overwriteCount > 1 ? 's' : ''} containing existing data
                </span>
            </div>
        `;

        // Build preview table
        const previewHTML = buildPreviewTable(previewData);
        previewTableContainer.innerHTML = previewHTML;

        // Show modal with animation
        modal.classList.remove('hidden');
        requestAnimationFrame(() => {
            modal.classList.add('active');
        });

        // Approve handler
        const handleApprove = () => {
            console.log('✅ User approved overwrite operation');
            
            modal.classList.remove('active');
            setTimeout(() => {
                modal.classList.add('hidden');
                cleanup();
                resolve({ approved: true });
            }, 200);
        };

        // Reject handler
        const handleReject = () => {
            console.log('❌ User rejected overwrite operation');
            
            // Build feedback message for backend
            const feedback = `User rejected the ${previewData.operation} operation. ` +
                `Target: ${previewData.sheetName}!${previewData.rangeAddress}. ` +
                `The user did not want to overwrite the existing data in these cells. ` +
                `Consider asking for a different location or verifying the data with the user first.`;
            
            modal.classList.remove('active');
            setTimeout(() => {
                modal.classList.add('hidden');
                cleanup();
                resolve({ approved: false, feedback });
            }, 200);
        };

        // Escape key handler
        const handleKeyDown = (e) => {
            if (e.key === 'Escape') {
                e.preventDefault();
                handleReject();
            } else if (e.key === 'Enter' && e.ctrlKey) {
                e.preventDefault();
                handleApprove();
            }
        };

        // Cleanup function
        const cleanup = () => {
            approveButton.removeEventListener('click', handleApprove);
            rejectButton.removeEventListener('click', handleReject);
            document.removeEventListener('keydown', handleKeyDown);
        };

        // Attach event listeners
        approveButton.addEventListener('click', handleApprove);
        rejectButton.addEventListener('click', handleReject);
        document.addEventListener('keydown', handleKeyDown);

        // Focus on approve button
        approveButton.focus();
    });
}

/**
 * Build HTML for the preview table showing before/after comparison
 */
function buildPreviewTable(previewData) {
    const { newData, existingData, rowCount, colCount, hasExistingData } = previewData;
    
    // Show ALL rows with scrolling (no truncation)
    let html = '<div class="edit-preview-scroll">';
    html += '<table class="edit-preview-table">';
    
    // Build table headers (column letters)
    html += '<thead><tr><th class="cell-header"></th>'; // Corner cell
    for (let col = 0; col < colCount; col++) {
        const colLetter = getColumnLetter(col);
        html += `<th class="cell-header">${colLetter}</th>`;
    }
    html += '</tr></thead>';
    
    // Build table body - SHOW ALL ROWS
    html += '<tbody>';
    for (let row = 0; row < rowCount; row++) {
        html += '<tr>';
        
        // Row number
        html += `<td class="cell-header">${row + 1}</td>`;
        
        // Data cells
        for (let col = 0; col < colCount; col++) {
            const newValue = newData[row]?.[col];
            const oldValue = existingData?.[row]?.[col];
            const isEmpty = (val) => val === '' || val === null || val === undefined;
            
            const hasOldData = hasExistingData && !isEmpty(oldValue);
            const hasNewData = !isEmpty(newValue);
            
            let cellClass = 'preview-cell';
            let cellContent = '';
            
            if (hasOldData && hasNewData) {
                // Overwriting existing data - show both
                cellClass += ' cell-modified';
                cellContent = `
                    <div class="cell-before">${formatCellValue(oldValue)}</div>
                    <div class="cell-arrow">→</div>
                    <div class="cell-after">${formatCellValue(newValue)}</div>
                `;
            } else if (hasNewData) {
                // Writing to empty cell - show new value
                cellClass += ' cell-new';
                cellContent = `<div class="cell-value">${formatCellValue(newValue)}</div>`;
            } else if (hasOldData) {
                // Not changing this cell (shouldn't happen, but handle it)
                cellClass += ' cell-unchanged';
                cellContent = `<div class="cell-value">${formatCellValue(oldValue)}</div>`;
            } else {
                // Empty cell, not writing anything
                cellClass += ' cell-empty';
                cellContent = `<div class="cell-value" style="color: #9ca3af;">—</div>`;
            }
            
            html += `<td class="${cellClass}">${cellContent}</td>`;
        }
        
        html += '</tr>';
    }
    html += '</tbody>';
    
    html += '</table>';
    html += '</div>';
    
    return html;
}

/**
 * Format cell value for display (handle formulas, numbers, etc.)
 */
function formatCellValue(value) {
    if (value === null || value === undefined || value === '') {
        return '<span style="color: #9ca3af; font-style: italic;">empty</span>';
    }
    
    // Truncate long values
    const str = String(value);
    if (str.length > 50) {
        return `<span title="${escapeHtml(str)}">${escapeHtml(str.substring(0, 47))}...</span>`;
    }
    
    // Highlight formulas
    if (str.startsWith('=')) {
        return `<span style="color: #059669; font-family: 'Courier New', monospace; font-size: 12px;">${escapeHtml(str)}</span>`;
    }
    
    return escapeHtml(str);
}

/**
 * Escape HTML special characters
 */
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

/**
 * Convert column index to letter (0 -> A, 25 -> Z, 26 -> AA)
 */
function getColumnLetter(index) {
    let letter = '';
    while (index >= 0) {
        letter = String.fromCharCode((index % 26) + 65) + letter;
        index = Math.floor(index / 26) - 1;
    }
    return letter;
}

/**
 * Show preview for multiple sheet operations with dropdown selector
 */
function showMultiSheetPreviewModal(sheetsData) {
    return new Promise((resolve) => {
        const modal = document.getElementById('editPreviewModal');
        const operationTypeElement = document.getElementById('editOperationType');
        const sheetNameElement = document.getElementById('editSheetName');
        const rangeInfoElement = document.getElementById('editRangeInfo');
        const warningElement = document.getElementById('editWarning');
        const previewTableContainer = document.getElementById('editPreviewTable');
        const approveButton = document.getElementById('editApprove');
        const rejectButton = document.getElementById('editReject');

        // Track current sheet index
        let currentSheetIndex = 0;
        let userDecisions = {}; // Track approval per sheet

        // Build operation description
        operationTypeElement.textContent = `⚠️ Overwriting in ${sheetsData.length} Sheet${sheetsData.length > 1 ? 's' : ''}`;

        // Create sheet dropdown if multiple sheets
        const sheetSelectorHTML = `
            <div style="display: flex; align-items: center; gap: 8px; margin-top: 6px;">
                <span style="font-size: 12px; color: #9ca3af;">Sheet:</span>
                <select id="editSheetSelector" style="font-size: 14px; font-weight: 600; color: #3b82f6; background: #eff6ff; padding: 6px 12px; border-radius: 6px; border: 1px solid #bfdbfe; cursor: pointer;">
                    ${sheetsData.map((sheet, idx) => 
                        `<option value="${idx}">${sheet.sheetName} (${sheet.rowCount}×${sheet.colCount})</option>`
                    ).join('')}
                </select>
            </div>
        `;
        
        // Replace sheet name element with dropdown
        sheetNameElement.outerHTML = sheetSelectorHTML;
        const sheetSelector = document.getElementById('editSheetSelector');

        // Function to render current sheet
        const renderCurrentSheet = () => {
            const currentSheet = sheetsData[currentSheetIndex];
            
            // Update range info
            rangeInfoElement.textContent = `${currentSheet.rangeAddress} (${currentSheet.rowCount} rows × ${currentSheet.colCount} columns)`;

            // Always show warning (modal only appears for overwrites)
            const overwriteCount = currentSheet.affectedCells?.length || (currentSheet.rowCount * currentSheet.colCount);
            warningElement.innerHTML = `
                <div style="display: flex; align-items: center; gap: 8px; padding: 12px; background: #fef3c7; border: 1px solid #f59e0b; border-radius: 6px; margin-top: 12px;">
                    <span style="font-size: 20px;">⚠️</span>
                    <span style="color: #92400e; font-weight: 500;">
                        This will overwrite ${overwriteCount} cell${overwriteCount > 1 ? 's' : ''} containing existing data
                    </span>
                </div>
            `;

            // Build and show preview table
            const previewHTML = buildPreviewTable(currentSheet);
            previewTableContainer.innerHTML = previewHTML;

            // Update button text to show progress
            const reviewedCount = Object.keys(userDecisions).length;
            const totalSheets = sheetsData.length;
            approveButton.textContent = reviewedCount === totalSheets 
                ? `✓ Approve All (${totalSheets} sheets)` 
                : `✓ Approve (${reviewedCount + 1}/${totalSheets})`;
        };

        // Sheet selector change handler
        sheetSelector.addEventListener('change', (e) => {
            currentSheetIndex = parseInt(e.target.value);
            renderCurrentSheet();
        });

        // Initial render
        renderCurrentSheet();

        // Show modal
        modal.classList.remove('hidden');
        requestAnimationFrame(() => {
            modal.classList.add('active');
        });

        // Approve handler - approve current sheet and move to next
        const handleApprove = () => {
            const currentSheet = sheetsData[currentSheetIndex];
            userDecisions[currentSheet.sheetName] = { approved: true };

            // If there are more sheets, move to next
            if (currentSheetIndex < sheetsData.length - 1) {
                currentSheetIndex++;
                sheetSelector.value = currentSheetIndex;
                renderCurrentSheet();
            } else {
                // All sheets reviewed and approved
                console.log('✅ User approved all sheet overwrite operations', { sheets: sheetsData.length });
                
                modal.classList.remove('active');
                setTimeout(() => {
                    modal.classList.add('hidden');
                    cleanup();
                    resolve({ approved: true });
                }, 200);
            }
        };

        // Reject handler - reject all remaining operations
        const handleReject = () => {
            console.log('❌ User rejected multi-sheet operation');
            
            const feedback = `User rejected write operations across ${sheetsData.length} sheet${sheetsData.length > 1 ? 's' : ''}. ` +
                `Sheets: ${sheetsData.map(s => s.sheetName).join(', ')}. ` +
                `The user did not approve these changes.`;
            
            modal.classList.remove('active');
            setTimeout(() => {
                modal.classList.add('hidden');
                cleanup();
                resolve({ approved: false, feedback });
            }, 200);
        };

        // Keyboard handlers
        const handleKeyDown = (e) => {
            if (e.key === 'Escape') {
                e.preventDefault();
                handleReject();
            } else if (e.key === 'Enter' && e.ctrlKey) {
                e.preventDefault();
                handleApprove();
            }
        };

        // Cleanup
        const cleanup = () => {
            approveButton.removeEventListener('click', handleApprove);
            rejectButton.removeEventListener('click', handleReject);
            document.removeEventListener('keydown', handleKeyDown);
            
            // Restore original sheet name element
            sheetSelector.outerHTML = '<span id="editSheetName" style="font-size: 14px; font-weight: 600; color: #3b82f6; background: #eff6ff; padding: 4px 12px; border-radius: 6px;"></span>';
        };

        // Attach listeners
        approveButton.addEventListener('click', handleApprove);
        rejectButton.addEventListener('click', handleReject);
        document.addEventListener('keydown', handleKeyDown);
        approveButton.focus();
    });
}

