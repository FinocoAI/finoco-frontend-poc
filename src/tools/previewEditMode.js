/**
 * previewEditMode.js
 * 
 * What-If Analysis with Cascade Preview
 * Frontend-only feature that lets users preview the impact of changing a cell value
 * before committing the change.
 * 
 * Flow:
 * 1. User clicks "Preview Edit" button
 * 2. User selects a cell and enters new value
 * 3. System temporarily applies the change
 * 4. Excel auto-calculates all dependent cells
 * 5. System shows preview of all affected cells
 * 6. User commits (keeps) or cancels (reverts)
 */

// State management for preview mode
let previewState = {
  isActive: false,
  originalCell: null,
  originalValue: null,
  newValue: null,
  affectedCells: [],
  originalDependentValues: new Map()
};

/**
 * Main entry point: Activate preview edit mode
 */
export async function activatePreviewEditMode() {
  console.log('🔍 Activating Preview Edit Mode...');
  
  try {
    return await Excel.run(async (context) => {
      // Get currently selected cell
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load("address, values, formulas");
      
      const worksheet = selectedRange.worksheet;
      worksheet.load("name");
      
      await context.sync();
      
      // Validate selection (must be single cell)
      const address = selectedRange.address;
      const cellAddress = address.includes('!') ? address.split('!')[1] : address;
      
      // Check if it's a single cell
      if (cellAddress.includes(':')) {
        showNotification('⚠️ Please select a single cell (not a range)', 'warning');
        return;
      }
      
      const cell = {
        sheet: worksheet.name,
        address: cellAddress,
        fullAddress: `${worksheet.name}!${cellAddress}`,
        value: selectedRange.values[0][0],
        formula: selectedRange.formulas[0][0]
      };
      
      console.log('📍 Selected cell:', cell);
      
      // Show input dialog to get new value
      const newValue = await showValueInputDialog(cell);
      
      if (newValue === null || newValue === undefined) {
        console.log('❌ User cancelled input');
        return;
      }
      
      // Execute the preview
      await executePreview(cell, newValue);
    });
    
  } catch (error) {
    console.error('❌ Error activating preview mode:', error);
    showNotification(`Error: ${error.message || error}`, 'error');
  }
}

/**
 * Show input dialog for new value
 */
async function showValueInputDialog(cell) {
  return new Promise((resolve) => {
    const modal = document.getElementById('previewInputModal');
    const input = document.getElementById('previewNewValueInput');
    const cellInfoDiv = document.getElementById('previewCellInfo');
    const confirmBtn = document.getElementById('previewInputConfirm');
    const cancelBtn = document.getElementById('previewInputCancel');
    
    // Set cell info
      cellInfoDiv.innerHTML = `
      <div class="preview-cell-info">
        <strong>Cell:</strong> ${cell.fullAddress}<br>
        <strong>Current Value:</strong> ${formatValue(cell.value)}<br>
        ${cell.formula && typeof cell.formula === 'string' && cell.formula.startsWith('=') ? `<strong>Formula:</strong> <code>${cell.formula}</code><br>` : ''}
      </div>
    `;
    
    // Pre-fill current value
    input.value = cell.value;
    input.select();
    
    // Show modal
    modal.style.display = 'flex';
    input.focus();
    
    // Handle confirm
    const handleConfirm = () => {
      const newValue = parseInputValue(input.value);
      cleanup();
      resolve(newValue);
    };
    
    // Handle cancel
    const handleCancel = () => {
      cleanup();
      resolve(null);
    };
    
    // Cleanup function
    const cleanup = () => {
      modal.style.display = 'none';
      confirmBtn.removeEventListener('click', handleConfirm);
      cancelBtn.removeEventListener('click', handleCancel);
      input.removeEventListener('keydown', handleKeydown);
    };
    
    // Handle Enter/Escape keys
    const handleKeydown = (e) => {
      if (e.key === 'Enter') {
        handleConfirm();
      } else if (e.key === 'Escape') {
        handleCancel();
      }
    };
    
    // Attach event listeners
    confirmBtn.addEventListener('click', handleConfirm);
    cancelBtn.addEventListener('click', handleCancel);
    input.addEventListener('keydown', handleKeydown);
  });
}

/**
 * Parse input value (handle numbers, percentages, formulas)
 */
function parseInputValue(input) {
  const trimmed = input.trim();
  
  // If it's a formula, return as-is
  if (trimmed.startsWith('=')) {
    return trimmed;
  }
  
  // If it's a percentage, convert to decimal
  if (trimmed.endsWith('%')) {
    const num = parseFloat(trimmed.slice(0, -1));
    return isNaN(num) ? trimmed : num / 100;
  }
  
  // Try to parse as number
  const num = parseFloat(trimmed);
  if (!isNaN(num)) {
    return num;
  }
  
  // Return as string
  return trimmed;
}

/**
 * Execute the preview flow
 */
async function executePreview(cell, newValue) {
  console.log('🎯 Executing preview:', cell.fullAddress, '→', newValue);
  
  // Show loading indicator
  showPreviewLoading();
  
  try {
    await Excel.run(async (context) => {
      // Step 1: Store original state
      previewState.originalCell = cell;
      previewState.originalValue = cell.value;
      previewState.newValue = newValue;
      previewState.originalDependentValues.clear();
      
      // Step 2: Get dependents BEFORE making changes (to capture original values)
      console.log('📊 Reading dependent cells...');
      const dependentAddresses = await getDependentCells(context, cell.sheet, cell.address);
      
      console.log(`   Found ${dependentAddresses.length} dependent cell(s)`);
      
      // Step 3: Read original values from dependents
      if (dependentAddresses.length > 0) {
        await readOriginalDependentValues(context, dependentAddresses);
      }
      
      // Step 4: Apply temporary change to source cell
      console.log('✏️ Applying temporary change...');
      const worksheet = context.workbook.worksheets.getItem(cell.sheet);
      const range = worksheet.getRange(cell.address);
      
      // Apply value or formula
      if (typeof newValue === 'string' && newValue.startsWith('=')) {
        range.formulas = [[newValue]];
      } else {
        range.values = [[newValue]];
      }
      
      await context.sync();
      
      // Step 5: Wait for Excel to recalculate
      await waitForCalculation(context);
      
      // Step 6: Read NEW values from dependents
      console.log('📈 Reading new calculated values...');
      const affectedCells = await readNewDependentValues(context, dependentAddresses);
      
      previewState.affectedCells = affectedCells;
      
      console.log(`✅ Preview ready: ${affectedCells.length} cells affected`);
      
      // Step 7: Show impact preview modal
      showImpactPreviewModal();
    });
    
  } catch (error) {
    console.error('❌ Error executing preview:', error);
    showNotification(`Error: ${error.message || error}`, 'error');
    hidePreviewLoading();
  }
}

/**
 * Get all dependent cells (cells that reference this cell)
 */
async function getDependentCells(context, sheetName, address) {
  try {
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const range = worksheet.getRange(address);
    
    range.load("address");
    await context.sync();
    
    // Get direct dependents
    const dependents = range.getDirectDependents();
    dependents.load("address, areas");
    await context.sync();
    
    // Extract all dependent addresses
    const areas = dependents.areas;
    areas.load("items");
    await context.sync();
    
    const dependentAddresses = [];
    for (let i = 0; i < areas.items.length; i++) {
      const area = areas.items[i];
      area.load("address");
      await context.sync();
      dependentAddresses.push(area.address);
    }
    
    return dependentAddresses;
    
  } catch (error) {
    // No dependents found (or error)
    if (error.message && error.message.includes('DirectDependentsNotFound')) {
      return [];
    }
    console.warn('Warning getting dependents:', error.message);
    return [];
  }
}

/**
 * Read original values from dependent cells (before change)
 */
async function readOriginalDependentValues(context, dependentAddresses) {
  for (const fullAddress of dependentAddresses) {
    const { sheet, address } = parseFullAddress(fullAddress);
    
    try {
      const worksheet = context.workbook.worksheets.getItem(sheet);
      const range = worksheet.getRange(address);
      range.load("values");
      await context.sync();
      
      const originalValue = range.values[0][0];
      previewState.originalDependentValues.set(fullAddress, originalValue);
      
    } catch (error) {
      console.warn(`Could not read original value for ${fullAddress}:`, error.message);
    }
  }
}

/**
 * Read NEW values from dependent cells (after change)
 */
async function readNewDependentValues(context, dependentAddresses) {
  const affectedCells = [];
  
  for (const fullAddress of dependentAddresses) {
    const { sheet, address } = parseFullAddress(fullAddress);
    
    try {
      const worksheet = context.workbook.worksheets.getItem(sheet);
      const range = worksheet.getRange(address);
      range.load("address, values, formulas");
      await context.sync();
      
      const newValue = range.values[0][0];
      const oldValue = previewState.originalDependentValues.get(fullAddress);
      const formula = range.formulas[0][0];
      
      // Only include if value actually changed
      if (newValue !== oldValue) {
        affectedCells.push({
          sheet: sheet,
          address: address,
          fullAddress: fullAddress,
          oldValue: oldValue,
          newValue: newValue,
          formula: formula,
          change: calculateChange(oldValue, newValue),
          changePercent: calculatePercentChange(oldValue, newValue)
        });
      }
      
    } catch (error) {
      console.warn(`Could not read new value for ${fullAddress}:`, error.message);
    }
  }
  
  return affectedCells;
}

/**
 * Show the impact preview modal
 */
function showImpactPreviewModal() {
  hidePreviewLoading();
  
  const modal = document.getElementById('previewImpactModal');
  const detailsDiv = document.getElementById('previewImpactDetails');
  const commitBtn = document.getElementById('previewCommitBtn');
  const cancelBtn = document.getElementById('previewCancelBtn');
  
  const { originalCell, originalValue, newValue, affectedCells } = previewState;
  
  let html = `
    <div class="preview-source-section">
      <div class="preview-section-title">📍 Source Cell</div>
      <div class="preview-source-cell">
        <div class="preview-cell-address">${originalCell.fullAddress}</div>
        <div class="preview-value-change">
          <span class="old-value">${formatValue(originalValue)}</span>
          <span class="arrow">→</span>
          <span class="new-value">${formatValue(newValue)}</span>
        </div>
      </div>
    </div>
  `;
  
  if (affectedCells.length === 0) {
    html += `
      <div class="preview-dependents-section">
        <div class="preview-section-title">📊 Impact Analysis</div>
        <div class="preview-no-impact">
          ℹ️ No dependent cells found. This cell is not referenced by any formulas.
        </div>
      </div>
    `;
  } else {
    html += `
      <div class="preview-dependents-section">
        <div class="preview-section-title">📊 Will Affect ${affectedCells.length} Cell${affectedCells.length > 1 ? 's' : ''}</div>
        <div class="preview-table-wrapper">
          <table class="preview-impact-table">
            <thead>
              <tr>
                <th>Cell</th>
                <th>Current Value</th>
                <th>New Value</th>
                <th>Change</th>
              </tr>
            </thead>
            <tbody>
    `;
    
    affectedCells.forEach(cell => {
      const changeClass = cell.change > 0 ? 'positive' : (cell.change < 0 ? 'negative' : 'neutral');
      const changePrefix = cell.change > 0 ? '+' : '';
      
      html += `
        <tr>
          <td class="cell-address">
            ${cell.fullAddress}
            ${cell.formula && cell.formula.startsWith('=') ? `<br><span class="cell-formula">${escapeHtml(cell.formula)}</span>` : ''}
          </td>
          <td class="cell-value">${formatValue(cell.oldValue)}</td>
          <td class="cell-value new">${formatValue(cell.newValue)}</td>
          <td class="cell-change ${changeClass}">
            ${changePrefix}${formatValue(cell.change)}
            ${cell.changePercent !== null ? `<br><span class="change-percent">(${changePrefix}${cell.changePercent.toFixed(1)}%)</span>` : ''}
          </td>
        </tr>
      `;
    });
    
    html += `
            </tbody>
          </table>
        </div>
      </div>
    `;
  }
  
  detailsDiv.innerHTML = html;
  modal.style.display = 'flex';
  
  // Attach event handlers
  commitBtn.onclick = commitChange;
  cancelBtn.onclick = cancelChange;
  
  // Keyboard shortcuts
  const handleKeydown = (e) => {
    if (e.key === 'Enter' && e.ctrlKey) {
      commitChange();
    } else if (e.key === 'Escape') {
      cancelChange();
    }
  };
  
  document.addEventListener('keydown', handleKeydown);
  
  // Store cleanup function
  modal.dataset.cleanup = 'keydownHandler';
  modal.keydownHandler = handleKeydown;
}

/**
 * Commit the change (keep it)
 */
async function commitChange() {
  console.log('✅ Committing change...');
  
  const modal = document.getElementById('previewImpactModal');
  modal.style.display = 'none';
  
  // Cleanup event listener
  if (modal.keydownHandler) {
    document.removeEventListener('keydown', modal.keydownHandler);
    modal.keydownHandler = null;
  }
  
  // Show success message
  const affectedCount = previewState.affectedCells.length;
  const message = affectedCount > 0 
    ? `✅ Change committed! ${affectedCount} cell${affectedCount > 1 ? 's' : ''} updated.`
    : '✅ Change committed!';
  
  showSuccessToast(message);
  
  // Clear state
  resetPreviewState();
}

/**
 * Cancel the change (revert to original)
 */
async function cancelChange() {
  console.log('❌ Cancelling change...');
  
  const modal = document.getElementById('previewImpactModal');
  modal.style.display = 'none';
  
  // Cleanup event listener
  if (modal.keydownHandler) {
    document.removeEventListener('keydown', modal.keydownHandler);
    modal.keydownHandler = null;
  }
  
  try {
    await Excel.run(async (context) => {
      // Revert to original value
      const worksheet = context.workbook.worksheets.getItem(previewState.originalCell.sheet);
      const range = worksheet.getRange(previewState.originalCell.address);
      
      // Restore original value
      if (previewState.originalCell.formula && typeof previewState.originalCell.formula === 'string' && previewState.originalCell.formula.startsWith('=')) {
        range.formulas = [[previewState.originalCell.formula]];
      } else {
        range.values = [[previewState.originalValue]];
      }
      
      await context.sync();
      
      showSuccessToast('↩️ Change cancelled. Original value restored.');
    });
    
  } catch (error) {
    console.error('Error reverting change:', error);
    showNotification('Error reverting change: ' + error.message, 'error');
  }
  
  // Clear state
  resetPreviewState();
}

/**
 * Reset preview state
 */
function resetPreviewState() {
  previewState = {
    isActive: false,
    originalCell: null,
    originalValue: null,
    newValue: null,
    affectedCells: [],
    originalDependentValues: new Map()
  };
}

/**
 * Show loading indicator
 */
function showPreviewLoading() {
  const modal = document.getElementById('previewLoadingModal');
  modal.style.display = 'flex';
}

/**
 * Hide loading indicator
 */
function hidePreviewLoading() {
  const modal = document.getElementById('previewLoadingModal');
  modal.style.display = 'none';
}

/**
 * Show success toast notification
 */
function showSuccessToast(message) {
  const toast = document.getElementById('previewToast');
  if (toast) {
    toast.textContent = message;
    toast.classList.add('show');
    
    setTimeout(() => {
      toast.classList.remove('show');
    }, 3000);
  } else {
    // Fallback to alert if toast doesn't exist
    alert(message);
  }
}

/**
 * Wait for Excel to finish calculating
 */
async function waitForCalculation(context) {
  // Small delay to let Excel recalculate
  await new Promise(resolve => setTimeout(resolve, 150));
  await context.sync();
}

/**
 * Calculate numeric change between two values
 */
function calculateChange(oldValue, newValue) {
  if (typeof oldValue === 'number' && typeof newValue === 'number') {
    return newValue - oldValue;
  }
  return 0;
}

/**
 * Calculate percent change
 */
function calculatePercentChange(oldValue, newValue) {
  if (typeof oldValue === 'number' && typeof newValue === 'number' && oldValue !== 0) {
    return ((newValue - oldValue) / Math.abs(oldValue)) * 100;
  }
  return null;
}

/**
 * Format value for display
 */
function formatValue(value) {
  if (value === null || value === undefined || value === '') {
    return '(empty)';
  }
  
  if (typeof value === 'number') {
    // Format large numbers with commas
    if (Math.abs(value) >= 1000) {
      return value.toLocaleString(undefined, { 
        minimumFractionDigits: 0,
        maximumFractionDigits: 2 
      });
    }
    // Format small decimals
    return value.toLocaleString(undefined, { 
      minimumFractionDigits: 0,
      maximumFractionDigits: 4 
    });
  }
  
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE';
  }
  
  return String(value);
}

/**
 * Parse full address like "Sheet1!A1" into {sheet, address}
 */
function parseFullAddress(fullAddress) {
  if (fullAddress.includes('!')) {
    const parts = fullAddress.split('!');
    return {
      sheet: parts[0].replace(/'/g, ''), // Remove quotes
      address: parts[1]
    };
  }
  return {
    sheet: null,
    address: fullAddress
  };
}

/**
 * Escape HTML to prevent XSS
 */
function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

/**
 * Show notification (replaces alert since alert() is not supported in Excel add-ins)
 */
function showNotification(message, type = 'info') {
  const toast = document.getElementById('previewToast');
  if (toast) {
    // Set appropriate styling based on type
    toast.style.background = type === 'error' ? '#ef4444' : 
                            type === 'warning' ? '#f59e0b' : 
                            type === 'success' ? '#10b981' : '#1f2937';
    
    toast.textContent = message;
    toast.classList.add('show');
    
    setTimeout(() => {
      toast.classList.remove('show');
    }, 4000);
  } else {
    // Fallback: log to console
    console.log(`[${type.toUpperCase()}] ${message}`);
  }
}

