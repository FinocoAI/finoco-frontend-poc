/**
 * Overwrite Approval Modal Module
 * Handles user approval for overwriting existing data in Excel
 */

/**
 * Show overwrite approval modal and return promise that resolves with user's decision
 * 
 * @param {string} rangeAddress - The range that would be overwritten (e.g., "A1:C10")
 * @param {number} cellsAffected - Number of cells with existing data that would be overwritten
 * @param {Array<string>} sampleCells - Sample of cells that would be affected (e.g., ["A1", "B2", "C3"])
 * @returns {Promise<boolean>} - true if user approves, false if user cancels
 */
export function showOverwriteModal(rangeAddress, cellsAffected, sampleCells = []) {
    return new Promise((resolve) => {
        const modal = document.getElementById('overwriteModal');
        const messageElement = document.getElementById('overwriteMessage');
        const detailsElement = document.getElementById('overwriteDetails');
        const approveButton = document.getElementById('overwriteApprove');
        const cancelButton = document.getElementById('overwriteCancel');

        // Build message
        const message = `This operation will overwrite ${cellsAffected} cell${cellsAffected > 1 ? 's' : ''} containing existing data in range <strong>${rangeAddress}</strong>.`;
        messageElement.innerHTML = message;

        // Build details (show sample of affected cells)
        if (sampleCells.length > 0) {
            const sampleText = sampleCells.slice(0, 10).join(', ');
            const moreText = sampleCells.length > 10 ? ` and ${sampleCells.length - 10} more...` : '';
            detailsElement.innerHTML = `<strong>Affected cells:</strong><br>${sampleText}${moreText}`;
        } else {
            detailsElement.innerHTML = `<strong>Range:</strong> ${rangeAddress}`;
        }

        // Show modal
        modal.classList.remove('hidden');

        // Approve handler
        const handleApprove = () => {
            modal.classList.add('hidden');
            cleanup();
            resolve(true);
        };

        // Cancel handler
        const handleCancel = () => {
            modal.classList.add('hidden');
            cleanup();
            resolve(false);
        };

        // Escape key handler
        const handleKeyDown = (e) => {
            if (e.key === 'Escape') {
                e.preventDefault();
                handleCancel();
            } else if (e.key === 'Enter' && e.ctrlKey) {
                e.preventDefault();
                handleApprove();
            }
        };

        // Cleanup function
        const cleanup = () => {
            approveButton.removeEventListener('click', handleApprove);
            cancelButton.removeEventListener('click', handleCancel);
            document.removeEventListener('keydown', handleKeyDown);
        };

        // Attach event listeners
        approveButton.addEventListener('click', handleApprove);
        cancelButton.addEventListener('click', handleCancel);
        document.addEventListener('keydown', handleKeyDown);

        // Focus on approve button
        approveButton.focus();
    });
}

