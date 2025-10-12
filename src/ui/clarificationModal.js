/**
 * Clarification Modal Module
 * Handles user clarification prompts during agent execution
 */

/**
 * Show clarification modal and return promise that resolves with user's answer
 */
export function showClarificationModal(question) {
    return new Promise((resolve) => {
        const modal = document.getElementById('clarificationModal');
        const questionElement = document.getElementById('clarificationQuestion');
        const inputElement = document.getElementById('clarificationInput');
        const submitButton = document.getElementById('clarificationSubmit');
        const cancelButton = document.getElementById('clarificationCancel');

        // Set question text
        questionElement.textContent = question;
        inputElement.value = '';

        // Show modal
        modal.classList.remove('hidden');
        inputElement.focus();

        // Submit handler
        const handleSubmit = () => {
            const answer = inputElement.value.trim();
            if (answer) {
                modal.classList.add('hidden');
                cleanup();
                resolve(answer);
            } else {
                inputElement.focus();
            }
        };

        // Cancel handler
        const handleCancel = () => {
            modal.classList.add('hidden');
            cleanup();
            resolve(null);
        };

        // Enter key handler
        const handleKeyDown = (e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                handleSubmit();
            } else if (e.key === 'Escape') {
                e.preventDefault();
                handleCancel();
            }
        };

        // Cleanup function
        const cleanup = () => {
            submitButton.removeEventListener('click', handleSubmit);
            cancelButton.removeEventListener('click', handleCancel);
            inputElement.removeEventListener('keydown', handleKeyDown);
        };

        // Attach event listeners
        submitButton.addEventListener('click', handleSubmit);
        cancelButton.addEventListener('click', handleCancel);
        inputElement.addEventListener('keydown', handleKeyDown);
    });
}

