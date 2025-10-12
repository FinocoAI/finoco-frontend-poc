/**
 * Helper Utilities Module
 * Common utility functions used across the application
 */

/**
 * Escape HTML to prevent XSS
 */
export function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

/**
 * Sleep function
 */
export function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Auto-resize textarea
 */
export function handleInputResize(event) {
    const textarea = event.target;
    textarea.style.height = 'auto';
    textarea.style.height = Math.min(textarea.scrollHeight, 120) + 'px';
}

/**
 * Handle Enter key (send message) and Shift+Enter (new line)
 */
export function handleInputKeydown(event, onEnter) {
    if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        onEnter();
    }
}

