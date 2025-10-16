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
 * Render markdown to safe HTML
 * Uses marked.js for parsing and DOMPurify for sanitization
 */
export function renderMarkdown(text) {
    if (!text) return '';
    
    // Check if marked and DOMPurify are available
    if (typeof marked === 'undefined' || typeof DOMPurify === 'undefined') {
        console.warn('Markdown libraries not loaded, falling back to escaped HTML');
        return escapeHtml(text);
    }
    
    try {
        // Configure marked for better rendering
        marked.setOptions({
            breaks: true, // Convert \n to <br>
            gfm: true, // GitHub Flavored Markdown
            headerIds: false, // Don't add IDs to headers
            mangle: false // Don't mangle email addresses
        });
        
        // Parse markdown to HTML
        const rawHtml = marked.parse(text);
        
        // Sanitize HTML to prevent XSS
        const cleanHtml = DOMPurify.sanitize(rawHtml, {
            ALLOWED_TAGS: [
                'p', 'br', 'strong', 'em', 'u', 'code', 'pre', 
                'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
                'ul', 'ol', 'li', 'blockquote', 'hr',
                'a', 'table', 'thead', 'tbody', 'tr', 'th', 'td',
                'span', 'div', 'del', 'ins'
            ],
            ALLOWED_ATTR: ['href', 'title', 'class', 'target', 'rel']
        });
        
        return cleanHtml;
    } catch (error) {
        console.error('Error rendering markdown:', error);
        return escapeHtml(text);
    }
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

