/**
 * Chat UI Module
 * Handles chat interface rendering and message display
 */

import { escapeHtml, renderMarkdown } from '../utils/helpers.js';

/**
 * Add message to chat
 */
export function addMessageToChat(sender, content, isSystem = false) {
    const chatContainer = document.getElementById('chatContainer');

    // Remove welcome message if it exists
    const welcomeMessage = chatContainer.querySelector('.welcome-message');
    if (welcomeMessage) {
        welcomeMessage.remove();
    }

    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${sender}`;

    const time = new Date().toLocaleTimeString('en-US', {
        hour: '2-digit',
        minute: '2-digit'
    });

    // Render content based on sender
    // AI messages support markdown, user messages are plain text
    const renderedContent = sender === 'ai' 
        ? renderMarkdown(content) 
        : escapeHtml(content);

    messageDiv.innerHTML = `
        <div class="message-header">
            <div class="message-avatar">${sender === 'user' ? 'You' : 'AI'}</div>
            <div class="message-sender">${sender === 'user' ? 'You' : 'AI Assistant'}</div>
        </div>
        <div class="message-content">${renderedContent}</div>
        <div class="message-time">${time}</div>
    `;

    chatContainer.appendChild(messageDiv);
    chatContainer.scrollTop = chatContainer.scrollHeight;
}

/**
 * Show typing indicator
 */
export function showTypingIndicator(text = 'Processing...') {
    const chatContainer = document.getElementById('chatContainer');

    // Remove existing indicator if present
    removeTypingIndicator();

    const typingDiv = document.createElement('div');
    typingDiv.className = 'message ai typing-indicator';
    typingDiv.id = 'typingIndicator';

    typingDiv.innerHTML = `
        <div class="message-header">
            <div class="message-avatar">AI</div>
            <div class="message-sender">AI Assistant</div>
        </div>
        <div class="message-content">
            <div class="typing-indicator-content">
                <span>${text}</span>
                <span class="typing-dots">
                    <span>.</span><span>.</span><span>.</span>
                </span>
            </div>
        </div>
    `;

    chatContainer.appendChild(typingDiv);
    chatContainer.scrollTop = chatContainer.scrollHeight;
}

/**
 * Update typing indicator text with progress
 */
export function updateTypingIndicator(text) {
    const indicator = document.getElementById('typingIndicator');
    if (indicator) {
        const contentDiv = indicator.querySelector('.typing-indicator-content');
        if (contentDiv) {
            const dotsSpan = contentDiv.querySelector('.typing-dots') || document.createElement('span');
            dotsSpan.className = 'typing-dots';
            dotsSpan.innerHTML = '<span>.</span><span>.</span><span>.</span>';

            contentDiv.innerHTML = `<span>${text}</span> `;
            contentDiv.appendChild(dotsSpan);
        }
    }
}

/**
 * Remove typing indicator
 */
export function removeTypingIndicator() {
    const typingIndicator = document.getElementById('typingIndicator');
    if (typingIndicator) {
        typingIndicator.remove();
    }
}

/**
 * Format data as table preview
 */
export function formatDataPreview(data, maxRows = 5) {
    if (!data || data.length === 0) return '';

    const displayData = data.slice(0, maxRows);
    let html = '<div class="data-preview"><table>';

    displayData.forEach((row, i) => {
        html += '<tr>';
        row.forEach(cell => {
            const tag = i === 0 ? 'th' : 'td';
            html += `<${tag}>${escapeHtml(String(cell))}</${tag}>`;
        });
        html += '</tr>';
    });

    if (data.length > maxRows) {
        html += `<tr><td colspan="${data[0].length}">... and ${data.length - maxRows} more rows</td></tr>`;
    }

    html += '</table></div>';
    return html;
}

