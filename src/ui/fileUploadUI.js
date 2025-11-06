/**
 * File Upload UI Module
 * Handles file upload interactions and displays upload status
 */

import { uploadFile, getConversationFile, getOrCreateConversation, deleteFile } from '../api/apiClient.js';
import { addMessageToChat } from './chatUI.js';

let currentConversationId = null;
let uploadedFile = null;

/**
 * Initialize file upload listeners (can be called without conversationId)
 */
export function initializeFileUpload(conversationId = null) {
    currentConversationId = conversationId;
    
    const uploadButton = document.getElementById('uploadButton');
    if (!uploadButton) {
        console.error('Upload button not found');
        return;
    }
    
    // Create hidden file inputs (only once)
    if (!document.getElementById('fileInput')) {
        createFileInputs();
    }
    
    // Remove old listener if exists, then add new one
    uploadButton.removeEventListener('click', handleUploadClick);
    uploadButton.addEventListener('click', handleUploadClick);
    
    // Check if file already exists (if we have conversationId)
    if (conversationId) {
        checkExistingFile();
    }
}

/**
 * Create hidden file input elements
 */
function createFileInputs() {
    // Main file input (PDF, Excel, etc)
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.id = 'fileInput';
    fileInput.accept = '.pdf,.xlsx,.xls,.csv';
    fileInput.style.display = 'none';
    document.body.appendChild(fileInput);
    
    // ToC JSON file input
    const tocInput = document.createElement('input');
    tocInput.type = 'file';
    tocInput.id = 'tocInput';
    tocInput.accept = '.json';
    tocInput.style.display = 'none';
    document.body.appendChild(tocInput);
    
    // Handle file selection
    fileInput.addEventListener('change', handleFileSelection);
}

/**
 * Handle upload button click
 */
function handleUploadClick() {
    const fileInput = document.getElementById('fileInput');
    fileInput.click();
}

/**
 * Handle file selection
 */
async function handleFileSelection(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    console.log(`📄 File selected: ${file.name}`);
    
    // Ask if user has a ToC file
    const hasToc = await askForTocFile();
    
    let tocFile = null;
    if (hasToc) {
        tocFile = await selectTocFile();
        if (!tocFile) {
            console.log('❌ ToC file selection cancelled');
            return;
        }
    }
    
    // Upload files
    await uploadFiles(file, tocFile);
    
    // Reset file inputs
    document.getElementById('fileInput').value = '';
    if (tocFile) {
        document.getElementById('tocInput').value = '';
    }
}

/**
 * Ask user if they have a ToC JSON file
 */
function askForTocFile() {
    return new Promise((resolve) => {
        // Create modal
        const modal = document.createElement('div');
        modal.className = 'clarification-modal';
        modal.innerHTML = `
            <div class="modal-overlay"></div>
            <div class="modal-content">
                <div class="modal-header">
                    <div class="modal-icon">📑</div>
                    <h3>Table of Contents</h3>
                </div>
                <div class="modal-body">
                    <p class="clarification-question">
                        Do you have a Table of Contents (ToC) JSON file for this document?
                    </p>
                    <p style="margin-top: 12px; font-size: 13px; color: #6b7280;">
                        ToC files help the AI navigate documents more efficiently and find specific sections faster.
                    </p>
                </div>
                <div class="modal-footer">
                    <button id="tocYesBtn" class="btn-primary">Yes, I have a ToC file</button>
                    <button id="tocNoBtn" class="btn-secondary">No, skip</button>
                </div>
            </div>
        `;
        
        document.body.appendChild(modal);
        
        // Remove hidden class to show
        setTimeout(() => modal.classList.remove('hidden'), 10);
        
        // Handle buttons
        const yesBtn = modal.querySelector('#tocYesBtn');
        const noBtn = modal.querySelector('#tocNoBtn');
        
        const cleanup = () => {
            modal.classList.add('hidden');
            setTimeout(() => modal.remove(), 300);
        };
        
        yesBtn.addEventListener('click', () => {
            cleanup();
            resolve(true);
        });
        
        noBtn.addEventListener('click', () => {
            cleanup();
            resolve(false);
        });
    });
}

/**
 * Let user select ToC JSON file
 */
function selectTocFile() {
    return new Promise((resolve) => {
        const tocInput = document.getElementById('tocInput');
        
        const handleChange = () => {
            const file = tocInput.files[0];
            tocInput.removeEventListener('change', handleChange);
            resolve(file);
        };
        
        const handleCancel = () => {
            // If user closes dialog without selecting
            setTimeout(() => {
                if (!tocInput.files.length) {
                    tocInput.removeEventListener('change', handleChange);
                    resolve(null);
                }
            }, 100);
        };
        
        tocInput.addEventListener('change', handleChange);
        tocInput.addEventListener('cancel', handleCancel);
        tocInput.click();
    });
}

/**
 * Upload files to backend
 */
async function uploadFiles(file, tocFile) {
    // Get or create conversation if needed
    if (!currentConversationId) {
        console.log('📤 No conversation yet, creating one...');
        const convResult = await getOrCreateConversation();
        if (!convResult.success) {
            showFileAttachmentBadge('error', `Failed to create conversation`);
            return;
        }
        currentConversationId = convResult.conversationId;
        console.log(`✓ Conversation created: ${currentConversationId}`);
    }
    
    // Show uploading indicator
    showUploadingIndicator(file.name);
    
    try {
        const result = await uploadFile(currentConversationId, file, tocFile);
        
        if (result.success) {
            uploadedFile = result.file;
            
            // Hide uploading indicator
            hideUploadingIndicator();
            
            // Update UI to show file is attached
            updateUploadButtonState(true, file.name);
            
            // Show success badge
            showFileAttachmentBadge('success', file.name, tocFile !== null);
            
        } else {
            // Hide uploading indicator
            hideUploadingIndicator();
            
            // Show error badge
            showFileAttachmentBadge('error', `Upload failed: ${result.error}`);
        }
        
    } catch (error) {
        // Hide uploading indicator
        hideUploadingIndicator();
        
        // Show error badge
        showFileAttachmentBadge('error', `Upload error: ${error.message}`);
    }
}

/**
 * Check if conversation already has a file
 */
async function checkExistingFile() {
    if (!currentConversationId) return;
    
    const result = await getConversationFile(currentConversationId);
    
    if (result.success && result.file) {
        uploadedFile = result.file;
        updateUploadButtonState(true, result.file.file_name);
        
        // Show badge if file exists
        const hasToc = result.file.toc && Object.keys(result.file.toc).length > 0;
        showFileAttachmentBadge('success', result.file.file_name, hasToc);
        
        console.log(`📄 Existing file found: ${result.file.file_name}`);
    }
}

/**
 * Update upload button visual state
 */
function updateUploadButtonState(hasFile, fileName) {
    const uploadButton = document.getElementById('uploadButton');
    const buttonText = uploadButton.querySelector('.upload-button-text');
    
    if (hasFile) {
        uploadButton.classList.add('has-file');
        uploadButton.title = `File attached: ${fileName}\nClick to replace`;
        buttonText.textContent = '📎 File attached';
    } else {
        uploadButton.classList.remove('has-file');
        uploadButton.title = 'Upload file';
        buttonText.textContent = 'Upload file';
    }
}

/**
 * Show uploading indicator with spinner
 */
function showUploadingIndicator(fileName) {
    // Remove any existing indicators
    hideUploadingIndicator();
    
    const indicator = document.createElement('div');
    indicator.id = 'fileUploadingIndicator';
    indicator.className = 'file-uploading-indicator';
    indicator.innerHTML = `
        <div class="upload-spinner"></div>
        <div class="upload-info">
            <div class="upload-filename">${fileName}</div>
            <div class="upload-status">Uploading... This may take a minute</div>
        </div>
    `;
    
    // Insert above input container
    const inputContainer = document.querySelector('.input-container');
    if (inputContainer) {
        inputContainer.insertBefore(indicator, inputContainer.firstChild);
    }
}

/**
 * Hide uploading indicator
 */
function hideUploadingIndicator() {
    const indicator = document.getElementById('fileUploadingIndicator');
    if (indicator) {
        indicator.remove();
    }
}

/**
 * Show file attachment badge above input
 */
function showFileAttachmentBadge(type, fileName, hasToc = false) {
    // Remove any existing badge
    const existingBadge = document.getElementById('fileAttachmentBadge');
    if (existingBadge) {
        existingBadge.remove();
    }
    
    if (type === 'error') {
        // Show error temporarily
        const badge = document.createElement('div');
        badge.id = 'fileAttachmentBadge';
        badge.className = 'file-attachment-badge error';
        badge.innerHTML = `
            <span class="badge-icon">❌</span>
            <span class="badge-text">${fileName}</span>
        `;
        
        const inputContainer = document.querySelector('.input-container');
        if (inputContainer) {
            inputContainer.insertBefore(badge, inputContainer.firstChild);
        }
        
        // Auto-remove after 5 seconds
        setTimeout(() => badge.remove(), 5000);
        return;
    }
    
    if (type === 'success') {
        // Show success badge (permanent)
        const badge = document.createElement('div');
        badge.id = 'fileAttachmentBadge';
        badge.className = 'file-attachment-badge success';
        badge.innerHTML = `
            <span class="badge-icon">📎</span>
            <div class="badge-content">
                <span class="badge-text">${fileName}</span>
                ${hasToc ? '<span class="badge-toc">ToC included</span>' : ''}
            </div>
            <button class="badge-close" title="Remove attachment">×</button>
        `;
        
        const inputContainer = document.querySelector('.input-container');
        if (inputContainer) {
            inputContainer.insertBefore(badge, inputContainer.firstChild);
        }
        
        // Handle close button
        const closeBtn = badge.querySelector('.badge-close');
        closeBtn.addEventListener('click', async () => {
            await handleFileDelete(badge);
        });
    }
}

/**
 * Handle file deletion
 */
async function handleFileDelete(badge) {
    if (!uploadedFile || !uploadedFile.id) {
        console.error('No uploaded file to delete');
        badge.remove();
        updateUploadButtonState(false);
        uploadedFile = null;
        return;
    }
    
    // Disable close button and show loading state
    const closeBtn = badge.querySelector('.badge-close');
    closeBtn.disabled = true;
    closeBtn.textContent = '⏳';
    
    try {
        console.log(`🗑️ Deleting file: ${uploadedFile.id}`);
        const result = await deleteFile(uploadedFile.id);
        
        if (result.success) {
            // Remove badge
            badge.remove();
            
            // Reset state
            updateUploadButtonState(false);
            uploadedFile = null;
            
            // Show temporary success message
            const tempBadge = document.createElement('div');
            tempBadge.className = 'file-attachment-badge success';
            tempBadge.style.opacity = '0.8';
            tempBadge.innerHTML = `
                <span class="badge-icon">✓</span>
                <span class="badge-text">File removed</span>
            `;
            
            const inputContainer = document.querySelector('.input-container');
            if (inputContainer) {
                inputContainer.insertBefore(tempBadge, inputContainer.firstChild);
            }
            
            // Auto-remove after 2 seconds
            setTimeout(() => tempBadge.remove(), 2000);
            
            console.log('✓ File deleted successfully');
        } else {
            // Show error
            closeBtn.disabled = false;
            closeBtn.textContent = '×';
            
            // Show error badge temporarily
            const errorBadge = document.createElement('div');
            errorBadge.className = 'file-attachment-badge error';
            errorBadge.innerHTML = `
                <span class="badge-icon">❌</span>
                <span class="badge-text">Failed to delete: ${result.error}</span>
            `;
            
            const inputContainer = document.querySelector('.input-container');
            if (inputContainer) {
                inputContainer.insertBefore(errorBadge, inputContainer.firstChild);
            }
            
            setTimeout(() => errorBadge.remove(), 5000);
        }
        
    } catch (error) {
        console.error('Error deleting file:', error);
        
        // Re-enable button
        closeBtn.disabled = false;
        closeBtn.textContent = '×';
        
        // Show error
        const errorBadge = document.createElement('div');
        errorBadge.className = 'file-attachment-badge error';
        errorBadge.innerHTML = `
            <span class="badge-icon">❌</span>
            <span class="badge-text">Error deleting file</span>
        `;
        
        const inputContainer = document.querySelector('.input-container');
        if (inputContainer) {
            inputContainer.insertBefore(errorBadge, inputContainer.firstChild);
        }
        
        setTimeout(() => errorBadge.remove(), 5000);
    }
}

/**
 * Get currently uploaded file info
 */
export function getUploadedFile() {
    return uploadedFile;
}

