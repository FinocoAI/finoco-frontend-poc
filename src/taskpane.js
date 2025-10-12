
// Import context capture modules
import { captureInitialContext } from './contextCapture.js';
import { captureUserSelection } from './selectionCapture.js';
import { buildQueryPayload, debugPrintPayload, validateQueryPayload } from './queryBuilder.js';

// Import tool system
import { executeTool, executeToolSequence, formatToolResult, getExecutionSummary } from './tools/executor.js';
import { getAllToolDefinitions, printToolRegistry } from './tools/registry.js';

const API_BASE_URL = 'https://41a36d0f8a03.ngrok-free.app';


// Legacy context structure (for UI display)
let excelContext = {
    workbookName: null,
    sheetName: null,
    allSheets: [],
    selectedRange: null,
    selectedData: null,
    headers: null,
    selectedRanges: [],
    isMultipleAreas: false
};

// NEW: Enhanced context structure (as per architecture)
let enhancedContext = {
    initialContext: null,    // From captureInitialContext()
    userSelection: null,     // From captureUserSelection()
};

// Initialize Office addins so we need this to check once add in is loaded -- then start matlab dont initializa app untill all done
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log('Office.js initialized');
        initializeApp();
    }
});

// Initialize the application
async function initializeApp() {
    // Set up event listeners
    document.getElementById('sendButton').addEventListener('click', handleSendMessage);
    document.getElementById('refreshButton').addEventListener('click', handleRefreshContext);
    document.getElementById('userInput').addEventListener('keydown', handleInputKeydown);
    document.getElementById('userInput').addEventListener('input', handleInputResize);

    // Load initial context (both legacy and new)
    await updateExcelContext();
    await updateEnhancedContext();

    // backen up or nott
    await checkAPIConnection();

    // Set up Excel event listeners
    await setupExcelEventListeners();

    // Add debug keyboard shortcuts
    document.addEventListener('keydown', (e) => {
        // Ctrl/Cmd + Shift + D: Context Capture
        if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key === 'D') {
            e.preventDefault();
            debugContextCapture();
        }

        // Ctrl/Cmd + Shift + T: Tool Testing
        if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key === 'K') {
            e.preventDefault();
            debugToolExecution();
        }
    });

    console.log('💡 Debug mode:');
    // Expose debug utilities globally for manual console testing
    window.debugToolExecution = debugToolExecution;
    window.debugContextCapture = debugContextCapture;
    console.log('🧩 Global debug functions available: debugToolExecution(), debugContextCapture()');

    console.log('  - Press Ctrl/Cmd + Shift + D to capture context');
    console.log('  - Press Ctrl/Cmd + Shift + T to test tools');
    console.log('');

    // Print available tools on startup
    printToolRegistry();
} 

// Update enhanced context using new architecture
async function updateEnhancedContext() {
    try {
        console.log('🔄 Updating enhanced context...');

        // Capture initial context (always)
        enhancedContext.initialContext = await captureInitialContext();

        // Capture user selection (always try - it returns null if nothing selected)
        // Don't use hasUserSelection() check - just let captureUserSelection() handle it
        enhancedContext.userSelection = await captureUserSelection();

        console.log('✅ Enhanced context updated');
    } catch (error) {
        console.error('❌ Error updating enhanced context:', error);
    }
}

// Debug function to capture and log all context
async function debugContextCapture() {
    console.group('🔍 DEBUG: Context Capture');
    console.log('Capturing context at:', new Date().toISOString());

    try {
        // Capture fresh context
        await updateEnhancedContext();

        // Build a test query payload
        const testQuery = "This is a test query for debugging";
        const payload = buildQueryPayload(
            testQuery,
            enhancedContext.initialContext,
            enhancedContext.userSelection
        );

        // Pretty print the payload
        debugPrintPayload(payload);

        // Validate
        const validation = validateQueryPayload(payload);
        console.log('✅ Validation result:', validation);

        // Show alert to user
        alert('Context captured! Check console for details (F12)');

    } catch (error) {
        console.error('❌ Debug capture failed:', error);
        alert(`Debug failed: ${error.message}`);
    }

    console.groupEnd();
}

// Debug function to test tool execution
async function debugToolExecution() {
    console.group('🔧 DEBUG: Tool Execution');
    console.log('Testing tools at:', new Date().toISOString());

    try {
        // Get active sheet name
        const activeSheetName = await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.load('name');
            await context.sync();
            return sheet.name;
        });

        console.log(`Active sheet: ${activeSheetName}`);

        // Test 1: getFullRangeData
        console.log('\n--- Test 1: getFullRangeData ---');
        const test1 = await executeTool('getFullRangeData', {
            sheetName: activeSheetName,
            address: 'A1:C10',
            includeFormulas: false
        });
        console.log(formatToolResult(test1));
        if (test1.success) {
            console.log('Sample data:', test1.result.values.slice(0, 3));
        }

        // Test 2: getColumnData
        console.log('\n--- Test 2: getColumnData ---');
        const test2 = await executeTool('getColumnData', {
            sheetName: activeSheetName,
            columns: ['A', 'B'],
            includeHeaders: true
        });
        console.log(formatToolResult(test2));
        if (test2.success) {
            console.log('Columns retrieved:', test2.result.columns);
            console.log('Row count:', test2.result.rowCount);
        }

        // Test 3: Multiple tools in sequence
        console.log('\n--- Test 3: Tool Sequence ---');
        const sequence = await executeToolSequence([
            {
                tool: 'getFullRangeData',
                params: { sheetName: activeSheetName, address: 'A1:B5' }
            },
            {
                tool: 'getColumnData',
                params: { sheetName: activeSheetName, columns: ['C'] }
            }
        ]);

        const summary = getExecutionSummary(sequence);
        console.log('Execution Summary:', summary);

        // Show tool definitions
        console.log('\n--- Available Tools ---');
        const toolDefs = getAllToolDefinitions();
        toolDefs.forEach(def => {
            console.log(`\n${def.name}:`);
            console.log(`  Description: ${def.description}`);
            console.log(`  Parameters:`, Object.keys(def.parameters));
        });

        console.log('Tool testing complete! Check console for details (F12)');

    } catch (error) {
        console.error('❌ Tool testing failed:', error);
        console.log(`Tool testing failed: ${error.message}`);
    }

    console.groupEnd();
}

// Handle refresh button click
async function handleRefreshContext() {
    console.log('Manual refresh triggered');
    await updateExcelContext();
    await updateEnhancedContext();
}

// Auto-resize textarea
function handleInputResize(event) {
    const textarea = event.target;
    textarea.style.height = 'auto';
    textarea.style.height = Math.min(textarea.scrollHeight, 120) + 'px';
}

// Handle Enter key (send message) and Shift+Enter (new line)
function handleInputKeydown(event) {
    if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        handleSendMessage();
    }
}

// Set up Excel event listeners
async function setupExcelEventListeners() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();

            // Listen to selection changes
            sheet.onSelectionChanged.add(async () => {
                await updateExcelContext();
                await updateEnhancedContext();
            });

            // Listen to sheet activation
            context.workbook.worksheets.onActivated.add(async () => {
                await updateExcelContext();
                await updateEnhancedContext();
            });

            await context.sync();
            console.log('Excel event listeners set up');
        });
    } catch (error) {
        console.error('Error setting up Excel event listeners:', error);
    }
}

// Update Excel context
async function updateExcelContext() {
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
            updateContextDisplay();
        });
    } catch (error) {
        console.error('Error updating Excel context:', error);
        updateContextDisplay();
    }
}

// Update context display in UI
function updateContextDisplay() {
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

// Check API connection
async function checkAPIConnection() {
    const statusDot = document.getElementById('statusDot');
    const statusText = document.getElementById('statusText');

    try {
        const response = await fetch(`${API_BASE_URL}/health`, {
            headers: {
                'ngrok-skip-browser-warning': 'true'
            }
        });
        if (response.ok) {
            statusDot.classList.add('connected');
            statusText.textContent = 'Connected';
        } else {
            throw new Error('API not responding');
        }
    } catch (error) {
        statusDot.classList.add('error');
        statusText.textContent = 'Disconnected';
        console.error('API connection error:', error);
    }
}

// Handle send message
// Generate unique request ID for deduplication
function generateRequestId() {
    return `req-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
}

async function handleSendMessage() {
    const input = document.getElementById('userInput');
    const message = input.value.trim();

    if (!message) return;

    // Clear input
    input.value = '';
    input.style.height = 'auto';

    // Disable send button
    const sendButton = document.getElementById('sendButton');
    sendButton.disabled = true;

    // Add user message to chat
    addMessageToChat('user', message);

    // Show typing indicator
    showTypingIndicator('Processing...');

    try {
        // Refresh context before sending
        await updateExcelContext();
        await updateEnhancedContext();

        // Build query payload
        const queryPayload = buildQueryPayload(
            message,
            enhancedContext.initialContext,
            enhancedContext.userSelection
        );

        // Log payload for debugging
        console.group('📤 Sending Query to Backend');
        debugPrintPayload(queryPayload);
        console.groupEnd();

        // Validate payload
        const validation = validateQueryPayload(queryPayload);
        if (!validation.isValid) {
            throw new Error(`Invalid payload: ${validation.errors.join(', ')}`);
        }

        // Generate unique request ID for deduplication
        const requestId = generateRequestId();
        console.log(`📋 Request ID: ${requestId}`);

        // Submit task to backend with retry logic (returns immediately with task_id)
        let submitResult;
        let taskId;
        const maxRetries = 3;
        let retryCount = 0;

        while (retryCount < maxRetries) {
            try {
                console.log(`📤 Submitting task (attempt ${retryCount + 1}/${maxRetries})...`);

                const submitResponse = await fetch(`${API_BASE_URL}/process`, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'ngrok-skip-browser-warning': 'true'
                    },
                    body: JSON.stringify({
                        task: message,
                        requestId: requestId,
                        enhancedPayload: queryPayload
                    })
                });

                if (!submitResponse.ok) {
                    throw new Error(`HTTP ${submitResponse.status}: ${submitResponse.statusText}`);
                }

                submitResult = await submitResponse.json();
                taskId = submitResult.task_id;

                console.log(`✓ Task submitted: ${taskId}`);
                console.log(`  Status: ${submitResult.status}`);
                break; // Success - exit retry loop

            } catch (error) {
                retryCount++;
                console.error(`❌ Submit attempt ${retryCount} failed:`, error.message);

                if (retryCount >= maxRetries) {
                    throw new Error(`Failed to submit task after ${maxRetries} attempts: ${error.message}`);
                }

                // Wait before retry (faster retry for ngrok issues)
                const waitTime = 500; // Fixed 500ms retry for ngrok instability
                console.log(`⏳ Retrying in ${waitTime}ms...`);
                updateTypingIndicator(`Connection error, retrying... (${retryCount}/${maxRetries})`);
                await sleep(waitTime);
            }
        }

        // If all retries failed, try to recover task using requestId
        if (!taskId) {
            console.warn('⚠️ Initial POST failed, attempting task recovery...');
            updateTypingIndicator('Recovering task...');

            try {
                const recoveryResponse = await fetch(`${API_BASE_URL}/tasks/by-request/${requestId}`, {
                    headers: {
                        'ngrok-skip-browser-warning': 'true'
                    }
                });

                if (recoveryResponse.ok) {
                    const recoveryData = await recoveryResponse.json();
                    taskId = recoveryData.task_id;
                    console.log(`✓ Task recovered: ${taskId}`);
                    console.log(`  Status: ${recoveryData.status}`);
                } else {
                    throw new Error('Task recovery failed');
                }
            } catch (recoveryError) {
                console.error('Task recovery failed:', recoveryError);
                throw new Error('Failed to submit task and recovery failed. Please try again.');
            }
        }

        // Start polling for task completion
        const result = await pollTaskStatus(taskId);

        // Remove typing indicator
        removeTypingIndicator();

        // Add AI response to chat - prioritize 'answer' field if present
        const displayMessage = result.answer && result.answer !== null && result.answer !== 'null'
            ? result.answer
            : result.message;
        addMessageToChat('ai', displayMessage);

        // Execute frontend tool calls if present
        if (result.toolCalls && Array.isArray(result.toolCalls) && result.toolCalls.length > 0) {
            console.log(`🔧 Executing ${result.toolCalls.length} frontend tool(s)`);

            for (const toolCall of result.toolCalls) {
                try {
                    const toolResult = await executeTool(toolCall.tool, toolCall.params);

                    if (toolResult.success) {
                        if (toolCall.tool === 'writeDataToRange') {
                            addMessageToChat('ai', `✓ Data written to ${toolResult.result.rangeAddress}`, true);
                        } else if (toolCall.tool === 'createChart') {
                            addMessageToChat('ai', `✓ Chart "${toolResult.result.title}" created`, true);
                        } else if (toolCall.tool === 'createTable') {
                            addMessageToChat('ai', `✓ Table "${toolResult.result.tableName}" created`, true);
                        } else {
                            addMessageToChat('ai', `✓ ${toolCall.tool} completed`, true);
                        }
                    } else {
                        addMessageToChat('ai', `⚠️ ${toolCall.tool} failed: ${toolResult.error}`, true);
                    }
                } catch (error) {
                    console.error(`Tool execution error:`, error);
                    addMessageToChat('ai', `⚠️ Failed to execute ${toolCall.tool}`, true);
                }
            }
        }

    } catch (error) {
        console.error('Error processing message:', error);
        removeTypingIndicator();
        addMessageToChat('ai', `Sorry, I encountered an error: ${error.message}`);
    } finally {
        sendButton.disabled = false;
        input.focus();
    }
}

// Poll task status until complete or failed
async function pollTaskStatus(taskId) {
    const pollInterval = 2000; // 2 seconds
    const maxPolls = 300; // 10 minutes (300 polls × 2s)
    let pollCount = 0;
    let consecutiveErrors = 0;
    const maxConsecutiveErrors = 5;

    while (pollCount < maxPolls) {
        try {
            const response = await fetch(`${API_BASE_URL}/tasks/${taskId}`, {
                headers: {
                    'ngrok-skip-browser-warning': 'true'
                }
            });

            if (!response.ok) {
                throw new Error(`Status check failed: HTTP ${response.status}`);
            }

            const status = await response.json();

            // Reset error counter on successful poll
            consecutiveErrors = 0;

            // Update typing indicator with progress and elapsed time
            const elapsed = status.elapsed_seconds || 0;
            updateTypingIndicator(`${status.progress} (${elapsed}s)`);

            console.log(`📊 Poll ${pollCount + 1}: ${status.status} - ${status.progress} (${elapsed}s)`);

            // Check if task completed
            if (status.status === 'complete') {
                console.log(`✓ Task completed after ${elapsed}s`);
                return status.result;
            }

            // Check if task failed
            if (status.status === 'failed') {
                throw new Error(status.error || 'Task processing failed');
            }

            // Wait before next poll
            await sleep(pollInterval);
            pollCount++;

        } catch (error) {
            consecutiveErrors++;
            console.error(`Polling error (${consecutiveErrors}/${maxConsecutiveErrors}):`, error.message);

            // If too many consecutive errors, give up
            if (consecutiveErrors >= maxConsecutiveErrors) {
                throw new Error(`Polling failed after ${maxConsecutiveErrors} consecutive errors: ${error.message}`);
            }

            // Otherwise, wait and retry
            updateTypingIndicator(`Connection issue, retrying... (${consecutiveErrors}/${maxConsecutiveErrors})`);
            await sleep(pollInterval);
            pollCount++;
        }
    }

    throw new Error('Task timeout - exceeded maximum wait time (10 minutes)');
}

// Update typing indicator text with progress
function updateTypingIndicator(text) {
    const indicator = document.querySelector('.typing-indicator');
    if (indicator) {
        const contentDiv = indicator.querySelector('.typing-indicator-content');
        if (contentDiv) {
            // Keep the animated dots, just update the text
            const dotsSpan = contentDiv.querySelector('.typing-dots') || document.createElement('span');
            dotsSpan.className = 'typing-dots';
            dotsSpan.innerHTML = '<span>.</span><span>.</span><span>.</span>';

            contentDiv.innerHTML = `<span>${text}</span> `;
            contentDiv.appendChild(dotsSpan);
        }
    }
}

// Helper: Sleep function
function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// Add message to chat
function addMessageToChat(sender, content, isSystem = false) {
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

    messageDiv.innerHTML = `
        <div class="message-header">
            <div class="message-avatar">${sender === 'user' ? 'You' : 'AI'}</div>
            <div class="message-sender">${sender === 'user' ? 'You' : 'AI Assistant'}</div>
        </div>
        <div class="message-content">${escapeHtml(content)}</div>
        <div class="message-time">${time}</div>
    `;

    chatContainer.appendChild(messageDiv);
    chatContainer.scrollTop = chatContainer.scrollHeight;
}

// Show typing indicator
function showTypingIndicator() {
    const chatContainer = document.getElementById('chatContainer');

    const typingDiv = document.createElement('div');
    typingDiv.className = 'message ai typing-message';
    typingDiv.id = 'typingIndicator';

    typingDiv.innerHTML = `
        <div class="message-header">
            <div class="message-avatar">AI</div>
            <div class="message-sender">AI Assistant</div>
        </div>
        <div class="message-content">
            <div class="typing-indicator">
                <div class="typing-dot"></div>
                <div class="typing-dot"></div>
                <div class="typing-dot"></div>
            </div>
        </div>
    `;

    chatContainer.appendChild(typingDiv);
    chatContainer.scrollTop = chatContainer.scrollHeight;
}

// Remove typing indicator
function removeTypingIndicator() {
    const typingIndicator = document.getElementById('typingIndicator');
    if (typingIndicator) {
        typingIndicator.remove();
    }
}

// Utility: Escape HTML
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// Utility: Format data as table preview
function formatDataPreview(data, maxRows = 5) {
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

