/**
 * Main Taskpane Module
 * Orchestrates the Excel AI Agent application
 */

// Import context modules
import { buildQueryPayload, debugPrintPayload, validateQueryPayload } from './queryBuilder.js';
import { 
    updateExcelContext, 
    updateEnhancedContext, 
    setupExcelEventListeners,
    excelContext,
    enhancedContext 
} from './context/contextManager.js';

// Import UI modules
import { 
    addMessageToChat, 
    showTypingIndicator, 
    updateTypingIndicator, 
    removeTypingIndicator 
} from './ui/chatUI.js';
import { updateAPIStatus } from './ui/contextUI.js';
import { showClarificationModal } from './ui/clarificationModal.js';

// Import API client
import { 
    checkAPIHealth, 
    submitTask, 
    recoverTask, 
    pollTaskStatus,
    generateRequestId 
} from './api/apiClient.js';

// Import tool system
import { executeTool } from './tools/executor.js';
import { printToolRegistry } from './tools/registry.js';

// Import debug utilities
import { setupDebugShortcuts } from './debug/debugUtils.js';

// Import helpers
import { handleInputResize, handleInputKeydown } from './utils/helpers.js';

// Initialize Office addins
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log('Office.js initialized');
        initializeApp();
    }
});

/**
 * Initialize the application
 */
async function initializeApp() {
    // Set up event listeners
    document.getElementById('sendButton').addEventListener('click', handleSendMessage);
    document.getElementById('refreshButton').addEventListener('click', handleRefreshContext);
    document.getElementById('userInput').addEventListener('keydown', (e) => 
        handleInputKeydown(e, handleSendMessage)
    );
    document.getElementById('userInput').addEventListener('input', handleInputResize);

    // Load initial context (lightweight UI context only)
    await updateExcelContext();

    // Check backend connection
    await checkAPIConnection();

    // Set up Excel event listeners
    await setupExcelEventListeners();

    // Add debug keyboard shortcuts
    setupDebugShortcuts();

    // Print available tools on startup
    printToolRegistry();
}

/**
 * Handle refresh button click
 */
async function handleRefreshContext() {
    console.log('Manual refresh triggered - updating UI context');
    await updateExcelContext();
}

/**
 * Check API connection
 */
async function checkAPIConnection() {
    const isConnected = await checkAPIHealth();
    updateAPIStatus(isConnected);
    
    if (!isConnected) {
        console.error('API connection error');
    }
}

/**
 * Handle send message
 */
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
        // Capture full context only when sending query
        console.log('🔄 Capturing context for query...');
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

        // Generate unique request ID
        const requestId = generateRequestId();
        console.log(`📋 Request ID: ${requestId}`);

        // Submit task with retry logic
        const submitResult = await submitTask(
            message, 
            queryPayload, 
            requestId,
            (retryCount, maxRetries) => {
                updateTypingIndicator(`Connection error, retrying... (${retryCount}/${maxRetries})`);
            }
        );

        let taskId = submitResult?.taskId;

        // If submission failed, try to recover task
        if (!submitResult.success) {
            console.warn('⚠️ Initial POST failed, attempting task recovery...');
            updateTypingIndicator('Recovering task...');

            const recoveryResult = await recoverTask(requestId);
            if (recoveryResult.success) {
                taskId = recoveryResult.taskId;
            } else {
                throw new Error(submitResult.error || 'Failed to submit task and recovery failed. Please try again.');
            }
        }

        // Poll for task completion
        const pollResult = await pollTaskStatus(
            taskId,
            // Progress callback
            (progress, elapsed, status) => {
                updateTypingIndicator(`${progress} (${elapsed}s)`);
            },
            // Clarification callback
            async (status) => {
                const question = status.result?.clarification_question || "I need more information to proceed.";
                
                // Remove typing indicator while waiting
                removeTypingIndicator();
                
                // Show question in chat
                addMessageToChat('ai', question);
                
                // Get user answer
                const userAnswer = await showClarificationModal(question);
                
                if (userAnswer) {
                    addMessageToChat('user', userAnswer);
                    showTypingIndicator('Resuming with your answer...');
                }
                
                return userAnswer;
            }
        );

        // Handle result
        if (!pollResult.success) {
            throw new Error(pollResult.error);
        }

        const result = pollResult.result;

        // Remove typing indicator
        removeTypingIndicator();

        // Add AI response to chat - prioritize 'answer' field if present
        const displayMessage = result.answer && result.answer !== null && result.answer !== 'null'
            ? result.answer
            : result.message;
        addMessageToChat('ai', displayMessage);

        // Execute frontend tool calls if present
        await executeFrontendTools(result.toolCalls);

    } catch (error) {
        console.error('Error processing message:', error);
        removeTypingIndicator();
        addMessageToChat('ai', `Sorry, I encountered an error: ${error.message}`);
    } finally {
        sendButton.disabled = false;
        input.focus();
    }
}

/**
 * Execute frontend tool calls returned by backend
 */
async function executeFrontendTools(toolCalls) {
    if (!toolCalls || !Array.isArray(toolCalls) || toolCalls.length === 0) {
        return;
    }

    console.log(`🔧 Executing ${toolCalls.length} frontend tool(s)`);

    for (const toolCall of toolCalls) {
        try {
            const toolResult = await executeTool(toolCall.tool, toolCall.params);

            if (toolResult.success) {
                // Show success message based on tool type
                if (toolCall.tool === 'writeDataToRange') {
                    addMessageToChat('ai', `✓ Data written to ${toolResult.result.rangeAddress}`, true);
                } else if (toolCall.tool === 'createChart') {
                    addMessageToChat('ai', `✓ Chart "${toolResult.result.title}" created`, true);
                } else if (toolCall.tool === 'createTable') {
                    addMessageToChat('ai', `✓ Table "${toolResult.result.tableName}" created`, true);
                } else if (toolCall.tool === 'applyFormula') {
                    addMessageToChat('ai', `✓ Formula applied to ${toolResult.result.rangeAddress}`, true);
                } else if (toolCall.tool === 'formatRange') {
                    addMessageToChat('ai', `✓ Formatting applied to ${toolResult.result.rangeAddress}`, true);
                } else if (toolCall.tool === 'createNewSheet') {
                    addMessageToChat('ai', `✓ Sheet "${toolResult.result.sheetName}" created`, true);
                } else if (toolCall.tool === 'insertRows' || toolCall.tool === 'deleteRows') {
                    addMessageToChat('ai', `✓ Rows modified successfully`, true);
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
