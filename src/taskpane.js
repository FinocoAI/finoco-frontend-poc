/**
 * Main Taskpane Module - Updated for simplified backend
 * Orchestrates the Excel AI Agent application with conversation-based chat
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
    API_BASE_URL,
    checkAPIHealth, 
    initiateChat, 
    pollConversation,
    respondToConversation 
} from './api/apiClient.js';

// Import tool system
import { executeTool } from './tools/executor.js';
import { printToolRegistry } from './tools/registry.js';
import { executeBatchedTools } from './tools/batchExecutor.js';

if (typeof window !== 'undefined') {
    window.executeBatchedTools = executeBatchedTools;
  }

// Import debug utilities
import { setupDebugShortcuts } from './debug/debugUtils.js';

// Import helpers
import { handleInputResize, handleInputKeydown } from './utils/helpers.js';

// Track if already initialized to prevent duplicate initialization
let isInitialized = false;

// Initialize Office addins
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        const timestamp = new Date().toISOString();
        console.log(`🔍 DEBUG: Office.onReady fired at ${timestamp}`);
        
        if (isInitialized) {
            console.warn(`⚠️ DUPLICATE INITIALIZATION DETECTED! Office.onReady called again at ${timestamp}`);
            console.warn(`This indicates Excel is reloading the task pane!`);
            return; // Prevent re-initialization
        }
        
        isInitialized = true;
        console.log('Office.js initialized');
        initializeApp();
    }
});

/**
 * Initialize the application
 */
async function initializeApp() {
    const timestamp = new Date().toISOString();
    console.log(`🔍 DEBUG: initializeApp called at ${timestamp}`);
    
    // Set up event listeners
    document.getElementById('sendButton').addEventListener('click', handleSendMessage);
    document.getElementById('refreshButton')?.addEventListener('click', handleRefreshContext);
    document.getElementById('userInput').addEventListener('keydown', (e) => 
        handleInputKeydown(e, handleSendMessage)
    );
    document.getElementById('userInput').addEventListener('input', handleInputResize);

    // Set up new UI button listeners
    document.getElementById('historyButton')?.addEventListener('click', handleHistoryClick);
    document.getElementById('menuButton')?.addEventListener('click', handleMenuClick);
    document.getElementById('uploadButton')?.addEventListener('click', handleUploadClick);

    // Set up suggestion click listeners
    const suggestionItems = document.querySelectorAll('.suggestion-item');
    suggestionItems.forEach(item => {
        item.addEventListener('click', () => {
            const query = item.getAttribute('data-query');
            if (query) {
                document.getElementById('userInput').value = query;
                handleSendMessage();
            }
        });
    });

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
 * Handle history button click
 */
function handleHistoryClick() {
    console.log('History button clicked - feature to be implemented');
    // TODO: Implement history feature
}

/**
 * Handle menu button click
 */
function handleMenuClick() {
    console.log('Menu button clicked - feature to be implemented');
    // TODO: Implement menu feature
}

/**
 * Handle upload button click
 */
function handleUploadClick() {
    console.log('Upload button clicked - feature to be implemented');
    // TODO: Implement file upload feature
}

/**
 * Handle send message
 */
async function handleSendMessage() {
    const sendButton = document.getElementById('sendButton');
    
    // DEBUG: Track handleSendMessage calls
    const timestamp = new Date().toISOString();
    const callStack = new Error().stack;
    console.log(`🔍 DEBUG: handleSendMessage called at ${timestamp}`);
    console.log(`🔍 DEBUG: Button disabled state: ${sendButton?.disabled}`);
    console.log(`🔍 DEBUG: Call stack:`, callStack);
    
    // GUARD: Prevent duplicate execution (race condition fix)
    if (sendButton.disabled) {
        console.log('⏸️ Already processing a message, ignoring duplicate call');
        return;
    }
    
    // IMMEDIATELY disable to prevent race condition
    sendButton.disabled = true;
    
    const input = document.getElementById('userInput');
    const message = input.value.trim();

    if (!message) {
        sendButton.disabled = false;  // Re-enable if no message
        return;
    }

    // Clear input
    input.value = '';
    input.style.height = 'auto';

    // Hide welcome screen on first message
    const welcomeScreen = document.getElementById('welcomeScreen');
    if (welcomeScreen && !welcomeScreen.classList.contains('hidden')) {
        welcomeScreen.classList.add('hidden');
    }

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

        // Initiate chat
        const initResult = await initiateChat(
            message, 
            queryPayload,
            (retryCount, maxRetries) => {
                updateTypingIndicator(`Connection error, retrying... (${retryCount}/${maxRetries})`);
            }
        );

        if (!initResult.success) {
            throw new Error(initResult.error || 'Failed to initiate chat');
        }

        const conversationId = initResult.conversationId;
        console.log(`💬 Conversation ID: ${conversationId}`);

        // Poll for conversation completion
        const pollResult = await pollConversation(
            conversationId,
            // Progress callback
            (explanation, elapsed, status) => {
                updateTypingIndicator(`${explanation} (${elapsed.toFixed(0)}s)`);
            },
            // Clarification callback
            async (question, message) => {
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
            },
            // Tool execution callback
            async (toolCalls) => {
                return await executeFrontendTools(toolCalls);
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
            : result.explanation;
        
        if (displayMessage) {
            addMessageToChat('ai', displayMessage);
        }

    } catch (error) {
        console.error('Error processing message:', error);
        removeTypingIndicator();
        
        // Show user-friendly error message
        let errorMessage = `Sorry, I encountered an error: ${error.message}`;
        
        // Provide specific guidance for connection errors
        if (error.message.includes('Failed to initiate chat') || 
            error.message.includes('Failed to fetch') ||
            error.message.includes('Connection error')) {
            errorMessage += '\n\n⚠️ Cannot reach backend server. Please check:\n' +
                           '• Backend is running (python backend/main.py)\n' +
                           '• Ngrok tunnel is active\n' +
                           '• API_BASE_URL in apiClient.js matches your ngrok URL';
        }
        
        addMessageToChat('ai', errorMessage);
    } finally {
        sendButton.disabled = false;
        input.focus();
    }
}

/**
 * Execute frontend tool calls returned by backend
 * Returns array of ALL tool results (READ and WRITE, successes and failures) to send back to backend
 * 
 * USES BATCHED EXECUTION: All WRITE tools are executed in a single Excel.run() context
 * for maximum efficiency and reliability. READ tools are executed sequentially.
 */
async function executeFrontendTools(toolCalls) {
    if (!toolCalls || !Array.isArray(toolCalls) || toolCalls.length === 0) {
        return [];
    }

    console.log(`🔧 Executing ${toolCalls.length} frontend tool(s)`);

    // List of READ tools that require continuation
    const READ_TOOLS = ['getFullRangeData', 'getColumnData', 'getTableData', 'getFormulasInRange', 
                        'getCellPrecedents', 'getCellDependents', 'getRelatedData', 'getChartSourceData',
                        'searchValues', 'getNamedRangeData'];
    
    // Separate READ and WRITE tools
    const readTools = toolCalls.filter(tc => READ_TOOLS.includes(tc.tool));
    const writeTools = toolCalls.filter(tc => !READ_TOOLS.includes(tc.tool));
    
    const toolResults = [];

    // PHASE 1: Execute WRITE tools in batched mode (single Excel.run)
    if (writeTools.length > 0) {
        console.log(`📦 Batching ${writeTools.length} WRITE tool(s) into single Excel.run()`);
        
        try {
            const batchResults = await executeBatchedTools(writeTools);
            
            // Display results AND collect for backend
            for (let i = 0; i < batchResults.length; i++) {
                const result = batchResults[i];
                const toolCall = writeTools[i];
                
                if (result.success) {
                    // Add to results for backend
                    toolResults.push({
                        tool: toolCall.tool,
                        params: toolCall.params,
                        result: result.result,
                        success: true
                    });
                    
                    // WRITE tool - show appropriate success message
                    if (toolCall.tool === 'writeDataToRange') {
                        addMessageToChat('ai', `✓ Data written to ${result.result.rangeAddress}`, true);
                    } else if (toolCall.tool === 'createChart') {
                        addMessageToChat('ai', `✓ Chart created`, true);
                    } else if (toolCall.tool === 'createTable') {
                        addMessageToChat('ai', `✓ Table "${result.result.tableName}" created`, true);
                    } else if (toolCall.tool === 'applyFormula') {
                        addMessageToChat('ai', `✓ Formula applied to ${result.result.address}`, true);
                    } else if (toolCall.tool === 'formatRange') {
                        addMessageToChat('ai', `✓ Formatting applied to ${result.result.address}`, true);
                    } else if (toolCall.tool === 'createNewSheet') {
                        addMessageToChat('ai', `✓ Sheet "${result.result.sheetName}" created`, true);
                    } else if (toolCall.tool === 'addCellNote') {
                        addMessageToChat('ai', `✓ Citation added to ${result.result.cellAddress}`, true);
                    } else if (toolCall.tool === 'insertRows' || toolCall.tool === 'deleteRows') {
                        addMessageToChat('ai', `✓ Rows modified successfully`, true);
                    } else {
                        addMessageToChat('ai', `✓ ${toolCall.tool} completed`, true);
                    }
                } else {
                    // Add error to results for backend
                    toolResults.push({
                        tool: toolCall.tool,
                        params: toolCall.params,
                        result: null,
                        success: false,
                        error: result.error,
                        errorType: result.errorType,
                        errorDetails: result.errorDetails
                    });
                    
                    addMessageToChat('ai', `⚠️ ${toolCall.tool} failed: ${result.error}`, true);
                }
            }
        } catch (error) {
            console.error(`Batch execution error:`, error);
            addMessageToChat('ai', `⚠️ Batch execution failed: ${error.message}`, true);
            
            // Add batch-level error for all write tools
            for (const toolCall of writeTools) {
                toolResults.push({
                    tool: toolCall.tool,
                    params: toolCall.params,
                    result: null,
                    success: false,
                    error: `Batch execution failed: ${error.message}`,
                    errorType: error.name || 'BatchExecutionError'
                });
            }
        }
    }

    // PHASE 2: Execute READ tools sequentially (need results)
    for (const toolCall of readTools) {
        try {
            const toolResult = await executeTool(toolCall.tool, toolCall.params);

            if (toolResult.success) {
                // Store result to send back to backend
                toolResults.push({
                    tool: toolCall.tool,
                    params: toolCall.params,
                    result: toolResult.result,
                    success: true
                });
                addMessageToChat('ai', `✓ ${toolCall.tool} completed`, true);
            } else {
                addMessageToChat('ai', `⚠️ ${toolCall.tool} failed: ${toolResult.error}`, true);
                toolResults.push({
                    tool: toolCall.tool,
                    params: toolCall.params,
                    result: null,
                    success: false,
                    error: toolResult.error
                });
            }
        } catch (error) {
            console.error(`Tool execution error:`, error);
            addMessageToChat('ai', `⚠️ Failed to execute ${toolCall.tool}`, true);
            toolResults.push({
                tool: toolCall.tool,
                params: toolCall.params,
                result: null,
                success: false,
                error: error.message,
                errorType: error.name || 'Error'
            });
        }
    }

    // Return ALL results (READ and WRITE, successes and failures)
    console.log(`📋 Returning ${toolResults.length} tool result(s) to backend`);
    return toolResults;
}
