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
    getOrCreateConversation,
    sendMessage,
    pollAgentRun
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

// Import preview edit mode
import { activatePreviewEditMode } from './tools/previewEditMode.js';

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
    document.getElementById('newConversationButton')?.addEventListener('click', handleNewConversation);
    document.getElementById('menuButton')?.addEventListener('click', handleMenuClick);
    document.getElementById('uploadButton')?.addEventListener('click', handleUploadClick);
    document.getElementById('traceDependenciesBtn')?.addEventListener('click', handleTraceDependencies);
    document.getElementById('previewEditBtn')?.addEventListener('click', handlePreviewEdit);
    document.getElementById('findErrorsBtn')?.addEventListener('click', handleFindErrors);
    document.getElementById('findCircularRefsBtn')?.addEventListener('click', handleFindCircularReferences);

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
    
    // Update trace button visibility on initial load
    await updateTraceButtonVisibility();

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
    await updateTraceButtonVisibility(); // Update trace button visibility on refresh
}

/**
 * Update trace button visibility based on selection
 * Shows button only if a single cell is selected
 */
export async function updateTraceButtonVisibility() {
    const traceBtn = document.getElementById('traceDependenciesBtn');
    if (!traceBtn) return;

    try {
        await Excel.run(async (context) => {
            const selectedRanges = context.workbook.getSelectedRanges();
            selectedRanges.load('areaCount');
            await context.sync();

            // Only show for single area selections
            if (selectedRanges.areaCount === 1) {
                const areas = selectedRanges.areas;
                areas.load('items');
                await context.sync();

                const singleArea = areas.items[0];
                singleArea.load('rowCount, columnCount');
                await context.sync();

                // Show button only for single cell (1x1 selection)
                if (singleArea.rowCount === 1 && singleArea.columnCount === 1) {
                    traceBtn.classList.remove('hidden');
                } else {
                    traceBtn.classList.add('hidden');
                }
            } else {
                traceBtn.classList.add('hidden');
            }
        });
    } catch (error) {
        console.error('Error updating trace button visibility:', error);
        traceBtn.classList.add('hidden');
    }
}

/**
 * Handle trace dependencies button click
 * Directly executes trace tools without going through LLM
 */
async function handleTraceDependencies() {
    console.log('📊 Trace Dependencies button clicked');
    
    const traceBtn = document.getElementById('traceDependenciesBtn');
    
    try {
        // Disable button during execution
        traceBtn.disabled = true;
        traceBtn.textContent = '⏳ Tracing...';

        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            const sheet = range.worksheet;
            
            range.load('address');
            sheet.load('name');
            await context.sync();
            
            // Extract cell address without sheet name
            let cellAddress = range.address;
            if (cellAddress.includes('!')) {
                cellAddress = cellAddress.split('!')[1];
            }
            const sheetName = sheet.name;
            
            console.log(`📊 Tracing dependencies for ${sheetName}!${cellAddress}`);
            
            // Import and execute traceDependencyGraph tool
            const { execute: traceExecute } = await import('./tools/read/traceDependencyGraph.js');
            const graphData = await traceExecute({
                sheetName: sheetName,
                address: cellAddress,  // Note: parameter name is 'address', not 'cellAddress'
                maxDepth: 3
            });
            
            console.log('✅ Dependency graph data retrieved:', graphData);
            
            // Import and execute displayDependencyGraph tool
            const { execute: displayExecute } = await import('./tools/write/displayDependencyGraph.js');
            await displayExecute({
                graphData: graphData,
                displayOptions: {
                    title: `Dependencies for ${sheetName}!${cellAddress}`,
                    highlightCritical: true
                }
            });
            
            console.log('✅ Dependency graph displayed successfully');
            
            // Show success message in chat
            addMessageToChat('ai', `✓ Dependency graph displayed for ${sheetName}!${cellAddress}`, true);
        });
        
    } catch (error) {
        console.error('❌ Failed to trace dependencies:', error);
        addMessageToChat('ai', `⚠️ Failed to trace dependencies: ${error.message}`, true);
    } finally {
        // Re-enable button
        traceBtn.disabled = false;
        traceBtn.textContent = '📊 Trace';
    }
}

/**
 * Handle preview edit button click
 * Activates What-If Analysis mode for previewing cell changes
 */
async function handlePreviewEdit() {
    console.log('🔍 Preview Edit (What-If) button clicked');
    
    const previewBtn = document.getElementById('previewEditBtn');
    
    try {
        // Disable button during execution
        const originalText = previewBtn.textContent;
        previewBtn.disabled = true;
        previewBtn.textContent = '⏳ Starting...';

        // Activate preview edit mode
        await activatePreviewEditMode();
        
    } catch (error) {
        console.error('❌ Failed to activate preview mode:', error);
        // Error notification is handled inside activatePreviewEditMode
    } finally {
        // Re-enable button
        previewBtn.disabled = false;
        previewBtn.textContent = '🔍 What-If';
    }
}

/**
 * Handle find errors button click
 * Directly executes error detection without LLM
 */
async function handleFindErrors() {
    console.log('⚠️ Find Errors button clicked');
    
    const findErrorsBtn = document.getElementById('findErrorsBtn');
    
    try {
        // Disable button during execution
        findErrorsBtn.disabled = true;
        findErrorsBtn.textContent = '⏳ Scanning...';

        // Import and execute findErrors tool
        const { execute } = await import('./tools/read/findErrors.js');
        const result = await execute({
            scope: 'sheet',  // Search current sheet
            tracePropagation: true
        });
        
        console.log('✅ Error scan complete:', result);
        
        // Build user-friendly message
        if (result.errorCount === 0) {
            addMessageToChat('ai', `✅ No errors found! Your worksheet is clean.`, true);
        } else {
            let message = `⚠️ Found ${result.errorCount} error(s) in current sheet:\n\n`;
            
            // Group by error type
            const typeBreakdown = [];
            for (const [type, count] of Object.entries(result.summary.byType)) {
                if (count > 0) {
                    typeBreakdown.push(`${type}: ${count}`);
                }
            }
            message += typeBreakdown.join(', ') + '\n\n';
            
            // Show first few errors
            const errorsToShow = result.errors.slice(0, 5);
            message += 'Error locations:\n';
            errorsToShow.forEach(err => {
                message += `• ${err.fullAddress} - ${err.errorType}`;
                if (err.propagationCount > 0) {
                    message += ` (affects ${err.propagationCount} cells)`;
                }
                message += '\n';
            });
            
            if (result.errorCount > 5) {
                message += `\n...and ${result.errorCount - 5} more error(s)`;
            }
            
            addMessageToChat('ai', message, true);
        }
        
    } catch (error) {
        console.error('❌ Failed to find errors:', error);
        addMessageToChat('ai', `⚠️ Failed to scan for errors: ${error.message}`, true);
    } finally {
        // Re-enable button
        findErrorsBtn.disabled = false;
        findErrorsBtn.textContent = '⚠️ Find Errors';
    }
}

/**
 * Handle find circular references button click
 * Directly executes circular reference detection without LLM
 */
async function handleFindCircularReferences() {
    console.log('🔄 Find Circular References button clicked');
    
    const findCircularBtn = document.getElementById('findCircularRefsBtn');
    
    try {
        // Disable button during execution
        findCircularBtn.disabled = true;
        findCircularBtn.textContent = '⏳ Scanning...';

        // Import and execute findCircularReferences tool
        const { execute } = await import('./tools/read/findCircularReferences.js');
        const result = await execute({
            scope: 'sheet'  // Search current sheet
        });
        
        console.log('✅ Circular reference scan complete:', result);
        
        // Build user-friendly message
        if (result.circularCount === 0) {
            addMessageToChat('ai', `✅ No circular references found!`, true);
        } else {
            let message = `🔄 Found ${result.circularCount} circular reference(s):\n\n`;
            
            result.circularReferences.forEach((circ, index) => {
                message += `${index + 1}. ${circ.chain.join(' → ')}\n`;
                message += `   Formula: ${circ.formula}\n\n`;
            });
            
            message += `💡 Tip: Circular references can be intentional (like iterative calculations) or errors. Review each one to ensure it's working as expected.`;
            
            addMessageToChat('ai', message, true);
        }
        
    } catch (error) {
        console.error('❌ Failed to find circular references:', error);
        addMessageToChat('ai', `⚠️ Failed to scan for circular references: ${error.message}`, true);
    } finally {
        // Re-enable button
        findCircularBtn.disabled = false;
        findCircularBtn.textContent = '🔄 Circular Refs';
    }
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
 * Handle new conversation button click
 * Clears the current conversation and starts fresh
 */
function handleNewConversation() {
    console.log('🔄 Starting new conversation...');
    
    // Clear the stored conversation ID (use correct key!)
    localStorage.removeItem('warren_conversation_id');
    
    // Clear the chat UI (keep it empty until user sends first message)
    const chatContainer = document.getElementById('chatContainer');
    if (chatContainer) {
        chatContainer.innerHTML = '';
    }
    
    console.log('✅ New conversation ready - warren_conversation_id cleared, UI reset');
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

        // Get or create conversation
        const convResult = await getOrCreateConversation();
        if (!convResult.success) {
            throw new Error(convResult.error || 'Failed to get conversation');
        }
        
        const conversationId = convResult.conversationId;
        console.log(`💬 Conversation ID: ${conversationId}`);

        // Send message to conversation
        const sendResult = await sendMessage(
            conversationId,
            message, 
            queryPayload,
            (retryCount, maxRetries) => {
                updateTypingIndicator(`Connection error, retrying... (${retryCount}/${maxRetries})`);
            }
        );

        if (!sendResult.success) {
            throw new Error(sendResult.error || 'Failed to send message');
        }

        const agentRunId = sendResult.agentRunId;
        console.log(`🤖 Agent Run ID: ${agentRunId}`);

        // Poll for agent run completion
        const pollResult = await pollAgentRun(
            agentRunId,
            // Progress callback
            (explanation, elapsed, status) => {
                updateTypingIndicator(`${explanation} (${elapsed.toFixed(0)}s)`);
            },
            // Clarification callback
            async (question, message) => {
                // Remove typing indicator while waiting
                removeTypingIndicator();
                
                // Show question in chat (display title for structured forms, full text for simple questions)
                let displayText = question;
                try {
                    if (typeof question === 'string' && question.trim().startsWith('{')) {
                        const parsed = JSON.parse(question);
                        if (parsed.title) {
                            displayText = parsed.title;
                        }
                    }
                } catch (e) {
                    // Not JSON, use as-is
                }
                addMessageToChat('ai', displayText);
                
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

        // Add AI response to chat
        // Backend returns answer at top level when status=SUCCESS
        const displayMessage = result.answer || result.explanation || result.message || 'Task completed';
        
        addMessageToChat('ai', displayMessage);

    } catch (error) {
        console.error('Error processing message:', error);
        removeTypingIndicator();
        
        // Show user-friendly error message
        let errorMessage = `Sorry, I encountered an error: ${error.message}`;
        
        // Provide specific guidance for connection errors
        if (error.message.includes('Failed to') || 
            error.message.includes('Failed to fetch') ||
            error.message.includes('Connection error')) {
            errorMessage += '\n\n⚠️ Cannot reach backend server. Please check:\n' +
                           '• BackendV2 is running (cd backendV2 && uvicorn app.main:app)\n' +
                           '• Ngrok tunnel is active and pointing to port 8000\n' +
                           '• API_BASE_URL in apiClient.js matches your ngrok URL + /api/v1';
        }
        
        addMessageToChat('ai', errorMessage);
    } finally {
        sendButton.disabled = false;
        input.focus();
    }
}

/**
 * Execute frontend tool calls returned by backend
 * Returns object with tool_results and updated_context (captured after execution)
 * 
 * EFFICIENCY: Tool results EXCLUDE params on success (agent already has them in context)
 * On failure, params ARE included to help with debugging
 * CRITICAL: Returns updated Excel context so backend has fresh state after tool execution
 * 
 * USES BATCHED EXECUTION: All WRITE tools are executed in a single Excel.run() context
 * for maximum efficiency and reliability. READ tools are executed sequentially.
 */
async function executeFrontendTools(toolCalls) {
    if (!toolCalls || !Array.isArray(toolCalls) || toolCalls.length === 0) {
        return {
            tool_results: [],
            updated_context: null
        };
    }

    console.log(`🔧 Executing ${toolCalls.length} frontend tool(s)`);
    
    // Transform backendV2 format to frontend format
    // Backend sends: { id, tool_name, params, target }
    // Frontend expects: { tool, params }
    const transformedToolCalls = toolCalls.map(tc => ({
        tool: tc.tool_name || tc.tool, // Support both formats
        params: tc.params,
        id: tc.id // Keep ID for tracking
    }));

    // List of READ tools that require continuation (execute individually, not batched)
    const READ_TOOLS = [
        // Progressive loading tools (preferred)
        'getSheetMetadata', 'getRangePreview',
        // Full data loading tools
        'getFullRangeData', 'getColumnData', 'getTableData', 'getFormulasInRange', 
        'getNamedRangeData', 'getChartSourceData',
        // Dependency and analysis tools
        'getCellPrecedents', 'getCellDependents', 'getRelatedData', 'traceDependencyGraph', 
        'searchValues', 'findErrors', 'findCircularReferences'
    ];
    
    // Separate READ and WRITE tools
    const readTools = transformedToolCalls.filter(tc => READ_TOOLS.includes(tc.tool));
    const writeTools = transformedToolCalls.filter(tc => !READ_TOOLS.includes(tc.tool));
    
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
                    // SUCCESS: Only send result, not params (agent already has params in context)
                    toolResults.push({
                        tool: toolCall.tool,
                        tool_call_id: toolCall.id,
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
                    } else if (toolCall.tool === 'displayDependencyGraph') {
                        addMessageToChat('ai', `✓ Dependency graph displayed`, true);
                    } else {
                        addMessageToChat('ai', `✓ ${toolCall.tool} completed`, true);
                    }
                } else {
                    // FAILURE: Include params for debugging
                    toolResults.push({
                        tool: toolCall.tool,
                        params: toolCall.params,  // Keep for debugging failures
                        tool_call_id: toolCall.id,
                        result: null,
                        success: false,
                        error: result.error,
                        errorType: result.errorType,
                        errorDetails: result.errorDetails,
                        userCancelled: result.userCancelled,
                        feedback: result.feedback  // Include feedback for agent
                    });
                    
                    // Show user-friendly message (use feedback if available for user cancellations)
                    const displayMessage = result.userCancelled 
                        ? `❌ Operation cancelled by user` 
                        : `⚠️ ${toolCall.tool} failed: ${result.error}`;
                    addMessageToChat('ai', displayMessage, true);
                }
            }
        } catch (error) {
            console.error(`Batch execution error:`, error);
            addMessageToChat('ai', `⚠️ Batch execution failed: ${error.message}`, true);
            
            // FAILURE: Add batch-level error for all write tools (params included for debugging)
            for (const toolCall of writeTools) {
                toolResults.push({
                    tool: toolCall.tool,
                    params: toolCall.params,  // Keep for debugging failures
                    tool_call_id: toolCall.id,
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
                // SUCCESS: Only send result, not params (agent already has params in context)
                toolResults.push({
                    tool: toolCall.tool,
                    tool_call_id: toolCall.id,
                    result: toolResult.result,
                    success: true
                });
                addMessageToChat('ai', `✓ ${toolCall.tool} completed`, true);
            } else {
                addMessageToChat('ai', `⚠️ ${toolCall.tool} failed: ${toolResult.error}`, true);
                // FAILURE: Include params for debugging
                toolResults.push({
                    tool: toolCall.tool,
                    params: toolCall.params,  // Keep for debugging failures
                    tool_call_id: toolCall.id,
                    result: null,
                    success: false,
                    error: toolResult.error
                });
            }
        } catch (error) {
            console.error(`Tool execution error:`, error);
            addMessageToChat('ai', `⚠️ Failed to execute ${toolCall.tool}`, true);
            // FAILURE: Include params for debugging
            toolResults.push({
                tool: toolCall.tool,
                params: toolCall.params,  // Keep for debugging failures
                tool_call_id: toolCall.id,
                result: null,
                success: false,
                error: error.message,
                errorType: error.name || 'Error'
            });
        }
    }

    // Capture updated Excel context after all tools executed
    console.log(`📸 Capturing updated Excel context after tool execution...`);
    let updatedContext = null;
    try {
        await updateExcelContext();
        await updateEnhancedContext();
        updatedContext = {
            initialContext: enhancedContext.initialContext,
            userSelection: enhancedContext.userSelection
        };
        console.log(`✓ Updated context captured: ${updatedContext.initialContext.sheets?.length || 0} sheets`);
    } catch (error) {
        console.error(`⚠️ Failed to capture updated context:`, error);
        // Continue without updated context - backend will use stale context with warning
    }

    // Return tool results AND updated context
    console.log(`📋 Returning ${toolResults.length} tool result(s) with updated context to backend`);
    return {
        tool_results: toolResults,
        updated_context: updatedContext
    };
}
