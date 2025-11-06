/**
 * API Client Module - BackendV2 Integration
 * Handles all backend communication with the new multi-agent architecture
 * Flow: Create Conversation → Send Message → Poll Agent Run → Respond with Tool Results
 */

export const API_BASE_URL = 'https://667df4e1a782.ngrok-free.app/api/v1';

/**
 * Check API health status
 */
export async function checkAPIHealth() {
    try {
        const response = await fetch(`${API_BASE_URL}/health/live`, {
            headers: {
                'ngrok-skip-browser-warning': 'true'
            }
        });
        return response.ok;
    } catch (error) {
        console.error('API health check error:', error);
        return false;
    }
}

/**
 * Get unique identifier for current workbook
 * This ensures each workbook has its own conversation (prevents context mixing)
 */
async function getWorkbookIdentifier() {
    try {
        return await Excel.run(async (context) => {
            const workbook = context.workbook;
            workbook.load('name');
            await context.sync();
            
            const workbookName = workbook.name;
            
            // Create a stable session ID for THIS workbook instance
            // Use sessionStorage so each browser tab/window gets unique ID
            const sessionKey = `warren_wb_session_${workbookName}`;
            let sessionId = sessionStorage.getItem(sessionKey);
            
            if (!sessionId) {
                sessionId = `${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
                sessionStorage.setItem(sessionKey, sessionId);
                console.log(`📝 New workbook session created: ${sessionId}`);
            }
            
            // Return unique identifier: sanitized workbook name + session ID
            const sanitizedName = workbookName.replace(/[^a-zA-Z0-9]/g, '_');
            const workbookId = `${sanitizedName}_${sessionId}`;
            return workbookId;
        });
    } catch (error) {
        console.error('Failed to get workbook identifier:', error);
        // Fallback: use timestamp-based ID
        return `fallback_${Date.now()}`;
    }
}

/**
 * Helper: Get current workbook name
 */
async function getWorkbookName() {
    try {
        return await Excel.run(async (context) => {
            const workbook = context.workbook;
            workbook.load('name');
            await context.sync();
            return workbook.name;
        });
    } catch (error) {
        console.error('Failed to get workbook name:', error);
        return 'Unknown Workbook';
    }
}

/**
 * Get or create a conversation for THIS specific workbook
 * Each workbook gets its own conversation stored in localStorage
 * This prevents context contamination when multiple Excel files are open
 */
export async function getOrCreateConversation() {
    try {
        // Step 1: Identify WHICH workbook we're in
        const workbookId = await getWorkbookIdentifier();
        const workbookName = await getWorkbookName();
        
        // Step 2: Use workbook-specific localStorage key
        const storageKey = `warren_conversation_${workbookId}`;
        let conversationId = localStorage.getItem(storageKey);
        
        console.log(`📂 Workbook: ${workbookName}`);
        console.log(`🔑 Workbook ID: ${workbookId}`);
        console.log(`💾 Storage key: ${storageKey}`);
        
        // Step 3: Check if conversation exists on backend
        if (conversationId) {
            console.log(`🔍 Found stored conversation: ${conversationId}, verifying...`);
            
            const response = await fetch(`${API_BASE_URL}/conversations/${conversationId}`, {
                headers: {
                    'ngrok-skip-browser-warning': 'true'
                }
            });
            
            if (response.ok) {
                const conv = await response.json();
                console.log(`✓ Using existing conversation: ${conversationId}`);
                console.log(`  📨 Messages: ${conv.message_count || 0}`);
                return { 
                    success: true, 
                    conversationId,
                    workbookId,
                    workbookName,
                    isExisting: true
                };
            } else {
                console.log(`⚠️ Conversation ${conversationId} not found on backend, creating new one`);
                localStorage.removeItem(storageKey);
                conversationId = null;
            }
        }
        
        // Step 4: Create new conversation via backend API
        console.log(`📤 Creating new conversation for: ${workbookName}`);
        
        const response = await fetch(`${API_BASE_URL}/conversations`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'ngrok-skip-browser-warning': 'true'
            },
            body: JSON.stringify({
                user_id: "1",
                title: "Untitled Conversation",  // Will be auto-updated from first user message
                meta: {
                    workbookName: workbookName,
                    workbookId: workbookId,
                    createdAt: new Date().toISOString()
                }
            })
        });

        if (!response.ok) {
            throw new Error(`Failed to create conversation: HTTP ${response.status}`);
        }

        const conversation = await response.json();
        conversationId = conversation.id;
        
        // Step 5: Store in workbook-specific key
        localStorage.setItem(storageKey, conversationId);
        
        console.log(`✓ Created new conversation: ${conversationId}`);
        console.log(`  💾 Stored in: ${storageKey}`);
        
        return { 
            success: true, 
            conversationId,
            workbookId,
            workbookName,
            isExisting: false
        };

    } catch (error) {
        console.error('❌ Failed to get/create conversation:', error);
        return { success: false, error: error.message };
    }
}

/**
 * Clear conversation for current workbook (for "New" button)
 * Only clears the conversation for THIS workbook, not others
 */
export async function clearCurrentConversation() {
    try {
        const workbookId = await getWorkbookIdentifier();
        const storageKey = `warren_conversation_${workbookId}`;
        
        const oldConversationId = localStorage.getItem(storageKey);
        localStorage.removeItem(storageKey);
        
        console.log(`✅ Cleared conversation for current workbook`);
        console.log(`  🔑 Workbook ID: ${workbookId}`);
        console.log(`  🗑️ Removed conversation: ${oldConversationId}`);
        
        return { success: true, workbookId };
    } catch (error) {
        console.error('Failed to clear conversation:', error);
        return { success: false, error: error.message };
    }
}

/**
 * Send a message to the conversation
 * Returns the agent_run that was created
 */
export async function sendMessage(conversationId, message, queryPayload, onRetry) {
    const maxRetries = 3;
    let retryCount = 0;

    while (retryCount < maxRetries) {
        try {
            console.log(`📤 Sending message (attempt ${retryCount + 1}/${maxRetries})...`);
            console.log(`  💬 Conversation: ${conversationId}`);
            console.log(`  📝 Message: "${message.substring(0, 50)}${message.length > 50 ? '...' : ''}"`);

            const response = await fetch(`${API_BASE_URL}/messages/${conversationId}`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'ngrok-skip-browser-warning': 'true'
                },
                body: JSON.stringify({
                    role: 'USER',
                    content: message,
                    meta: {
                        initialContext: queryPayload.initialContext,
                        userSelection: queryPayload.userSelection
                    }
                })
            });

            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }

            const result = await response.json();
            const agentRunId = result.agent_run?.id;

            if (!agentRunId) {
                throw new Error('No agent_run returned from message');
            }

            console.log(`✓ Message sent successfully`);
            console.log(`  🤖 Agent Run ID: ${agentRunId}`);
            console.log(`  💬 Conversation ID: ${conversationId}`);
            console.log(`  📊 Initial Status: ${result.agent_run.status}`);
            
            return { 
                success: true, 
                agentRunId,
                conversationId,
                message: result.message,
                agentRun: result.agent_run
            };

        } catch (error) {
            retryCount++;
            console.error(`❌ Send message attempt ${retryCount} failed:`, error.message);

            if (retryCount >= maxRetries) {
                return { 
                    success: false, 
                    error: `Failed to send message after ${maxRetries} attempts: ${error.message}` 
                };
            }

            // Wait before retry
            const waitTime = 500;
            console.log(`⏳ Retrying in ${waitTime}ms...`);
            if (onRetry) {
                onRetry(retryCount, maxRetries);
            }
            await sleep(waitTime);
        }
    }
}

/**
 * Poll agent run status
 */
export async function pollAgentRun(agentRunId, onProgress, onClarificationNeeded, onToolExecutionNeeded) {
    const pollInterval = 2000; // 2 seconds
    const maxPolls = 300; // 10 minutes
    let pollCount = 0;
    let consecutiveErrors = 0;
    const maxConsecutiveErrors = 5;

    console.log(`🔄 Starting to poll agent run: ${agentRunId}`);

    while (pollCount < maxPolls) {
        try {
            const response = await fetch(`${API_BASE_URL}/agents/${agentRunId}/poll`, {
                headers: {
                    'ngrok-skip-browser-warning': 'true'
                }
            });

            if (!response.ok) {
                throw new Error(`Status check failed: HTTP ${response.status}`);
            }

            const data = await response.json();
            consecutiveErrors = 0; // Reset on success

            // Calculate elapsed time (estimate based on poll count)
            const elapsed = pollCount * (pollInterval / 1000);
            
            // Update progress with current status
            if (onProgress) {
                const statusMessage = data.status === 'RUNNING' ? 'Processing...' : data.status;
                onProgress(statusMessage, elapsed, data.status);
            }

            console.log(`📊 Poll ${pollCount + 1} [${agentRunId.substring(0, 8)}...]: ${data.status}`);

            // Handle different statuses
            if (data.status === 'SUCCESS') {
                console.log(`✓ Agent run completed successfully`);
                // Backend returns answer at top level, not in final_output
                return { success: true, result: data };
            }

            if (data.status === 'ERROR') {
                console.log(`❌ Agent run failed`);
                return { 
                    success: false, 
                    error: data.error_message || 'Agent run failed'
                };
            }

            if (data.status === 'NEEDS_CLARIFICATION') {
                console.log(`⏸️ Agent needs clarification`);
                
                if (!onClarificationNeeded) {
                    return { 
                        success: false, 
                        error: 'Clarification needed but no handler provided' 
                    };
                }

                // Extract question from output_data
                const question = data.output_data?.question || "I need more information to proceed.";
                
                const answer = await onClarificationNeeded(question, data);
                
                if (!answer) {
                    return { success: false, error: 'Cancelled by user' };
                }

                // Submit answer and continue polling
                const submitSuccess = await respondToAgentRun(agentRunId, null, answer);
                if (!submitSuccess) {
                    return { success: false, error: 'Failed to submit clarification' };
                }

                // Continue polling
                pollCount++;
                await sleep(500); // Short delay before resuming
                continue;
            }

            if (data.status === 'NEEDS_TOOL_EXECUTION') {
                console.log(`🔧 Frontend tools needed`);
                
                if (!onToolExecutionNeeded) {
                    return { 
                        success: false, 
                        error: 'Tool execution needed but no handler provided' 
                    };
                }

                // Execute tools and get results + updated context
                const toolCalls = data.tool_calls || [];
                const executionResult = await onToolExecutionNeeded(toolCalls);
                
                // Extract tool results and updated context
                const toolResults = executionResult.tool_results || [];
                const updatedContext = executionResult.updated_context || null;
                
                // Always send results back to backend (with updated context)
                console.log(`📊 Tool execution complete: ${toolResults.length} results, context updated: ${updatedContext !== null}`);
                
                const submitSuccess = await respondToAgentRun(agentRunId, toolResults, null, updatedContext);
                
                if (!submitSuccess) {
                    return { success: false, error: 'Failed to submit tool results' };
                }
                
                // Continue polling for next response
                pollCount++;
                await sleep(500);
                continue;
            }

            // Still processing (RUNNING or PENDING)
            await sleep(pollInterval);
            pollCount++;

        } catch (error) {
            consecutiveErrors++;
            console.error(`Polling error (${consecutiveErrors}/${maxConsecutiveErrors}):`, error.message);

            if (consecutiveErrors >= maxConsecutiveErrors) {
                return { 
                    success: false, 
                    error: `Polling failed after ${maxConsecutiveErrors} consecutive errors: ${error.message}` 
                };
            }

            // Update UI and retry
            if (onProgress) {
                onProgress(
                    `Connection issue, retrying... (${consecutiveErrors}/${maxConsecutiveErrors})`,
                    0,
                    'retrying'
                );
            }
            await sleep(pollInterval);
            pollCount++;
        }
    }

    return { 
        success: false, 
        error: 'Timeout - exceeded maximum wait time (10 minutes)' 
    };
}

/**
 * Respond to agent run (tool results OR clarification answer)
 * @param {string} agentRunId - The agent run ID
 * @param {Array} toolResults - Tool execution results (if any)
 * @param {string} clarificationAnswer - Clarification answer (if any)
 * @param {Object} updatedContext - Updated Excel context after tool execution (if any)
 */
export async function respondToAgentRun(agentRunId, toolResults = null, clarificationAnswer = null, updatedContext = null) {
    try {
        const body = {};
        
        if (toolResults) {
            body.tool_results = toolResults;
        }
        
        if (clarificationAnswer) {
            body.clarification_answer = clarificationAnswer;
        }
        
        if (updatedContext) {
            body.updated_context = updatedContext;
            console.log(`📸 Including updated context with ${updatedContext.initialContext?.sheets?.length || 0} sheets`);
        }

        const response = await fetch(`${API_BASE_URL}/agents/${agentRunId}/respond`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'ngrok-skip-browser-warning': 'true'
            },
            body: JSON.stringify(body)
        });

        if (!response.ok) {
            throw new Error(`Failed to respond: HTTP ${response.status}`);
        }

        console.log('✓ Response submitted, agent resuming...');
        return true;

    } catch (error) {
        console.error('Failed to respond:', error);
        return false;
    }
}

/**
 * Upload a file (with optional ToC JSON) to a conversation
 * @param {string} conversationId - Conversation ID
 * @param {File} file - Main file (PDF, Excel, etc)
 * @param {File} tocFile - Optional ToC JSON file
 */
export async function uploadFile(conversationId, file, tocFile = null) {
    try {
        console.log(`📤 Uploading file: ${file.name}`);
        
        const formData = new FormData();
        formData.append('conversation_id', conversationId);
        formData.append('file', file);
        
        if (tocFile) {
            console.log(`📄 Including ToC file: ${tocFile.name}`);
            formData.append('toc_file', tocFile);
        }
        
        const response = await fetch(`${API_BASE_URL}/files/upload`, {
            method: 'POST',
            headers: {
                'ngrok-skip-browser-warning': 'true'
            },
            body: formData
        });
        
        if (!response.ok) {
            const errorData = await response.json().catch(() => ({}));
            throw new Error(errorData.detail || `Upload failed: HTTP ${response.status}`);
        }
        
        const result = await response.json();
        console.log(`✓ File uploaded: ${result.id}`);
        console.log(`  File name: ${result.file_name}`);
        console.log(`  ToC provided: ${result.toc_provided}`);
        
        return { success: true, file: result };
        
    } catch (error) {
        console.error('❌ File upload failed:', error);
        return { success: false, error: error.message };
    }
}

/**
 * Get file info for a conversation
 * @param {string} conversationId - Conversation ID
 */
export async function getConversationFile(conversationId) {
    try {
        const response = await fetch(`${API_BASE_URL}/files/conversation/${conversationId}`, {
            headers: {
                'ngrok-skip-browser-warning': 'true'
            }
        });
        
        if (response.status === 404) {
            return { success: true, file: null }; // No file uploaded yet
        }
        
        if (!response.ok) {
            throw new Error(`Failed to get file: HTTP ${response.status}`);
        }
        
        const file = await response.json();
        return { success: true, file };
        
    } catch (error) {
        console.error('❌ Failed to get file:', error);
        return { success: false, error: error.message };
    }
}

/**
 * Delete a file by ID
 * @param {string} fileId - File ID to delete
 */
export async function deleteFile(fileId) {
    try {
        console.log(`🗑️ Deleting file: ${fileId}`);
        
        const response = await fetch(`${API_BASE_URL}/files/${fileId}`, {
            method: 'DELETE',
            headers: {
                'ngrok-skip-browser-warning': 'true'
            }
        });
        
        if (!response.ok) {
            const errorData = await response.json().catch(() => ({}));
            throw new Error(errorData.detail || `Delete failed: HTTP ${response.status}`);
        }
        
        const result = await response.json();
        console.log(`✓ File deleted: ${result.file_id}`);
        
        return { success: true, result };
        
    } catch (error) {
        console.error('❌ File deletion failed:', error);
        return { success: false, error: error.message };
    }
}

/**
 * Helper: Sleep function
 */
function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}
