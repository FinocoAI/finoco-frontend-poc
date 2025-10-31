/**
 * API Client Module - BackendV2 Integration
 * Handles all backend communication with the new multi-agent architecture
 * Flow: Create Conversation → Send Message → Poll Agent Run → Respond with Tool Results
 */

export const API_BASE_URL = 'https://f27d29966c3f.ngrok-free.app/api/v1';

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
 * Get or create a conversation for the user
 * For now, we'll use a single conversation per user (stored in localStorage)
 */
export async function getOrCreateConversation() {
    try {
        // Check if we have a stored conversation ID
        let conversationId = localStorage.getItem('warren_conversation_id');
        
        if (conversationId) {
            // Verify it still exists
            const response = await fetch(`${API_BASE_URL}/conversations/${conversationId}`, {
                headers: {
                    'ngrok-skip-browser-warning': 'true'
                }
            });
            
            if (response.ok) {
                console.log(`✓ Using existing conversation: ${conversationId}`);
                return { success: true, conversationId };
            }
        }
        
        // Create new conversation
        console.log('📤 Creating new conversation...');
        const response = await fetch(`${API_BASE_URL}/conversations`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'ngrok-skip-browser-warning': 'true'
            },
            body: JSON.stringify({
                user_id: "1", // String format as expected by schema
                title: 'Excel Analysis Session'
            })
        });

        if (!response.ok) {
            throw new Error(`Failed to create conversation: HTTP ${response.status}`);
        }

        const conversation = await response.json();
        conversationId = conversation.id;
        
        // Store for future use
        localStorage.setItem('warren_conversation_id', conversationId);
        
        console.log(`✓ Created new conversation: ${conversationId}`);
        return { success: true, conversationId };

    } catch (error) {
        console.error('❌ Failed to get/create conversation:', error);
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

            console.log(`✓ Message sent, agent_run created: ${agentRunId}`);
            console.log(`  Status: ${result.agent_run.status}`);
            
            return { 
                success: true, 
                agentRunId,
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

            console.log(`📊 Poll ${pollCount + 1}: ${data.status}`);

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

                // Execute tools and get ALL results (READ and WRITE, successes and failures)
                const toolCalls = data.tool_calls || [];
                const toolResults = await onToolExecutionNeeded(toolCalls);
                
                // Always send results back to backend
                console.log(`📊 Tool execution complete: ${toolResults.length} results`);
                
                const submitSuccess = await respondToAgentRun(agentRunId, toolResults, null);
                
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
 */
export async function respondToAgentRun(agentRunId, toolResults = null, clarificationAnswer = null) {
    try {
        const body = {};
        
        if (toolResults) {
            body.tool_results = toolResults;
        }
        
        if (clarificationAnswer) {
            body.clarification_answer = clarificationAnswer;
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
 * Helper: Sleep function
 */
function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}
