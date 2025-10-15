/**
 * API Client Module - Simplified for new backend architecture
 * Handles all backend communication with conversation-based chat
 */

export const API_BASE_URL = 'https://finoco-dev-lb-1271513380.us-east-1.elb.amazonaws.com/engine';

/**
 * Check API health status
 */
export async function checkAPIHealth() {
    try {
        const response = await fetch(`${API_BASE_URL}/health`, {
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
 * Initiate a new conversation/chat
 */
export async function initiateChat(message, queryPayload, onRetry) {
    const maxRetries = 3;
    let retryCount = 0;
    const idempotencyKey = crypto.randomUUID(); // Generate a unique key for this operation
    
    // DEBUG: Add timestamp and stack trace to track duplicate calls
    const callTimestamp = new Date().toISOString();
    const callStack = new Error().stack;
    console.log(`🔍 DEBUG: initiateChat called at ${callTimestamp}`);
    console.log(`🔍 DEBUG: Call stack:`, callStack);

    while (retryCount < maxRetries) {
        try {
            console.log(`📤 Initiating chat (attempt ${retryCount + 1}/${maxRetries})...`);
            console.log(`   Idempotency Key: ${idempotencyKey}`);

            const response = await fetch(`${API_BASE_URL}/chat/initiate`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'ngrok-skip-browser-warning': 'true',
                    'Idempotency-Key': idempotencyKey // Send the key in the headers
                },
                body: JSON.stringify({
                    query: message,
                    enhancedPayload: queryPayload
                })
            });

            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }

            const result = await response.json();
            const conversationId = result.conversation_id;

            console.log(`✓ Chat initiated: ${conversationId}`);
            console.log(`  Status: ${result.status}`);
            
            return { success: true, conversationId, result };

        } catch (error) {
            retryCount++;
            console.error(`❌ Initiate attempt ${retryCount} failed:`, error.message);

            if (retryCount >= maxRetries) {
                return { 
                    success: false, 
                    error: `Failed to initiate chat after ${maxRetries} attempts: ${error.message}` 
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
 * Poll conversation status
 */
export async function pollConversation(conversationId, onProgress, onClarificationNeeded, onToolExecutionNeeded) {
    const pollInterval = 2000; // 2 seconds
    const maxPolls = 300; // 10 minutes
    let pollCount = 0;
    let consecutiveErrors = 0;
    const maxConsecutiveErrors = 5;

    while (pollCount < maxPolls) {
        try {
            const response = await fetch(`${API_BASE_URL}/chat/${conversationId}/poll`, {
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
            
            // Update progress
            if (onProgress && data.message) {
                onProgress(data.message.explanation || 'Processing...', elapsed, data.status);
            }

            console.log(`📊 Poll ${pollCount + 1}: ${data.status} - ${data.message?.explanation || 'Processing...'}`);

            // Handle different statuses
            if (data.status === 'complete') {
                console.log(`✓ Chat completed`);
                return { success: true, result: data.message };
            }

            if (data.status === 'needs_clarification') {
                console.log(`⏸️ Chat paused for clarification`);
                
                if (!onClarificationNeeded) {
                    return { 
                        success: false, 
                        error: 'Clarification needed but no handler provided' 
                    };
                }

                // Extract question from askUser tool
                const question = data.message?.toolCalls?.[0]?.params?.question || "I need more information to proceed.";
                
                const answer = await onClarificationNeeded(question, data.message);
                
                if (!answer) {
                    return { success: false, error: 'Chat cancelled by user' };
                }

                // Submit answer and continue polling
                const submitSuccess = await respondToConversation(conversationId, answer);
                if (!submitSuccess) {
                    return { success: false, error: 'Failed to submit clarification' };
                }

                // Continue polling
                pollCount++;
                await sleep(500); // Short delay before resuming
                continue;
            }

            if (data.status === 'needs_tool_execution') {
                console.log(`🔧 Frontend tools needed`);
                
                if (!onToolExecutionNeeded) {
                    return { 
                        success: false, 
                        error: 'Tool execution needed but no handler provided' 
                    };
                }

                // Execute tools and get ALL results (READ and WRITE, successes and failures)
                const toolResults = await onToolExecutionNeeded(data.message.toolCalls || []);
                
                // Always send results back to backend (even if empty or write-only)
                // This allows the agent to see failures and retry/fix them
                const resultsSummary = {
                    executed: toolResults.length,
                    successful: toolResults.filter(r => r.success).length,
                    failed: toolResults.filter(r => !r.success).length,
                    results: toolResults
                };
                
                console.log(`📊 Tool execution summary: ${resultsSummary.successful} succeeded, ${resultsSummary.failed} failed`);
                
                const submitSuccess = await respondToConversation(
                    conversationId, 
                    JSON.stringify(resultsSummary, null, 2)
                );
                
                if (!submitSuccess) {
                    return { success: false, error: 'Failed to submit tool results' };
                }
                
                // Continue polling for next response
                pollCount++;
                await sleep(500);
                continue;
            }

            if (data.status === 'failed') {
                return { 
                    success: false, 
                    error: data.error || 'Chat processing failed' 
                };
            }

            // Still processing (status: "processing")
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
        error: 'Chat timeout - exceeded maximum wait time (10 minutes)' 
    };
}

/**
 * Respond to conversation (clarification answer OR tool results)
 */
export async function respondToConversation(conversationId, content) {
    try {
        const response = await fetch(`${API_BASE_URL}/chat/${conversationId}/respond`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'ngrok-skip-browser-warning': 'true'
            },
            body: JSON.stringify({ content })
        });

        if (!response.ok) {
            throw new Error(`Failed to respond: HTTP ${response.status}`);
        }

        console.log('✓ Response submitted, resuming chat...');
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
