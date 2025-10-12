/**
 * API Client Module
 * Handles all backend communication including task submission and polling
 */

const API_BASE_URL = 'https://41a36d0f8a03.ngrok-free.app';

/**
 * Generate unique request ID for deduplication
 */
export function generateRequestId() {
    return `req-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
}

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
 * Submit task to backend with retry logic
 */
export async function submitTask(message, queryPayload, requestId, onRetry) {
    const maxRetries = 3;
    let retryCount = 0;

    while (retryCount < maxRetries) {
        try {
            console.log(`📤 Submitting task (attempt ${retryCount + 1}/${maxRetries})...`);

            const response = await fetch(`${API_BASE_URL}/process`, {
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

            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }

            const result = await response.json();
            const taskId = result.task_id;

            console.log(`✓ Task submitted: ${taskId}`);
            console.log(`  Status: ${result.status}`);
            
            return { success: true, taskId, result };

        } catch (error) {
            retryCount++;
            console.error(`❌ Submit attempt ${retryCount} failed:`, error.message);

            if (retryCount >= maxRetries) {
                return { 
                    success: false, 
                    error: `Failed to submit task after ${maxRetries} attempts: ${error.message}` 
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
 * Recover task using request ID
 */
export async function recoverTask(requestId) {
    try {
        console.warn('⚠️ Attempting task recovery...');
        
        const response = await fetch(`${API_BASE_URL}/tasks/by-request/${requestId}`, {
            headers: {
                'ngrok-skip-browser-warning': 'true'
            }
        });

        if (response.ok) {
            const data = await response.json();
            console.log(`✓ Task recovered: ${data.task_id}`);
            return { success: true, taskId: data.task_id, data };
        } else {
            return { success: false, error: 'Task recovery failed' };
        }
    } catch (error) {
        console.error('Task recovery failed:', error);
        return { success: false, error: error.message };
    }
}

/**
 * Poll task status until complete or failed
 */
export async function pollTaskStatus(taskId, onProgress, onClarificationNeeded) {
    const pollInterval = 2000; // 2 seconds
    const maxPolls = 300; // 10 minutes
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
            consecutiveErrors = 0; // Reset on success

            // Update progress
            const elapsed = status.elapsed_seconds || 0;
            if (onProgress) {
                onProgress(status.progress, elapsed, status.status);
            }

            console.log(`📊 Poll ${pollCount + 1}: ${status.status} - ${status.progress} (${elapsed}s)`);

            // Check completion
            if (status.status === 'complete') {
                console.log(`✓ Task completed after ${elapsed}s`);
                return { success: true, result: status.result };
            }

            // Check for clarification
            if (status.status === 'needs_clarification') {
                console.log(`⏸️ Task paused for clarification`);
                
                if (!onClarificationNeeded) {
                    return { 
                        success: false, 
                        error: 'Clarification needed but no handler provided' 
                    };
                }

                const answer = await onClarificationNeeded(status);
                
                if (!answer) {
                    return { success: false, error: 'Task cancelled by user' };
                }

                // Submit answer and continue polling
                const submitSuccess = await submitClarification(taskId, answer);
                if (!submitSuccess) {
                    return { success: false, error: 'Failed to submit clarification' };
                }

                // Continue polling
                pollCount++;
                await sleep(pollInterval);
                continue;
            }

            // Check failure
            if (status.status === 'failed') {
                return { 
                    success: false, 
                    error: status.error || 'Task processing failed' 
                };
            }

            // Wait before next poll
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
        error: 'Task timeout - exceeded maximum wait time (10 minutes)' 
    };
}

/**
 * Submit clarification answer
 */
export async function submitClarification(taskId, answer) {
    try {
        const response = await fetch(`${API_BASE_URL}/tasks/${taskId}/respond`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'ngrok-skip-browser-warning': 'true'
            },
            body: JSON.stringify({ answer })
        });

        if (!response.ok) {
            throw new Error(`Failed to submit clarification: HTTP ${response.status}`);
        }

        console.log('✓ Clarification submitted, resuming task...');
        return true;

    } catch (error) {
        console.error('Failed to submit clarification:', error);
        return false;
    }
}

/**
 * Helper: Sleep function
 */
function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

