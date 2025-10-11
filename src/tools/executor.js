/**
 * executor.js
 *
 * Tool Execution Engine
 * Handles validation, execution, and error handling for all tools
 *
 * Features:
 * - Parameter validation
 * - Error handling and reporting
 * - Execution logging
 * - Result formatting
 */

import { getTool, hasTool } from './registry.js';

/**
 * Execute a single tool
 *
 * @param {string} toolName - Name of the tool to execute
 * @param {Object} params - Parameters for the tool
 * @returns {Promise<Object>} Tool execution result
 */
export async function executeTool(toolName, params) {
  console.group(`🔧 Executing Tool: ${toolName}`);
  console.log('Parameters:', params);

  const startTime = performance.now();

  try {
    // Check if tool exists
    if (!hasTool(toolName)) {
      throw new Error(`Tool "${toolName}" not found in registry`);
    }

    // Get the tool
    const tool = getTool(toolName);

    // Validate parameters
    validateParameters(tool.toolDefinition, params);

    // Execute the tool
    console.log(`⚡ Executing...`);
    const result = await tool.execute(params);

    const duration = (performance.now() - startTime).toFixed(2);
    console.log(`✅ Success in ${duration}ms`);
    console.groupEnd();

    return {
      success: true,
      tool: toolName,
      result: result,
      duration: duration,
    };

  } catch (error) {
    const duration = (performance.now() - startTime).toFixed(2);
    console.error(`❌ Failed in ${duration}ms:`, error);
    console.groupEnd();

    return {
      success: false,
      tool: toolName,
      error: error.message || String(error),
      duration: duration,
    };
  }
}

/**
 * Execute multiple tools in sequence
 *
 * @param {Array<Object>} toolCalls - Array of {tool, params} objects
 * @returns {Promise<Array<Object>>} Array of execution results
 */
export async function executeToolSequence(toolCalls) {
  console.group(`🔧 Executing Tool Sequence (${toolCalls.length} tools)`);

  const results = [];

  for (let i = 0; i < toolCalls.length; i++) {
    const { tool, params } = toolCalls[i];
    console.log(`\n[${i + 1}/${toolCalls.length}] ${tool}`);

    const result = await executeTool(tool, params);
    results.push(result);

    // Stop if a tool fails (optional - can be configurable)
    if (!result.success) {
      console.warn(`⚠️ Tool ${tool} failed, stopping sequence`);
      break;
    }
  }

  console.groupEnd();
  return results;
}

/**
 * Execute multiple tools in parallel
 *
 * @param {Array<Object>} toolCalls - Array of {tool, params} objects
 * @returns {Promise<Array<Object>>} Array of execution results
 */
export async function executeToolsParallel(toolCalls) {
  console.group(`🔧 Executing Tools in Parallel (${toolCalls.length} tools)`);

  const promises = toolCalls.map(({ tool, params }) => executeTool(tool, params));
  const results = await Promise.all(promises);

  console.groupEnd();
  return results;
}

/**
 * Validate tool parameters against tool definition
 *
 * @param {Object} toolDefinition - Tool definition with parameter schema
 * @param {Object} params - Actual parameters provided
 * @throws {Error} If validation fails
 */
function validateParameters(toolDefinition, params) {
  const { parameters } = toolDefinition;

  // Check required parameters
  for (const [paramName, paramSchema] of Object.entries(parameters)) {
    const isRequired = paramSchema.required !== false;
    const value = params[paramName];

    // Check if required parameter is missing
    if (isRequired && (value === undefined || value === null)) {
      throw new Error(
        `Missing required parameter: ${paramName} (${paramSchema.description || 'no description'})`
      );
    }

    // Skip validation if parameter not provided and not required
    if (value === undefined || value === null) {
      continue;
    }

    // Validate parameter type
    const expectedType = paramSchema.type;
    const actualType = Array.isArray(value) ? 'array' : typeof value;

    if (expectedType && actualType !== expectedType) {
      throw new Error(
        `Invalid type for parameter "${paramName}": expected ${expectedType}, got ${actualType}`
      );
    }

    // Validate array contents if specified
    if (expectedType === 'array' && paramSchema.items) {
      for (const item of value) {
        const itemType = typeof item;
        if (itemType !== paramSchema.items.type) {
          throw new Error(
            `Invalid array item type in "${paramName}": expected ${paramSchema.items.type}, got ${itemType}`
          );
        }
      }
    }
  }
}

/**
 * Format tool results for display
 *
 * @param {Object} result - Tool execution result
 * @returns {string} Formatted result string
 */
export function formatToolResult(result) {
  if (result.success) {
    return `✅ ${result.tool} completed in ${result.duration}ms`;
  } else {
    return `❌ ${result.tool} failed: ${result.error}`;
  }
}

/**
 * Get summary of tool execution results
 *
 * @param {Array<Object>} results - Array of tool execution results
 * @returns {Object} Summary statistics
 */
export function getExecutionSummary(results) {
  const successful = results.filter(r => r.success).length;
  const failed = results.filter(r => !r.success).length;
  const totalDuration = results.reduce((sum, r) => sum + parseFloat(r.duration), 0).toFixed(2);

  return {
    total: results.length,
    successful,
    failed,
    totalDuration: `${totalDuration}ms`,
  };
}
