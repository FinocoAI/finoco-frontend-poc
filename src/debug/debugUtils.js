/**
 * Debug Utilities Module
 * Provides debugging and testing utilities
 */

import { buildQueryPayload, debugPrintPayload, validateQueryPayload } from '../queryBuilder.js';
import { updateEnhancedContext, enhancedContext } from '../context/contextManager.js';
import { executeTool, executeToolSequence, formatToolResult, getExecutionSummary } from '../tools/executor.js';
import { getAllToolDefinitions, printToolRegistry } from '../tools/registry.js';

/**
 * Debug function to capture and log all context
 */
export async function debugContextCapture() {
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

/**
 * Debug function to test tool execution
 */
export async function debugToolExecution() {
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

/**
 * Setup debug keyboard shortcuts
 */
export function setupDebugShortcuts() {
    document.addEventListener('keydown', (e) => {
        // Ctrl/Cmd + Shift + D: Context Capture
        if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key === 'D') {
            e.preventDefault();
            debugContextCapture();
        }

        // Ctrl/Cmd + Shift + K: Tool Testing
        if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key === 'K') {
            e.preventDefault();
            debugToolExecution();
        }
    });

    console.log('💡 Debug mode:');
    console.log('  - Press Ctrl/Cmd + Shift + D to capture context');
    console.log('  - Press Ctrl/Cmd + Shift + K to test tools');
    
    // Expose debug utilities globally for manual console testing
    window.debugToolExecution = debugToolExecution;
    window.debugContextCapture = debugContextCapture;
    console.log('🧩 Global debug functions available: debugToolExecution(), debugContextCapture()');
    console.log('');
}

