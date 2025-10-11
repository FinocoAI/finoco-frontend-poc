/**
 * registry.js
 *
 * Central Tool Registry
 * Imports and exports all available tools for the Excel agent
 *
 * Organization:
 * - read/ - Tools that fetch data from Excel (no modifications)
 * - write/ - Tools that modify Excel data
 */

// Import read tools
import * as getFullRangeData from './read/getFullRangeData.js';
import * as getColumnData from './read/getColumnData.js';
import * as getTableData from './read/getTableData.js';
import * as getNamedRangeData from './read/getNamedRangeData.js';
import * as searchValues from './read/searchValues.js';
import * as getFormulasInRange from './read/getFormulasInRange.js';
import * as getCellDependents from './read/getCellDependents.js';
import * as getCellPrecedents from './read/getCellPrecedents.js';
import * as getRelatedData from './read/getRelatedData.js';
import * as getChartSourceData from './read/getChartSourceData.js';

// Import write tools
import * as writeDataToRange from './write/writeDataToRange.js';
import * as createChart from './write/createChart.js';
import * as insertRows from './write/insertRows.js';
import * as deleteRows from './write/deleteRows.js';
import * as createTable from './write/createTable.js';
import * as applyFormula from './write/applyFormula.js';
import * as formatRange from './write/formatRange.js';
import * as createNewSheet from './write/createNewSheet.js';

/**
 * Tool Registry
 * Organized by category (read/write)
 */
export const toolRegistry = {
  read: {
    // Basic data retrieval
    getFullRangeData,
    getColumnData,

    // Structured data
    getTableData,
    getNamedRangeData,

    // Search and filter
    searchValues,

    // Formula analysis
    getFormulasInRange,
    getCellDependents,
    getCellPrecedents,

    // Relationships and lookups
    getRelatedData,

    // Charts
    getChartSourceData,
  },
  write: {
    // Data modification
    writeDataToRange,

    // Row operations
    insertRows,
    deleteRows,

    // Structure
    createTable,
    createNewSheet,

    // Formulas and formatting
    applyFormula,
    formatRange,

    // Charts
    createChart,
  }
};

/**
 * Get all tool definitions (for LLM/Backend)
 * Returns array of tool definitions that can be sent to backend
 *
 * @returns {Array<Object>} Array of tool definitions
 */
export function getAllToolDefinitions() {
  const definitions = [];

  // Collect read tools
  for (const [name, tool] of Object.entries(toolRegistry.read)) {
    definitions.push(tool.toolDefinition);
  }

  // Collect write tools
  for (const [name, tool] of Object.entries(toolRegistry.write)) {
    definitions.push(tool.toolDefinition);
  }

  return definitions;
}

/**
 * Get a specific tool by name
 *
 * @param {string} toolName - Name of the tool (e.g., "getFullRangeData")
 * @returns {Object|null} Tool object with definition and execute function
 */
export function getTool(toolName) {
  // Search in read tools
  if (toolRegistry.read[toolName]) {
    return toolRegistry.read[toolName];
  }

  // Search in write tools
  if (toolRegistry.write[toolName]) {
    return toolRegistry.write[toolName];
  }

  return null;
}

/**
 * Check if a tool exists
 *
 * @param {string} toolName - Name of the tool
 * @returns {boolean} True if tool exists
 */
export function hasTool(toolName) {
  return getTool(toolName) !== null;
}

/**
 * Get tools by category
 *
 * @param {string} category - Category name ("read" or "write")
 * @returns {Object} Object with tool name as key and tool as value
 */
export function getToolsByCategory(category) {
  if (category === 'read') {
    return toolRegistry.read;
  }
  if (category === 'write') {
    return toolRegistry.write;
  }
  return {};
}

/**
 * Get tool categories
 *
 * @returns {Array<string>} Array of category names
 */
export function getCategories() {
  return ['read', 'write'];
}

/**
 * Print tool registry summary (for debugging)
 */
export function printToolRegistry() {
  console.group('🔧 Tool Registry');

  console.log('📖 Read Tools:');
  for (const [name, tool] of Object.entries(toolRegistry.read)) {
    console.log(`  - ${name}: ${tool.toolDefinition.description}`);
  }

  console.log('✏️ Write Tools:');
  for (const [name, tool] of Object.entries(toolRegistry.write)) {
    console.log(`  - ${name}: ${tool.toolDefinition.description}`);
  }

  console.log(`\nTotal: ${Object.keys(toolRegistry.read).length + Object.keys(toolRegistry.write).length} tools`);

  console.groupEnd();
}
