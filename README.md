# Excel Agent Frontend

Frontend and tools to be executed in frontend.

## Architecture

### Core Components

**[src/taskpane.js](src/taskpane.js)** - Main application entry point
- Initializes Office.js and UI event listeners
- Manages chat interface and user interactions
- Coordinates context capture and tool execution
- Handles backend API communication

**[src/contextCapture.js](src/contextCapture.js)** - Workbook state capture
- Captures initial workbook context (sheets, named ranges, tables)
- Provides workbook-level metadata

**[src/selectionCapture.js](src/selectionCapture.js)** - User selection tracking
- Captures current user selection and surrounding data
- Handles multi-area selections

**[src/queryBuilder.js](src/queryBuilder.js)** - Query payload construction
- Builds structured payloads for backend API
- Validates and formats context data

### Tool System

**[src/tools/registry.js](src/tools/registry.js)** - Central tool registry
- Imports and registers all available tools
- Organizes tools by category (read/write)
- Provides tool lookup and discovery functions

**[src/tools/executor.js](src/tools/executor.js)** - Tool execution engine
- Validates tool parameters
- Executes tools with error handling
- Supports sequential and parallel execution
- Provides execution summaries

### Available Tools

**Read Tools** ([src/tools/read/](src/tools/read/))
- `getFullRangeData` - Fetch complete range data
- `getColumnData` - Retrieve specific columns
- `getTableData` - Read Excel table data
- `getNamedRangeData` - Access named ranges
- `searchValues` - Search for values in sheets
- `getFormulasInRange` - Extract formulas
- `getCellDependents` - Find dependent cells
- `getCellPrecedents` - Find precedent cells
- `getRelatedData` - Fetch related data
- `getChartSourceData` - Get chart source data

**Write Tools** ([src/tools/write/](src/tools/write/))
- `writeDataToRange` - Write data to ranges
- `createChart` - Create charts
- `insertRows` - Insert rows
- `deleteRows` - Delete rows
- `createTable` - Create Excel tables
- `applyFormula` - Apply formulas
- `formatRange` - Format ranges
- `createNewSheet` - Create worksheets

## Adding New Tools

### 1. Create Tool File

Create a new file in `src/tools/read/` or `src/tools/write/`:

```javascript
// src/tools/read/myNewTool.js

/**
 * Tool definition for backend/LLM
 */
export const toolDefinition = {
  name: "myNewTool",
  description: "Brief description of what this tool does",
  executor: "frontend",
  parameters: {
    paramName: {
      type: "string",         // string, number, boolean, array
      required: true,
      description: "Parameter description"
    },
    optionalParam: {
      type: "number",
      required: false,
      default: 10,
      description: "Optional parameter"
    }
  },
  returns: {
    field1: "string",
    field2: "number"
  }
};

/**
 * Execute function
 */
export async function execute(params) {
  console.log(`🔧 Executing myNewTool:`, params);

  return Excel.run(async (context) => {
    try {
      const { paramName, optionalParam = 10 } = params;

      // Validate parameters
      if (!paramName) {
        throw new Error('paramName is required');
      }

      // Your Excel.js logic here
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      // ... perform operations ...

      await context.sync();

      // Return result object
      const result = {
        field1: "value",
        field2: 42
      };

      console.log(`  ✅ myNewTool completed`);
      return result;

    } catch (error) {
      console.error(`  ❌ Error in myNewTool:`, error);
      throw {
        tool: "myNewTool",
        error: error.message,
        params: params
      };
    }
  });
}
```

### 2. Register Tool

Add import and registration in [src/tools/registry.js](src/tools/registry.js):

```javascript
// Import your tool
import * as myNewTool from './read/myNewTool.js';

export const toolRegistry = {
  read: {
    // ... existing tools ...
    myNewTool,  // Add here
  },
  write: {
    // ... existing tools ...
  }
};
```

### 3. Test Tool

Use debug shortcuts in the add-in:
- `Ctrl/Cmd + Shift + T` - Test tool execution
- `Ctrl/Cmd + Shift + D` - Debug context capture

Or use console:
```javascript
// In browser console
await debugToolExecution()
```

## Development

**Debug Mode**: Console functions available at runtime
- `debugToolExecution()` - Test all tools
- `debugContextCapture()` - Inspect context data

**Keyboard Shortcuts**:
- `Ctrl/Cmd + Shift + K` - Run tool tests
- `Ctrl/Cmd + Shift + D` - Capture and log context

## API Integration

Backend endpoint: `POST /process`

Payload structure:
```json
{
  "task": "User query text",
  "enhancedPayload": {
    "query": "User query",
    "initialContext": { /* workbook context */ },
    "userSelection": { /* selection context */ }
  }
}
```

Backend can request tool execution via response:
```json
{
  "message": "AI response",
  "toolCalls": [
    {
      "tool": "writeDataToRange",
      "params": { /* tool parameters */ }
    }
  ]
}
```
