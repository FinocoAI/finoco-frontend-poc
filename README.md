# Warren by Finoco.ai - Frontend

AI that understands your spreadsheets. Frontend application and tools to be executed in frontend.

**✨ Updated:** Now using simplified conversation-based architecture with `conversation_id` instead of `task_id`.

## Recent Changes

- **Simplified API**: Now using `/chat/initiate`, `/chat/{conversation_id}/poll`, `/chat/{conversation_id}/respond`
- **Conversation-based**: Single conversation ID tracks entire chat flow
- **Unified Response**: Both clarifications and tool results use same respond endpoint
- **See [FRONTEND_CHANGES.md](./FRONTEND_CHANGES.md) for detailed migration guide**

## Architecture

### Core Components

**[src/taskpane.js](src/taskpane.js)** - Main application entry point
- Initializes Office.js and UI event listeners
- Manages chat interface and user interactions
- Coordinates context capture and tool execution
- Handles backend API communication (updated for conversation-based flow)

**[src/api/apiClient.js](src/api/apiClient.js)** - Backend communication (Updated)
- `initiateChat()` - Start new conversation
- `pollConversation()` - Check conversation status
- `respondToConversation()` - Send clarifications or tool results

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
- `writeDataToRange` - Write data to ranges (with formula support)
- `createChart` - Create charts
- `insertRows` - Insert rows
- `deleteRows` - Delete rows
- `createTable` - Create Excel tables
- `applyFormula` - Apply formulas
- `formatRange` - Format ranges
- `createNewSheet` - Create worksheets
- `addCellNote` - Add citations and notes to cells

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

## API Integration (Updated)

### Initialize Conversation
```javascript
POST /chat/initiate
{
  "query": "User query text",
  "enhancedPayload": {
    "initialContext": { /* workbook context */ },
    "userSelection": { /* selection context */ }
  }
}

→ Returns: { conversation_id, status: "processing" }
```

### Poll Conversation
```javascript
GET /chat/{conversation_id}/poll

→ Returns: {
  "status": "processing|needs_clarification|needs_tool_execution|complete",
  "message": {
    "explanation": "...",
    "tool_call_type": "...",
    "toolCalls": [...],
    "answer": "..."
  }
}
```

### Respond to Conversation
```javascript
POST /chat/{conversation_id}/respond
{
  "content": "user answer" OR "Tool results: {...}"
}

→ Returns: { status: "processing" }
```

## Conversation Flow

1. **User sends message** → `initiateChat()` → get `conversation_id`
2. **Frontend polls** → `pollConversation()` → check status
3. **If needs clarification** → show modal → send answer → resume polling
4. **If needs tools** → execute tools → send results (if READ tools) → resume polling
5. **If complete** → show final message

---

**For detailed migration information, see [FRONTEND_CHANGES.md](./FRONTEND_CHANGES.md)**
