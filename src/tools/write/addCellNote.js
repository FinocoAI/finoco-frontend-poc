/**
 * addCellNote.js
 *
 * Tool: Add a comment/note to a cell with citations and references
 * Executor: Frontend (requires Office.js)
 * 
 * Uses Excel's Comment API (threaded comments) which is fully supported in both Excel Online and Desktop.
 * Perfect for adding web search citations or file parser references to specific cells.
 * 
 * Note: Uses modern "Comments" instead of legacy "Notes" for better Excel Online compatibility.
 */

export const toolDefinition = {
  name: "addCellNote",
  description: "Add a comment (citation, reference, or note) to a specific cell. Perfect for adding web search citations or document references. Fully supported in Excel Online and Desktop.",
  executor: "frontend",
  parameters: {
    sheetName: {
      type: "string",
      required: true,
      description: "Name of the sheet containing the cell"
    },
    cellAddress: {
      type: "string",
      required: true,
      description: "Cell address to add the comment to (e.g., 'A1', 'B5')"
    },
    noteText: {
      type: "string",
      required: true,
      description: "The comment/citation text to add to the cell. Can include URLs, references, or any explanatory text."
    },
    author: {
      type: "string",
      required: false,
      description: "Optional author name for the comment (defaults to 'Warren')"
    }
  },
  returns: {
    sheetName: "string",
    cellAddress: "string",
    noteAdded: "boolean",
    noteText: "string"
  }
};

export async function execute(params) {
  console.log(`🔧 Executing addCellNote:`, params);

  return Excel.run(async (context) => {
    try {
      const { sheetName, cellAddress, noteText, author = "Warren" } = params;

      // Validate parameters
      if (!sheetName || typeof sheetName !== 'string') {
        throw new Error('Invalid sheetName');
      }
      if (!cellAddress || typeof cellAddress !== 'string') {
        throw new Error('Invalid cellAddress');
      }
      if (!noteText || typeof noteText !== 'string') {
        throw new Error('Invalid noteText');
      }

      // Get worksheet and cell
      const worksheet = context.workbook.worksheets.getItem(sheetName);
      const cell = worksheet.getRange(cellAddress);
      
      // Load cell address
      cell.load("address");
      await context.sync();

      // Format the comment text
      const timestamp = new Date().toLocaleString();
      const formattedText = `${noteText}\n\n—${author}, ${timestamp}`;

      try {
        // Primary method: Use modern Comments API (fully supported in Excel Online)
        // First, check if there's already a comment and delete it
        const commentCollection = worksheet.comments;
        commentCollection.load("items");
        await context.sync();

        // Find and delete existing comment on this cell if it exists
        for (let i = 0; i < commentCollection.items.length; i++) {
          const comment = commentCollection.items[i];
          comment.load("cellAddress");
          await context.sync();
          
          // Check if this comment is on our target cell
          if (comment.cellAddress === cell.address) {
            comment.delete();
            await context.sync();
            console.log(`  🗑️ Deleted existing comment on ${cell.address}`);
            break;
          }
        }

        // Add new comment
        const newComment = commentCollection.add(cell, formattedText, Excel.ContentType.plain);
        await context.sync();

        console.log(`  ✅ Added comment to cell ${cell.address}`);
        
        return {
          sheetName: sheetName,
          cellAddress: cell.address,
          noteAdded: true,
          noteText: noteText
        };

      } catch (commentError) {
        console.error(`  ❌ Comment API failed:`, commentError);
        
        // Fallback: Try using Note API (legacy, limited Excel Online support)
        try {
          const formattedNote = `[${author}] ${timestamp}\n${noteText}`;
          
          await cell.setCellProperties([[{
            note: {
              content: formattedNote
            }
          }]]);
          
          await context.sync();
          
          console.log(`  ✅ Added note to cell ${cell.address} (using legacy Note API)`);
          
          return {
            sheetName: sheetName,
            cellAddress: cell.address,
            noteAdded: true,
            noteText: noteText
          };
          
        } catch (noteError) {
          console.error(`  ❌ Note API also failed:`, noteError);
          throw {
            tool: "addCellNote",
            error: noteError.message || commentError.message,
            params: params
          };
        }
      }

    } catch (error) {
      console.error(`  ❌ Error in addCellNote:`, error);
      throw {
        tool: "addCellNote",
        error: error.message,
        params: params
      };
    }
  });
}

