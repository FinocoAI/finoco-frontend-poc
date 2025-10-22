/**
 * getCellPrecedents.js
 *
 * Tool: Find all cells that this cell references (precedent tracing)
 * Executor: Frontend (requires Office.js)
 */

import { parseFormulaReferences, getRepresentativeAddress } from '../../utils/formulaParser.js';

export const toolDefinition = {
  name: "getCellPrecedents",
  description: "Find all cells that this cell depends on/references (trace precedents)",
  executor: "frontend",
  parameters: {
    sheetName: {
      type: "string",
      required: true,
      description: "Name of the sheet"
    },
    address: {
      type: "string",
      required: true,
      description: "Cell address (e.g., 'C5')"
    }
  },
  returns: {
    sheetName: "string",
    sourceCell: "string",
    precedents: "array"  // Array of precedent cell addresses
  }
};

export async function execute(params) {
  console.log(`🔧 Executing getCellPrecedents:`, params);

  return Excel.run(async (context) => {
    try {
      const { sheetName, address } = params;

      if (!sheetName || typeof sheetName !== 'string') {
        throw new Error('Invalid sheetName');
      }
      if (!address || typeof address !== 'string') {
        throw new Error('Invalid address');
      }

      // Get worksheet and range
      const worksheet = context.workbook.worksheets.getItem(sheetName);
      const range = worksheet.getRange(address);

      range.load("address, formulas");
      await context.sync();

      const formula = range.formulas[0][0];
      let precedentAddresses = [];
      let method = 'none';

      // APPROACH 1: Try Excel API first
      try {
        const precedents = range.getDirectPrecedents();
        precedents.load("address, areas");
        await context.sync();

        // Extract precedent addresses
        const areas = precedents.areas;
        areas.load("items");
        await context.sync();

        for (let i = 0; i < areas.items.length; i++) {
          const area = areas.items[i];
          area.load("address");
          await context.sync();
          precedentAddresses.push(area.address);
        }

        method = 'api';
        console.log(`  📍 API found ${precedentAddresses.length} precedent(s) for ${address}`);

      } catch (apiError) {
        // API failed - expected for complex formulas
        console.log(`  ⚠️ API failed for ${address}, trying formula parsing`);
      }

      // APPROACH 2: If API returned nothing or failed, parse formula
      if (precedentAddresses.length === 0 && formula && formula.startsWith('=')) {
        console.log(`  📝 Using formula parsing for ${address}`);
        
        const parsedRefs = parseFormulaReferences(formula, sheetName);
        
        // Convert parsed references to full addresses
        for (const ref of parsedRefs) {
          let addr = ref.address;
          
          // For full column/row references, use representative cell
          if (ref.isFullColumn || ref.isFullRow) {
            addr = getRepresentativeAddress(ref.address);
            console.log(`  🔄 Converted ${ref.address} to ${addr}`);
          } else if (ref.isRange && !ref.isFullColumn && !ref.isFullRow) {
            // For regular ranges, use first cell
            addr = ref.address.split(':')[0];
          }
          
          // Create full address with sheet
          const fullAddr = ref.sheet === sheetName ? addr : `${ref.sheet}!${addr}`;
          precedentAddresses.push(fullAddr);
        }
        
        method = 'parsed';
        console.log(`  📝 Parsed ${precedentAddresses.length} precedent(s) from formula`);
      }

      const result = {
        sheetName: sheetName,
        sourceCell: range.address,
        precedents: precedentAddresses,
        precedentCount: precedentAddresses.length,
        method: method, // 'api', 'parsed', or 'none'
        formula: formula || null
      };

      console.log(`  ✅ Found ${precedentAddresses.length} precedent(s) for ${address} (method: ${method})`);

      return result;

    } catch (error) {
      console.error(`  ❌ Error in getCellPrecedents:`, error);
      throw {
        tool: "getCellPrecedents",
        error: error.message,
        params: params
      };
    }
  });
}
