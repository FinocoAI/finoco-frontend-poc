/**
 * formatRange.js
 *
 * Tool: Apply formatting to a range
 * Executor: Frontend (requires Office.js)
 */

export const toolDefinition = {
  name: "formatRange",
  description: "Apply formatting (font, fill, number format) to a range. IMPORTANT: Use range notation (e.g., 'A1:K1') to format multiple cells with the same formatting in a single call instead of making separate calls for each cell.",
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
      description: "Range address - use range notation like 'A1:D10' to batch multiple cells with same formatting"
    },
    format: {
      type: "object",
      required: true,
      description: "Format object with properties: numberFormat, fontBold, fontItalic, fontSize, fontColor, fillColor, horizontalAlignment, verticalAlignment, borders"
    }
  },
  returns: {
    sheetName: "string",
    address: "string",
    formatted: "boolean"
  }
};

export async function execute(params) {
  console.log(`🔧 Executing formatRange:`, params);

  return Excel.run(async (context) => {
    try {
      const { sheetName, address, format } = params;

      // Validate parameters
      if (!sheetName || typeof sheetName !== 'string') {
        throw new Error('Invalid sheetName');
      }
      if (!address || typeof address !== 'string') {
        throw new Error('Invalid address');
      }
      if (!format || typeof format !== 'object') {
        throw new Error('Invalid format: must be an object');
      }

      // Get worksheet
      const worksheet = context.workbook.worksheets.getItem(sheetName);

      // Get the range
      const range = worksheet.getRange(address);
      range.load("address");

      await context.sync();

      // Apply number format
      if (format.numberFormat) {
        range.numberFormat = format.numberFormat;
      }

      // Apply font formatting
      if (format.fontBold !== undefined) {
        range.format.font.bold = format.fontBold;
      }
      if (format.fontItalic !== undefined) {
        range.format.font.italic = format.fontItalic;
      }
      if (format.fontSize) {
        range.format.font.size = format.fontSize;
      }
      if (format.fontColor) {
        range.format.font.color = format.fontColor;
      }

      // Apply fill color
      if (format.fillColor) {
        range.format.fill.color = format.fillColor;
      }

      // Apply alignment
      if (format.horizontalAlignment) {
        // Excel API requires lowercase alignment values
        range.format.horizontalAlignment = format.horizontalAlignment.toLowerCase();
      }
      if (format.verticalAlignment) {
        // Excel API requires lowercase alignment values
        range.format.verticalAlignment = format.verticalAlignment.toLowerCase();
      }

      // Apply borders
      if (format.borders) {
        const borderTypes = ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'];
        borderTypes.forEach(type => {
          const borderValue = format.borders[type.toLowerCase()];
          if (borderValue) {
            // If boolean true, use "Continuous" style; otherwise use the provided style string
            const borderStyle = typeof borderValue === 'boolean' ? 'Continuous' : borderValue;
            range.format.borders.getItem(type).style = borderStyle;
          }
        });
      }

      await context.sync();

      const result = {
        sheetName: sheetName,
        address: range.address,
        formatted: true,
        appliedFormats: Object.keys(format)
      };

      console.log(`  ✅ Applied formatting to ${range.address}`);

      return result;

    } catch (error) {
      console.error(`  ❌ Error in formatRange:`, error);
      throw {
        tool: "formatRange",
        error: error.message,
        params: params
      };
    }
  });
}
